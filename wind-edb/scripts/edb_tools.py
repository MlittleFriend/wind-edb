"""Read-only OOXML audit and version-10 Wind EDB comment codec (stdlib)."""
import argparse
import base64
import hashlib
import json
from pathlib import Path
import posixpath
import re
import struct
import xml.etree.ElementTree as ET
import zipfile
import zlib

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
MAX_PAYLOAD = 8 * 1024 * 1024


def validate_config(config):
    if not isinstance(config, dict):
        raise ValueError('EDB configuration must be an object')
    for key in ('metric', 'exportConfig', 'exportTitleOption', 'timeFrameSetting'):
        if key not in config:
            raise ValueError('Missing configuration field: ' + key)
    if not isinstance(config['metric'], list) or not config['metric']:
        raise ValueError('metric must be a nonempty list')
    for metric in config['metric']:
        if not isinstance(metric, dict) or not metric.get('code'):
            raise ValueError('Each metric must have a verified code')
    for key in ('exportConfig', 'exportTitleOption', 'timeFrameSetting'):
        if not isinstance(config[key], dict):
            raise ValueError(key + ' must be an object')
    return config


def pack_config(config):
    validate_config(config)
    raw = b'EDB\0' + struct.pack('<I', 10) + json.dumps(
        config, ensure_ascii=False, separators=(',', ':')).encode('utf-8')
    if len(raw) > MAX_PAYLOAD:
        raise ValueError('Configuration too large')
    compressor = zlib.compressobj(6, zlib.DEFLATED, -15)
    data = compressor.compress(raw) + compressor.flush()
    header = struct.pack('<IHHHHHIIIHH', 0x04034b50, 20, 0x808, 8, 0, 0, 0, 0, 0, 1, 0)
    trailer = struct.pack('<IIII', 0x08074b50, zlib.crc32(raw), len(data), len(raw))
    return base64.b64encode(header + b'0' + data + trailer).decode('ascii')


def unpack_config(text):
    if len(text) > MAX_PAYLOAD * 2:
        raise ValueError('Encoded comment too large')
    data = base64.b64decode(''.join(text.split()), validate=True)
    if len(data) < 47:
        raise ValueError('Not a complete supported EDB comment')
    h = struct.unpack('<IHHHHHIIIHH', data[:30])
    if h[0] != 0x04034b50 or h[2] != 0x808 or h[3] != 8:
        raise ValueError('Unsupported ZIP header/flags/method')
    start = 30 + h[-2] + h[-1]
    if data[30:30+h[-2]] != b'0':
        raise ValueError('Unsupported EDB member name')
    inflater = zlib.decompressobj(-15)
    raw = inflater.decompress(data[start:], MAX_PAYLOAD + 1)
    if len(raw) > MAX_PAYLOAD or not inflater.eof:
        raise ValueError('Oversized or incomplete deflate stream')
    tail = inflater.unused_data
    if len(tail) != 16:
        raise ValueError('Unsupported descriptor/trailing data')
    signature, crc, compressed, expanded = struct.unpack('<IIII', tail)
    if (signature != 0x08074b50 or crc != zlib.crc32(raw)
            or expanded != len(raw) or compressed != len(data) - start - 16):
        raise ValueError('EDB checksum/length mismatch')
    if raw[:8] != b'EDB\0' + struct.pack('<I', 10):
        raise ValueError('Unsupported EDB signature/version')
    return validate_config(json.loads(raw[8:].decode('utf-8')))


def target_part(source, target):
    return target.lstrip('/') if target.startswith('/') else posixpath.normpath(
        posixpath.join(posixpath.dirname(source), target))


def relationships(z, source):
    part = posixpath.join(posixpath.dirname(source), '_rels', posixpath.basename(source)+'.rels')
    if part not in z.namelist():
        return []
    return [r.attrib for r in ET.fromstring(z.read(part))]


def inspect_book(path, configs_dir=None):
    path = Path(path)
    result = {'path': str(path.resolve()), 'sha256': hashlib.sha256(path.read_bytes()).hexdigest(),
              'anchors': [], 'errors': [], 'external_links': [], 'warnings': [],
              'status': 'offline_inspection_only_not_refresh_verification'}
    if configs_dir:
        Path(configs_dir).mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path) as z:
        workbook = ET.fromstring(z.read('xl/workbook.xml'))
        calc = workbook.find('s:calcPr', NS)
        result['calcPr'] = calc.attrib if calc is not None else {}
        rels = {r['Id']: r for r in relationships(z, 'xl/workbook.xml')}
        result['assets'] = {n: hashlib.sha256(z.read(n)).hexdigest() for n in z.namelist()
                            if n.startswith('xl/media/') or re.fullmatch(r'xl/charts/chart\d+\.xml', n)}
        result['sheets'] = []
        for sheet in workbook.find('s:sheets', NS):
            part = target_part('xl/workbook.xml', rels[sheet.get(RID)]['Target'])
            root = ET.fromstring(z.read(part))
            comments = {}
            for rel in relationships(z, part):
                if rel['Type'].endswith('/comments'):
                    cr = ET.fromstring(z.read(target_part(part, rel['Target'])))
                    for c in cr.findall('.//s:comment', NS):
                        comments[c.get('ref')] = ''.join(t.text or '' for t in c.findall('.//s:t', NS))
            result['sheets'].append({'name': sheet.get('name'), 'part': part})
            for c in root.findall('.//s:sheetData/s:row/s:c', NS):
                value = c.findtext('s:v', default='', namespaces=NS)
                if c.get('t') == 'e':
                    result['errors'].append({'sheet': sheet.get('name'), 'cell': c.get('r'), 'value': value})
                formula = c.findtext('s:f', default='', namespaces=NS)
                if not re.search(r'\b(?:edb|wset)\s*\(', formula, re.I):
                    continue
                a = {'sheet': sheet.get('name'), 'cell': c.get('r'), 'formula': formula,
                     'cached_value': value, 'comment_present': c.get('r') in comments}
                if re.search(r'\bedb\s*\(', formula, re.I):
                    try:
                        config = unpack_config(comments.get(c.get('r'), ''))
                        a['config'] = config
                        a['comment_status'] = 'decoded_v10'
                        if configs_dir:
                            output = Path(configs_dir) / ('connection-%d.json' % (len(result['anchors'])+1))
                            write_new_json(output, config)
                            a['config_file'] = str(output.resolve())
                    except (ValueError, KeyError, zlib.error, UnicodeError) as ex:
                        a['comment_status'] = 'unverified: ' + str(ex)
                result['anchors'].append(a)
        for n in z.namelist():
            if n.endswith('.rels') and n.startswith('xl/externalLinks/'):
                result['external_links'].extend(r.attrib for r in ET.fromstring(z.read(n)))
        result['warnings'].append('Shared-formula followers and non-OOXML formats require native Excel inspection.')
    return result


def write_new_json(path, data):
    with Path(path).open('x', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    p = sub.add_parser('inspect'); p.add_argument('workbook'); p.add_argument('--out'); p.add_argument('--configs-dir')
    for cmd in ('pack', 'unpack'):
        p = sub.add_parser(cmd); p.add_argument('source'); p.add_argument('destination')
    args = parser.parse_args()
    if args.command == 'inspect':
        data = inspect_book(args.workbook, args.configs_dir)
        if args.out:
            write_new_json(args.out, data)
        else:
            print(json.dumps(data, ensure_ascii=True, indent=2))
    elif args.command == 'pack':
        value = pack_config(json.loads(Path(args.source).read_text(encoding='utf-8-sig')))
        with Path(args.destination).open('x', encoding='ascii') as f:
            f.write(value)
    else:
        write_new_json(args.destination, unpack_config(Path(args.source).read_text(encoding='utf-8-sig')))


if __name__ == '__main__':
    main()
