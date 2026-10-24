"""Find usable EDB metrics in a workbook, then prepare a dated native query."""
import argparse
import json
from edb_tools import inspect_book, write_new_json
from prepare_query import prepare


def candidates(book, terms):
    found = []
    for anchor in book['anchors']:
        for metric in anchor.get('config', {}).get('metric', []):
            searchable = ' '.join(str(metric.get(k, '')) for k in
                                  ('code', 'name', 'funcName', 'enName', 'unit')).casefold()
            if all(term.casefold() in searchable for term in terms):
                found.append({'sheet': anchor['sheet'], 'anchor': anchor['cell'],
                              'code': metric['code'], 'name': metric.get('name'),
                              'unit': metric.get('unit'), 'freq_enum': metric.get('freq'),
                              'changeRecord': metric.get('changeRecord', []),
                              'metric': metric})
    return found


def main():
    p = argparse.ArgumentParser(description=__doc__)
    sub = p.add_subparsers(dest='command', required=True)
    search = sub.add_parser('search')
    search.add_argument('workbook')
    search.add_argument('--terms', nargs='*', default=[])
    search.add_argument('--out')
    select = sub.add_parser('prepare')
    select.add_argument('workbook')
    select.add_argument('--sheet', required=True)
    select.add_argument('--anchor', required=True)
    select.add_argument('--codes', nargs='+', required=True)
    select.add_argument('--start')
    select.add_argument('--end')
    select.add_argument('--last', type=int)
    select.add_argument('--output', required=True)
    a = p.parse_args()
    book = inspect_book(a.workbook)
    if a.command == 'search':
        result = {'matches': candidates(book, a.terms),
                  'unresolved_anchors': [{k: x.get(k) for k in ('sheet', 'cell', 'formula', 'comment_status')}
                                         for x in book['anchors'] if not x.get('config')],
                  'note': 'Literal candidate search, not semantic verification or a complete Wind catalogue.'}
        if a.out:
            write_new_json(a.out, result)
        print(json.dumps(result, ensure_ascii=True, indent=2))
    else:
        matches = [x for x in book['anchors'] if x['sheet'] == a.sheet and x['cell'] == a.anchor.upper()]
        if len(matches) != 1 or not matches[0].get('config'):
            raise ValueError('Choose one decoded EDB anchor from search results')
        query = prepare(matches[0]['config'], a.codes, a.start, a.end, a.last)
        write_new_json(a.output, query)
        print(json.dumps({'status': 'prepared_refresh_required', 'output': a.output,
                          'source_sha256': book['sha256'], 'sheet': a.sheet,
                          'anchor': a.anchor, 'codes': a.codes}))


if __name__ == '__main__':
    main()
