"""Select verified metrics and an explicit observation-date window; no code guessing."""
import argparse
import copy
import datetime as dt
import json
from pathlib import Path
from edb_tools import validate_config, write_new_json


def date_string(value):
    try:
        return dt.datetime.strptime(value, '%Y-%m-%d').strftime('%Y%m%d')
    except ValueError as ex:
        raise ValueError('Use a real calendar date YYYY-MM-DD: ' + value) from ex


def prepare(config, codes, start=None, end=None, last=None):
    validate_config(config)
    if not codes or len(codes) != len(set(codes)):
        raise ValueError('Provide unique verified indicator codes')
    if last is not None:
        if start is not None or end is not None or not isinstance(last, int) or last < 1:
            raise ValueError('last must be positive and cannot combine with dates')
    elif start is None or end is None:
        raise ValueError('Provide both start and end observation dates')
    out = copy.deepcopy(config)
    selected = []
    for code in codes:
        matches = [m for m in out['metric'] if m['code'] == code]
        if len(matches) != 1:
            raise ValueError('Code missing or ambiguous in verified template: ' + code)
        selected.append(matches[0])
    out['metric'] = selected
    setting = out['timeFrameSetting']
    if last is not None:
        setting.update(periodType='ptLastObs', lastType='ltObservations', lastObs=last)
    else:
        first, final = date_string(start), date_string(end)
        if first > final:
            raise ValueError('start must not exceed end')
        setting.update(periodType='ptObsBetween', obsFrom=first, obsTo=final)
    return out


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--template', required=True)
    p.add_argument('--codes', nargs='+', required=True)
    p.add_argument('--start', help='Observation date YYYY-MM-DD, not publication cutoff')
    p.add_argument('--end')
    p.add_argument('--last', type=int, help='Latest N observations, not latest N days')
    p.add_argument('--output', required=True)
    a = p.parse_args()
    config = json.loads(Path(a.template).read_text(encoding='utf-8-sig'))
    out = prepare(config, a.codes, a.start, a.end, a.last)
    write_new_json(a.output, out)
    print(json.dumps({'codes': a.codes, 'timeFrameSetting': out['timeFrameSetting'],
                      'status': 'prepared_requires_native_return_validation'}, ensure_ascii=True))


if __name__ == '__main__':
    main()
