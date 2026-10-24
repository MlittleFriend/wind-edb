"""Create a NEW Excel/Wind probe; never edits an existing workbook."""
import argparse
import json
import os
from pathlib import Path
import site
import time
from edb_tools import pack_config, inspect_book, write_new_json


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--config', required=True)
    p.add_argument('--output', required=True)
    p.add_argument('--addin', required=True)
    p.add_argument('--timeout', type=int, default=90)
    p.add_argument('--dependency-dir')
    a = p.parse_args()
    output = Path(a.output).resolve()
    evidence_path = output.with_suffix('.evidence.json')
    if output.suffix.lower() != '.xlsx' or output.exists() or evidence_path.exists():
        raise ValueError('Choose an unused .xlsx output and evidence path')
    if not 1 <= a.timeout <= 600:
        raise ValueError('timeout must be 1..600 seconds')
    addin = Path(a.addin).resolve(strict=True)
    config = json.loads(Path(a.config).read_text(encoding='utf-8-sig'))
    comment = pack_config(config)
    dll_handles = []
    if a.dependency_dir:
        deps = Path(a.dependency_dir).resolve(strict=True)
        site.addsitedir(str(deps))
        dll = deps / 'pywin32_system32'
        if dll.is_dir():
            dll_handles.append(os.add_dll_directory(str(dll)))
    import pythoncom
    import win32com.client
    output.parent.mkdir(parents=True, exist_ok=True)
    evidence = {'status': 'started', 'business_validation': 'not_performed', 'output': str(output)}
    excel = None
    book = None
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        evidence['owned_excel_hwnd'] = excel.Hwnd
        print(json.dumps({'owned_excel_hwnd': excel.Hwnd}), flush=True)
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        book = excel.Workbooks.Add()
        excel.Calculation = -4105
        excel.CalculateBeforeSave = True
        excel.Workbooks.Open(str(addin), 0, True)
        sheet = book.Worksheets(1)
        sheet.Name = 'EDB_probe'
        cell = sheet.Range('A4')
        cell.AddComment(comment)
        cell.Formula = "='" + addin.name.replace("'", "''") + "'!edb()"
        sheet.Calculate()
        deadline = time.monotonic() + a.timeout
        observed = False
        while time.monotonic() < deadline:
            pythoncom.PumpWaitingMessages()
            values = sheet.Range('B5:Z100').Value2
            numeric_count = sum(isinstance(v, (float, int)) and not isinstance(v, bool)
                                and not (isinstance(v, int) and v < -2000000000)
                                for row in values for v in row)
            evidence['calculation_state'] = excel.CalculationState
            evidence['numeric_cells_in_sample'] = numeric_count
            if numeric_count and excel.CalculationState == 0:
                observed = True
                break
            time.sleep(1)
        evidence['anchor_formula'] = cell.Formula
        evidence['used_range'] = sheet.UsedRange.Address
        evidence['sample_A4_H14'] = sheet.Range('A4:H14').Value2
        if not observed:
            raise TimeoutError('No numerical output plus completed calculation within deadline')
        book.SaveAs(str(output), 51)
        book.Close(False)
        book = None
        evidence['disk_inspection'] = inspect_book(output)
        if evidence['disk_inspection']['errors']:
            raise ValueError('Saved workbook contains Excel errors; inspect evidence')
        anchors = evidence['disk_inspection']['anchors']
        if not anchors or anchors[0].get('config') != config:
            raise ValueError('Saved anchor/config did not round-trip')
        evidence['status'] = 'technical_return_observed_business_validation_pending'
    except Exception as ex:
        evidence['status'] = 'failed'
        evidence['error'] = str(ex)
        raise
    finally:
        try:
            if book is not None:
                book.Close(False)
            if excel is not None:
                excel.Quit()
        finally:
            write_new_json(evidence_path, evidence)
            pythoncom.CoUninitialize()
    print(json.dumps({'status': evidence['status'], 'evidence': str(evidence_path)}), flush=True)


if __name__ == '__main__':
    main()
