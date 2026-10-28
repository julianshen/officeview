"""
Writes tests/fixtures/xlsx-sparklines/xlsxwriter-written.xlsx with XlsxWriter, an independent writer, so the viewer's
resident and paged sparkline parsers are checked against XML this repository did not generate (spec section 3,
"files written by other writers"). The output is committed; XlsxWriter is not a project dependency.

    python3 -m venv /tmp/xw && /tmp/xw/bin/pip install xlsxwriter==3.3.0
    /tmp/xw/bin/python -I scripts/sparkline-refs/xlsxwriter-case.py

Layout matches cases.ts so Excel captures line up: hosts in column J (40 characters), host rows 60 pt, gridlines off.
Data lives on Sheet2 (cross-sheet ranges), except row 11's host also holds text.
"""
import os
import sys

import xlsxwriter

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '../../tests/fixtures/xlsx-sparklines/xlsxwriter-written.xlsx')

DATA = [
    [-2, 2, 3, -1, 0, 5, 4, 1],            # 1 line, markers, style 1
    [-3, 2, -1, 4, -4, 3, 1, -2],          # 2 column, every point flag + negative, style 3
    [1, -1, -1, 1, -1, 1, 1, -1],          # 3 win/loss, negative points + axis, style 20
    [4, 7, 3, 9, 2, 8, 5, 6],              # 4 line reversed (rightToLeft) with first/last
    [3, 7, None, 4, None, None, 8, 5],     # 5 line, empty cells connected (span)
    [4, 7, 3, 9, 2, 8, 5, 6],              # 6 line, date axis on row 13 dates, weight 2.25
    [50, 52, 51, 55, 53, 54, 50, 56],      # 7 column, custom min/max 0..100
    [4, 7, 3, 9, 2, 8, 5, 6],              # 8 line, show_hidden with column C hidden
    [1, 2, 3, 4, 5, 6, 7, 8],              # 9 grouped (with row 10), style 6, group max/min
    [10, 20, 30, 40, 50, 60, 70, 80],      # 10
    [2, 5, 3, 8, 6, 9, 4, 7],              # 11 line in a host that also holds text
]
DATES = [45000, 45001, 45003, 45010, 45011, 45020, 45021, 45030]  # row 13 of Sheet2


def rng(row: int) -> str:
    return f'Sheet2!A{row}:H{row}'


def main() -> int:
    wb = xlsxwriter.Workbook(OUT)
    host = wb.add_worksheet('Sheet1')
    data = wb.add_worksheet('Sheet2')
    for r, values in enumerate(DATA):
        for c, v in enumerate(values):
            if v is not None:
                data.write_number(r, c, v)
    for c, d in enumerate(DATES):
        data.write_number(12, c, d)
    data.set_column('C:C', None, None, {'hidden': True})

    host.hide_gridlines(2)
    host.set_column('J:J', 40)
    for r in range(11):
        host.set_row(r, 60)
    host.write_string('J11', 'Growth')

    host.add_sparkline('J1', {'range': rng(1), 'markers': True, 'style': 1})
    host.add_sparkline('J2', {'range': rng(2), 'type': 'column', 'style': 3, 'high_point': True, 'low_point': True,
                              'first_point': True, 'last_point': True, 'negative_points': True})
    host.add_sparkline('J3', {'range': rng(3), 'type': 'win_loss', 'style': 20, 'negative_points': True, 'axis': True})
    host.add_sparkline('J4', {'range': rng(4), 'reverse': True, 'markers': True, 'first_point': True, 'last_point': True,
                              'series_color': '#E965E0'})
    host.add_sparkline('J5', {'range': rng(5), 'empty_cells': 'connect', 'markers': True})
    host.add_sparkline('J6', {'range': rng(6), 'date_axis': 'Sheet2!A13:H13', 'weight': 2.25, 'markers': True})
    host.add_sparkline('J7', {'range': rng(7), 'type': 'column', 'min': 0, 'max': 100})
    host.add_sparkline('J8', {'range': rng(8), 'show_hidden': True, 'markers': True})
    host.add_sparkline('J9', {'location': ['J9', 'J10'], 'range': [rng(9), rng(10)], 'style': 6, 'markers': True,
                              'max': 'group', 'min': 'group'})
    host.add_sparkline('J11', {'range': rng(11), 'high_point': True, 'low_point': True})
    wb.close()
    print(f'wrote {os.path.normpath(OUT)}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
