from enum import Enum
from typing import List

from pandex import Table
from pandex.sheet import Cursor

CHART_AREA_PATTERN = {
    'pattern': 'light_downward_diagonal',
    'fg_color': '#dddddd',
    'bg_color': 'white',
}


class Unit(Enum):
    PERCENT = 1
    PIECE = 2


class ModeColor(Enum):
    POS = '#92D050'
    NEG = '#CB3C39'
    NEU = '#31859C'
    BRO = 'yellow'
    UNK = '#715197'


class PieChart:
    def __init__(self, name: str, table: Table, target_row: int = 0, unit: Unit = Unit.PERCENT):
        self._name = name

        self._table = table
        self._target_row = target_row

        self._unit = unit

    def write(self, workbook, worksheet, cursor: Cursor):
        worksheet_name = worksheet.get_name()
        chart_data = self._get_chart_data(worksheet_name)

        chart = workbook.add_chart({'type': 'pie'})
        chart.add_series(chart_data['series'])

        chart.set_title({'name': chart_data['name']})
        chart.set_legend({'position': 'top'})
        chart.set_chartarea({
            'pattern': CHART_AREA_PATTERN
        })

        worksheet.insert_chart(cursor.row, cursor.col, chart)

        cursor.row += 14
        cursor.col += 7

    def _get_chart_data(self, sheet_name):
        data_labels = {
            'leader_lines': True
        }

        if self._unit == Unit.PERCENT:
            data_labels['percentage'] = True
        elif self._unit == Unit.PIECE:
            data_labels['value'] = True

        header = self._table.header
        header_mapping = header.cell_mapping

        data = self._table.data
        data_mapping = data.cell_mapping

        return {
            'name': self._name,
            'series': {
                'categories': [sheet_name] + header_mapping[-1][0] + header_mapping[-1][-1],
                'values': [sheet_name] + data_mapping[self._target_row][0] + data_mapping[self._target_row][-1],
                'data_labels': data_labels,
                'points': [
                    {'fill': {'color': mc.value}} for mc in ModeColor
                ]
            }
        }


class ColumnChart:
    def __init__(self, name, table, unit=Unit.PERCENT):
        self._name = name
        self._table = table
        self._unit = unit

    def write(self, workbook, worksheet, cursor):
        worksheet_name = worksheet.get_name()
        chart_data = self._get_chart_data(worksheet_name)
        for data in chart_data:
            chart = workbook.add_chart(
                {'type': 'column', 'subtype': 'stacked' if self._unit == Unit.PIECE else 'percent_stacked'})

            for series in data['series']:
                chart.add_series(series)

            chart.set_title({'name': self._name})
            chart.set_chartarea({
                'pattern': CHART_AREA_PATTERN
            })

            chart.set_plotarea({
                'pattern': CHART_AREA_PATTERN
            })

            chart.show_hidden_data()

            worksheet.insert_chart(cursor.row, cursor.col, chart)

            cursor.row += 14
            cursor.col += 7

    def _get_chart_data(self, sheet_name):
        mode_colors = [mc.value for mc in ModeColor]

        header = self._table.header
        header_mapping = header.cell_mapping

        index = self._table.index
        index_mapping = index.cell_mapping

        data = self._table.data
        data_mapping = data.cell_mapping

        return [{
            'name': self._name,
            'series': [{
                'name': [sheet_name] + header_mapping[-1][i],
                'categories': [sheet_name] + index_mapping[-1][0] + index_mapping[-1][-1],
                'values': [sheet_name] + data_mapping[0][i] + data_mapping[-1][i],
                'fill': {'color': mode_colors[i]},
                'data_labels': {'value': True}
            } for i in range(0, len(header_mapping[-1]))]
        }]


class LineChart:
    def __init__(self, name: str, table: Table, target_rows: List[int], skip_columns: int = 0):
        self._name = name

        self._table = table
        self._target_rows = target_rows
        self._skip_columns = skip_columns

    def write(self, workbook, worksheet, cursor: Cursor):
        worksheet_name = worksheet.get_name()

        chart = workbook.add_chart({'type': 'line'})

        chart_data = self._get_chart_data(worksheet_name)
        for series in chart_data:
            chart.add_series(series)

        chart.set_title({'name': self._name})
        chart.set_legend({'position': 'top'})
        chart.set_chartarea({
            'pattern': CHART_AREA_PATTERN
        })

        worksheet.insert_chart(cursor.row, cursor.col, chart)

        cursor.row += 14
        cursor.col += 7

    def _get_chart_data(self, sheet_name):
        header = self._table.header
        header_mapping = header.cell_mapping

        index = self._table.index
        index_mapping = index.cell_mapping

        data = self._table.data
        data_mapping = data.cell_mapping

        return [{
            'name': [sheet_name] + index_mapping[-1][row],
            'categories': [sheet_name] + header_mapping[-1][0 + self._skip_columns] + header_mapping[-1][-1],
            'values': [sheet_name] + data_mapping[row][0 + self._skip_columns] + data_mapping[row][-1],
        } for row in self._target_rows]
