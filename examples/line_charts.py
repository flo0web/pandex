import pandas as pd
import xlsxwriter

from pandex import Sheet, Table, Side, LineChart


class ForecastTable(Table):
    def write_table_title(self, cursor, worksheet, cell_format):
        columns_num = len(self._df.columns)
        worksheet.merge_range(cursor.row, cursor.col, cursor.row, cursor.col + columns_num, self._name, cell_format)

        cursor.row += 1
        return cursor

    def set_columns_width(self, worksheet):
        worksheet.set_column(0, 0, 24)

        columns_num = len(self._df.columns)
        worksheet.set_column(1, columns_num, 4)


if __name__ == '__main__':
    df = pd.DataFrame(
        [
            [1, 1, 1, 1, 1],
            [3.5, 3.5, 3.5, 3.5, 3.5],
            [1, 2, 1, 1, 3],
            [5, 5, 5, 5, 5],
            [3.5, 3.6, 3.4, 3.7, 3.9],
        ],
        index=[
            'Отзывы из проекции',
            'Рейтинг из проекции',
            'Отзывы из расписания',
            'Рейтинг из расписания',
            'Итоговый рейтинг'
        ],
        columns=['1', '2', '3', '4', '5']
    )

    file_name = 'line.xlsx'

    workbook = xlsxwriter.Workbook(file_name)
    sheet = Sheet(workbook, 'Line')

    tables = []
    tables_group = sheet.create_shape()
    for i in range(0, 3):
        table = ForecastTable(df, name='Site %s' % i)
        tables_group.add(table, side=Side.BOTTOM, margin_rows=12 if i > 0 else 0)
        tables.append(table)

    charts_group = sheet.create_shape(side=Side.RIGHT, margin_cols=1)
    for i in range(0, 3):
        chart = LineChart('Прогноз рейтинга', tables[i], target_row=4)
        charts_group.add(chart, side=Side.BOTTOM, margin_rows=5 if i > 0 else 0)

    workbook.close()
