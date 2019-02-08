from itertools import groupby

import pandas as pd


class TableHeader:
    def __init__(self, index, columns):
        self._index = index
        self._columns = columns

        # в mapping заключаются координаты ячеек с данными,
        # т.к. при многоуровневых заголовках между колонками с данными могут быть промежутки.
        # Например, для
        # location     Moscow                          Krasnoyarsk
        # stats        Population     Young       Old  Population     Young       Old
        # mapping будет иметь вид
        # [
        #   [[0, 1], [0, 4]
        # ], [
        #   [1, 1], [1, 2], [1, 3], [1, 4], [1, 5], [1, 6]]
        # ]
        self._cell_mapping = []

    @property
    def cell_mapping(self):
        return self._cell_mapping

    def write(self, worksheet, cursor, cell_format):
        if isinstance(self._columns, pd.MultiIndex):
            self._write_multi_header(worksheet, cursor, cell_format)
        else:
            self._write_flat_header(worksheet, cursor, cell_format)

    def _write_flat_header(self, worksheet, cursor, cell_format):
        """Записывает одноуровневые заголовок"""

        # Записываются сначала заголовки индексов
        col_index = cursor.col
        for name in self._index.names:
            worksheet.write(cursor.row, col_index, name, cell_format)
            col_index += 1

        # Затем заголовки данных
        self._cell_mapping.append([])
        for name in list(self._columns):
            worksheet.write(cursor.row, col_index, u'%s' % name, cell_format)
            self._cell_mapping[0].append([cursor.row, col_index])
            col_index += 1

        # К исходному курсору прибавляется одна строка
        cursor.row += 1

    def _write_multi_header(self, worksheet, cursor, cell_format):
        """Записывает многоуровневые заголовок"""
        levels_count = len(self._columns.levels)

        self._cell_mapping = [[] for _ in range(0, levels_count)]

        row_index = cursor.row
        col_index = cursor.col

        for name in self._index.names:
            worksheet.merge_range(row_index, col_index, row_index + (levels_count - 1), col_index, name, cell_format)
            col_index += 1

        # Многоуровневые заголовки (pd.MultiIndex) можно представить как коллекцию комбинаций возможных уровней:
        #
        # header = pd.MultiIndex.from_product(
        #             [['Moscow', 'Krasnoyarsk'],
        #              ['Population', 'Young', 'Old']],
        #             names=['location', 'stats'])
        #
        # header.get_values()
        # [('Moscow', 'Population') ('Moscow', 'Young') ('Moscow', 'Old')
        #  ('Krasnoyarsk', 'Population') ('Krasnoyarsk', 'Young')
        #  ('Krasnoyarsk', 'Old')]
        #
        # Индексы каждой комбинации отражают уровень вложенности заголовка.
        # Из нашего примера ('Moscow', 'Population') ('Moscow', 'Young'):
        # Moscow соответствует уровню 0, Population и Young - уровню 1.
        #
        # Итерации с группировкой позволяют посчитать сколько подуровней соответствует каждому родительскому уровню,
        # что дает возможность выполнить объединение ячеек с заголовком в таблице excel
        values = self._columns.get_values()
        for level in range(0, levels_count):
            current_col_index = col_index
            for level_tree, group in groupby(values, lambda col: col[:level + 1]):
                group_count = len(list(group))
                current_row_col_index_end = current_col_index + (group_count - 1)

                level_name = level_tree[level]
                if current_col_index == current_row_col_index_end:
                    worksheet.write(row_index, current_col_index, level_name, cell_format)
                else:
                    worksheet.merge_range(row_index, current_col_index, row_index, current_row_col_index_end,
                                          str(level_name), cell_format)

                self._cell_mapping[level].append([row_index, current_col_index])

                current_col_index = current_row_col_index_end + 1

            # К индексу строки прибавляется одна строка на каждый уровень
            row_index += 1

        cursor.row += levels_count


class TableIndex:
    def __init__(self, index):
        self._index = index

        self._cell_mapping = []

    @property
    def cell_mapping(self):
        return self._cell_mapping

    def write(self, worksheet, cursor, cell_format):
        if isinstance(self._index, pd.MultiIndex):
            self.__write_multi_index(worksheet, cursor, cell_format)
        else:
            self.__write_flat_index(worksheet, cursor, cell_format)

    def __write_flat_index(self, worksheet, cursor, cell_format):
        row_index = cursor.row

        self._cell_mapping.append([])
        for name in self._index:
            worksheet.write(row_index, cursor.col, str(name), cell_format)
            self._cell_mapping[0].append([row_index, cursor.col])
            row_index += 1

        cursor.col += 1

    def __write_multi_index(self, worksheet, cursor, cell_format):
        levels_count = len(self._index.levels)

        self._cell_mapping = [[] for _ in range(0, levels_count)]  # create shape levels

        row_index = cursor.row
        col_index = cursor.col

        values = self._index.get_values()
        for level in range(0, levels_count):
            current_row_index = row_index
            for level_tree, group in groupby(values, lambda col: col[:level + 1]):
                group_count = len(list(group))
                current_row_index_end = current_row_index + (group_count - 1)

                level_name = level_tree[level]
                if current_row_index == current_row_index_end:
                    worksheet.write(current_row_index, col_index, level_name, cell_format)
                else:
                    worksheet.merge_range(current_row_index, col_index, current_row_index_end, col_index,
                                          str(level_name), cell_format)

                self._cell_mapping[level].append([current_row_index, col_index])

                current_row_index = current_row_index_end + 1

            col_index += 1

        cursor.col += levels_count


class TableData:
    def __init__(self, df):
        self._df = df

        self._cell_mapping = []

    @property
    def cell_mapping(self):
        return self._cell_mapping

    def write(self, worksheet, cursor, cell_format):
        row_index = cursor.row

        for row in self._df.itertuples():
            current_col_index = cursor.col
            mapping_row = []
            for col in row[1:]:  # row[1:] data without indexes
                worksheet.write(row_index, current_col_index, col, cell_format)
                mapping_row.append([row_index, current_col_index])

                current_col_index += 1

            self._cell_mapping.append(mapping_row)

            row_index += 1

        cursor.row = row_index
        cursor.col += len(self._df.columns)


class Table:
    header_class = TableHeader
    index_class = TableIndex
    data_class = TableData

    cells_format = {
        'index': {'align': 'left', 'valign': 'top', 'border': 1},
        'header': {'align': 'center', 'valign': 'vcenter', 'fg_color': '#D7E4BC', 'border': 6},
        'data': {'align': 'center', 'border': 1}
    }

    def __init__(self, df, name=None):
        self._df = df
        self._name = name

        self._header: TableHeader = None
        self._index: TableIndex = None
        self._data: TableData = None

    @property
    def header(self):
        return self._header

    @property
    def index(self):
        return self._index

    @property
    def data(self):
        return self._data

    def write(self, workbook, worksheet, cursor):
        xl_format = self._get_format(workbook)

        self.write_table_title(cursor, worksheet, xl_format['header'])

        self._header = self.header_class(self._df.index, self._df.columns)
        self._header.write(worksheet, cursor, xl_format['header'])

        self._index = self.index_class(self._df.index)
        self._index.write(worksheet, cursor, xl_format['index'])

        self._data = self.data_class(self._df)
        self._data.write(worksheet, cursor, xl_format['data'])

        self.set_columns_width(worksheet)

        return cursor

    def set_columns_width(self, worksheet):
        """Определяется в дочерник классах для настройки ширины конкретных колонок"""
        pass

    def write_table_title(self, cursor, worksheet, cell_format):
        """Определяется в дочерник классах для записи заголовка таблицы"""
        pass

    def _get_format(self, workbook):
        return {
            'header': workbook.add_format(self.cells_format['header']),
            'index': workbook.add_format(self.cells_format['index']),
            'data': workbook.add_format(self.cells_format['data'])
        }
