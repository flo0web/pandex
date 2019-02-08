from unittest import TestCase
from unittest.mock import Mock

import numpy as np
import pandas as pd

from pandex.sheet import Cursor
from pandex.table import TableHeader, TableIndex, TableData, Table


class TableHeaderTestCase(TestCase):
    def setUp(self):
        header = pd.MultiIndex.from_product(
            [
                ['Moscow', 'Krasnoyarsk'],
                ['Population', 'Young', 'Old']
            ],
            names=['location', 'stats']
        )

        self.df = pd.DataFrame(
            np.random.randn(5, 6),
            index=['a', 'b', 'c', 'd', 'e'],
            columns=header
        )

        self.worksheet = Mock()
        self.cell_format = Mock()

    def test_multi_headers(self):
        cursor = Cursor()

        header = TableHeader(self.df.index, self.df.columns)
        header.write(self.worksheet, cursor, self.cell_format)

        self.assertListEqual(
            [[[0, 1], [0, 4]], [[1, 1], [1, 2], [1, 3], [1, 4], [1, 5], [1, 6]]],
            header._cell_mapping
        )

        self.assertEqual(2, cursor.row)
        self.assertEqual(0, cursor.col)


class TableIndexTestCase(TestCase):
    def setUp(self):
        index = pd.MultiIndex.from_product(
            [
                ['Moscow', 'Krasnoyarsk'],
                ['Population', 'Young', 'Old']
            ],
            names=['location', 'stats']
        )

        self.df = pd.DataFrame(
            np.random.randn(6, 5),
            index=index,
            columns=['a', 'b', 'c', 'd', 'e']
        )

        self.worksheet = Mock()
        self.cell_format = Mock()

    def test_multi_index(self):
        cursor = Cursor()

        index = TableIndex(self.df.index)
        index.write(self.worksheet, cursor, self.cell_format)

        self.assertListEqual(
            [[[0, 0], [3, 0]], [[0, 1], [1, 1], [2, 1], [3, 1], [4, 1], [5, 1]]],
            index._cell_mapping
        )

        self.assertEqual(0, cursor.row)
        self.assertEqual(2, cursor.col)


class TableDataTestCase(TestCase):
    def setUp(self):
        self.df = pd.DataFrame(
            np.random.randn(3, 3),
            index=['1', '2', '3'],
            columns=['a', 'b', 'c']
        )

        self.worksheet = Mock()
        self.cell_format = Mock()

    def test_data(self):
        cursor = Cursor()

        data = TableData(self.df)
        data.write(self.worksheet, cursor, self.cell_format)

        self.assertListEqual(
            [[[0, 0], [0, 1], [0, 2]], [[1, 0], [1, 1], [1, 2]], [[2, 0], [2, 1], [2, 2]]],
            data._cell_mapping
        )

        self.assertEqual(3, cursor.row)
        self.assertEqual(3, cursor.col)


class TableTestCase(TestCase):
    def setUp(self):
        self.df = pd.DataFrame(
            np.random.randn(3, 3),
            index=['1', '2', '3'],
            columns=['a', 'b', 'c']
        )

        self.workbook = Mock()
        self.worksheet = Mock()
        self.cell_format = Mock()

    def test_table(self):
        cursor = Cursor()

        table = Table(self.df)
        table.write(self.workbook, self.worksheet, cursor)

        self.assertEqual(4, cursor.row)
        self.assertEqual(4, cursor.col)
