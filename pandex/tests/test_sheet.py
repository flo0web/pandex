from unittest import TestCase
from unittest.mock import Mock

import numpy as np
import pandas as pd

from pandex import Sheet, Table
from sheet import Side


class SheetTestCase(TestCase):
    def setUp(self):
        self.df = pd.DataFrame(
            np.random.randn(3, 3),
            index=['1', '2', '3'],
            columns=['a', 'b', 'c']
        )

        self.workbook = Mock()
        self.worksheet = Mock()
        self.cell_format = Mock()

    def test_group_single(self):
        sheet = Sheet(self.workbook, 'Test')

        group = sheet.create_shape()

        table = Table(self.df)
        group.add(table)

        self.assertEqual(4, sheet.cursor.row)
        self.assertEqual(4, sheet.cursor.col)

    def test_group_right(self):
        sheet = Sheet(self.workbook, 'Test')

        first_group = sheet.create_shape()

        first_table = Table(self.df)
        first_group.add(first_table)

        second_group = sheet.create_shape()

        _, second_group_cursor_end = second_group.get_cursors()
        self.assertEqual(0, second_group_cursor_end.row)
        self.assertEqual(4, second_group_cursor_end.col)

        second_table = Table(self.df)
        second_group.add(second_table)

        self.assertEqual(4, sheet.cursor.row)
        self.assertEqual(8, sheet.cursor.col)

        third_table = Table(self.df)
        second_group.add(third_table, side=Side.BOTTOM)

        self.assertEqual(8, sheet.cursor.row)
        self.assertEqual(8, sheet.cursor.col)
