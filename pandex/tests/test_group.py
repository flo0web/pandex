from unittest import TestCase
from unittest.mock import Mock

import numpy as np
import pandas as pd

from pandex import Sheet, Table
from sheet import Side


class GroupTestCase(TestCase):
    def setUp(self):
        self.df = pd.DataFrame(
            np.random.randn(3, 3),
            index=['1', '2', '3'],
            columns=['a', 'b', 'c']
        )

        self.workbook = Mock()
        self.worksheet = Mock()
        self.cell_format = Mock()

    def test_write_table(self):
        sheet = Sheet(self.workbook, 'Test')

        group = sheet.create_shape()

        table = Table(self.df)
        group.add(table)

        _, group_cursor_end = group.get_cursors()

        self.assertEqual(4, group_cursor_end.row)
        self.assertEqual(4, group_cursor_end.col)

    def test_write_table_right(self):
        sheet = Sheet(self.workbook, 'Test')

        group = sheet.create_shape()

        table = Table(self.df)
        group.add(table)

        new_table = Table(self.df)
        group.add(new_table)

        _, group_cursor_end = group.get_cursors()

        self.assertEqual(4, group_cursor_end.row)
        self.assertEqual(8, group_cursor_end.col)

    def test_write_table_bottom(self):
        sheet = Sheet(self.workbook, 'Test')

        group = sheet.create_shape()

        table = Table(self.df)
        group.add(table)

        new_table = Table(self.df)
        group.add(new_table, side=Side.BOTTOM)

        _, group_cursor_end = group.get_cursors()

        self.assertEqual(8, group_cursor_end.row)
        self.assertEqual(4, group_cursor_end.col)
