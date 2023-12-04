# turn off the formatter
# fmt: off

'''
this object is used to load table into object
by different way to define the table and provide different way to process
1. get or set value
2. table information (where is this table from)
(contain the shape, sheet, book, and file trace)
3. table comparison or plot
4. column or row access for different comparison method
(need to use column obj? or not? => need to think about)

'''

# whatch out the index of excel is y,x ; which is row, column

import xlwings as xw
import win32api as win_ap

# pandas and matplot is data science's good friend~
# import pandas as pd


class table_gen():
    def __init__(self, table_type0=0, **kwargs):
        '''
        table object:
        for table index cell input: 'ind_row', 'ind_column'
        => this is single cell and auto scan for table

        '''
        self.table_type = table_type0
        self.ctrl_para_dict = {'ind_row':0, 'ind_col':0, }

        prog_fake = 0
        if prog_fake == 1:
            self.wb = xw.Book()
            self.sh_comp = self.wb.sheets("temp")
            ind_range0 = self.sh_comp.range((1, 1))

        '''
        table object should contain the
        '''
        # record the range for this table
        self.ind_range = ind_range0
        # knowing the dimenstion and starting index
        self.sheet_name = ind_range0.shape
        self.book_name = ind_range0.sheet
        self.file_trace = ind_range0.sheet.book


        # dimension of the range
        self.ind_row = self.ran_shape[0]
        self.ind_col = self.ran_shape[1]

        # starting index (normalize)
        self.nor_Row = ind_range0.row
        self.nor_Col = ind_range0.column


        pass
