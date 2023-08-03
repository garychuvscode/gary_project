"""
this object used to handle all the report arragement after raw dat input
handling the combing of raw data from different sheet or area

try to make good summary to make report easier for others to understand
"""

# import related tool needed for report arragement
import xlwings as xw
import logging as log

# mainly for only process the report


class report_arragement:
    def __init__(self, file_name0=""):
        """
        this class used to process some regular copy, paste and plot of report \n
        file input need to be .xlsm
        """
        # input file name of report
        self.file_name = str(file_name0)

        self.wb = xw.books(self.file_name)
        print(f"select {self.file_name} as report file")

        pass

    def Buck_eff_load_regulation(self):
        pass

    def LDO_load_regulation(self):
        pass

    pass


if __name__ == "__main__":
    #  the testing code for this file object

    pass
