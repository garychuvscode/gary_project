# turn off the formatter
# fmt: off

'''
231210 function done 
refer to the file: 20220226_python call VBA marco for more settings note
'''

import win32com.client as win32

# file trace for VBA related is not same format with python
# file_name0 = r"c:\py_gary\test_excel\VBA_temp.xlsm"

# these two method of open app is ok 
# xl = win32.gencache.EnsureDispatch('Excel.Application')
xl = win32.Dispatch('Excel.Application')
xl.Visible = True
print(xl)
ss = xl.Workbooks.Add()
print(ss)

# 231210: no need for saving, all the code setting are in 
# the text file 

# ss.name = 'VBA_temp.xlsm'
# ss.SaveAs(file_name0)

xlmodule = ss.VBProject.VBComponents.Add(1)
xlmodule.Name = 'testing123'
code = '''
sub TestMacro()
    msgbox "Testing 1 2 3"
    end sub
'''
xlmodule.CodeModule.AddFromString(code)
ss.Application.Run('testing123.TestMacro')
# if not going to record the VBA code, just need a window
#  for processing

# close the window and app after finished 
ss.Close(SaveChanges=False)
xl.Quit()


# need to solve trust issue of excel 
# think about to integrate into the object or control sheet 
