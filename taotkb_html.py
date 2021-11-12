
import win32com.client
import os 
import sys

dir_path = os.path.dirname(os.path.realpath(__file__))
path = os.getcwd()+'\\tkb.xlsm'

print(os.getcwd())

xl=win32com.client.Dispatch('Excel.Application')
xl.Workbooks.Open(Filename=path, ReadOnly=1)
xl.Application.Run('Sheet1.ToHTML')
xl.Application.Quit()
del xl