'''

Author   :  AnurajPilanku

Use case : Alphine GSR Incident Quality Review

'''
import win32com.client
#import os, os.path
import time
try:
    doc = ""
    xl1 = ""
    path = r"\\acprd01\E\3M_CAC\Alphine\alphine.xlsm"

    xl1 = win32com.client.Dispatch("Excel.Application")
    xl1.Visible = False
    xl1.DisplayAlerts = False
    time.sleep(5)
    doc = xl1.Workbooks.Open(path)
    time.sleep(5)
    xl1.Application.Run("UpdateNotificationV2")
    time.sleep(20)
    doc.Save()
    doc.Close()
    xl1.Quit()
    print("success")
except:
    import os
    import win32com.shell.shell as shell
    import time

    # commands='taskkill /f /im EXCEL.EXE'
    # shell.ShellExecuteEx(lpVerb='runas',lpFile='cmd.exe',lpParameters='/c'+commands)
    os.system('taskkill /f /im EXCEL.EXE')
    time.sleep(10)
    doc = ""
    xl1 = ""
    path = r"\\acprd01\E\3M_CAC\Alphine\alphine.xlsm"

    xl1 = win32com.client.Dispatch("Excel.Application")
    xl1.Visible = False
    xl1.DisplayAlerts = False
    time.sleep(5)
    doc = xl1.Workbooks.Open(path)
    time.sleep(5)
    xl1.Application.Run("UpdateNotificationV2")
    time.sleep(20)
    doc.Save()
    doc.Close()
    xl1.Quit()
    print("success")
