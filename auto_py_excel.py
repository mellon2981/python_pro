import os, os.path, win32com.client, shutil, time
import tentaclio as tentaclio
from stat import S_IREAD, S_IWUSR
from datetime import timedelta, datetime
import psutil
import pythoncom
import threading

def kill_proc_win_(proc_: str) -> None:
    """
    A function to terminate processes with the specified name in the Windows operating system.

    Args:
        proc_ (str): The name of the process to terminate.

    Returns:
        None
    """
    # Iterate through all processes and terminate processes with the specified name
    for process in (process for process in psutil.process_iter() if process.name() == proc_):
        process.kill()


def error_mail(error_: Exception, description_: str) -> None:
    """
    Function to send an error message in Outlook.

    Args:
        error_ (Exception): The error object that occurred.
        description_ (str): Description of the action that caused the error.

    Returns:
        None
    """
    app = win32com.client.Dispatch('Outlook.Application')
    ml = app.CreateItem(0)
    ml.To = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI').Session.Accounts[0]
    ml.Subject = 'Error'
    ml.Body = f'An error occurred when: {description_} :{error_}'
    ml.Send()

def excel_macro(folder: str, file: str, mud: str, macro: str, kill_excel: bool = True) -> None:
    """
    Function to run a macro in an Excel file using win32com.

    Args:
        folder (str): Path to the folder containing the Excel file.
        file (str): The name of the Excel file in which to run the macro.
        mud (str): The name of the module containing the macro.
        macro (str): The name of the macro to run.
        kill_excel (bool, optional): Flag to end the Excel process before executing the macro. Default is True.

    Returns:
        None
    """
    if kill_excel:
        kill_proc_win_('EXCEL.EXE')
    pythoncom.CoInitialize()
    exl = win32com.client.DispatchEx('Excel.Application')
    # exl.DisplayAlerts = False
    exl.Visible = False
    wbk = exl.Workbooks.Open(folder + file)
    try:
        exl.Application.Run(f'{mud}.{macro}')
        time.sleep(5)
        wbk.Save()
        wbk.Close()
        exl.Quit()
        print(f'Макрос {macro} в {folder}{file} отработал!')
    except Exception as _:
        print(f'ERROR: Failed to run macro {macro} в {folder}{file}!')
        error_mail(_, f'{folder}{file}')
        wbk.Close()
        exl.Quit()
