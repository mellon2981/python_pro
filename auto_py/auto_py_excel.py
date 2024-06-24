import os, os.path, win32com.client, shutil, time
import tentaclio as tentaclio
from stat import S_IREAD, S_IWUSR
from datetime import timedelta, datetime
import psutil
import pythoncom
from UliPlot.XLSX import auto_adjust_xlsx_column_width
import importlib.util
import threading
from auto_py_outlook import error_mail


# If you need to load a module from another folder
# def module_from_file(module_name, file_path):
#     spec = importlib.util.spec_from_file_location(module_name, file_path)
#     module = importlib.util.module_from_spec(spec)
#     spec.loader.exec_module(module)
#     return module


# auto_py_outlook_agro = module_from_file('~module_name', r'~file_path')


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
        print(f'Macro {macro} in {folder}{file} worked!')
    except Exception as _:
        print(f'ERROR: Failed to run macro {macro} in {folder}{file}!')
        error_mail(_, f'{folder}{file}')
        wbk.Close()
        exl.Quit()


def refresh_excel(folder: str, file: str, kill_excel: bool = False) -> None:
    """
    Function to update a specific Excel file.

    Args:
        folder (str): Path to the folder containing the Excel file.
        file (str): The name of the Excel file to update.
        kill_excel (bool): Flag to end the Excel process after updating.

    Returns:
        None
    """
    if kill_excel:
        kill_proc_win_('EXCEL.EXE')
    pythoncom.CoInitialize()
    exl = win32com.client.DispatchEx('Excel.Application')
    exl.DisplayAlerts = False
    exl.Visible = False
    wbk = exl.Workbooks.Open(f'{folder}{file}')
    try:
        wbk.RefreshAll()
        exl.CalculateUntilAsyncQueriesDone()
        time.sleep(5)
        wbk.Save()
        wbk.Close()
        exl.Quit()
        print(f'The file {folder}{file} has been updated!')
    except Exception as _:
        print(f'ERROR: The file {folder}{file} has not been updated!')
        error_mail(_, f'{folder}{file}')
        wbk.Close()
        exl.Quit()


def refresh_excel_files(folder_path: str, excel_files: list, kill_excel: bool = False) -> None:
    """
    Function to update Excel files in a specified folder using multi-threading.

    Args:
        folder_path (str): Path to the folder containing Excel files.
        excel_files (list): List of Excel files to update.
        kill_excel (bool, optional): Flag to end the Excel process after updating. Default is False.

    Returns:
        None
    """
    try:
        threads = []
        for i in excel_files:
            t = threading.Thread(target=refresh_excel, args=(folder_path, i, kill_excel))
            threads.append(t)
            t.start()
        # Waiting for all threads to complete
        for t in threads:
            t.join()
    except Exception as _:
        error_mail(_, f'Error when refresh excel files!')


def copy_excel(path_from: str, path_to: str, file: str) -> None:
    """
    Function for copying an Excel file from one directory to another.

    Args:
        path_from (str): Path to the source directory from where you want to copy the file.
        path_to (str): Path to the target directory where you want to copy the file.
        v_file (str): The name of the Excel file to copy.

    Returns:
        None
    """
    try:
        os.chmod(f'{path_to}{file}', S_IWUSR)
        os.remove(os.path.join(path_to, file))
        shutil.copy(f'{path_from}{file}', f'{path_to}{file}')
        os.chmod(f'{path_to}{file}', S_IREAD)
    except Exception as _:
        error_mail(_, f'Error when copy {file} to {path_to}!')


def file_date(path: str, file: str, all_files: bool = False) -> str:
    """
    Get the minimum or last update date of a file in a specified path.

    Args:
        path (str): The path to the folder containing the file(s).
        file (str): The file name to check for the last update date.
        all_files (bool): If True, find the file with the minimum update date in the folder.

    Returns:
        str: The last update date of the file in the specified path.

    Raises:
        Exception: If there is an error while getting the file update date.
    """
    try:
        d = None
        if all_files:
            for f in os.listdir(path):
                p = os.path.join(path, f)
                if os.path.isfile(p):
                    f_date = datetime.fromtimestamp(os.path.getmtime(p)).strftime('%Y-%m-%d')
                    if d is None or f_date < d:
                        d = f_date
            return d
        else:
            d = datetime.fromtimestamp(os.path.getmtime(os.path.join(path, file))).strftime('%Y-%m-%d')
            return d
    except Exception as _:
        error_mail(_, 'Error getting file update date!')


def filepath_date(path: str) -> str:
    """
    Get last update date of a file in a specified path.

    Args:
        path (str): The path to the folder containing the file(s).

    Returns:
        str: The last update date of the file in the specified path.

    Raises:
        Exception: If there is an error while getting the file update date.
    """
    try:
        d = datetime.fromtimestamp(os.path.getmtime(path)).strftime('%Y-%m-%d')
        return d
    except Exception as _:
        error_mail(_, 'Error getting file update date!')


def last_date_sent_mail(subject):
    """
    Get last date of mail in outlook

    Args:
        path (str): The path to the folder containing the file(s).

    Returns:
        str: The last update date of the file in the specified path.

    Raises:
        Exception: If there is an error while getting the file update date.
    """
    try:
        app = win32com.client.Dispatch("Outlook.Application")
        folder = app.Session.GetDefaultFolder(5)
        item = folder.Items.Restrict(f"[Subject] = '{subject}'")
        if item.Count == 0:
            return (datetime.today() - timedelta(days=1)).strftime('%Y-%m-%d')
        item.Sort('[SentOn]', True)
        return item[0].SentOn.strftime('%Y-%m-%d')
    except Exception as _:
        print(_)
