import win32com.client
from datetime import timedelta, datetime


def error_mail(error: Exception, description: str) -> None:
    """
    Function to send an error message via Outlook.

    Args:
        error (Exception): The error object that occurred.
        description (str): Description of the action that caused the error.

    Returns:
        None
    """
    app = win32com.client.Dispatch('Outlook.Application')
    ml = app.CreateItem(0)
    ml.To = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI').Session.Accounts[0]
    ml.Subject = 'Error'
    ml.Body = f'An error occurred when {description} :{error}'
    ml.Send()


def send_mail_outlook(recip: str = None, mail_theme: str = None, body: str = None, attachment: str = None, send_me: bool = True) -> None:
    """
    Function for sending email via Outlook.

    Args:
        recip (str): Recipients of the letter.
        mail_theme (str): Letter subject.
        body (str): Text of the letter.
        attachment (str): Path to the attached file.
        send_me (bool): Flag to send a copy to yourself.

    Returns:
        None
    """
    try:
        my_acc = f'{win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI").Session.Accounts[0]}; ' if send_me else ''
        recip_other = recip if recip is not None else ''
        if f'{my_acc}{recip_other}' != '':
            app = win32com.client.Dispatch('Outlook.Application')
            ml = app.CreateItem(0)
            ml.To = f'{my_acc}{recip_other}'.rstrip('; ')
            ml.Subject = mail_theme if mail_theme is not None else ''
            if body is not None:
                ml.HTMLBody = body
            if attachment is not None:
                ml.Attachments.Add(attachment)
            ml.Send()
    except Exception as _:
        error_mail(_, 'send outlook mail')


def clear_folder_outlook(folder: str, moveto: str) -> None:
    """
    A feature to clean up a folder in Outlook by moving items to another folder.

    Args:
        folder (str): The name of the folder to be cleared.
        moveto (str): The name of the folder where you want to move the items.

    Returns:
        None
    """
    try:
        app = win32com.client.Dispatch("Outlook.Application")
        f = app.GetNamespace('MAPI').GetDefaultFolder(6).Folders[moveto]
        while str(app.GetNamespace("MAPI").GetDefaultFolder(6).Folders[folder].Items.GetFirst()) != 'None':
            app.GetNamespace("MAPI").GetDefaultFolder(6).Folders[folder].Items.GetFirst().Move(f)
    except Exception as _:
        error_mail(_, 'clear_folder_outlook()')


def delete_mail_outlook() -> None:
    """
    Deletes old emails from specified Outlook folders based on predefined conditions.

    This function connects to Outlook, accesses various folders, and deletes emails
    based on their age and certain conditions specific to each folder. The folders
    include Inbox, Outbox, and several custom folders.

    Note:
    - The folder names 'folder 1', 'folder 2', 'folder 3', 'folder 4', 'folder 5' are specific and may need localization.

    Exception handling is in place to ensure the function continues running even if errors occur while processing individual emails or folders.
    """

    try:
        # Connect to Outlook application
        app = win32com.client.Dispatch("Outlook.Application")

        # Access default folders and custom folders
        inbox = app.GetNamespace("MAPI").GetDefaultFolder(6)
        outbox = app.GetNamespace("MAPI").GetDefaultFolder(5)
        locdel = app.GetNamespace('MAPI').GetDefaultFolder(6).Folders['folder 1']
        arhord = app.GetNamespace('MAPI').GetDefaultFolder(6).Folders['folder 2']
        arh = app.GetNamespace('MAPI').GetDefaultFolder(6).Folders['folder 3']
        err = app.GetNamespace('MAPI').GetDefaultFolder(6).Folders['folder 4']
        post = app.GetNamespace('MAPI').GetDefaultFolder(6).Folders['folder 5']
        arhkd = app.GetNamespace('MAPI').GetDefaultFolder(6).Folders['folder 5'].Folders['folder 5_1']
        delmail = app.GetNamespace('MAPI').GetDefaultFolder(3)

        # Calculate date thresholds
        dnm = (datetime.today() - timedelta(days=30)).strftime('%Y-%m-%d')
        dnw = (datetime.today() - timedelta(days=14)).strftime('%Y-%m-%d')

        # List of folders to process
        folders = [inbox, outbox, locdel, arhord, arh, err, post, arhkd, delmail]

        # Iterate over each folder
        for k, folder in enumerate(folders):
            i = 0
            try:
                while str(folder.Items[i]) != 'None':
                    email = str(folder.Items[i])

                    # Inbox conditions
                    if ('COMObject <unknown>' in email or ' - User unknown' in email) and k == 0:
                        folder.Items[i].Delete()
                        i -= 1
                    else:
                        date_sent = folder.Items[i].SentOn.strftime('%Y-%m-%d')
                        if date_sent < dnm and k == 0:
                            folder.Items[i].Delete()
                            i -= 1

                    # Outbox conditions
                    if 'Error' in email and k == 1:
                        folder.Items[i].Delete()
                        i -= 1
                    elif date_sent < dnm and k == 1:
                        folder.Items[i].Delete()
                        i -= 1

                    # Locdel conditions
                    if k == 2 and date_sent <= dnw:
                        folder.Items[i].Delete()
                        i -= 1
                    elif k == 2 and date_sent > dnm:
                        break

                    # Arhord conditions
                    if k == 3 and date_sent <= dnm:
                        folder.Items[i].Delete()
                        i -= 1
                    elif k == 3 and date_sent > dnm:
                        break

                    # Arh conditions
                    if k == 4 and date_sent <= dnm:
                        folder.Items[i].Delete()
                        i -= 1
                    elif k == 4 and date_sent > dnm:
                        break

                    # Error conditions
                    if k == 5 and date_sent <= dnw:
                        folder.Items[i].Delete()
                        i -= 1
                    elif k == 5 and date_sent > dnw:
                        break

                    # Post conditions
                    if k == 6 and date_sent <= dnw:
                        folder.Items[i].Delete()
                        i -= 1
                    elif k == 6 and date_sent > dnm:
                        break

                    # Arhkd conditions
                    if k == 7 and date_sent <= dnm:
                        folder.Items[i].Delete()
                        i -= 1
                    elif k == 7 and date_sent > dnm:
                        break

                    # Delmail conditions
                    if k == 8 and date_sent <= dnm:
                        folder.Items[i].Delete()
                        i -= 1
                    elif k == 8 and date_sent > dnm:
                        break

                    i += 1
            except Exception as e:
                print(e)

    except Exception as e:
        print(e, 'delete_mail_outlook()')
