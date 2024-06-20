import os
import os.path
import win32com.client
from datetime import date, timedelta, datetime

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
                    date_sent = folder.Items[i].SentOn.strftime('%Y-%m-%d')

                    # Inbox conditions
                    if ('COMObject <unknown>' in email or ' - User unknown' in email) and k == 0:
                        folder.Items[i].Delete()
                        i -= 1
                    elif date_sent < dnm and k == 0:
                        folder.Items[i].Delete()
                        i -= 1

                    # Outbox conditions
                    elif 'Error' in email and k == 1:
                        folder.Items[i].Delete()
                        i -= 1
                    elif date_sent < dnm and k == 1:
                        folder.Items[i].Delete()
                        i -= 1

                    # Locdel conditions
                    elif k == 2 and date_sent <= dnw:
                        folder.Items[i].Delete()
                        i -= 1
                    elif k == 2 and date_sent > dnm:
                        break

                    # Arhord conditions
                    elif k == 3 and date_sent <= dnm:
                        folder.Items[i].Delete()
                        i -= 1
                    elif k == 3 and date_sent > dnm:
                        break

                    # Arh conditions
                    elif k == 4 and date_sent <= dnm:
                        folder.Items[i].Delete()
                        i -= 1
                    elif k == 4 and date_sent > dnm:
                        break

                    # Error conditions
                    elif k == 5 and date_sent <= dnw:
                        folder.Items[i].Delete()
                        i -= 1
                    elif k == 5 and date_sent > dnw:
                        break

                    # Post conditions
                    elif k == 6 and date_sent <= dnw:
                        folder.Items[i].Delete()
                        i -= 1
                    elif k == 6 and date_sent > dnm:
                        break

                    # Arhkd conditions
                    elif k == 7 and date_sent <= dnm:
                        folder.Items[i].Delete()
                        i -= 1
                    elif k == 7 and date_sent > dnm:
                        break

                    # Delmail conditions
                    elif k == 8 and date_sent <= dnm:
                        folder.Items[i].Delete()
                        i -= 1
                    elif k == 8 and date_sent > dnm:
                        break

                    i += 1
            except Exception as e:
                print(e)

    except Exception as e:
        print(e, 'delete_mail_outlook()')
