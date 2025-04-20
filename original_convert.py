import sys
import mailbox
import os
import email
import win32com.client
import shutil

def extract_emails_from_mbox(mbox_file):
    mbox = mailbox.mbox(mbox_file)
    emails = []
    for message in mbox:
        emails.append(message.as_string())
    return emails

def get_folder_by_name(parent_folder, folder_name):
    for folder in parent_folder.Folders:
        if folder.Name == folder_name:
            return folder
    return None

def import_emails_to_outlook(emails, pst_file):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    namespace.AddStoreEx(pst_file, 3)
    pst_folder = namespace.Folders.GetLast()

    inbox_folder = get_folder_by_name(pst_folder, "Inbox")
    if inbox_folder is None:
        inbox_folder = pst_folder.Folders.Add("Inbox")

    for raw_email in emails:
        msg = email.message_from_string(raw_email)

        mail_item = outlook.CreateItem(0)
        mail_item.Subject = msg['subject']
        mail_item.BodyFormat = 1

        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                payload = part.get_payload(decode=True)
                if content_type == 'text/plain':
                    mail_item.Body = payload.decode('utf-8', errors='ignore')
                elif content_type == 'text/html':
                    mail_item.HTMLBody = payload.decode('utf-8', errors='ignore')
        else:
            content_type = msg.get_content_type()
            payload = msg.get_payload(decode=True)
            if content_type == 'text/plain':
                mail_item.Body = payload.decode('utf-8', errors='ignore')
            elif content_type == 'text/html':
                mail_item.HTMLBody = payload.decode('utf-8', errors='ignore')

        mail_item.Save()
        mail_item.Move(inbox_folder)

if __name__ == "__main__":
    # Check if the correct number of arguments are provided
    if len(sys.argv) != 2:
        print("Usage: python convert.py <mbox_file>")
        sys.exit(1)

    # Get the mbox_file path from the command-line argument
    mbox_file = sys.argv[1]

    # Check if the mbox_file exists
    if not os.path.exists(mbox_file):
        print(f"Error: File '{mbox_file}' does not exist.")
        sys.exit(1)

    # Determine the directory of mbox_file to create pst_file in the same directory
    pst_file = os.path.join(os.path.dirname(mbox_file), 'emails.pst')

    # Extract emails from the mbox file
    emails = extract_emails_from_mbox(mbox_file)
    # Import emails into the PST
    import_emails_to_outlook(emails, pst_file)
