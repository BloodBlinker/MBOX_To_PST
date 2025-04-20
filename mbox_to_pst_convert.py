import sys
import mailbox
import os
import email
import win32com.client
import tempfile
import logging
import re

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def sanitize_filename(filename):
    """Sanitize the filename by replacing invalid characters with underscores."""
    return re.sub(r'[\/:*?"<>|]', '_', filename)

def extract_emails_from_mbox(mbox_file):
    """Generator function to yield email message objects from an MBOX file one at a time."""
    try:
        mbox = mailbox.mbox(mbox_file)
        for message in mbox:
            yield message
    except Exception as e:
        logging.error(f"Error reading MBOX file '{mbox_file}': {e}")
        raise

def get_folder_by_name(parent_folder, folder_name):
    """Return a folder by name within a parent folder, or None if not found."""
    for folder in parent_folder.Folders:
        if folder.Name == folder_name:
            return folder
    return None

def import_emails_to_outlook(emails_iter, pst_file, folder_name="Inbox"):
    """Import emails from an iterable of message objects into a PST file under the specified folder."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except Exception as e:
        logging.error(f"Failed to initialize Outlook: {e}")
        raise

    logging.info(f"Starting import to {pst_file}, folder: {folder_name}")

    # Add the PST store
    namespace.AddStoreEx(pst_file, 3)  # 3 corresponds to olStoreUnicode
    pst_folder = namespace.Folders.GetLast()

    # Get or create the target folder
    target_folder = get_folder_by_name(pst_folder, folder_name)
    if target_folder is None:
        target_folder = pst_folder.Folders.Add(folder_name)

    # Process each email
    for i, msg in enumerate(emails_iter, 1):
        logging.info(f"Processing email {i}")
        try:
            mail_item = outlook.CreateItem(0)  # 0 is olMailItem
            mail_item.Subject = msg['subject'] if msg['subject'] else "No Subject"

            # Set email headers to preserve metadata like sender and recipient
            headers = '\r\n'.join(f"{key}: {value}" for key, value in msg.items())
            mail_item.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E", headers)

            # Handle email body
            html_body = None
            plain_body = None
            for part in msg.walk():
                if part.get_content_type() == 'text/html' and not html_body:
                    charset = part.get_content_charset() or 'utf-8'
                    html_body = part.get_payload(decode=True).decode(charset, errors='ignore')
                elif part.get_content_type() == 'text/plain' and not plain_body:
                    charset = part.get_content_charset() or 'utf-8'
                    plain_body = part.get_payload(decode=True).decode(charset, errors='ignore')

            if html_body:
                mail_item.HTMLBody = html_body
            elif plain_body:
                mail_item.Body = plain_body
            else:
                mail_item.Body = "No body content"

            # Handle attachments
            attachments = [part for part in msg.walk() if part.get('Content-Disposition') and 'attachment' in part.get('Content-Disposition')]
            for attachment in attachments:
                filename = attachment.get_filename()
                if filename:
                    filename = sanitize_filename(filename)
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(filename)[1]) as temp_file:
                        temp_file.write(attachment.get_payload(decode=True))
                        temp_path = temp_file.name
                    try:
                        mail_item.Attachments.Add(temp_path)
                    finally:
                        try:
                            os.remove(temp_path)
                        except Exception as e:
                            logging.warning(f"Failed to delete temp file {temp_path}: {e}")

            mail_item.Save()
            mail_item.Move(target_folder)
        except Exception as e:
            logging.error(f"Error importing email {i}: {e}")
            continue

if __name__ == "__main__":
    # Parse command-line arguments
    if len(sys.argv) < 2 or len(sys.argv) > 4:
        print("Usage: python convert.py <mbox_file> [pst_file] [folder_name]")
        sys.exit(1)

    mbox_file = sys.argv[1]

    # Determine PST file path
    if len(sys.argv) > 2:
        pst_file = sys.argv[2]
    else:
        pst_file = os.path.join(os.path.dirname(mbox_file), 'emails.pst')

    # Determine folder name
    if len(sys.argv) > 3:
        folder_name = sys.argv[3]
    else:
        folder_name = "Inbox"

    # Check if MBOX file exists
    if not os.path.exists(mbox_file):
        print(f"Error: File '{mbox_file}' does not exist.")
        sys.exit(1)

    # Perform the conversion
    try:
        emails_iter = extract_emails_from_mbox(mbox_file)
        import_emails_to_outlook(emails_iter, pst_file, folder_name)
        logging.info("Conversion completed successfully.")
    except Exception as e:
        logging.error(f"Conversion failed: {e}")
        sys.exit(1)
