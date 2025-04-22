import sys
import mailbox
import os
import email
import win32com.client
import shutil
import tempfile
from email.utils import parseaddr, getaddresses, parsedate_to_datetime
from datetime import datetime

def sanitize_filename(filename):
    """Sanitizes the filename by replacing invalid characters and truncating if necessary."""
    if not filename:
        return 'attachment.bin'
    invalid_chars = r'\/:*?"<>|'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    filename = filename.strip().lstrip('.')
    if not filename:
        filename = 'attachment'
    max_length = 255
    if len(filename) > max_length:
        name, ext = os.path.splitext(filename)
        name = name[:max_length - len(ext)] if ext else name[:max_length]
        filename = name + ext
    return filename

def extract_emails_from_mbox(mbox_file):
    """Extracts emails from an mbox file, retaining all metadata and folder labels."""
    mbox = mailbox.mbox(mbox_file)
    emails = []
    for message in mbox:
        email_data = {
            'raw': message.as_string(),
            'labels': message.get('X-Gmail-Labels', '').split(',') if message.get('X-Gmail-Labels') else ['Inbox']
        }
        emails.append(email_data)
    return emails

def get_folder_by_name(parent_folder, folder_name):
    """Finds or creates a subfolder by name within the parent folder."""
    for folder in parent_folder.Folders:
        if folder.Name == folder_name:
            return folder
    return parent_folder.Folders.Add(folder_name)

def check_outlook_accessible():
    """Checks if Outlook is running and accessible."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        namespace.Folders  # Test access
        return True
    except Exception as e:
        return False

def import_emails_to_outlook(emails, pst_file):
    """Imports emails into an Outlook PST file with progress tracking and folder structure."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        print(f"Error: Failed to initialize Outlook. Ensure Outlook is installed and running. Details: {str(e)[:100]}")
        sys.exit(1)

    try:
        namespace = outlook.GetNamespace("MAPI")
        namespace.AddStoreEx(pst_file, 3)  # 3 = olStoreUnicode
        pst_folder = namespace.Folders.GetLast()
    except Exception as e:
        print(f"Error: Failed to add PST store. Check permissions and disk access. Details: {str(e)[:100]}")
        sys.exit(1)

    total_emails = len(emails)
    processed_emails = 0
    total_attachments = 0

    print(f"\n{'='*50}")
    print(f"Starting import of {total_emails} emails")
    print(f"{'='*50}\n")

    try:
        for idx, email_data in enumerate(emails, 1):
            try:
                raw_email = email_data['raw']
                labels = email_data['labels']
                msg = email.message_from_string(raw_email)

                # Determine folder (use first label, default to Inbox)
                folder_name = labels[0].strip() if labels and labels[0].strip() else 'Inbox'
                target_folder = get_folder_by_name(pst_folder, folder_name)

                mail_item = outlook.CreateItem(0)  # 0 = olMailItem
                mail_item.Subject = msg.get('Subject', '')

                # Set To, CC, BCC
                for header in ['To', 'CC', 'BCC']:
                    value = msg.get(header, '')
                    if value:
                        addresses = getaddresses([value])
                        formatted = '; '.join([email.utils.formataddr((name, addr)) for name, addr in addresses])
                        setattr(mail_item, header, formatted)

                # Set From using PropertyAccessor
                from_header = msg.get('From')
                if from_header:
                    name, addr = parseaddr(from_header)
                    pa = mail_item.PropertyAccessor
                    if name:
                        pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1A001E", name)
                    if addr:
                        pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1F001E", addr)

                # Set Date using PropertyAccessor
                date_header = msg.get('Date')
                if date_header:
                    try:
                        parsed_date = parsedate_to_datetime(date_header)
                        pa = mail_item.PropertyAccessor
                        pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E060040", parsed_date)
                        pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x00390040", parsed_date)
                    except Exception as e:
                        print(f"\nError setting date for email {idx}: {e}")

                # Process body and attachments
                text_body, html_body = None, None
                email_attachments = 0
                temp_files = []
                try:
                    for part in msg.walk():
                        if part.is_multipart():
                            continue
                        content_type = part.get_content_type()
                        disposition = str(part.get('Content-Disposition', '')).lower()
                        filename = part.get_filename()

                        # Handle attachments
                        if filename or 'attachment' in disposition:
                            data = part.get_payload(decode=True)
                            if not data:
                                continue
                            sanitized = sanitize_filename(filename) if filename else 'attachment.bin'
                            try:
                                with tempfile.NamedTemporaryFile(delete=False) as temp_file:
                                    temp_file.write(data)
                                    temp_path = temp_file.name
                                    temp_files.append(temp_path)
                                mail_item.Attachments.Add(temp_path, 1, 0, sanitized)
                                email_attachments += 1
                                total_attachments += 1
                            except Exception as e:
                                print(f"\nFailed to attach {sanitized} in email {idx}: {e}")
                        else:
                            # Handle body content
                            data = part.get_payload(decode=True)
                            charset = part.get_content_charset('utf-8')
                            try:
                                decoded = data.decode(charset, errors='ignore') if data else ''
                            except (LookupError, UnicodeDecodeError):
                                decoded = data.decode('utf-8', errors='ignore') if data else ''
                            if content_type == 'text/plain':
                                text_body = decoded
                            elif content_type == 'text/html':
                                html_body = decoded
                finally:
                    # Clean up temporary files
                    for temp_path in temp_files:
                        try:
                            if os.path.exists(temp_path):
                                os.unlink(temp_path)
                        except Exception as e:
                            print(f"\nFailed to clean up temporary file {temp_path}: {e}")

                if text_body:
                    mail_item.Body = text_body
                if html_body:
                    mail_item.HTMLBody = html_body

                # Update progress display
                percent_complete = (idx / total_emails) * 100
                sys.stdout.write(
                    f"\rProcessing: [{('#' * int(percent_complete//2)).ljust(50)}] "
                    f"{percent_complete:.1f}% ({idx}/{total_emails}) | "
                    f"Attachments: {total_attachments}"
                )
                sys.stdout.flush()

                mail_item.Save()
                mail_item.Move(target_folder)
                processed_emails += 1

            except Exception as e:
                print(f"\nError processing email {idx}: {str(e)[:100]}")

        # Final progress update
        sys.stdout.write(
            f"\rProcessing: [{'#'*50}] 100.0% ({total_emails}/{total_emails}) | "
            f"Attachments: {total_attachments}"
        )
        sys.stdout.flush()

    finally:
        print(f"\n\n{'='*50}")
        print(f"Processed: {processed_emails}/{total_emails} emails")
        print(f"Attachments saved: {total_attachments}")
        print(f"PST file location: {os.path.abspath(pst_file)}")
        print(f"{'='*50}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python convert.py <mbox_file>")
        sys.exit(1)

    mbox_file = sys.argv[1]
    if not os.path.exists(mbox_file):
        print(f"Error: File '{mbox_file}' does not exist.")
        sys.exit(1)

    pst_file = os.path.join(os.path.dirname(mbox_file), 'emails.pst')

    # Check if Outlook is accessible
    if not check_outlook_accessible():
        print("Error: Outlook is not running or accessible. Please start Outlook and try again.")
        sys.exit(1)

    # Extract emails and estimate size
    emails = extract_emails_from_mbox(mbox_file)
    total_size = 0
    for email_data in emails:
        raw_email = email_data['raw']
        msg = email.message_from_string(raw_email)
        total_size += len(raw_email)
        for part in msg.walk():
            if 'attachment' in str(part.get('Content-Disposition', '')).lower() or part.get_filename():
                data = part.get_payload(decode=True)
                if data:
                    total_size += len(data)
    total_size = int(total_size * 1.5)  # Adjust multiplier for overhead

    pst_dir = os.path.dirname(pst_file)
    try:
        disk_usage = shutil.disk_usage(pst_dir)
    except FileNotFoundError:
        print(f"Directory {pst_dir} does not exist.")
        sys.exit(1)

    # Check PST size limit (50GB default for modern Outlook)
    pst_size_limit = 50 * 1024 * 1024 * 1024  # 50GB in bytes
    if total_size > pst_size_limit:
        print(f"Warning: Estimated PST size ({total_size/1024/1024/1024:.1f} GB) exceeds Outlook's default limit (50 GB). Import may fail.")
        sys.exit(1)

    print(f"\n{'='*50}")
    print(f"Found {len(emails)} emails in {os.path.basename(mbox_file)}")
    print(f"Estimated PST size: {total_size/1024/1024:.1f} MB")
    print(f"Available space: {disk_usage.free/1024/1024:.1f} MB")
    print(f"{'='*50}")

    if disk_usage.free < total_size:
        print(f"Error: Insufficient disk space. Required: {total_size/1024/1024:.1f} MB, Available: {disk_usage.free/1024/1024:.1f} MB")
        sys.exit(1)

    import_emails_to_outlook(emails, pst_file)
