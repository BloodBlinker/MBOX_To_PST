import sys
import mailbox
import os
import email
import win32com.client
from win32com.client import constants
import tempfile
import logging
import re
import argparse
import shutil
import pywintypes
import charset_normalizer
from email.header import decode_header
from datetime import datetime
from email.utils import getaddresses, parsedate_to_datetime
from dateutil import parser as date_parser
import uuid
import traceback

# Configure logging only once
if not logging.getLogger().handlers:
    logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s')

def setup_logging(verbose=False, quiet=False):
    if quiet:
        logging.getLogger().setLevel(logging.ERROR)
    elif verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    else:
        logging.getLogger().setLevel(logging.INFO)

def sanitize_filename(filename, max_length=100):
    """Sanitize a filename by replacing invalid characters and truncating to a safe length."""
    clean = re.sub(r'[\\/:*?"<>|]', '_', filename)
    if len(clean) > max_length:
        clean = clean[:max_length]
    return clean

def decode_mime_header(value):
    """Decode an RFC-2047 encoded header into a readable string."""
    if not value:
        return ""
    parts = decode_header(value)
    decoded = []
    for text, charset in parts:
        if isinstance(text, bytes):
            try:
                decoded.append(text.decode(charset or 'utf-8', errors='ignore'))
            except (LookupError, TypeError):
                decoded.append(text.decode('latin-1', errors='ignore'))
        else:
            decoded.append(text)
    return ''.join(decoded)

def extract_emails_from_mbox(mbox_file):
    """Yield email messages from an MBOX file one at a time."""
    try:
        mbox = mailbox.mbox(mbox_file)
    except (OSError, mailbox.NoSuchMailbox, mailbox.Error) as e:
        logging.error(f"Failed to open MBOX file '{mbox_file}': {e}")
        raise
    for message in mbox:
        yield message

def find_or_add_pst(namespace, pst_path):
    """Locate an existing PST or add a new one in Outlook, handling locked files."""
    pst_name = os.path.basename(pst_path)
    for store in namespace.Stores:
        if store.FilePath and os.path.basename(store.FilePath).lower() == pst_name.lower():
            logging.info(f"Using existing PST: {pst_path}")
            return store.GetRootFolder()
    logging.info(f"Creating new PST: {pst_path}")
    try:
        namespace.AddStoreEx(pst_path, constants.olStoreUnicode)
    except pywintypes.com_error as e:
        logging.error(f"Failed to add PST (possibly locked or in use): {e}")
        logging.info("Please ensure the PST file is not open in Outlook or another application.")
        raise
    for store in namespace.Stores:
        if store.FilePath and os.path.basename(store.FilePath).lower() == pst_name.lower():
            return store.GetRootFolder()
    raise RuntimeError(f"Could not mount PST: {pst_path}")

def get_folder_by_name(parent_folder, folder_name):
    """Find or create a folder in the PST by name, with sanitized name."""
    name = sanitize_filename(folder_name)
    for folder in parent_folder.Folders:
        if folder.Name.lower() == name.lower():
            return folder
    return parent_folder.Folders.Add(name)

def extract_bodies(msg, fallback_encoding='utf-8'):
    """Extract all HTML and plain text parts from an email message, including nested parts."""
    html_bodies = []
    plain_bodies = []
    for part in msg.walk():
        ctype = part.get_content_type()
        if ctype == 'text/html':
            raw = part.get_payload(decode=True) or b''
            text = decode_text(raw, fallback_encoding)
            if text:
                html_bodies.append(text)
        elif ctype == 'text/plain':
            raw = part.get_payload(decode=True) or b''
            text = decode_text(raw, fallback_encoding)
            if text:
                plain_bodies.append(text)
    return html_bodies, plain_bodies

def decode_text(raw, fallback_encoding):
    """Decode raw bytes to text using charset_normalizer or fallback encoding."""
    detected = charset_normalizer.from_bytes(raw).best()
    if detected:
        try:
            return detected.output().decode('utf-8', errors='ignore')
        except Exception:
            pass
    try:
        return raw.decode(fallback_encoding, errors='ignore')
    except (LookupError, TypeError):
        return raw.decode('latin-1', errors='ignore')

def format_address_list(msg, header_name):
    """Convert email addresses from a header into a comma-separated string."""
    addrs = getaddresses(msg.get_all(header_name, []))
    return ', '.join(addr for _, addr in addrs)

def set_headers(mail, msg):
    """Set email headers like subject, sender, and recipients in Outlook."""
    mail.Subject = decode_mime_header(msg.get('Subject')) or "No Subject"
    mail.SenderEmailAddress = decode_mime_header(msg.get('From'))
    mail.To = format_address_list(msg, 'To')
    mail.CC = format_address_list(msg, 'Cc')
    date_hdr = msg.get('Date')
    if date_hdr:
        try:
            mail.SentOn = parsedate_to_datetime(date_hdr)
        except Exception:
            try:
                mail.SentOn = date_parser.parse(date_hdr)
            except Exception:
                logging.warning(f"Could not parse date '{date_hdr}'. Using current time.")
                mail.SentOn = datetime.now()
    else:
        mail.SentOn = datetime.now()

def add_body(mail, msg, fallback_encoding='utf-8'):
    """Add the email body to the Outlook item, preferring HTML if available."""
    html_bodies, plain_bodies = extract_bodies(msg, fallback_encoding)
    if html_bodies:
        mail.HTMLBody = '\n'.join(html_bodies)
        return 'HTML'
    if plain_bodies:
        mail.Body = '\n'.join(plain_bodies)
        return 'Plain'
    mail.Body = "No body content"
    return 'None'

def add_attachments(mail, msg, temp_dir):
    """Add attachments to the email, keeping temp files until email is saved."""
    temp_files = []
    counter = 1
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        filename = part.get_filename()
        if not filename:
            filename = f"attachment_{counter}.dat"
            counter += 1
        else:
            filename = decode_mime_header(filename)
        clean = sanitize_filename(filename)
        unique_id = uuid.uuid4().hex[:4]  # Shortened to 4 chars
        clean = f"{clean}_{unique_id}"
        path = os.path.join(temp_dir, clean)
        try:
            data = part.get_payload(decode=True)
            with open(path, 'wb') as f:
                f.write(data)
            mail.Attachments.Add(path)
            temp_files.append(path)
        except OSError as e:
            logging.error(f"Failed to handle attachment '{clean}': {e}")
    return temp_files

def import_emails_to_outlook(emails_iter, pst_file, folder_name, dry_run, temp_dir, profile=None, fallback_encoding='utf-8'):
    """Import MBOX emails into an Outlook PST file."""
    total = 0
    failures = 0
    outlook = None
    namespace = None
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        if profile:
            namespace.Logon(profile)
        root = find_or_add_pst(namespace, pst_file)
        target_folder = get_folder_by_name(root, folder_name)
        logging.info(f"Importing to folder: {target_folder.Name} in PST: {pst_file}")

        for idx, msg in enumerate(emails_iter, start=1):
            total += 1
            try:
                if dry_run:
                    attachments = [p.get_filename() for p in msg.walk() if p.get_filename()]
                    body_type = 'HTML' if any(p.get_content_type() == 'text/html' for p in msg.walk()) else 'Plain'
                    logging.info(f"[Dry-run #{idx}] Subject: {decode_mime_header(msg.get('Subject'))}, Body: {body_type}, Attachments: {len(attachments)}, Target Folder: {target_folder.Name}")
                    continue

                if idx % 10 == 0:
                    logging.info(f"Processed {idx} messages...")

                temp_files = []
                mail = outlook.CreateItem(0)  # Create a new email item
                set_headers(mail, msg)
                add_body(mail, msg, fallback_encoding)
                temp_files = add_attachments(mail, msg, temp_dir)
                mail.Save()
                mail.Move(target_folder)
                for temp_file in temp_files:
                    try:
                        os.remove(temp_file)
                    except OSError:
                        pass
            except Exception as e:
                failures += 1
                logging.error(f"Email #{idx} failed: {e}")
                logging.debug(traceback.format_exc())

        logging.info(f"Completed: {total - failures}/{total} emails imported, {failures} failed.")
    except KeyboardInterrupt:
        logging.error("Import cancelled by user.")
        sys.exit(1)
    finally:
        try:
            if namespace:
                namespace.Logoff()
            if outlook:
                outlook.Quit()
        except Exception as e:
            logging.error(f"Failed to cleanly shut down Outlook: {e}")

def check_disk_space(path, mbox_size, multiplier=10):
    """Ensure sufficient disk space is available for the conversion."""
    total, used, free = shutil.disk_usage(path)
    required = mbox_size * multiplier
    if free < required:
        logging.error(f"Insufficient disk space at {path}. Need: {required // (1024*1024)}MB, Free: {free // (1024*1024)}MB.")
        sys.exit(1)

def main():
    """Parse arguments and run the MBOX to PST conversion."""
    parser = argparse.ArgumentParser(description="Convert an MBOX file to an Outlook PST.")
    parser.add_argument("mbox", help="Path to the .mbox file")
    parser.add_argument("--pst", "-p", help="Path to output PST file (default: <mbox_dir>/emails.pst)")
    parser.add_argument("--folder", "-f", default="Inbox", help="Folder name in PST (default: Inbox)")
    parser.add_argument("--dry-run", action="store_true", help="Preview actions without importing")
    parser.add_argument("--profile", help="Outlook profile name (optional)")
    parser.add_argument("--temp-dir", help="Custom temporary directory (optional)")
    parser.add_argument("--fallback-encoding", default="utf-8", help="Encoding if detection fails (default: utf-8)")
    parser.add_argument("--verbose", "-v", action="store_true", help="Enable debug logging")
    parser.add_argument("--quiet", "-q", action="store_true", help="Suppress non-error logs")
    parser.add_argument("--space-multiplier", type=int, default=10, help="Disk space multiplier (default: 10)")
    args = parser.parse_args()

    setup_logging(args.verbose, args.quiet)

    if not os.path.isfile(args.mbox):
        logging.error(f"MBOX file not found: {args.mbox}")
        sys.exit(1)

    pst = args.pst or os.path.join(os.path.dirname(args.mbox), "emails.pst")
    mbox_size = os.path.getsize(args.mbox)
    check_disk_space(os.path.dirname(pst), mbox_size, args.space_multiplier)

    temp_dir = args.temp_dir or tempfile.mkdtemp(prefix="mbox2pst_")
    try:
        emails = extract_emails_from_mbox(args.mbox)
        import_emails_to_outlook(emails, pst, args.folder, args.dry_run, temp_dir, args.profile, args.fallback_encoding)
    except Exception as e:
        logging.error(f"Conversion failed: {e}")
        sys.exit(1)
    finally:
        if not args.temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)

if __name__ == "__main__":
    main()
