# MBOX to PST Converter

**Effortlessly convert MBOX files to PST format with this robust Python script for Windows.**

This open-source tool streamlines the conversion of MBOX email files to PST format, compatible with Microsoft Outlook. Designed for both individual users and organizations, it offers advanced features to ensure reliable email migration and archiving, trusted by clients worldwide.

## Features

- **Seamless Conversion**: Convert MBOX files to PST with a single command.
- **Metadata Preservation**: Retains sender, recipient, dates, and other email metadata.
- **Attachment Support**: Preserves email attachments with sanitized filenames.
- **Dry-Run Mode**: Preview conversion details without modifying data.
- **Customizable Options**: Specify PST file path, folder name, Outlook profile, and temporary directory.
- **Robust Error Handling**: Continues processing despite individual email failures, with detailed logging.
- **Disk Space Check**: Ensures sufficient disk space before conversion.
- **Flexible Encoding**: Handles various email encodings with fallback support.
- **Open Source**: Free to use, modify, and distribute.

## Installation

Follow these steps to set up and use the MBOX to PST Converter on your Windows machine:

1. **Set Up Outlook**:

   - Ensure Microsoft Outlook is installed and configured on your device.

2. **Install Python**:

   - Download and install the latest version of Python from the official Python website.

3. **Install Dependencies**:

   - Open a terminal and install required packages:

     ```bash
     pip install pywin32 charset-normalizer python-dateutil
     ```

4. **Clone the Repository**:

   - Clone this repository to your local machine:

     ```bash
     git clone https://github.com/BloodBlinker/MBOX_To_PST.git
     ```

5. **Navigate to the Project Directory**:

   - Change into the project directory:

     ```bash
     cd MBOX_To_PST
     ```

6. **Prepare Your MBOX File**:

   - Place the MBOX file you wish to convert in a convenient location.

7. **Run the Conversion**:

   - Execute the conversion with the following command:

     ```bash
     python convert.py your_mbox_file.mbox
     ```

   - Replace `your_mbox_file.mbox` with the path to your MBOX file.

   - Optional arguments:

     - `--pst <path>`: Specify the output PST file path (default: `emails.pst` in MBOX directory).
     - `--folder <name>`: Set the target folder in the PST (default: `Inbox`).
     - `--dry-run`: Preview actions without importing.
     - `--profile <name>`: Use a specific Outlook profile.
     - `--temp-dir <path>`: Specify a custom temporary directory.
     - `--verbose`: Enable detailed debug logging.
     - `--quiet`: Suppress non-error logs.
     - `--fallback-encoding <encoding>`: Set fallback encoding (default: `utf-8`).
     - `--space-multiplier <number>`: Adjust disk space requirement (default: 10).

   Example with options:

   ```bash
   python convert.py sample.mbox --pst output.pst --folder "Archive" --dry-run --verbose
   ```

8. **Post-Conversion Steps**:

   - After conversion, move the generated PST file to your desired location.
   - If running the script again with the same Outlook account:
     - Open Outlook.
     - Go to `File > Account Settings > Account Settings`.
     - Under the `Data Files` tab, remove the previously created PST file to avoid conflicts.

## Important Notes

- **Original Code Preservation**: The original version of the converter is preserved as `original_convert.py` for reference and testing. The latest, enhanced version is in `convert.py`.
- **Testing the New Version**: We recommend testing `convert.py` with a sample MBOX file before removing the original code.
- **International Clients**: This tool supports diverse email formats and encodings, ensuring compatibility for global users.
- **Logging**: Detailed logs are generated to help troubleshoot issues, especially useful in verbose mode (`--verbose`).

## Thank You for Using MBOX to PST Converter

We value your trust in our tool. For questions, support, or contributions, please open an issue in this repository or contact us directly.

---

*This project is open source. Feel free to fork the repository and submit pull requests to enhance its functionality.*
