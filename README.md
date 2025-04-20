# MBOX to PST Converter

**Convert MBOX files to PST format easily with this Python script for Windows.**

This open-source tool simplifies the process of converting MBOX email files to PST format, making it accessible for users with varying levels of technical expertise. Ideal for individuals and organizations needing to migrate or archive emails in a format compatible with Microsoft Outlook.

## Features

- **Effortless Conversion**: Convert MBOX files to PST with a single command.
- **User-Friendly Interface**: Command-line tool with clear, straightforward instructions.
- **Attachment Support**: Preserves email attachments during conversion.
- **Metadata Preservation**: Retains sender, recipient, and other email metadata.
- **Open Source**: Free to use, modify, and distribute.

## Installation

Follow these steps to set up and use the MBOX to PST Converter on your Windows machine:

1. **Set Up Outlook**:
   - Ensure Microsoft Outlook is installed and configured on your device.

2. **Install Python**:
   - Download and install the latest version of Python from the [official Python website](https://www.python.org/downloads/).

3. **Install Dependencies**:
   - Open a terminal and run:
     ```bash
     pip install pywin32
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
   - Move the MBOX file you wish to convert into this directory.

7. **Run the Conversion**:
   - Execute the following command in the terminal:
     ```bash
     python convert.py your_mbox_file.mbox
     ```
   - Replace `your_mbox_file.mbox` with the name of your MBOX file.

8. **Post-Conversion Steps**:
   - After conversion, you can move the generated PST file to a desired location.
   - If you plan to run the script again using the same Outlook account, follow these steps:
     - Open Outlook.
     - Go to `File > Account Settings > Account Settings`.
     - Under the `Data Files` tab, remove the last created PST file.

## Important Notes

- **Original Code Preservation**: The original version of the converter is preserved in this repository as `original_convert.py` for reference and testing purposes. The new, improved version is available in the `convert.py` file.
- **Testing the New Version**: We recommend testing the new version with a sample MBOX file before removing the original code.
- **International Clients**: This tool is designed to be user-friendly for clients worldwide, with clear instructions and support for various email formats.

## Thank You for Using MBOX to PST Converter

We appreciate your trust in our tool. For any questions or support, please open an issue in this repository or contact us directly.

---

*This project is open source and contributions are welcome. Feel free to fork the repository and submit pull requests.*
