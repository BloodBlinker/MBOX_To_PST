#  MBOX to PST Converter

**Reliable. Efficient. Ready for enterprise use.**

This powerful open-source Python tool enables seamless conversion of MBOX email files to PST format—fully compatible with Microsoft Outlook. Designed with robustness and flexibility in mind, it's ideal for individuals, system administrators, and organizations migrating or archiving mailboxes.

##  Key Features

-  **Full Metadata Extraction** – Retains subject, sender, recipients, and timestamps  
-  **Preserves Folder Hierarchies** – Converts Gmail label structures to Outlook folders  
-  **Attachment Support** – Handles embedded files with smart filename sanitization  
-  **Progress Indicator** – Displays real-time conversion progress  
-  **Disk & PST Size Checks** – Prevents conversion issues due to space limits  
-  **Resilient Outlook Integration** – Built-in error handling for stability  
-  **Auto Cleanup** – Deletes temporary files post-conversion  
-  **Dual Email Format Support** – Maintains both HTML and plain text bodies  

##  Installation & Setup (Windows)

### 1. Prerequisites

- **Microsoft Outlook** must be installed and configured and when opening the Outlook run it as administrator.  
- **Python**: Download and install from [python.org](https://www.python.org/downloads/)

### 2. Install Required Python Packages

```bash
pip install pywin32 charset-normalizer python-dateutil
```

### 3. Clone the Repository

```bash
git clone https://github.com/BloodBlinker/MBOX_To_PST.git
cd MBOX_To_PST
```

### 4. Prepare for Conversion

- Place your `.mbox` file in an easily accessible folder

### 5. Run the Converter

```bash
python convert.py path/to/your_file.mbox
```

> Replace `path/to/your_file.mbox` with the actual file path.

### 6. After Conversion

- Move the generated `.pst` file to your preferred location  
- If converting multiple times using the same Outlook profile:
  - Open Outlook
  - Navigate to `File > Account Settings > Account Settings`
  - Under the **Data Files** tab, remove the previous PST to avoid duplication or conflict

##  Development Notes

- **Legacy Code**: The original script (`original_convert.py`) is preserved for historical reference  
- **Enhanced Version**: Use `mbox_to_pst.py` for improved performance and error handling  
- **Test First**: Always test with a sample file before running on production data

##  Contributing & Support

We welcome feedback, bug reports, and contributions.  
-  Open an [issue](https://github.com/BloodBlinker/MBOX_To_PST/issues) for help or feature requests  
-  Fork the project and submit pull requests for improvements

---

###  Thank You

Thank you for choosing **MBOX to PST Converter**. We’re honored to be part of your email migration journey.

---

> **License**: This project is open-source and licensed under [MIT](LICENSE).  
> Build with confidence. Fork freely. Contribute proudly.
