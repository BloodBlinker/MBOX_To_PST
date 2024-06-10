# MBOX_To_PST
This is a python script that can be executed in Windows to convert your MBOX files to PST files.


## Features

- Simple Conversion: Easily convert MBOX files to PST format.
- User-Friendly: Command-line interface with straightforward commands.
- Open Source: Free to use and modify.

## Installation

  To run timer on your device, follow these steps:

  1. set-up and open Outlook. 

  2. Ensure you have the latest Python installed. If not, refer to the official Python
     documentation or installation instructions.

  3. Install pywin32 using the following command in the terminal :

    pip install pywin32
    
  4. Clone this repository to your local machine using the following command:

   ```bash
   https://github.com/BloodBlinker/MBOX_To_PST.git
   ```

  5. Change into the project directory:
  
    cd MBOX_To_PST
    
  6. Move the mbox file which you want to convert in to this directory

  7. Run the following command in the terminal to bgin the conversion :

    python MBOX_to_PST_Converter.py mbox_filename.mbox
  
  

After successful completion,copy/move the pst output to any other place, and if you want to run the script again using the same Outlook account, follow these steps: Go to Outlook, navigate to `File > Account Settings > Account Settings`, and then under the `Data Files` tab, remove the last created PST file.


### Thank you for Using
