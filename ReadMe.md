Prerequisites

|IMPORTANT| Make Sure your internet signal and speed is stable. 

In order to run the python script, your system must have the following programs/packages installed.

Chrome : https://www.google.com/chrome/

Python : https://www.python.org/downloads/

Add Python to PATH : 

Open windows - Type "Edit the system environtment variables" - click environtment variables - Double Click on PATH on user variables - NEW - Paste this C:\Users\akbar\AppData\Local\Programs\Python\Python311\Scripts

or you can see this following link : https://www.javatpoint.com/how-to-set-python-path

Install Library

Open CMD - type this following text

pip install selenium

pip install webdriver_manager

pip install pandas

pip install loging

pip install xlrd

pip install openpyxl


Approach

Place The Data on Recipients data.xlsx

you dont have to put +62 (Indonesian) on Recipients data. If not, you can edit and put your country code on script.py

Run python script script.py using py script.py in the terminal

The script opens WhatsApp web using chrome.

You needs to scan QR code from phone.

Enter in command prompt to execute further.

The script hit url with contact number and message from excel sheet.

Once all the message will be sent chrome driver will automatically closed.

The output result of the program will export to result.xlsx

Legal

This code is in no way affiliated with, authorized, maintained, sponsored or endorsed by WhatsApp or any of its affiliates or subsidiaries. This is an independent and unofficial software. Use at your own risk. Commercial use of this code/repo is strictly prohibited.
