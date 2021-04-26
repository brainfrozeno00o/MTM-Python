import os # imported for file-related logic
import ntpath # imported for file-related logic
import win32com.client # to have this library, do "!pip install pywin32" first; imported this for Excel-related logic
import getpass # imported for password-related logic
import time # imported for delay

# Please input the absolute file path of the encrypted file itself and not its directory!
print(f"\nWelcome to MTM Password Changer! Please input the absolute path of the encrypted file. If no input is put, then the program will exit.\n")
encrypted_file = input(f"Please input the absolute path of the encrypted file here: ")

# no input = nothing to rename, thus program will close
if encrypted_file == "":
    print(f"\nYou have provided no input... thus the program will not continue.")
    exit()

# currently masking the passwords here... you will not see what you have inputted so be careful
old_password = getpass.getpass(f"\nPlease input the old password of your encrypted file: ")
new_password = getpass.getpass(f"\nPlease input the new password of your encrypted file: ")

# initialize the Excel app in Python
xlApp = win32com.client.Dispatch("Excel.Application")
xlWb = None

try:                
    # parameters used in WorkBooks.Open are FileName, UpdateLinks, ReadOnly, Format, and Password
    # for more information, please visit: https://docs.microsoft.com/en-us/office/vba/api/excel.workbooks.open
    xlWb = xlApp.Workbooks.Open(encrypted_file, False, False, None, Password=old_password)
    print(f"\nSuccessfully opened file... now changing password")
except Exception as e:
    xlApp.Application.Quit()
    print("\nFailed to open the encrypted file due to the error below... exiting the program now.\n")
    print(e)
    exit()

# changing the password using Workbook.Password property
# to learn more, visit: https://docs.microsoft.com/en-us/office/vba/api/excel.workbook.password
xlWb.Password = new_password
xlWb.Save()
xlWb.Close()
xlApp.Application.Quit()
print(f"\nSuccessfully changed the password!")
time.sleep(5)