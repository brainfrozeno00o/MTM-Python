import os # imported this for file-related logic
import datetime # imported this for date-related logic
import re # imported this for file-related and date-related logic
import win32com.client # to have this library, do "!pip install pywin32" first; imported this for Excel-related logic
import smtplib # imported this for email-sending-related logic
import time # imported this for setting delays after sending an email
from email.mime.multipart import MIMEMultipart # imported this for setting up message of the entire email
from email.mime.base import MIMEBase # imported this for setting up attachments to the message
from email import encoders # imported this for encoding the attachment upon sending

def change_dates(file_list, counterparty):
    # if there are files in that specific folder, then change its names
    if file_list:
        for current_file in file_list:
            # changing the filename for CS counterparty is a special case as the date string is in this format: MMMddyyyy
            if counterparty == "CS":
                new_cs_filename = re.sub(cs_pattern, CS_date_string, current_file)
                os.rename(current_file, new_cs_filename)
            # for changing the date of Nomura, you have to open its encrypted workbook
            elif "Nomura" in counterparty:
                xlApp = win32com.client.Dispatch("Excel.Application")
                xlWb = None
                xlWs = None

                # parameters used in WorkBooks.Open are FileName, UpdateLinks, ReadOnly, Format, and Password
                # for more information, please visit: https://docs.microsoft.com/en-us/office/vba/api/excel.workbooks.open
                if "HK" in counterparty:
                    xlWb = xlApp.Workbooks.Open(os.getcwd() + "\\" + current_file, False, False, None, Password="NOMA1735")
                else:
                    xlWb = xlApp.Workbooks.Open(os.getcwd() + "\\" + current_file, False, False, None, Password="9712034")

                xlWs = xlWb.Sheets(1) # first sheet starts at 1, and date is in the first sheet
                # xlWs.Cells(24, 2) refers to Cell B24 -> where the date is
                xlWs.Cells(24, 2).Value = Nomura_date_string 
                xlWs.Cells(24, 2).NumberFormat = "dd-MMM-yyyy" # Blue Prism would actually extract this as dd/mm/yyyy though... but it should not be a problem
                xlWb.Save()
                xlWb.Close()
                xlApp.Application.Quit()
            # changing the filename for other counterparties is pretty straightforward
            else:
                new_filename = re.sub(general_pattern, general_date_string, current_file)
                os.rename(current_file, new_filename)
    # call the sending email logic after renaming all the files, take note that you have to pass the updated list as you have renamed/updated the files
    send_emails(os.listdir(), folder)

def send_emails(file_list, counterparty):
    # formulating the from email address
    formulate_from_email = ""
    # CA has a ca2 on its email address
    if counterparty == "CA":
        formulate_from_email = "ca2.efg.mtm@gmail.com"
    # special cases for GS, JPM, Nomura, SOCG as they both have SG and HK Valuation Files
    elif "GS" in counterparty:
        formulate_from_email = "gs.efg.mtm@gmail.com" 
    elif "JPM" in counterparty:
        formulate_from_email = "jpm.efg.mtm@gmail.com"
    elif "Nomura" in counterparty:
        formulate_from_email = "nomura.efg.mtm@gmail.com"
    elif "SOCG" in counterparty:
        formulate_from_email = "socg.efg.mtm@gmail.com"
    # pretty straightforward for other counterparties as well
    else:
        formulate_from_email = counterparty.lower() + ".efg.mtm@gmail.com"

    to_email = ["john.lingad@synpulse.com", "louis.delarosa@synpulse.com"] # has to be a list if multiple recipients, always has to be a list from now on...
    password = "synpulse" # this is for all emails set up by Louis

    # setting up messsage here
    message = MIMEMultipart()

    # setting up sender and recipient
    message["From"] = counterparty + "<" + formulate_from_email + ">" # sample would be CA <ca2.efg.mtm@gmail.com>
    message["To"] = ", ".join(to_email) if len(to_email) > 1 else "".join(to_email) # must be a string; did a join if it would be multiple recipients

    # setting up subject
    overall_subject = ""
    # GS has its own subject, thus made logic for formulating subject if counterparty is currently GS
    if "GS" in counterparty:
        overall_subject = "Valuation Report for EFG"
        if "HK" in counterparty:
            overall_subject = overall_subject + " Private Bank SA - Hong Kong - Daily FX on " + GS_subject_date_string
        else:
            overall_subject = overall_subject + " Bank AG (EFG Bank SA) (EFG Bank Ltd) - Daily FX on " + GS_subject_date_string
    # JPM has its own subject, thus made logic for formulating subject if counterparty is currently JPM
    elif "JPM" in counterparty:
        overall_subject = "Valuation Statement for Counterparty 00029432200"
        if "HK" in counterparty:
            overall_subject = overall_subject + "2 Profile 79134 - ASIA"
        else:
            overall_subject = overall_subject + "3 Profile 78830 - ASIA"
    # Nomura has its own subject, thus made logic for formulating subject if counterparty is currently Nomura
    elif "Nomura" in counterparty:
        overall_subject = "Valuation Request-EFG"
        if "HK" in counterparty:
            overall_subject = overall_subject + " PRIVATE BANK SA-HONGKONG - Dynamic"
        else:
            overall_subject = overall_subject + " BANK AG - SINGAPORE BRANCH"
    # SOCG has its own subject, thus made logic for formulating subject if counterparty is currently SOCG
    elif "SOCG" in counterparty:
        overall_subject = "SG - EFG BANK AG : FX Option Valuation Report(s)_" + SOCG_subject_date_string + " ["
        if "HK" in counterparty:
            overall_subject = overall_subject + "100559]"
        else:
            overall_subject = overall_subject + "100560]"
    # just added "Test" after the counterparty code
    else:
        overall_subject = counterparty + " Test"
    message["Subject"] = overall_subject

    # putting one or all attachments; only do this if there are files found in the folder
    # for more details on doing this, please visit: https://www.tutorialspoint.com/send-mail-with-attachment-from-your-gmail-account-using-python
    if file_list:
        for one_file in file_list:
            payload = MIMEBase("application", "octet-stream")
            payload.set_payload(open(one_file, "rb").read())
            encoders.encode_base64(payload)
            payload.add_header("Content-Disposition", f"attachment; filename={one_file}")
            message.attach(payload)

    # sending email via SMTP
    # for more details on doing this, please visit: https://automatetheboringstuff.com/chapter16/
    session = smtplib.SMTP("smtp.gmail.com", 587)
    session.starttls()
    session.login(formulate_from_email, password)
    overall_message = message.as_string()
    session.sendmail(formulate_from_email, to_email, overall_message)
    session.quit()
    print(f"Mail Sent for this counterparty: {counterparty}")
    # UBS is the last folder, thus don't show this in the console if it's the last folder already... can make this generic though
    if counterparty != "UBS":
        time.sleep(15) # 15 second delay
        print(f"Next email...")

day_today = datetime.date.today() # get the current day today
real_business_day = day_today # initialize it temporarily to the date today as a date type

if day_today.weekday() == 0: # 0 is Monday
    real_business_day = real_business_day - datetime.timedelta(days=3)
# this else block actually includes Saturday and Sunday as well
else:
    real_business_day = real_business_day - datetime.timedelta(days=1)

# setting up here date-related strings
CS_date_string = real_business_day.strftime("%b%d%Y") # sample would be Sep182020
Nomura_date_string = day_today.strftime("%d-%b-%Y") # sample would be 19-Sep-2020
general_date_string = real_business_day.strftime("%Y%m%d") # sample would be 20201809
GS_subject_date_string = real_business_day.strftime("%d%b%y") # sample would be 23Sep20
SOCG_subject_date_string = real_business_day.strftime("%Y%d%m") # sample would be 20200923

# reason why I used re.VERBOSE is to put comments in compiling RegEx and it will not consider spaces/tabs/newlines
# regex for this sample date: 20200919
general_pattern = re.compile("""
    2020    # current year, should be changed eventually for next year
    \d      # followed by any digit
    {4}     # limit it to 4 digits for MMdd
""", re.VERBOSE)

# regex for this sample date: Sep192020
# this is a bad RegEx tbh for formulating the short form of a month, please suggest a better RegEx for this one
cs_pattern = re.compile("""
    [JFMASOND]    # first letter of the short form of a month
    [aepuco]      # second letter of the short form of a month
    [nbrylgptvc]  # third letter of the short form of a month
    \d            # followed by any digit
    {6}           # limit it to 6 digits for MMyyyy
""", re.VERBOSE)

# Welcome message here... and then proceeding with inputting the full path of the parent folder containing all the folders of the counterparties
print(f"\nWelcome to the MTM File Renamer and Sending Email Script! Please input the parent folder that contains all the folders of all counterparties. If left blank, the program will proceed to its default folder.\n")
parent_folder = input(f"Input the full file path here: ")

# proceed to the default folder I created if there is no input (basically pressing Enter)
if parent_folder == "":
    print(f"\nNow proceeding with the default folder \"C:\PythonScriptAttachments\"... now proceeding with the main process\n")
    # created a new folder that contains all the folders for all attachments by the counterparties as I did not want to mess with the Sharepoint folder
    os.chdir(f"C:\\PythonScriptAttachments") # change the current working directory to the created folder in C: that has all the counterparty's folders and files
# if there is an input
else:
    # try if it's an absolute path
    try:
        print(f"\nNow proceeding with the inputted folder {parent_folder}... now proceeding with the main process\n")
        os.chdir(parent_folder)
    # show the error and then exit the program
    except Exception as e:
        print(f"\nYou have inputted a wrong file name... full error details below\n")
        print(e)
        exit()

current_working_directory = os.getcwd() # this basically just gets the current working directory
current_folders = os.listdir() # getting all the counterparty folders

for folder in current_folders:
    # find out the importance of "desktop.ini" in this article: https://www.computerhope.com/issues/ch001060.htm
    if folder != "desktop.ini":
        # formulating the path for the counterparty's folder
        new_path = os.getcwd() + "\\" + folder
        # change current working directory to that counterparty's folder
        os.chdir(new_path)
        # gets all the files found in the current counterparty folder
        get_current_files = os.listdir()
        # call the method that changes the filenames of all files in the counterparty folder
        change_dates(get_current_files, folder)
        # change the current working directory to the parent folder
        os.chdir("..")

print(f"\nSent out all emails successfully!")
time.sleep(5)