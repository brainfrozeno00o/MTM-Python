import os # imported for file-related logic
import ntpath # imported for file-related logic
import re # imported for date-related logic
import datetime # imported for date-related logic
import time # imported for delay

# Please input the absolute file path of the Pivot Price file itself and not its directory!
print(f"\nWelcome to MTM Pivot Price File Renamer! Please input the absolute path containing the Pivot Price file. If no input is put, then the program will exit.\n")
pivot_price_file = input(f"Please input the absolute path of the pivot price file here: ")

# no input = nothing to rename, thus program will close
if pivot_price_file == "":
    print(f"\nYou have provided no input... thus the program will not continue.")
    exit()

directory = ""
file_name = ""

# not sure if the try-except block is needed anymore...
try:
    directory, file_name = ntpath.split(pivot_price_file)
except Exception as e:
    print("\nYou have provided a wrong file path.. error will be show below and the program will exit.\n")
    print(e)
    exit()

# pivot price files have the "PIVOT price update" in their names...
if "PIVOT price update" not in file_name:
    print("\nThis is not a Pivot Price file! The program will now exit...")
    exit()
else:
    print("\nProceeding to renaming...")
    # formulating the RegEx pattern
    date_pattern = re.compile("""
        \d            # digit character
        {2}           # limit to only 2 for the day
        \s            # match a space character
        [JFMASOND]    # first letter of the short form of a month
        [aepuco]      # second letter of the short form of a month
        [nbrylgptvc]  # third letter of the short form of a month
        \s            # match a space character
        \d            # digit character
        {4}           # limit to only 4 for the year
    """, re.VERBOSE)

    # formulating the closure of business day format
    day_today = datetime.date.today() # get the current day today
    real_business_day = day_today # initialize it temporarily to the date today as a date type

    if day_today.weekday() == 0: # 0 is Monday
        real_business_day = real_business_day - datetime.timedelta(days=3)
    # this else block actually includes Saturday and Sunday as well
    else:
        real_business_day = real_business_day - datetime.timedelta(days=1)

    date_string = real_business_day.strftime("%d %b %Y") # sample would be 19 Sep 2020

    os.chdir(directory) # change working directory to the directory of the file

    new_filename = re.sub(date_pattern, date_string, file_name) # formulate the new file name of the file

    os.rename(file_name, new_filename) # rename the file

    print("\nPivot Price File has been renamed successfully!")
    time.sleep(5)