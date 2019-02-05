# Accessing-and-collecting-data-from-excel-files-in-different-folders

To get this to run on Windows simply: Download the whole folder. SHIFT + RIGHT CLICK inside the main containing the .py file, PRESS 'open PowerShell window here', TYPE & ENTER: python results_extractor.py




The purpose of the script is to the following:

--> CanEye softeware was used to generate excel data for multiple image sets, each excel file was placed in a directory (time) and many subdirectories (person, and other junk ones).

--> The Post-Doc researcher I was working with told me to go through the files one by one to extract 9 key data points and put them all in one excel file with pages for each time.

--> The 9 data points are: 4 FAPAR values, 3 PAI - ALA values, the directory name (time), the subdirectory name (person).
    * Note that directory and subdirectory names are not from the excel file itsself. *

--> There were over 70 files so this work would be extremely tedious to do by hand.

--> To overcome this I decided to automate the processes.

--> This script also had to be run on OTHER people's CANEYE data so it had to be generalizable to over 500 different excel files from different people.

--> This script had to be given to a person with no computer science training, so it had to be simple for them to run.

--> Lots of 'print' statements are included so non-computer science users can know whats happening just from reading the terminal.

--> For the sake of keeping our lab data confidential, I will make excel files with fake values.

