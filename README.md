# Features
To be added soon.

# Understanding the Code
Some users may be a little nervous about what is happening, especially after being asked to enable Macros and (for MacOS users) getting a bunch of file & folder permission pop-ups and a password request. To help assuage any worries, I have uploaded the source code here for anyone to see.

   1. ThisWorkbook.vba - This is the code that takes all the data you entered and generates unique PDFs for each student.
   2. Worksheet.vba - This is the code that automatically coverts numbers and lowercase letters into the expected letter grades
   3. SpeakingEvals.applescript - This is the code that extends what Excel can do on MacOS.
   4. Dialog_Toolkit.zip - This is the script for MacOS that allows the nicer dialog windows. It is a bit complex, but you are welcome to read it. Note that it isn't my code,
      but the original author has released it as freeware.

You do not need to download any of these files. The first two are embedded in the Excel file, and SpeakingEvals.applescript is the human-readable version of SpeakingEvals.scpt, which has been compiled into a format that MacOS natively understands. These files are just here for your reference and curiousity. Additional comments & notes will gradually be added to the code to make the various steps and processes easier to understand. The goal is for you to feel safe about running the code, and if you are interested, hopefully learn a little about programming scripts for MS Office and/or MacOS.

# SpeakingEvals.scpt (MacOS ONLY)
While it is technically optional, it is STRONGLY RECOMMENDED that you install it. NO SUPPORT is offered for using the Excel file without it. This file enables many important functions and improvements to make generating the reports more efficient and much more pleasant for users. It solves many bugs with using VBA (MS Office macros) on MacOS and addresses some quirks related to Apple's overly strict security policies. (I'm all for security, but some of their decisions make life difficult for developers.) See below for how to easily install this file.

## DOWNLOADING & INSTALLATION
### Automated Method (Recommend)
Open a Terminal window. You can press `Command(⌘) + Spacebar` to open Spotlight search and type "Terminal". Then, you can press return to open it. If you prefer to use Finder, you can press `Command(⌘) + Shift + U` to open the Utilities folder. Double-click on Terminal to open it. With a Terminal window open, copy and paste the following command and press return. It is one long command, so copy the whole thing, even if it appears as two or more lines.

`curl -L -o ~/Library/Application\ Scripts/com.microsoft.Excel/SpeakingEvals.scpt https://github.com/papercutter0324/SpeakingEvals/raw/main/SpeakingEvals.scpt`

### Manual Method
You first need to manually download SpeakingEvals.scpt to your Downloads folder. There are four ways to easily do this.
   1. [Click here](https://github.com/papercutter0324/SpeakingEvals/raw/main/SpeakingEvals.scpt) and save to your Downloads folder.
   2. Right-click on SpeakingEvals.scpt above, select 'Save As', and save it to your Downloads folder.
   3. Click on SpeakingEvals.scpt, which will take you to a new page. On the right, you should see a small button labeled 'Raw'. Click on it and save the file to your Downloads folder.
   4. On the right on this page is the Releases section. Click on either "Releases" or the most recent release (you should see a green tag and a 'Latest' label next to it). On the new page,
      click on and download "Source code (zip)". You may need to click on 'Assets' to see it. Once downloaded, open the zip file and extract SpeakingEvals.scpt to your Downloads folder.

Next, open a Terminal window (using one of the methods mentioned in the "Automated Method" section) and run the following command.

`mv ~/Downloads/SpeakingEvals.scpt ~/Library/Application\ Scripts/com.microsoft.Excel`

As you can see, this is not nearly as simple of the automated method, so it is highly recommended to use it instead.
