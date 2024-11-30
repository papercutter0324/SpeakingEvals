# Understanding What Is Happening
Some users may be a little nervous about what is happening, especially after being asked to enable Macros and (for MacOS users) getting a bunch of file & folder permission pop-ups. For those who are curious, I have uploaded the source code here.

ThisWorkbook.vba - This is the code that takes all the data you entered and generates unique PDFs for each student.
Worksheet.vba - This is the code that automatically coverts numbers and lowercase letters into the expected letter grades
SpeakingEvals.applescript - (Will be uploaded soon.) This is the code that extends what Excel can do on MacOS.

Note that you do not need these files; the first two are embedded in the Excel file, and SpeakingEvals.applescript is converted into a format that MacOS will natively understand. These files are just here for your reference and curiousity.

# SpeakingEvals.scpt
This is an optional file that will extend what Excel can do when generating the reports. It will give various options such as loading 'Speaking Evaluation Template.docx' from a different location, saving the generated PDFs to a different folder, and other features to make the experience smoother for MacOS users. Again, it is completely optional.
If you wish to install the file, it needs to be installed in a directory that is usually hidden from most users. The good news, though, is that it is very easy to move the file into the correct directory.

## DOWNLOADING & INSTALLATION
### Automated Method (Recommend)
Open a Terminal window. You can press `Command(⌘) + Spacebar` to open Spotlight search and type "Terminal". Then, you can press return to open it. If you prefer to use Finder, you can press `Command(⌘) + Shift + U` to open the Utilities folder. Double-click on Terminal to open it. With a Terminal window open, copy and paste the following command and press return. It is one long command, so copy the whole thing, even if it appears as two or more lines.

`curl -L -o ~/Library/Application\ Scripts/com.microsoft.Excel/SpeakingEvals.scpt https://github.com/papercutter0324/SpeakingEvals/raw/main/SpeakingEvals.scpt`

### Manual Method
You need to manually download AngryBirds.scpt to your Downloads folder. There are four ways to easily do this.
   1. [Click here](https://github.com/papercutter0324/SpeakingEvals/raw/main/SpeakingEvals.scpt) and save to your Downloads folder.
   2. Right-click on SpeakingEvals.scpt above, select 'Save As', and save it to your Downloads folder.
   3. Click on SpeakingEvals.scpt, which will take you to a new page. On the right, you should see a small button labeled 'Raw'. Click on it and save the file to your Downloads folder.
   4. On the right on this page is the Releases section. Click on either "Releases" or the most recent release (you should see a green tag and a 'Latest' label next to it). On the new page, click on and download "Source code (zip)". You mau need to click on 'Assets' to see it. Once downloaded, open the zip file and extract AngryBirds.scpt to your Downloads folder.

Next, open a Terminal window (using one of the methods mentioned in the "Automated Method" section) and run the following command.

`mv ~/Downloads/SpeakingEvals.scpt ~/Library/Application\ Scripts/com.microsoft.Excel`
