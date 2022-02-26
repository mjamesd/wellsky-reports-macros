# WellSky Resource Manager Reports Macros
Excel VBA macros for making WellSky Resource Management reports useable.

![Licensed under GNU GPLv3](https://img.shields.io/badge/license-GNU%20GPLv3-brightgreen)

## How to Install 
### Importing a Macro Module
1. First you will have to create your personal macro workbook. 

   1.1. Open Excel. Click on the “Tell me what you want to do” field in the top-middle of the app window (next to the “Help” tab on the ribbon). Type “rec” and select “Record macro” from the list of options. 

   1.2. Change “Store macro in” to “Personal Macro Workbook”. Click OK. You are now recording all actions you take in Excel as a macro.

   1.3. Press the Enter key once then click the square in the bottom left corner of the screen to “stop” the recording. This creates your personal macro workbook and saves the recording you just made as a macro in it. The personal macro workbook will open every time you open Excel, so you will always have access to these macros. 

2. Next you will import the macro(s) you wish to use. 

   2.1. First download the macro .bas file. 

   2.2. In Excel, press the Alt+F11 keys. This will open the “Microsoft Visual Basic for Applications” window. On the left, it lists some “VBAProject” files and folders. Open the project for “PERSONAL.XLSB”, right-click on the “Modules” folder, and choose “Import File…” 

   2.3. Navigate to the macro file you want, select it, and click “Open”. This will import the module into Excel.

   2.4. Close the “Microsoft Visual Basic for Applications” window and exit Excel completely to ensure it saved the imported macro. You will be asked whether you want to save the changes you made to the Personal Macro Workbook. Choose “Save All”.

## Usage 
To use a macro, press the Alt+F8 keys to open the list of macros you have in your personal macro workbook and any macros stored in the file(s) you currently have open. Select the macro you wish to run and click Run. Please be aware that it may take a minute for the macro to complete its work, during which time you will not be able to use Excel. Typically the larger the file, the longer it will take to run the macro. Please also be aware that you cannot “undo” the actions of a macro, so make sure to save your work prior to running it – in case the macro causes an error or causes Excel to crash (i.e., if the file is too long).

This macro file contains two main macros and one unfinished and commented-out function.
1. WellSky_Create_Participant_Information_Report
   - This report will combine information from the WellSky "Patient Birthdays" and "Patient List" reports to make a master Patient report with these columns:
     - Patient Name (raw) <-- **this is their full name unformatted from WellSky if needed for matching in other reports** *(this column is hidden)*
     - MR Number
     - first name
     - last name
     - Date of Birth
     - Age (calculated from `TODAY()`)
     - Sex At Birth
     - Primary phone number
     - Address 1
     - Address 2
     - City
     - State
     - ZIP Code
   - Run the "Patient Birthdays" report for birthdays from Jan 1 to Dec 31. This will get all patients. Export as "Excel SpreadSheet (XLS)".
   - Run the "Patient List" report -- not the "Active Patient List" or "New Patient List", just the plain old "Patient List" report. It is in Patient Reports > Patient List Reports. Do not select "List New Patients Only". The Sort Order doesn't matter but I usually choose "mr number". Export as "Excel SpreadSheet (XLS)".
   - Open both reports. Drag and drop Sheet 1 from the Patients List report to *after* Sheet 1 on the Patient Birthdays report. This will cause the Patient List report window to close.
   - Press the Alt+F8 keys to open the macro list window. Choose "PERSONAL.XLSB!WellSky_Create_Participant_Information_Report". The macro will run and will show message boxes with what it is doing.
   - Once it is complete, you will have a report formatted in an Excel table. If patients' addresses were missing or malformed those fields for those patients will say "ERROR:" and then an appropriate error message.
   - The report does **not** de-duplicate the raw data so if you have a patient entered multiple times, the report will show all instances of the patient. Use [Excel's conditional formatting](https://www.ablebits.com/office-addins-blog/2016/03/09/how-to-highlight-duplicates-excel/) to find these duplicates.
2. WellSky_CleanReport_AppointmentActivity
   - This report will clean up the Appointment Activity report.
   - Run the report in WellSky, open it in Excel, and run the macro. No special instructions.
   - If an appointment has multiple procedure codes, the macro will copy the patient's info onto new rows with the procedure codes.
   - The macro shows message boxes wiht what it is doing (and is pretty snarky).

## Contributing
DO NOT EDIT THE CODE. This will most likely cause the macro to not work, or worse, to work incorrectly. There are some comments in the code (the green text) to explain what it is doing, but please do not edit it unless you know what you’re doing.

## Support
Contact Mark Drummond (mjamesd@gmail.com) for support.

## License
This work is licensed under GNU GPLv3.
