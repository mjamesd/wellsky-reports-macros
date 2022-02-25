Attribute VB_Name = "WellSkyMacros"
Function formatWellSkyDate(dateToFormat As String, cellToPlace As Range)
    ' Use this formula to format cell [@Date] to, like, an actual date format
    ' =TEXT(CONCAT(TRIM(LEFT(dateToFormat,FIND("/",dateToFormat)-1)),"/",TRIM(MID(dateToFormat,FIND("/",dateToFormat)+1,2)),"/",TRIM(RIGHT(dateToFormat,4))),"mm/dd/yyyy")
    ' btw, WellSky sucks
End Function

Sub WellSky_Create_Participant_Information_Report()
'
' WellSky_Create_Participant_Information_Report Macro
'

    ' Declare variables
    Dim dateOfReport, newFileName, newFullFileName, tempFname, tempLname, Address1, Address2, CSZ, City, State, Zipcode As String
    Dim currentRow, lastRowNum As Long
    Const lastColNum As Long = 5 ' column #5 = column E -- we know the report only has 5 columns.
    Const shtName As String = "Participant Info Report"
    Const shtLinkName As String = "linkSheet"
    Const tblName As String = "tblParticipantInformation"
    
    If Application.Sheets.Count <> 2 Then
        MsgBox "You must move the sheet from the WellSky Patient List report to this workbook to continue.", , "Warning: Second sheet not detected"
        MsgBox "EXITING WITHOUT DOING ANYTHING.", , "Warning: Second sheet not detected"
        Exit Sub
    Else
        Sheets(1).Name = shtName
        Sheets(2).Name = shtLinkName
    End If
    
    ' ==============================================================================================================================
    MsgBox "The file will be renamed to """ & shtName & "_RunOn_" & """ and then the date the report was run.", , shtName & ": Rename File"
    MsgBox "First we will format the participants' information:" & vbNewLine _
        & "1. First and last names" & vbNewLine _
        & "2. Calculate their age based on the date the report is viewed" & vbNewLine _
        & "3. Add the Participant IDs (Ptt ID)" & vbNewLine _
        & "4. Add sexes at birth.", , shtName & ": Step 1"
    ' ==============================================================================================================================
    
    
    ' First we select the Participant Report sheet and get the date on which the report was run, and add it to the new file name.
    Sheets(shtName).Select
    
    ' Get the date that the report was run from cell U12 and set `dateOfReport`
    Range("V12").Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[-1],FIND(""/"",RC[-1])-1)"
    Range("W12").Select
    ActiveCell.FormulaR1C1 = _
        "=SUBSTITUTE(MID(RC[-2],FIND(""/"",RC[-2])+1,2),""/"","""")"
    Range("X12").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[-3],4)"
    Range("Y12").Select
    ActiveCell.FormulaR1C1 = _
        "=TEXT(CONCAT(RC[-1],""-"",TRIM(RC[-3]),""-"",TRIM(RC[-2])),""yyyy-mm-dd"")"
    dateOfReport = Selection
    
    ' Set the new file name, saveAs the file, and open it back up
    newFileName = shtName & "_RunOn_" & dateOfReport & ".xlsx"
    newFullFileName = Left(ActiveWorkbook.FullName, Len(ActiveWorkbook.Name) + 1) & newFileName ' <-- <-- <-- <-- <-- THIS IS WRONG
    ActiveWorkbook.SaveAs fileName:=newFullFileName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks.Open fileName:=newFullFileName
    'Workbooks(newFileName).Activate

    ' Next we format the "link sheet" so we can use it later
    Sheets(shtLinkName).Select
    ' Clean up report
    ' Remove the WellSky header, unmerge cells, delete extra columns and rows
    Rows("1:11").Select
    Selection.Delete Shift:=xlUp
    Range("A1:U1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.UnMerge
    ' We don't need to clean it up any more than than this because we are just going to look _
        up values then delete the entire sheet

    ' Clean up report
    ' Select the main sheet again
    Sheets(shtName).Select
    ' Remove the WellSky header, unmerge cells, delete extra columns and rows
    Rows("1:13").Select
    Selection.Delete Shift:=xlUp
    Range("K1").Value = "Primary phone"
    Range("L1").Value = ""
    Range("A1:E1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.UnMerge
    Cells(1, 1).Select
    Selection.Value = "Participant Name (Raw)"
    Rows("1:1").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireColumn.Delete
    Columns("E:E").Select ' Use column E in case a patient's address takes more than one line.
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    
    ' Make report into an Excel table so we can filter, etc.
    Range("E2").Select ' use column E because there are more rows here than column A.
    With Application.ActiveSheet
        lastRowNum = .Cells(.Rows.Count, "E").End(xlUp).Row
    End With
    Range(Cells(1, 1), Cells(lastRowNum, lastColNum)).Select
    'Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = tblName
    ActiveSheet.ListObjects(tblName).TableStyle = "TableStyleMedium15"
    Selection.Columns.AutoFit
    Range("A1").Select

    ' Import and/or format Participant (Ptt) ID, First and Last Names, and Sex At Birth
    ' We want to keep the unformatted name in the report, though, so we can match to other WS reports that _
        don 't have the participant's ID
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1").Value = "Ptt ID"
    Range("C1").Value = "First Name"
    Range("D1").Value = "Last Name"
    ' Formula for Ptt ID: Index/Match the Participant ID from linkSheet (column A)
    Range("B2").Select
    ActiveCell.Formula = _
        "=INDEX(linkSheet!A:A,MATCH([@[Participant Name (Raw)]],linkSheet!C:C,0))"
    Range("B2:B" & lastRowNum).Select
    ' Copy column B and **PASTE VALUES** into same position, overwriting the formulas with the formula results
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ' Formula for first name
    Range("C2").Select
    ActiveCell.Formula = _
        "=LET(" & Chr(10) & "   fnameTEST1, MID([@[Participant Name (Raw)]],FIND("","",[@[Participant Name (Raw)]])+2, 99)," & Chr(10) & _
        "   fnameTEST2,LEFT(fnameTEST1, FIND("" "",fnameTEST1)-1)," & Chr(10) & "   IF(" & Chr(10) & _
        "      NOT(ISERROR(fnameTEST2))," & Chr(10) & "      fnameTEST2," & Chr(10) & "      fnameTEST1" & _
        Chr(10) & "   )" & Chr(10) & ")"
    ' Formula for last name
    Range("D2").Select
    ActiveCell.Formula = "=LEFT([@[Participant Name (Raw)]], FIND("","",[@[Participant Name (Raw)]])-1)"
    ' Copy columns C & D and **PASTE VALUES** into same position, overwriting the formulas with the formula results
    Range("C2:D" & lastRowNum).Select
    Selection.Copy
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    ' Delete cell values (not entire rows or columns) in columns D with the value "#VALUE!"
    ' We do this here and not on other columns because we are going to use column D/#4/Last Name to check whether the row _
        contains a name or Address2/CSZ info
    ActiveSheet.ListObjects(tblName).Range.AutoFilter Field:=4, Criteria1:="#VALUE!" ' Field:=4 means it is filtering on column D/#4/Last Name
    Range("D2:D" & lastRowNum).Select
    Selection.ClearContents ' will only delete data matching the filter criteria
    ActiveSheet.ListObjects(tblName).Range.AutoFilter Field:=4 ' remove filter
    ' Formula for Sex At Birth: Index/Match the participant's sex from linkSheet (column H)
    Columns("G:G").Select
    ' Add column between "Age" and "Primary Phone", rename it "Sex At Birth"
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("G1").Select
    ActiveCell.Value = "Sex At Birth"
    Range("G2").Select
    ActiveCell.Formula = _
        "=INDEX(linkSheet!H:H,MATCH([@[Participant Name (Raw)]],linkSheet!C:C,0))"
    Range("G2:G" & lastRowNum).Select
    ' Copy column G and **PASTE VALUES** into same position, overwriting the formulas with the formula results
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ' Formula for Age: Calculate the Age column based on DoB and today's date
    Range("F2:F" & lastRowNum).Formula = "=ROUNDDOWN((TODAY()-[@[Date of Birth]])/365.25,0)"
    
    ' Done importing info from linkSheet, now delete it. Suppress deletion confimration alert, then _
        re-enable them after deletion
    Sheets(shtLinkName).Select
    Application.DisplayAlerts = False
    ActiveWindow.SelectedSheets.Delete
    Application.DisplayAlerts = True
    'Sheets(shtName).Select ' not needed?
    
    ' Keep column A so we can match it in other reports, but hide it
    Columns("A:A").Select
    Selection.EntireColumn.Hidden = True
    Range("B1").Select
    
    ' ==============================================================================================================================
    Dim seconds, minutes As Integer
    seconds = Round(lastRowNum / 32) ' macro runs at 32 lines per second
    minutes = (seconds / 60) - 1
    seconds = (seconds Mod 60) - 1
    MsgBox "Next we will format the participants' names, sexes at birth, and addresses." & vbNewLine & vbNewLine _
        & "***NOTE***" & vbNewLine & "This may take awhile depending on the number of participants." & vbNewLine & vbNewLine _
        & "This report has " & lastRowNum & " lines so it will take approximately " & minutes _
        & " minutes and " & seconds & " seconds.", , shtName & ": Step 2"
    ' ==============================================================================================================================
    
    ' Add columns for address transformation
    Range("I1").Select
    Selection.ListObject.ListColumns.Add
    Selection.ListObject.ListColumns.Add
    Selection.ListObject.ListColumns.Add
    Selection.ListObject.ListColumns.Add
    Range("I1").Value = "Address 1"
    Range("J1").Value = "Address 2"
    Range("K1").Value = "City"
    Range("L1").Value = "State"
    Range("M1").Value = "ZIP Code"
    
    ' Loop through all rows and perform transform on names and addresses:
    '   First and Last names
    '       Remove any text including and after the first instance of an asterisk ("*")
    '       Remove any text including and after the first instance of an open parentheses ("(")
    '       Remove any double quotation marks (there are 4 characters representing these _
                quotation marks: character #s 34, 132, 147, and 148)
    '      Trim whitespace from beginning and end
    '   Move the Address2, City, State, ZIP to their respective columns
    '       CSZ = City/State/Zipcode
    
    ' Select column D in the first row with raw address data
    Range("D2").Select
    Do While ActiveCell.Row <= lastRowNum
        ' Do While loop will continue until it reaches the last row
        ' TRY TO KEEP `ActiveCell` IN COLUMN D (#4) and on the participant's last name, only move the ActiveCell pointer at the end of the loop

        ' ============================
        ' TRANSFORM FIRST & LAST NAMES
        ' ============================
        
        ' Transform first name (offset one column to the left)
        tempFname = ActiveCell.Offset(0, -1).Value
        ' Detect and remove non-name info
        If InStr(1, tempFname, "*", vbTextCompare) Then
            tempFname = Left(tempFname, InStr(1, tempFname, "*", vbTextCompare) - 1)
        End If
        If InStr(1, tempFname, "(", vbTextCompare) Then
            tempFname = Left(tempFname, InStr(1, tempFname, "(", vbTextCompare) - 1)
        End If
        ' Detect and remove double quotes & after
        If InStr(tempFname, Chr(34)) > 0 Then
            tempFname = Replace(tempFname, Chr(34), "")
        End If
        If InStr(tempFname, Chr(132)) > 0 Then
            tempFname = Replace(tempFname, Chr(132), "")
        End If
        If InStr(tempFname, Chr(147)) > 0 Then
            tempFname = Replace(tempFname, Chr(147), "")
        End If
        If InStr(tempFname, Chr(148)) > 0 Then
            tempFname = Replace(tempFname, Chr(148), "")
        End If
        ' Trim whitespace from beginning and end
        ActiveCell.Offset(0, -1).Value = Trim(tempFname)
        
        ' Now transform last name (current ActiveCell)
        tempLname = ActiveCell.Value
        ' Detect and remove non-name info
        If InStr(1, tempLname, "*", vbTextCompare) Then
            tempLname = Left(tempLname, InStr(1, tempLname, "*", vbTextCompare) - 1)
        End If
        If InStr(1, tempLname, "(", vbTextCompare) Then
            tempLname = Left(tempLname, InStr(1, tempLname, "(", vbTextCompare) - 1)
        End If
        ' Detect and remove double quotes & after
        If InStr(tempLname, Chr(34)) > 0 Then
            tempLname = Replace(tempLname, Chr(34), "")
        End If
        If InStr(tempLname, Chr(132)) > 0 Then
            tempLname = Replace(tempLname, Chr(132), "")
        End If
        If InStr(tempLname, Chr(147)) > 0 Then
            tempLname = Replace(tempLname, Chr(147), "")
        End If
        If InStr(tempLname, Chr(148)) > 0 Then
            tempLname = Replace(tempLname, Chr(148), "")
        End If
        ' Trim whitespace from beginning and end
        ActiveCell.Value = Trim(tempLname)

        ' ============================
        ' TRANSFORM SEX AT BIRTH
        ' ============================
        If ActiveCell.Offset(0, 3).Text = "#N/A" Then
            ActiveCell.Offset(0, 3).Value = "Unknown"
        ElseIf ActiveCell.Offset(0, 3) = "" Then
            ActiveCell.Offset(0, 3).Value = "Unknown"
        ElseIf ActiveCell.Offset(0, 3) = "M" Then
            ActiveCell.Offset(0, 3).Value = "Male"
        ElseIf ActiveCell.Offset(0, 3) = "F" Then
            ActiveCell.Offset(0, 3).Value = "Female"
        Else
            ActiveCell.Offset(0, 3).Value = "Other"
        End If
        
        
        ' ===================
        ' TRANSFORM ADDRESSES
        ' ===================

        ' No transform needed for Address1 data in column I (5 columns offset from column D)
        Address1 = ActiveCell.Offset(0, 5).Value
        
        ' There might be an Address2, i.e., an apartment number; let's look at the next row down ("Offset(1)")
        ' If column D on this row is blank (i.e., not another Last Name), that means there is address data in column I (5 columns offset from column D)
        If ActiveCell.Offset(1).Value = "" Then
            If InStr(1, ActiveCell.Offset(1, 5).Value, ",") Then
                ' If there is a comma character in this cell, then it probably isn't an apartment number; rather, it _
                    is the CSZ
                Address2 = ""
                CSZ = ActiveCell.Offset(1, 5).Value
            Else
                ' Otherwise, check if this is a data entry error
                ' If city/state does not match zipcode in WellSky, it only returns zipcode -- which _
                    doesn't have a comma either!
                ' Check if column D of the next line down starts with a First Name
                ' If not, then there is Address2 and CSZ
                If ActiveCell.Offset(2).Value = "" Then
                    Address2 = Trim(ActiveCell.Offset(1, 5).Value)
                    CSZ = Trim(ActiveCell.Offset(2, 5).Value)
                Else
                    ' Otherwise, there is an error in the data and we need to manually correct it; in the meantime, however, _
                        we set CSZ as that incorrect data and tell the user (below) that it needs attention.
                    CSZ = Trim(ActiveCell.Offset(1, 5).Value)
                    Address2 = ""
                End If
            End If
            ' Now we break apart city, state, and zip code
            If Len(CSZ) > 0 Then
                If InStr(1, CSZ, ",") Then
                    City = Trim(Left(CSZ, InStr(1, CSZ, ",") - 1))
                    State = Trim(Mid(CSZ, InStr(1, CSZ, ",") + 2, 2))
                    Zipcode = Trim(Mid(CSZ, Len(City) + Len(State) + 3))
                Else
                    ' If there is *not* a comma in the CSZ, this is the signal from several lines above that _
                        there is an error in the data
                    City = CSZ
                    State = "ERROR: City/State/Zip not formatted correctly."
                    Zipcode = "ERROR:  City/State/Zip not formatted correctly."
                End If
            Else
                ' Otherwise, there is no data in CSZ -- this would be a data error from WellSky and should be reported as such.
                City = CSZ
                State = "ERROR: City/State/Zip not formatted correctly. It's blank."
                Zipcode = "ERROR: City/State/Zip not formatted correctly. It's blank."
            End If
        Else
            Address2 = ""
            City = "ERROR: No second line of address data."
            State = "ERROR: No second line of address data."
            Zipcode = "ERROR: No second line of address data."
        End If
        ' Now we enter Address1, Address2, City, State, and Zipcode in the correct cells
        ActiveCell.Offset(0, 5).Value = Address1    ' column I
        ActiveCell.Offset(0, 6).Value = Address2    ' column J
        ActiveCell.Offset(0, 7).Value = City        ' column K
        ActiveCell.Offset(0, 8).Value = State       ' column L
        ActiveCell.Offset(0, 9).Value = Zipcode     ' column M
        
        ' Now select the next non-blank row in column D/#4/"Last Name" to continue the loop
        ActiveCell.Offset(1).Select
        ' If this cell in column D does not contain a value (i.e., a Last Name), delete it since _
            we already have moved the data out of that row
        Do While ActiveCell.Value = "" And ActiveCell.Row <= lastRowNum
            ActiveCell.EntireRow.Delete
            lastRowNum = lastRowNum - 1
            ' ActiveCell.Offset(1).Select
        Loop
        ' Make sure we are starting the next loop in the Last Name column (column D/#4) of this non-blank cell's row
        Cells(ActiveCell.Row, 4).Select
    Loop
    ' go to top
    Range("B1").Select
    
    ' ==============================================================================================================================
    MsgBox "Lastly, we will format the report and add our own header.", , shtName & ": Step 3"
    ' ==============================================================================================================================
    
    ' Next, format the table...
    Range("A1:M" & lastRowNum).Select
    Range("A1").Activate
    With Selection.Font
        .Name = "Calibri"
        .Size = 12
    End With
    ' ...and autofit the row heights and column widths
    Range("B1:M" & lastRowNum).Select
    For Each c In Selection.Columns
        c.ColumnWidth = 255
    Next c
    Selection.Columns.AutoFit ' address columns will be wide because of the error messages
    Selection.Rows.AutoFit
    
    ' Finally, put our own header in columns B-D because we hid column A... and we're done!
    Rows("1:1").Select
    Range("B1").Activate
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1").Select
    ActiveCell.Value = "Report Run on:"
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
    End With
    Range("C1").Select
    ActiveCell.Value = dateOfReport
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
    End With
    Range("D1").Select
    ActiveCell.Value = "Participant Information Report"
    With Selection
        .Font.Bold = True
        .Font.Size = 20
        .EntireRow.AutoFit
    End With
    
    'Save the workbook
    ActiveWorkbook.Save
    
    ' The original report is just so awful.
    ' Did they /intentionally/ do it so you can't get your own data out of their system?
    ' Or are they just idiots?
    MsgBox "The process is complete. Filter the City, State, or ZIP Code columns to find any errors " & _
        "(will say ""Error!"" with an error message). You will have to manually solve these errors.", , _
        shtName & ": WellSky Report Formatted for, Like, _Actual_ Use"

End Sub


Sub WellSky_CleanReport_AppointmentActivity()
'
' WellSky_AppointmentActivity Macro
'
    ' Declare variables
    Const tblName As String = "tblAppointmentActivity"
    Const shtName As String = "Appointment Activity Report"
    Dim dateOfReport, newFileName, newFullFileName As String
    Dim currentRow As Long
    Dim lastRowNum As Long
    Dim lastRowNum_preDeDup As Long
    
    ' Get the date that the report was run from cel Q12. Set the table name to _
        "AppointmentActivity_RunOn_{date}" to use later
    Range("R12").Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[-1],FIND(""/"",RC[-1])-1)"
    Range("S12").Select
    ActiveCell.FormulaR1C1 = "=SUBSTITUTE(MID(RC[-2],FIND(""/"",RC[-2])+1,2),""/"","""")"
    Range("T12").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[-3],4)"
    Range("U12").Select
    ActiveCell.FormulaR1C1 = "=TEXT(CONCAT(RC[-1],""-"",TRIM(RC[-3]),""-"",TRIM(RC[-2])),""yyyy-mm-dd"")"
    dateOfReport = Selection
    
    ' ==============================================================================================================================
    MsgBox "The file will be renamed to """ & shtName & "_RunOn_" & dateOfReport & _
        ".xlsx"" and then the date the report was run.", , shtName & ": Rename File"
    ' ==============================================================================================================================
    
    ' Set the new file name, saveAs the file, and open it back up
    Sheets(1).Name = shtName
    newFileName = shtName & "_RunOn_" & dateOfReport & ".xlsx"
    newFullFileName = ActiveWorkbook.Path & "\" & newFileName
    ActiveWorkbook.SaveAs fileName:=newFullFileName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Workbooks.Open fileName:=newFullFileName
    Workbooks(newFileName).Activate
    
    ' ==============================================================================================================================
    MsgBox "First we will format report.", , shtName & ": Step 1"
    ' ==============================================================================================================================
    
    ' Clean up report
    Rows("1:13").Select
    Selection.Delete Shift:=xlUp
    Range("A1:C1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.UnMerge
    Rows("1:1").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireColumn.Delete
    Columns("B:B").EntireColumn.Delete ' Delete participant names
    Columns("D:D").Select ' Use column D in case an appointment has more than one procedure code.
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    ' Make report into an Excel table so we can filter, etc.
    Range("A1").Select
    ActiveCell.Value = "Ptt ID"
    With Application.ActiveSheet
        lastRowNum = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    Range(Cells(1, 1), Cells(lastRowNum, 4)).Select ' columns A-D
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = tblName
    ActiveSheet.ListObjects(tblName).TableStyle = "TableStyleMedium15"
    Selection.Columns.AutoFit
    Range("A1").Select
    
    ' ==============================================================================================================================
    MsgBox "Next we'll check for multiple procedure codes per visit. We'll combine appointments with multiple procedure codes. " _
        & "Note: This may take awhile depending on how many entries the report has." & vbNewLine _
        & "(The report has " & lastRowNum - 1 & " entries)" _
        , , shtName & ": Step 2"
    ' ==============================================================================================================================
    
    ' Add multiple procedure codes to one row
    Range("A2").Select
    lastRowNum_preDeDup = lastRowNum
    Do While ActiveCell.Row <= lastRowNum
        currentRow = ActiveCell.Row
        If IsEmpty(ActiveCell.Value) Then
            Cells(currentRow - 1, 4).Value = Replace(Cells(currentRow - 1, 4).Value, ",", "") & " | " & Replace(Cells(currentRow, 4).Value, ",", "")
            With Cells(currentRow - 1, 4).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            ActiveCell.EntireRow.Delete
            lastRowNum = lastRowNum - 1
        Else
            ActiveCell.Offset(1, 0).Select
        End If
    Loop
    MsgBox "Reconciled " & lastRowNum_preDeDup - lastRowNum & " multiple procedure codes.", , shtName & ": Step 2"
    
    ' ==============================================================================================================================
    MsgBox "Now going to format the dates, because WellSky couldn't be bothered to give us properly-formatted dates " _
        & " even though we pay them over $10,000 a year.", , shtName & ": Step 3"
    ' ==============================================================================================================================
    
    ' Format Dates
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C2").Select
    ActiveCell.Formula2R1C1 = _
        "=TEXT(CONCAT(TRIM(LEFT(RC[-1],FIND(""/"",RC[-1])-1)),""/"",TRIM(MID(RC[-1],FIND(""/"",RC[-1])+1,2)),""/"",TRIM(RIGHT(RC[-1],4))),""mm/dd/yyyy"")"
    Range("C2:C" & lastRowNum).Select
    Selection.Copy
    Range("B2:B" & lastRowNum).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    With ActiveSheet.ListObjects(tblName).Sort
        .Header = xlYes
        .SortFields.Clear
        .SortFields.Add Key:=Range("B1:B" & lastRowNum), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        .Apply
    End With
    Range("A1").Select
    
    ' ==============================================================================================================================
    MsgBox "Next, we will conditionally format the ""Status"" column.", , shtName & ": Step 4"
    ' ==============================================================================================================================
    
    ' Conditional formatting for Status column (column C/#3)
    
    Range("C2:C" & lastRowNum).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Pending"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16754788
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Cancel"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ' ==============================================================================================================================
    MsgBox "Lastly, we will format the report and add our own header.", , shtName & ": Step 5"
    ' ==============================================================================================================================
    
    ' Next, format the table...
    Range("A1:E" & lastRowNum).Select
    Range("A1").Activate
    With Selection.Font
        .Name = "Calibri"
        .Size = 12
    End With
    ' ...and autofit the row heights and column widths
    Range("A1:E" & lastRowNum).Select
    For Each c In Selection.Columns
        c.ColumnWidth = 255
    Next c
    Selection.Columns.AutoFit ' address columns will be wide because of the error messages
    Selection.Rows.AutoFit
    
    ' Finally, put our own header in columns A-C... and we're done!
    Rows("1:1").Select
    Range("A1").Activate
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.Value = "Report Run on:"
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
    End With
    Range("B1").Select
    ActiveCell.Value = dateOfReport
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
    End With
    Range("C1").Select
    ActiveCell.Value = shtName
    With Selection
        .Font.Bold = True
        .Font.Size = 20
        .EntireRow.AutoFit
    End With
    
    'Save the workbook
    ActiveWorkbook.Save
    
    ' The original report is just so awful.
    ' Did they /intentionally/ do it so you can't get your own data out of their system?
    ' Or are they just idiots?
    MsgBox "The process is complete.", , shtName & ": Formatted for, Like, _Actual_ Use"
    
    
End Sub


