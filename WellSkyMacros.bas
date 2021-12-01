Attribute VB_Name = "WellSkyMacros"
Sub formatWellSkyDate()
    ' Use this formula to format cell [@Date] to, like, an actual date format
    ' =TEXT(CONCAT(TRIM(LEFT([@Date],FIND("/",[@Date])-1)),"/",TRIM(MID([@Date],FIND("/",[@Date])+1,2)),"/",TRIM(RIGHT([@Date],4))),"mm/dd/yyyy")
    ' btw, WellSky sucks
End Sub

Sub WellSky_CleanReport_AppointmentActivity()
'
' WellSky_AppointmentActivity Macro
'
    ' Declare variables
    Dim tblName As String
    Dim currentRow As Long
    Dim lastRowNum As Long
    Dim lastRowNum_preDeDup As Long
    Const lastColNum As Long = 5 ' column #5 = column E -- we know the report only has 5 columns.
    
    
    ' Get the date that the report was run from cel Q12. Set the table name to _
        "AppointmentActivity_RunOn_{date}" to use later
    Range("R12").Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[-1],2)"
    Range("S12").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-2],4,2)"
    Range("T12").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[-3],4)"
    Range("U12").Select
    ActiveCell.FormulaR1C1 = "=TEXT(CONCAT(RC[-1],""-"",TRIM(RC[-3]),""-"",TRIM(RC[-2])),""yyyy-mm-dd"")"
    tblName = "AppointmentActivity_RunOn_" & Selection
    
    ' Clean up report
    Rows("1:13").Select
    Selection.Delete Shift:=xlUp
    Range("A1:C1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.UnMerge
    Rows("1:1").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireColumn.Delete
    Columns("E:E").Select ' Use column E in case an appointment has more than one procedure code.
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    
    ' Make report into an Excel table so we can filter, etc.
    Range("A1").Select
    With Application.ActiveSheet
        lastRowNum = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    Range(Cells(1, 1), Cells(lastRowNum, lastColNum)).Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = tblName
    ActiveSheet.ListObjects(tblName).TableStyle = "TableStyleMedium15"
    Selection.Columns.AutoFit
    Range("A1").Select
    
    ' Add multiple procedure codes to one row
    Range("B2").Select
    lastRowNum_preDeDup = lastRowNum
    MsgBox "Now going to check for multiple procedure codes per visit." & vbNewLine & _
        "Note: This may take awhile depending on how many entries the report has." & vbNewLine & _
        "(The report has " & lastRowNum - 1 & " entries)"
    Do While ActiveCell.Row <= lastRowNum
        currentRow = ActiveCell.Row
        If IsEmpty(ActiveCell.Value) Then
            Cells(currentRow - 1, 5).Value = Cells(currentRow - 1, 5).Value & " " & Cells(currentRow, 5).Value
            ActiveCell.EntireRow.Delete
            lastRowNum = lastRowNum - 1
        Else
            ActiveCell.Offset(1, 0).Select
        End If
    Loop
    MsgBox "Reconciled " & lastRowNum_preDeDup - lastRowNum & " multiple procedure codes."
    
    ' Format Dates
    MsgBox "Now going to format the dates, because WellSky couldn't be bothered to give us properly-formatted dates even though we pay them over $10,000 a year."
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D2").Select
    ActiveCell.Formula2R1C1 = _
        "=TEXT(CONCAT(TRIM(LEFT(RC[-1],FIND(""/"",RC[-1])-1)),""/"",TRIM(MID(RC[-1],FIND(""/"",RC[-1])+1,2)),""/"",TRIM(RIGHT(RC[-1],4))),""mm/dd/yyyy"")"
    Range("D2:D" & lastRowNum).Select
    Selection.Copy
    Range("C2:C" & lastRowNum).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Columns("D:D").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("C1").Select
    Do While ActiveCell.Row <= lastRowNum
        Selection.NumberFormat = "mm/dd/yyyy;@"
        ActiveCell.FormulaR1C1 = Selection.Value
        ActiveCell.Offset(1, 0).Select
    Loop
    With ActiveSheet.ListObjects(tblName).Sort
        .Header = xlYes
        .SortFields.Clear
        .SortFields.Add Key:=Range("C1:C" & lastRowNum), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        .Apply
    End With
    Range("A1").Select
End Sub


Sub WellSky_CleanReport_PatientMailingLabels()
'
' This report converts the WellSky report to a table of data that can be used for Mail Merge
' Run WellSky Report "Patient Mailing Labels" using appropriate dates
' Open the report, convert it to the new version of Excel files, then run this macro
'
    
    ' First, clean up the report
    ' unmerge cells & append column data into Column A
    Dim Last_Row_ColA As Integer
    Dim Last_Row_ColB As Integer
    Dim Last_Row_ColC As Integer

    ' unmerge all cells and delete blank columns
    Cells.Select
    Selection.UnMerge
    Columns("B:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:C").Select
    Selection.Columns.AutoFit
    Cells(1, 1).Select
    
    ' calculate last row of data per column
    Last_Row_ColA = Cells(Rows.Count, 1).End(xlUp).Row + 2 'add 2 blank cells below
    Last_Row_ColB = Cells(Rows.Count, 2).End(xlUp).Row
    Last_Row_ColC = Cells(Rows.Count, 3).End(xlUp).Row
    
    ' append data
    Range(Cells(1, 2), Cells(Last_Row_ColB, 2)).Select
    Selection.Cut
    Cells(Last_Row_ColA, 1).Select
    ActiveSheet.Paste
    Last_Row_ColA = Cells(Rows.Count, 1).End(xlUp).Row + 2 'add 2 blank cells below
    Range(Cells(1, 3), Cells(Last_Row_ColC, 3)).Select
    Selection.Cut
    Cells(Last_Row_ColA, 1).Select
    ActiveSheet.Paste
    
    ' clear clipboard, go to cell 1,1
    Application.CutCopyMode = False
    Cells(1, 1).Select
    
    ' Make column headers
    ' write the column headers for the formatted address data. we will delete column A in the table_PatientMailingLabels() function.
    Range("B1").Select
    ActiveCell.Value = "First Name"
    Range("C1").Select
    ActiveCell.Value = "Last Name"
    Range("D1").Select
    ActiveCell.Value = "Address1"
    Range("E1").Select
    ActiveCell.Value = "Address2"
    Range("F1").Select
    ActiveCell.Value = "City"
    Range("G1").Select
    ActiveCell.Value = "State"
    Range("H1").Select
    ActiveCell.Value = "Zipcode"
    ' select cell A1
    Cells(1, 1).Select
    
    ' For each address block, slice data & insert into table
    ' first name, last name, address1, address2, city, state, zipcode
    Dim sliceRow, blankRowCount As Long
    Dim Name, fName, lName, Address1, Address2, CSZ, City, State, Zipcode As String
    ' CSZ = City/State/Zipcode
    ' select first cell with raw address data
    Cells(2, 1).Select
    Do While ActiveCell.Value <> ""
        ' Do While loop will continue until there are no more non-empty cells
        Name = ActiveCell.Value
        ' first name is first word before first space character
        fName = Trim(Left(Name, InStr(1, Name, " ") - 1))
        lName = Trim(Name)
        Do While InStr(1, lName, " ") > 0
            ' keep trimming the name until there are no more spaces
            lName = Right(lName, Len(lName) - InStr(1, lName, " "))
        Loop
        lName = Trim(lName)
        ActiveCell.Offset(1).Select
        Do While ActiveCell.Value = Name
            ' the report is stupid, it lists the patient's name twice
            ActiveCell.Offset(1).Select
        Loop
        ' no transform needed for Address1
        Address1 = ActiveCell.Value
        ActiveCell.Offset(1).Select
        ' there might be an Address2, i.e., an apartment number; let's look
        If InStr(1, ActiveCell.Value, ",") Then
            ' if there is a comma character in this cell, then it probably isn't an apartment number; rather it is the city/state/zipcode
            Address2 = ""
            CSZ = ActiveCell.Value
        Else
            Address2 = Trim(ActiveCell.Value)
            ActiveCell.Offset(1).Select
            ' sometimes there are more random blank rows here, so we skip through them with this loop.
            Do While ActiveCell.Value = ""
                ActiveCell.Offset(1).Select
            Loop
            CSZ = Trim(ActiveCell.Value)
        End If
        ' now we break apart city, state, and zipcode
        City = Trim(Left(CSZ, InStr(1, CSZ, ",") - 1))
        State = Trim(Mid(CSZ, InStr(1, CSZ, ",") + 2, 2))
        Zipcode = Trim(Mid(CSZ, Len(City) + Len(State) + 3))
        ' the insert_PatientMailingLabels() function will move the selected cell, so we need to keep track of our place with sliceRow
        sliceRow = ActiveCell.Row
        ' now we call the function that actually inserts the information into the table.
        'Call insert_PatientMailingLabels(fName, lName, Address1, Address2, City, State, Zipcode, sliceRow)
        Dim LastRow As Long
        ' find last row, select first blank row for entering data
        LastRow = Cells(Rows.Count, 2).End(xlUp).Row
        Cells(LastRow, 2).Offset(1, 0).Select
        ' insert data
        ActiveCell.Value = fName
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = lName
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = Address1
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = Address2
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = City
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = State
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = Zipcode
        ' select working row from slice_PatientMailingLabels() function
        Cells(sliceRow, 1).Select
        ' are we at the end of the address list? let's check.
        blankRowCount = 0
        ActiveCell.Offset(1).Select
        Do While ActiveCell.Value = "" And blankRowCount < 10
            ' make sure there are no more rows with address data
            ActiveCell.Offset(1).Select
            blankRowCount = blankRowCount + 1
        Loop
    Loop
    
    ' Finally, make the data an Excel table
    Dim LastCell As Range
    ' delete column with raw data
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    ' find last cell in address data
    Set LastCell = Range("A1").End(xlToRight).End(xlDown)
    Range(Cells(1, 1), Cells(LastCell.Row, LastCell.Column)).Select
    Application.CutCopyMode = False
    ' make data into an Excel table named "PatientAddresses"
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(LastCell.Row, LastCell.Column)), , xlYes).Name = "PatientAddresses"
    Range(Cells(1, 1), Cells(LastCell.Row, LastCell.Column)).Select
    ActiveSheet.ListObjects("PatientAddresses").TableStyle = "TableStyleMedium15"
    Selection.Columns.AutoFit
    Selection.Rows.AutoFit
    Cells(1, 1).Select
    
    
End Sub
'
'Sub table_PatientMailingLabels()
'    Dim LastCell As Range
'    ' delete column with raw data
'    Columns("A:A").Select
'    Selection.Delete Shift:=xlToLeft
'    ' find last cell in address data
'    Set LastCell = Range("A1").End(xlToRight).End(xlDown)
'    Range(Cells(1, 1), Cells(LastCell.Row, LastCell.Column)).Select
'    Application.CutCopyMode = False
'    ' make data into an Excel table named "PatientAddresses"
'    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(LastCell.Row, LastCell.Column)), , xlYes).Name = "PatientAddresses"
'    Range(Cells(1, 1), Cells(LastCell.Row, LastCell.Column)).Select
'    ActiveSheet.ListObjects("PatientAddresses").TableStyle = "TableStyleMedium15"
'    Selection.Columns.AutoFit
'    Selection.Rows.AutoFit
'    Cells(1, 1).Select
'
'End Sub
'
'
'Sub insert_PatientMailingLabels(fName, lName, Address1, Address2, City, State, Zipcode, sliceRow)
'    Dim LastRow As Long
'    ' find last row, select first blank row for entering data
'    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
'    Cells(LastRow, 2).Offset(1, 0).Select
'    ' insert data
'    ActiveCell.Value = fName
'    ActiveCell.Offset(0, 1).Select
'    ActiveCell.Value = lName
'    ActiveCell.Offset(0, 1).Select
'    ActiveCell.Value = Address1
'    ActiveCell.Offset(0, 1).Select
'    ActiveCell.Value = Address2
'    ActiveCell.Offset(0, 1).Select
'    ActiveCell.Value = City
'    ActiveCell.Offset(0, 1).Select
'    ActiveCell.Value = State
'    ActiveCell.Offset(0, 1).Select
'    ActiveCell.Value = Zipcode
'    ' select working row from slice_PatientMailingLabels() function
'    Cells(sliceRow, 1).Select
'End Sub
'
'
'Sub slice_PatientMailingLabels()
'    ' first name, last name, address1, address2, city, state, zipcode
'    Dim sliceRow, blankRowCount As Long
'    Dim Name, fName, lName, Address1, Address2, CSZ, City, State, Zipcode As String
'    ' CSZ = City/State/Zipcode
'    ' select first cell with raw address data
'    Cells(2, 1).Select
'    Do While ActiveCell.Value <> ""
'        ' Do While loop will continue until there are no more non-empty cells
'        Name = ActiveCell.Value
'        ' first name is first word before first space character
'        fName = Trim(Left(Name, InStr(1, Name, " ") - 1))
'        lName = Trim(Name)
'        Do While InStr(1, lName, " ") > 0
'            ' keep trimming the name until there are no more spaces
'            lName = Right(lName, Len(lName) - InStr(1, lName, " "))
'        Loop
'        lName = Trim(lName)
'        ActiveCell.Offset(1).Select
'        Do While ActiveCell.Value = Name
'            ' the report is stupid, it lists the patient's name twice
'            ActiveCell.Offset(1).Select
'        Loop
'        ' no transform needed for Address1
'        Address1 = ActiveCell.Value
'        ActiveCell.Offset(1).Select
'        ' there might be an Address2, i.e., an apartment number; let's look
'        If InStr(1, ActiveCell.Value, ",") Then
'            ' if there is a comma character in this cell, then it probably isn't an apartment number; rather it is the city/state/zipcode
'            Address2 = ""
'            CSZ = ActiveCell.Value
'        Else
'            Address2 = Trim(ActiveCell.Value)
'            ActiveCell.Offset(1).Select
'            ' sometimes there are more random blank rows here, so we skip through them with this loop.
'            Do While ActiveCell.Value = ""
'                ActiveCell.Offset(1).Select
'            Loop
'            CSZ = Trim(ActiveCell.Value)
'        End If
'        ' now we break apart city, state, and zipcode
'        City = Trim(Left(CSZ, InStr(1, CSZ, ",") - 1))
'        State = Trim(Mid(CSZ, InStr(1, CSZ, ",") + 2, 2))
'        Zipcode = Trim(Mid(CSZ, Len(City) + Len(State) + 3))
'        ' the insert_PatientMailingLabels() function will move the selected cell, so we need to keep track of our place with sliceRow
'        sliceRow = ActiveCell.Row
'        ' now we call the function that actually inserts the information into the table.
'        Call insert_PatientMailingLabels(fName, lName, Address1, Address2, City, State, Zipcode, sliceRow)
'        ' are we at the end of the address list? let's check.
'        blankRowCount = 0
'        ActiveCell.Offset(1).Select
'        Do While ActiveCell.Value = "" And blankRowCount < 10
'            ' make sure there are no more rows with address data
'            ActiveCell.Offset(1).Select
'            blankRowCount = blankRowCount + 1
'        Loop
'    Loop
'End Sub
'
'
'Sub colHeaders_PatientMailingLabels()
'    ' write the column headers for the formatted address data. we will delete column A in the table_PatientMailingLabels() function.
'    Range("B1").Select
'    ActiveCell.Value = "First Name"
'    Range("C1").Select
'    ActiveCell.Value = "Last Name"
'    Range("D1").Select
'    ActiveCell.Value = "Address1"
'    Range("E1").Select
'    ActiveCell.Value = "Address2"
'    Range("F1").Select
'    ActiveCell.Value = "City"
'    Range("G1").Select
'    ActiveCell.Value = "State"
'    Range("H1").Select
'    ActiveCell.Value = "Zipcode"
'    ' select cell A1
'    Cells(1, 1).Select
'End Sub
'
'
'Sub CleanupReport_PatientMailingLabels()
'    ' unmerge cells & append column data into Column A
'    Dim Last_Row_ColA As Integer
'    Dim Last_Row_ColB As Integer
'    Dim Last_Row_ColC As Integer
'
'    ' unmerge all cells and delete blank columns
'    Cells.Select
'    Selection.UnMerge
'    Columns("B:F").Select
'    Selection.Delete Shift:=xlToLeft
'    Columns("C:F").Select
'    Selection.Delete Shift:=xlToLeft
'    Columns("A:C").Select
'    Selection.Columns.AutoFit
'    Cells(1, 1).Select
'
'    ' calculate last row of data per column
'    Last_Row_ColA = Cells(Rows.Count, 1).End(xlUp).Row + 2 'add 2 blank cells below
'    Last_Row_ColB = Cells(Rows.Count, 2).End(xlUp).Row
'    Last_Row_ColC = Cells(Rows.Count, 3).End(xlUp).Row
'
'    ' append data
'    Range(Cells(1, 2), Cells(Last_Row_ColB, 2)).Select
'    Selection.Cut
'    Cells(Last_Row_ColA, 1).Select
'    ActiveSheet.Paste
'    Last_Row_ColA = Cells(Rows.Count, 1).End(xlUp).Row + 2 'add 2 blank cells below
'    Range(Cells(1, 3), Cells(Last_Row_ColC, 3)).Select
'    Selection.Cut
'    Cells(Last_Row_ColA, 1).Select
'    ActiveSheet.Paste
'
'    ' clear clipboard, go to cell 1,1
'    Application.CutCopyMode = False
'    Cells(1, 1).Select
'
'End Sub
