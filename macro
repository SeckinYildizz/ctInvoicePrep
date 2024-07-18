Sub CT_Billing()
    
    'Set required variables
    Dim newBook As Workbook
    Dim ws, ts, output As Worksheet
    Dim aDate, aServ, aCode, aCI, aCO, aNumU, aURate, aTotal As Range
    Dim filePath, rawData, monthName As String
    Dim i, ii, lastRow1, lastRow2 As Integer
    Dim folderDialog As FileDialog
    
    ' Ask user to select the target folder on Windows
    ' Create a FileDialog object as a Folder Picker dialog
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    ' Set properties for the folder dialog
    With folderDialog
        .Title = "Select a Folder"
        .AllowMultiSelect = False ' Allow the user to select only one folder
        If .Show = -1 Then ' If user clicks OK
            filePath = .SelectedItems(1) & "\" ' Get the selected folder path
        Else
            MsgBox "No folder selected. Exiting..."
            Exit Sub
        End If
    End With
    
    ' Ask user to select the target folder on Mac
    'filePath = MacScript("return POSIX path of (choose folder with prompt ""Select a folder"") as string")
    'If filePath = "" Then
    '    MsgBox "No folder selected. Exiting..."
    '    Exit Sub
    'End If
    
    'Ask the month
    monthName = InputBox("Type the name of the month you would like to bill.")
    monthName = UCase(Left(monthName, 1)) & Mid(monthName, 2)
    
    'Set the raw data sheet and find the last row in it
    Set ws = ActiveWorkbook.Sheets(2)
    lastRow1 = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    'Add a new sheet for the incoming loop
    Sheets.Add after:=Sheets(Sheets.Count)
    Set ts = Sheets(Sheets.Count)
    ws.Range("A2:A" & lastRow1).AdvancedFilter _
        Action:=xlFilterCopy, CopyToRange:=ts.Range("A1"), Unique:=True
    
    'Find the last row of the ts to loop over it
    lastRow2 = ts.Cells(ts.Rows.Count, "A").End(xlUp).Row
    
    For ii = 1 To lastRow2
        Sheets(4).Copy after:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).Name = Trim(ts.Range("A" & ii).Value)
        Set output = Worksheets(Trim(ts.Range("A" & ii).Value))
        output.Range("G5").Value = ts.Range("A" & ii).Value
        output.Range("F11").Value = monthName
        Set aDate = output.Range("A14")
        Set aServ = output.Range("D14")
        Set aCode = output.Range("K14")
        Set aCI = output.Range("P14")
        Set aCO = output.Range("Q14")
        Set aNumU = output.Range("S14")
        Set aURate = output.Range("V14")
        Set aTotal = output.Range("Y14")
        For i = 1 To lastRow1
            If Trim(ws.Range("A" & i).Value) = Trim(ts.Range("A" & ii).Value) And Trim(ws.Range("K" & i).Value) = "Approved" Then
                aDate.Value = ws.Range("B" & i).Value
                aCode.Value = Left(Trim(ws.Range("C" & i).Value), 5)
                On Error Resume Next
                aServ.Value = WorksheetFunction.VLookup(aCode.Value, Sheets(3).UsedRange, 2, False)
                On Error GoTo 0
                aNumU.Value = ws.Range("F" & i).Value
                aURate.Value = ws.Range("G" & i).Value
                aTotal.Value = aNumU.Value * aURate.Value
                
                'Shift the active lines one row
                Set aDate = aDate.Offset(1, 0)
                Set aServ = aServ.Offset(1, 0)
                Set aCode = aCode.Offset(1, 0)
                Set aCI = aCI.Offset(1, 0)
                Set aCO = aCO.Offset(1, 0)
                Set aNumU = aNumU.Offset(1, 0)
                Set aURate = aURate.Offset(1, 0)
                Set aTotal = aTotal.Offset(1, 0)
            End If
        Next i
        Set newBook = Workbooks.Add
        output.Copy Before:=newBook.Sheets(1)
        Application.DisplayAlerts = False
        newBook.Sheets(2).Delete
        Application.DisplayAlerts = True
        newBook.SaveAs Filename:=filePath & monthName & " " & Year(Now) & " " & ts.Range("A" & ii).Value & ".xlsx", FileFormat:=xlOpenXMLStrictWorkbook
        newBook.Close SaveChanges:=False
        Application.DisplayAlerts = False
        output.Delete
        Application.DisplayAlerts = True
    Next ii
    
    Application.DisplayAlerts = False
    Sheets(Sheets.Count).Delete
    Application.DisplayAlerts = True
    Sheets(1).Select
    MsgBox ("Process is finished. Check the target location to see the output.")
End Sub
