Private Sub processFiles(folder)

    Dim file As String
    Dim wb As Workbook
    
    If Right(folder, 1) <> "\" Then
        folder = folder & "\"
    End If
    
    file = Dir(folder & "CIP*.xls")

    Do While Len(file) > 0
        Set wb = Workbooks.Open(folder & file)
        doWork wb
        'wb.Close SaveChanges:=True
        file = Dir()
    Loop
    
End Sub
Private Sub doWork(wb)
    
    'find last row
    Dim lastRow As Integer
    lastRow = 2
    Do While Range("A" & (lastRow + 1)).Value <> ""
        lastRow = lastRow + 1
    Loop
    
    'insert columns
    Columns("A:B").Insert Shift:=xlToRight
    Columns("E:F").Insert Shift:=xlToRight
    
    'formatting
    Range("C2").Copy
    Range("B2").PasteSpecial Paste:=xlPasteFormats
    Range("C1").Copy
    Range("B1").PasteSpecial Paste:=xlPasteAll
    Range("C" & lastRow + 2).Copy
    Range("B" & lastRow + 2).PasteSpecial Paste:=xlPasteAll
    Range("C" & lastRow + 5).Copy
    Range("B" & lastRow + 5).PasteSpecial Paste:=xlPasteAll
    Range("B2:M2").Interior.Color = RGB(189, 215, 238)
    With Range("B2:M" & lastRow).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 0
    End With
    With Range("B2:M" & lastRow).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 0
    End With
    With Range("B2:M" & lastRow).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 0
    End With
    With Range("B2:M" & lastRow).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 0
    End With
    With Range("B2:M" & lastRow).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 0
    End With
    With Range("B2:M" & lastRow).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 0
    End With
   
    'change headers
    Range("B2").Value = "Item No."
    Range("E2").Value = "Award_Current"
    Range("F2").Value = "FY_End"
    Range("G2").Value = "Project"
    Range("H2").Value = "Description"
    Range("I2").Value = "Typ"
    Range("J2").Value = "Cost_Current"
    Range("K2").Value = "Program Manager"
    Range("L2").Value = "Primary Drivers"
    Range("M2").Value = "Ad Memo or Source"
    
    'item no
    Dim i As Integer
    For i = 1 To lastRow - 2
        Range("B" & i + 2).Value = i
    Next i
    
    'program manager
    For i = 1 To lastRow - 2
        If Trim(Range("I" & i + 2).Value) = "4-PROP" Then Range("K" & i + 2).Value = "x"
        If Trim(Range("I" & i + 2).Value) = "2-ROW" Then Range("K" & i + 2).Value = "x"
        If Trim(Range("I" & i + 2).Value) = "10-INSP" Then Range("K" & i + 2).Value = "x"
    Next i
    
    'date formats
    Range("E3:F" & lastRow).Value = "=DATEVALUE(RC[-2])"
    Range("E3:F" & lastRow).NumberFormat = "m/yyyy"
    
    'currency formats
    Range("J3:J" & lastRow).NumberFormat = "$#,##0;($#,##0)"
    Range("J" & lastRow + 2).Value = "=SUM(J3:J" & lastRow & ")"
    Range("J" & lastRow + 2).NumberFormat = "$#,##0;($#,##0)"
    Range("J" & lastRow + 3).Clear
    
    'formatting
    Range("B3:B" & lastRow).HorizontalAlignment = xlCenter
    Range("E3:F" & lastRow).HorizontalAlignment = xlCenter
    Range("K3:K" & lastRow).HorizontalAlignment = xlCenter
    Range("M3:M" & lastRow).HorizontalAlignment = xlCenter
    Sheets(1).Columns.AutoFit
    Columns("C:D").Hidden = True
    Columns("B").ColumnWidth = 8.67
    Columns("F").ColumnWidth = 10.44
    Columns("I").ColumnWidth = 16
    If Columns("J").ColumnWidth < 15.33 Then Columns("J").ColumnWidth = 15.33
    If Columns("K").ColumnWidth < 17.11 Then Columns("K").ColumnWidth = 17.11
    If Columns("L").ColumnWidth < 31.78 Then Columns("L").ColumnWidth = 31.78
    Columns("A").Delete
    
    'page setup
    With Worksheets(1).PageSetup
        .PrintArea = "A1:L" & lastRow + 5
        .PrintTitleRows = "$2:$2"
        .Orientation = xlLandscape
        .LeftFooter = "DATE PRINTED: &D"
        .RightFooter = "PAGE &P OF &N"
        .FitToPagesTall = False
    End With

End Sub
Sub CIPReport()

    Dim folder As String
    
    folder = InputBox("Enter folder path (Ex. 'H:\CIP'): ")
    
    processFiles folder
    
End Sub
