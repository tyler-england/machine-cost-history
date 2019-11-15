VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub GetModelData()

    Dim fileNames As New Collection
    Dim file As Variant, fileExt As String
    Dim wBook1 As Workbook, wBook2 As Workbook, wSheet As Worksheet
    Dim i As Integer, rowNum As Integer
    Dim wkbkDir As String, bomName As String
    
    '''''hardcoded values
    wkbkDir = "\\PSACLW02\Home\Shared\Mateer\Past Machine Costs\"
    '''''
    
    Set wBook1 = ActiveWorkbook
    
    Range("A5:J1000").ClearContents
    If Right(wkbkDir, 1) <> "\" Then
        wkbkDir = wkbkDir & "\"
    End If
    filepath = wkbkDir & Range("C2").Value & "\"
    file = Dir$(filepath & "*.xls*")
    
    Do While file <> ""
        fileNames.Add (Replace(file, ".xls", ""))
        file = Dir$
    Loop

    Set wSheet = Sheet6
    i = 1
    
    For Each file In fileNames
            wSheet.Range("A" & i + 4).Value = fileNames(i)
            i = i + 1
    Next file

    i = 5 'row 5
    
    Range("A5").Select
    Selection.End(xlDown).Select
    
    If ActiveCell.row < 500 Then
    
        rowNum = ActiveCell.row
        Range("A5").Select
        
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        For i = 5 To rowNum
            fileExt = Dir$(filepath & Range("A" & i).Value & "*")
            
            If fileExt > "" Then
            
                Range("B" & i).Formula = "=right(A" & i & ",len(A" & i & ")-6)"
                Range("C" & i).Formula = "=right('" & filepath & "[" & fileExt & "]Sheet1'!$C$4,6)"
                
                'FINAL? toggle --- uncomment as application requires
                'Range("D4").value="Final?"
                'Range("D" & i).Formula = "=IF('" & filepath & "[" & fileExt & "]Sheet1'!$I$1=" & """Final""" & "," & """Yes""" & "," & """No""" & ")"
                Range("D4").Value = "Job Name"
                Range("D" & i).Formula = "=trim('" & filepath & "[" & fileExt & "]Sheet1'!$B$4)"
                
                
                Range("E" & i).Formula = "=trim(right('" & filepath & "[" & fileExt & "]Sheet1'!$B$5,LEN('" & filepath & "[" & fileExt & "]Sheet1'!$B$5)-10))"
                Range("F" & i).Formula = "=VLOOKUP(" & """TOTAL COST TO DATE""" & ",'" & filepath & "[" & fileExt & "]Sheet1'!$B$21:$I$150,4,FALSE)"
                Range("G" & i).Formula = "=IFERROR(VALUE(IFERROR(RIGHT(VLOOKUP(" & """MANUFACTURING LABOR & OVERHEAD ADJUSTMENT""" & ",'" & filepath & "[" & fileExt & "]Sheet1'!$B$21:$I$150,7,FALSE),3),RIGHT(VLOOKUP(" & """MANUFACTURING LABOR & OVERHEAD:ADJUSTMENT""" & ",'" & filepath & "[" & fileExt & "]Sheet1'!$B$21:$I$150,7,FALSE),3))),0)"
                Range("H" & i).Formula = "=IFERROR(VALUE(VLOOKUP(" & """ENGINEERING """ & ",'" & filepath & "[" & fileExt & "]Sheet1'!$B$21:$I$150,7,FALSE))-VALUE(G" & i & "),0)"
                Range("I" & i).Formula = "=IFERROR(VALUE(VLOOKUP(" & """ASSEMBLY """ & ",'" & filepath & "[" & fileExt & "]Sheet1'!$B$21:$I$150,7,FALSE)),0)"
                Range("J" & i).Formula = "=VALUE(trim('" & filepath & "[" & fileExt & "]Sheet1'!$E$11))"
    
                If Application.WorksheetFunction.IsErr(Range("C" & i).Value) Then
                    MsgBox "Workbook error, row " & i
                    Set wBook2 = Workbooks.Open(filepath & fileExt)
                    wBook1.Activate
                    wBook2.Close False
                End If
                
            End If
            
        Next i
        
        Range("A1").Select
        
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True

    Else:
    Range("A5").Formula = "No folder found with more than 1 machine"
    Range("A6").Select
    
    End If
    
    Sheet6.Columns.AutoFit

End Sub
