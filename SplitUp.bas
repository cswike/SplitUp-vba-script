Attribute VB_Name = "SplitUp"
Sub SplitUp()
    'Filters active worksheet by a specified column, then splits the data into a new sheet for each unique value in specified column.
    'Good way to split out data into tabs for each division.
    'Column to be sorted (i.e. division column) must go all the way to the last row of the data.
    'Assumes row 1 is a header.
    'Currently copies all data starting from column A, this could be changed but it didn't seem useful enough to bother.
    
    Dim sht As Worksheet
    Set sht = ThisWorkbook.ActiveSheet
    
    ' Get column with data to be sorted
    Dim sortCol As String
    sortCol = UCase(InputBox("Which column would you like to split into sheets? (column letter)", "Column to split", "A"))
    If sortCol = "" Then Exit Sub
    While Not IsLetter(sortCol)
        sortCol = InputBox("Error: not a letter. Enter the column you need to sort (column letter)", "Column to sort", "A")
        If sortCol = "" Then Exit Sub
    Wend
    
    Dim sortColNo As Integer
    sortColNo = GetColNo(sortCol)
    
    Dim lastCol As String
    Dim vArr
    vArr = Split(Cells(1, sht.UsedRange.Columns(sht.UsedRange.Columns.Count).Column).Address(True, False), "$")
    lastCol = InputBox("What's the last column (i.e. furthest right) in your spreadsheet that you'd like to include? (column letter)", "Last column", vArr(0))
    
    If Not IsLetter(lastCol) Then Exit Sub
    
    
    Application.ScreenUpdating = False
    Dim rng As Range
    Dim cl As Range
    Dim dict As Object
    Dim ky As Variant
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    With sht
        Set rng = .Range(.Range(sortCol & "2"), .Range(sortCol & .Rows.Count).End(xlUp))
    End With
    
    For Each cl In rng
        If Not dict.exists(cl.Value) Then
            dict.Add cl.Value, cl.Value
        End If
    Next cl
    
    For Each ky In dict.keys
        Dim new_name As String
        new_name = dict(ky)
        sht.Range("A1:" & lastCol & sht.Rows.Count).AutoFilter Field:=sortColNo, Criteria1:=ky
        Dim LR As Long
        LR = Range(sortCol & Rows.Count).End(xlUp).Row
        Range("A2:" & lastCol & LR).SpecialCells(xlCellTypeVisible).Copy
        Sheets.Add(After:=Sheets(Sheets.Count)).name = new_name
        Sheets(new_name).Cells(2, 1).PasteSpecial
        sht.Range("A1:" & lastCol & "1").Copy
        Sheets(new_name).Cells(1, 1).PasteSpecial
        Sheets(new_name).Cells.EntireColumn.AutoFit
        Sheets(new_name).Cells(1, 1).Select
        sht.Select
    Next ky
    
    sht.AutoFilterMode = False
    
    With Application
        .CutCopyMode = False
        .ScreenUpdating = True
    End With
    
End Sub

Function IsLetter(r As String) As Boolean
    Dim x As String
    Dim Counter As Integer
    For Counter = 1 To Len(r)
        x = UCase(Mid(r, Counter, 1))
        IsLetter = Asc(x) > 64 And Asc(x) < 91
        If IsLetter = False Then Exit For
    Next
End Function

Function GetColNo(Col As String) As Integer
' Returns 0 if column is greater than ZZ (even though Excel can go up to XFD)
' ...but really, isn't 702 columns enough for you?
    
    If Len(Col) = 1 Then
        GetColNo = Asc(UCase(Col)) - 64
    ElseIf Len(Col) = 2 Then
        Dim L1 As Integer
        Dim L2 As Integer
        L1 = Asc(UCase(Left(Col, 1))) - 64
        L2 = Asc(UCase(Right(Col, 1))) - 64
        GetColNo = (26 * L1) + L2
    Else
        GetColNo = 0
    End If

End Function
