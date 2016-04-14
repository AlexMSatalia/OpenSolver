Attribute VB_Name = "OpenSolverImportExport"
Option Explicit

Private Const HIDDEN_SHEET_NAME As String = "__OpenSolver__"

Private Function GetHiddenSheet() As Worksheet
    Dim sheet As Worksheet
    On Error GoTo doesntExist
    Set sheet = ActiveWorkbook.Sheets(HIDDEN_SHEET_NAME)
    Set GetHiddenSheet = sheet
    Exit Function
    
doesntExist:
    Set sheet = ActiveWorkbook.Sheets.Add
    sheet.Name = HIDDEN_SHEET_NAME
    sheet.Visible = xlSheetHidden
    Set GetHiddenSheet = sheet
End Function

Sub FillHiddenSheet()
    Application.Interactive = False
    Application.ScreenUpdating = False

    Dim HiddenSheet As Worksheet
    Set HiddenSheet = GetHiddenSheet()
    HiddenSheet.UsedRange.ClearContents
    
    Dim modelSheet As Worksheet, ColIndex As Long
    ColIndex = 1
    For Each modelSheet In ActiveWorkbook.Sheets
        If modelSheet.Visible = xlSheetVisible And modelSheet.Name <> HIDDEN_SHEET_NAME Then
            Dim StartingCell As Range
            Set StartingCell = HiddenSheet.Cells(1, ColIndex)
            ExportModel modelSheet, StartingCell
            ColIndex = ColIndex + 1
        End If
    Next modelSheet
    
    Application.ScreenUpdating = True
    Application.Interactive = True
End Sub

Sub ExportModel(modelSheet As Worksheet, StartingCell As Range)
    Dim NamesToWrite() As String, RowCount As Long
    ReDim NamesToWrite(1 To modelSheet.Names.Count + 1) As String
    RowCount = 0
    AddNameEntry "ModelSheet", EscapeSheetName(modelSheet) & "A:Z", NamesToWrite, RowCount

    Dim ModelName As Name
    For Each ModelName In modelSheet.Names
        Dim NameKey As String, SplitValues() As String
        SplitValues = Split(ModelName.Name, "!")
        NameKey = SplitValues(UBound(SplitValues))
        
        If InStr(NameKey, "OpenSolver_") > 0 Or InStr(NameKey, "solver_") > 0 Then
            AddNameEntry NameKey, RemoveEquals(ModelName.RefersTo), NamesToWrite, RowCount
        End If
    Next ModelName
    
    ' Copy out the names into a RowCount-by-1 variant for writing to the sheet
    ' We can't use this as the original array because we can't redim the first dim of a 2D array
    Dim FormulaeForWriting() As Variant, i As Long
    ReDim FormulaeForWriting(1 To RowCount, 1 To 1) As Variant
    For i = 1 To RowCount
        FormulaeForWriting(i, 1) = NamesToWrite(i)
    Next i
    
    Dim RangeToWrite As Range
    Set RangeToWrite = StartingCell.Resize(RowCount, 1)
    RangeToWrite.Formula = FormulaeForWriting
    
End Sub

Private Sub AddNameEntry(Key As String, value As String, NamesToWrite() As String, RowCount As Long)
    RowCount = RowCount + 1
    NamesToWrite(RowCount) = "=" & Key & "=" & value
End Sub
