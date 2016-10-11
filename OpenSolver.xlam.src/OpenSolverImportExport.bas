Attribute VB_Name = "OpenSolverImportExport"
Option Explicit

Private Const HIDDEN_SHEET_NAME As String = "__OpenSolver__"

Private Function GetHiddenSheet() As Worksheet
          Dim sheet As Worksheet
1         On Error GoTo doesntExist
2         Set sheet = ActiveWorkbook.Sheets(HIDDEN_SHEET_NAME)
3         Set GetHiddenSheet = sheet
4         Exit Function
          
doesntExist:
5         Set sheet = ActiveWorkbook.Sheets.Add
6         sheet.Name = HIDDEN_SHEET_NAME
7         sheet.Visible = xlSheetHidden
8         Set GetHiddenSheet = sheet
End Function

Sub FillHiddenSheet()
1         Application.Interactive = False
2         Application.ScreenUpdating = False

          Dim HiddenSheet As Worksheet
3         Set HiddenSheet = GetHiddenSheet()
4         HiddenSheet.UsedRange.ClearContents
          
          Dim modelSheet As Worksheet, ColIndex As Long
5         ColIndex = 1
6         For Each modelSheet In ActiveWorkbook.Sheets
7             If modelSheet.Visible = xlSheetVisible And modelSheet.Name <> HIDDEN_SHEET_NAME Then
                  Dim StartingCell As Range
8                 Set StartingCell = HiddenSheet.Cells(1, ColIndex)
9                 ExportModel modelSheet, StartingCell
10                ColIndex = ColIndex + 1
11            End If
12        Next modelSheet
          
13        Application.ScreenUpdating = True
14        Application.Interactive = True
End Sub

Sub ExportModel(modelSheet As Worksheet, StartingCell As Range)
          Dim NamesToWrite() As String, RowCount As Long
1         ReDim NamesToWrite(1 To modelSheet.Names.Count + 1) As String
2         RowCount = 0
3         AddNameEntry "ModelSheet", EscapeSheetName(modelSheet) & "A:Z", NamesToWrite, RowCount

          Dim ModelName As Name
4         For Each ModelName In modelSheet.Names
              Dim NameKey As String, SplitValues() As String
5             SplitValues = Split(ModelName.Name, "!")
6             NameKey = SplitValues(UBound(SplitValues))
              
7             If InStr(NameKey, "OpenSolver_") > 0 Or InStr(NameKey, "solver_") > 0 Then
8                 AddNameEntry NameKey, RemoveEquals(ModelName.RefersTo), NamesToWrite, RowCount
9             End If
10        Next ModelName
          
          ' Copy out the names into a RowCount-by-1 variant for writing to the sheet
          ' We can't use this as the original array because we can't redim the first dim of a 2D array
          Dim FormulaeForWriting() As Variant, i As Long
11        ReDim FormulaeForWriting(1 To RowCount, 1 To 1) As Variant
12        For i = 1 To RowCount
13            FormulaeForWriting(i, 1) = NamesToWrite(i)
14        Next i
          
          Dim RangeToWrite As Range
15        Set RangeToWrite = StartingCell.Resize(RowCount, 1)
16        RangeToWrite.Formula = FormulaeForWriting
          
End Sub

Private Sub AddNameEntry(Key As String, value As String, NamesToWrite() As String, RowCount As Long)
1         RowCount = RowCount + 1
2         NamesToWrite(RowCount) = "=" & Key & "=" & value
End Sub
