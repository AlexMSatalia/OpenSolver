Attribute VB_Name = "SolverLinear"
Option Explicit

Sub WriteConstraintListToSheet(r As Range, s As COpenSolver)
          ' Write a list of all the constraints in a column at cell r
          ' TODO: This will not correctly handle constraints on another sheet
          
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         r.Cells(1, 1).Value2 = "Cons"
4         r.Cells(1, 2).Value2 = "SP"
5         r.Cells(1, 3).Value2 = "Inc"
6         r.Cells(1, 4).Value2 = "Dec"
          
          Dim constraint As Long, row As Long, instance As Long, i As Long
7         i = 1
8         For row = 1 To s.NumRows
9             constraint = s.RowToConstraint(row)
10            instance = s.GetConstraintInstance(row, constraint)
              
              Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
11            s.GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
12            If Not RHSCellRange Is Nothing Then
13                RHSstring = RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
14            Else
15                RHSstring = ConvertToCurrentLocale(StripWorksheetNameAndDollars(RHSstring, s.sheet))
16            End If
              
              Dim summary As String
17            summary = LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & _
                        RelationEnumToString(s.Relation(constraint)) & RHSstring
          
18            r.Cells(row + 1, 1).value = summary
19            r.Cells(row + 1, 2).Value2 = ZeroIfSmall(s.ConShadowPrice(row))
20            r.Cells(row + 1, 3).Value2 = ZeroIfSmall(s.ConIncrease(row))
21            r.Cells(row + 1, 4).Value2 = ZeroIfSmall(s.ConDecrease(row))
22        Next row

          'Write the variable duals
23        row = row + 2
24        r.Cells(row, 1).Value2 = "Vars"
25        r.Cells(row, 2).Value2 = "RC"
26        r.Cells(row, 3).Value2 = "Inc"
27        r.Cells(row, 4).Value2 = "Dec"
28        row = row + 1

29        For i = 1 To s.NumVars
30            r.Cells(row, 1).Value2 = s.VarCellName(i)
31            r.Cells(row, 2).Value2 = ZeroIfSmall(s.VarReducedCost(i))
32            r.Cells(row, 3).Value2 = ZeroIfSmall(s.VarIncrease(i))
33            r.Cells(row, 4).Value2 = ZeroIfSmall(s.VarDecrease(i))
34            row = row + 1
35        Next i

ExitSub:
36        If RaiseError Then RethrowError
37        Exit Sub

ErrorHandler:
38        If Not ReportError("SolverModelDiff", "WriteConstraintListToSheet") Then Resume
39        RaiseError = True
40        GoTo ExitSub

End Sub
Sub WriteConstraintSensitivityTable(sheet As Worksheet, s As COpenSolver)
      'Writes out the sensitivity table on a new page (like the solver sensitivity report)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim Column As Long, row As Long, i As Long
          
3         sheet.Cells(1, 1) = "OpenSolver Sensitivity Report - " & s.Solver.ShortName
4         sheet.Cells(2, 1) = "Worksheet: [" & sheet.Parent.Name & "] " & sheet.Name
5         sheet.Cells(3, 1) = "Report Created: " & Now()
          
6         Column = 2
7         row = 6
          
          Dim sheetName As String
8         sheetName = sheet.Name


9         sheet.Cells(row - 1, Column - 1) = "Decision Variables"

          Dim headings() As Variant
10        headings = Array("Cells", "Name", "Final Value", "Reduced Costs", "Objective Value", "Allowable Increase", "Allowable Decrease")
11        sheet.Cells(row, Column).Resize(1, UBound(headings) - LBound(headings) + 1).value = headings

          Dim CostVector() As Double
12        CostVector = s.CostCoeffs.AsVector(s.NumVars)

13        row = row + 1
          
          'put the values into the variable table
14        For i = 1 To s.NumVars
15            sheet.Cells(row, Column) = s.VarCellName(i)
16            sheet.Cells(row, Column + 2) = ZeroIfSmall(s.VarFinalValue(i))
17            sheet.Cells(row, Column + 3) = ZeroIfSmall(s.VarReducedCost(i))
18            sheet.Cells(row, Column + 4) = ZeroIfSmall(CostVector(i))
19            sheet.Cells(row, Column + 5) = ZeroIfSmall(s.VarIncrease(i))
20            sheet.Cells(row, Column + 6) = ZeroIfSmall(s.VarDecrease(i))
21            sheet.Cells(row, Column + 1) = findName(s.sheet, s.VarCellName(i))
22            row = row + 1
23        Next i
          
24        row = row + 2
          
25        headings(LBound(headings) + 3) = "Shadow Price"
26        headings(LBound(headings) + 4) = "RHS Value"

          'Headings for constraint table
27        sheet.Cells(row - 1, Column - 1) = "Constraints"
28        sheet.Cells(row, Column).Resize(1, UBound(headings) - LBound(headings) + 1).value = headings

29        row = row + 1

          'Values for constraint table
30        For i = 1 To s.NumRows
31            sheet.Cells(row, Column + 2) = ZeroIfSmall(s.ConFinalValue(i))
32            sheet.Cells(row, Column + 3) = ZeroIfSmall(s.ConShadowPrice(i))
33            sheet.Cells(row, Column + 4) = ZeroIfSmall(s.RHS(i))
34            sheet.Cells(row, Column + 5) = ZeroIfSmall(s.ConIncrease(i))
35            sheet.Cells(row, Column + 6) = ZeroIfSmall(s.ConDecrease(i))

              Dim constraint As Long, instance As Long
36            constraint = s.RowToConstraint(i)
37            instance = s.GetConstraintInstance(i, constraint)
              
              Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
38            s.GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
39            If Not RHSCellRange Is Nothing Then
40                RHSstring = RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
41            Else
42                RHSstring = ConvertToCurrentLocale(StripWorksheetNameAndDollars(RHSstring, s.sheet))
43            End If
              Dim summary As String
44            summary = LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & _
                        RelationEnumToString(s.Relation(constraint)) & RHSstring
              'Cell Range for each constraint
45            sheet.Cells(row, Column).value = summary
              'Finds the nearest name for the constraint
46            sheet.Cells(row, Column + 1) = findName(s.sheet, LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False))
47            row = row + 1
48        Next i

          'Format the sensitivity table
49        FormatSensitivityTable sheet, row, s.NumVars

ExitSub:
50        If RaiseError Then RethrowError
51        Exit Sub

ErrorHandler:
52        If Not ReportError("SolverModelDiff", "WriteConstraintSensitivityTable") Then Resume
53        RaiseError = True
54        GoTo ExitSub
          
End Sub

Sub FormatSensitivityTable(sheet As Worksheet, row As Long, NumVars As Double)
      'Formats the sensitivity table on the new page with borders and bold writing
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim ScreenStatus As Boolean
3         ScreenStatus = Application.ScreenUpdating
4         Application.ScreenUpdating = False

          Dim currentSheet As Worksheet
5         GetActiveSheetIfMissing currentSheet
          
6         sheet.Select
7         ActiveWindow.DisplayGridlines = False

          Dim startRow As String
8         startRow = 6
          
9         sheet.Cells.EntireColumn.AutoFit
10        Columns("A:A").ColumnWidth = 5
11        With sheet.Range(Cells(2, 2), Cells(row, 8))
12            .HorizontalAlignment = xlCenter
13        End With
          
          'Create the borders for the constraint table
14        sheet.Range(Cells(startRow, 2), Cells(startRow + NumVars, 8)).Select
15        With Selection.Borders(xlEdgeLeft)
16            .LineStyle = xlContinuous
17            .ColorIndex = 0
18            .TintAndShade = 0
19            .Weight = xlMedium
20        End With
21        With Selection.Borders(xlEdgeTop)
22            .LineStyle = xlContinuous
23            .ColorIndex = 0
24            .TintAndShade = 0
25            .Weight = xlMedium
26        End With
27        With Selection.Borders(xlEdgeBottom)
28            .LineStyle = xlContinuous
29            .ColorIndex = 0
30            .TintAndShade = 0
31            .Weight = xlMedium
32        End With
33        With Selection.Borders(xlEdgeRight)
34            .LineStyle = xlContinuous
35            .ColorIndex = 0
36            .TintAndShade = 0
37            .Weight = xlMedium
38        End With
39        With Selection.Borders(xlInsideVertical)
40            .LineStyle = xlContinuous
41            .Weight = xlThin
42        End With
          
          'Create the borders for the variable table
43        sheet.Range(Cells(NumVars + startRow + 3, 2), Cells(row - 1, 8)).Select
44        With Selection.Borders(xlEdgeLeft)
45            .LineStyle = xlContinuous
46            .ColorIndex = 0
47            .TintAndShade = 0
48            .Weight = xlMedium
49        End With
50        With Selection.Borders(xlEdgeTop)
51            .LineStyle = xlContinuous
52            .ColorIndex = 0
53            .TintAndShade = 0
54            .Weight = xlMedium
55        End With
56        With Selection.Borders(xlEdgeBottom)
57            .LineStyle = xlContinuous
58            .ColorIndex = 0
59            .TintAndShade = 0
60            .Weight = xlMedium
61        End With
62        With Selection.Borders(xlEdgeRight)
63            .LineStyle = xlContinuous
64            .ColorIndex = 0
65            .TintAndShade = 0
66            .Weight = xlMedium
67        End With
68        With Selection.Borders(xlInsideVertical)
69            .LineStyle = xlContinuous
70            .Weight = xlThin
71        End With
          
          'Bold the constraint table headings and make them blue as well as put a border around them
72        sheet.Range(Cells(NumVars + startRow + 3, 2), Cells(NumVars + startRow + 3, 8)).Select
73        With Selection.Borders(xlEdgeBottom)
74            .LineStyle = xlContinuous
75            .ColorIndex = 0
76            .TintAndShade = 0
77            .Weight = xlMedium
78        End With
79        With Selection.Font
80            .Bold = True
81            .ThemeColor = xlThemeColorLight2
82        End With
          
          'Bold the variable table headings and make them blue as well as put a border around them
83        sheet.Range(Cells(startRow, 2), Cells(startRow, 8)).Select
84        With Selection.Borders(xlEdgeBottom)
85            .LineStyle = xlContinuous
86            .ColorIndex = 0
87            .TintAndShade = 0
88            .Weight = xlMedium
89        End With
90        With Selection.Font
91            .Bold = True
92            .ThemeColor = xlThemeColorLight2
93        End With
          
          'Bold the headings
94        With sheet.Range("A:A").Font
95            .Bold = True
96        End With
          
97        currentSheet.Select

ExitSub:
98        Application.ScreenUpdating = ScreenStatus
99        If RaiseError Then RethrowError
100       Exit Sub

ErrorHandler:
101       If Not ReportError("SolverModelDiff", "FormatSensitivityTable") Then Resume
102       RaiseError = True
103       GoTo ExitSub
End Sub

Function findName(searchSheet As Worksheet, cell As String) As String
      'Finds the name of a constraint or variable by finding the nearest strings to the left and above the cell and putting these together
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim NotFoundStringLeft As Boolean, NotFoundStringTop As Boolean
          Dim row As Long, col As Long, i As Long, j As Long
          Dim CellValue As String, LHSName As String, AboveName As String

3         row = searchSheet.Range(cell).row
4         col = searchSheet.Range(cell).Column
5         i = col - 1
6         j = row - 1
          
7         NotFoundStringLeft = True
8         NotFoundStringTop = True
          'Loop through to the left and above the cell to find the first non-numeric cell
9         While (NotFoundStringLeft And i > 0) Or (NotFoundStringTop And j > 0)
              'Find the nearest name to the left of the variable or constraint if one exists
10            If i > 0 And NotFoundStringLeft Then
11                CellValue = CStr(searchSheet.Cells(row, i).value)
                  'x = IsAmericanNumber(CellValue)
12                If Not (IsNumeric(CellValue) Or (Left(CellValue, 1) = "=") Or (Len(CellValue) = 0)) Then
13                    NotFoundStringLeft = False
14                End If
15                i = i - 1
16            End If
              'Find the nearest name above the variable or constraint if it exists
17            If j > 0 And NotFoundStringTop Then
18                CellValue = searchSheet.Cells(j, col)
19                If Not (IsNumeric(CellValue) Or (Left(CellValue, 1) = "=") Or (Len(CellValue) = 0)) Then
20                    NotFoundStringTop = False
21                End If
22                j = j - 1
23            End If
24        Wend
25        LHSName = CStr(searchSheet.Cells(row, i + 1).value)
26        AboveName = CStr(searchSheet.Cells(j + 1, col).value)
          
          'Put the names together
27        If Len(AboveName) = 0 Then
28            findName = LHSName
29        ElseIf Len(LHSName) = 0 Then
30            findName = AboveName
31        Else
32            findName = LHSName & " " & AboveName
33        End If

ExitFunction:
34        If RaiseError Then RethrowError
35        Exit Function

ErrorHandler:
36        If Not ReportError("SolverModelDiff", "findName") Then Resume
37        RaiseError = True
38        GoTo ExitFunction

End Function

