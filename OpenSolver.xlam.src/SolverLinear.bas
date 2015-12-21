Attribute VB_Name = "SolverLinear"
Option Explicit

Sub WriteConstraintListToSheet(r As Range, s As COpenSolver)
          ' Write a list of all the constraints in a column at cell r
          ' TODO: This will not correctly handle constraints on another sheet
          
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

1916      r.Cells(1, 1).Value2 = "Cons"
1917      r.Cells(1, 2).Value2 = "SP"
1918      r.Cells(1, 3).Value2 = "Inc"
1919      r.Cells(1, 4).Value2 = "Dec"
          
          Dim constraint As Long, row As Long, instance As Long
1921      For row = 1 To s.NumRows
              constraint = s.RowToConstraint(row)
1922          instance = s.GetConstraintInstance(row, constraint)

              Dim UnusedConstraint As Boolean
1923          UnusedConstraint = s.SparseA(row).Count = 0
              
              Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
1924          s.GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
1925          If Not RHSCellRange Is Nothing Then
1926              RHSstring = RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
1927          Else
1928              RHSstring = ConvertToCurrentLocale(StripWorksheetNameAndDollars(RHSstring, s.sheet))
1930          End If
              
              Dim summary As String
1931          summary = IIf(UnusedConstraint, "", "") & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & _
                        RelationEnumToString(s.Relation(constraint)) & RHSstring & IIf(UnusedConstraint, "", "")
          
1932          r.Cells(row + 1, 1).value = summary
1933          r.Cells(row + 1, 2).Value2 = ZeroIfSmall(s.ShadowPrice(row))
1934          r.Cells(row + 1, 3).Value2 = ZeroIfSmall(s.IncreaseCon(row))
1935          r.Cells(row + 1, 4).Value2 = ZeroIfSmall(s.DecreaseCon(row))
1936      Next row

          'Write the variable duals
1937      row = row + 2
1938      r.Cells(row, 1).Value2 = "Vars"
1939      r.Cells(row, 2).Value2 = "RC"
1940      r.Cells(row, 3).Value2 = "Inc"
1941      r.Cells(row, 4).Value2 = "Dec"
1942      row = row + 1
          Dim i As Long
1943      For i = 1 To s.NumVars
1944          r.Cells(row, 1).Value2 = s.VarCell(i)
1945          r.Cells(row, 2).Value2 = ZeroIfSmall(s.ReducedCosts(i))
1946          r.Cells(row, 3).Value2 = ZeroIfSmall(s.IncreaseVar(i))
1947          r.Cells(row, 4).Value2 = ZeroIfSmall(s.DecreaseVar(i))
1948          row = row + 1
1949      Next i

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("SolverModelDiff", "WriteConstraintListToSheet") Then Resume
          RaiseError = True
          GoTo ExitSub

End Sub
Sub WriteConstraintSensitivityTable(sheet As Worksheet, s As COpenSolver)
'Writes out the sensitivity table on a new page (like the solver sensitivity report)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim Column As Long, row As Long, i As Long
          
2273      sheet.Cells(1, 1) = "OpenSolver Sensitivity Report - " & s.Solver.ShortName
2274      sheet.Cells(2, 1) = "Worksheet: [" & sheet.Parent.Name & "] " & sheet.Name
2275      sheet.Cells(3, 1) = "Report Created: " & Now()
          
2276      Column = 2
2277      row = 6
          
          Dim sheetName As String
2278      sheetName = sheet.Name


2280      sheet.Cells(row - 1, Column - 1) = "Decision Variables"

          Dim headings() As Variant
2279      headings = Array("Cells", "Name", "Final Value", "Reduced Costs", "Objective Value", "Allowable Increase", "Allowable Decrease")
          sheet.Cells(row, Column).Resize(1, UBound(headings) - LBound(headings) + 1).value = headings

2284      row = row + 1
          
          'put the values into the variable table
2286      For i = 1 To s.NumVars
2287          sheet.Cells(row, Column) = s.VarCell(i)
2288          sheet.Cells(row, Column + 2) = ZeroIfSmall(s.FinalVarValue(i))
2289          sheet.Cells(row, Column + 3) = ZeroIfSmall(s.ReducedCosts(i))
2290          sheet.Cells(row, Column + 4) = ZeroIfSmall(s.CostCoeffs(i))
2291          sheet.Cells(row, Column + 5) = ZeroIfSmall(s.IncreaseVar(i))
2292          sheet.Cells(row, Column + 6) = ZeroIfSmall(s.DecreaseVar(i))
2293          sheet.Cells(row, Column + 1) = findName(s.sheet, s.VarCell(i))
2294          row = row + 1
2295      Next i
          
2296      row = row + 2
          
2297      headings(LBound(headings) + 3) = "Shadow Price"
2298      headings(LBound(headings) + 4) = "RHS Value"

          'Headings for constraint table
2299      sheet.Cells(row - 1, Column - 1) = "Constraints"
          sheet.Cells(row, Column).Resize(1, UBound(headings) - LBound(headings) + 1).value = headings

2303      row = row + 1

          'Values for constraint table
2305      For i = 1 To s.NumRows
2306          sheet.Cells(row, Column + 2) = ZeroIfSmall(s.FinalValue(i))
2307          sheet.Cells(row, Column + 3) = ZeroIfSmall(s.ShadowPrice(i))
2308          sheet.Cells(row, Column + 4) = ZeroIfSmall(s.RHS(i))
2309          sheet.Cells(row, Column + 5) = ZeroIfSmall(s.IncreaseCon(i))
2310          sheet.Cells(row, Column + 6) = ZeroIfSmall(s.DecreaseCon(i))

              Dim constraint As Long, instance As Long
              constraint = s.RowToConstraint(i)
              instance = s.GetConstraintInstance(i, constraint)
              
              Dim UnusedConstraint As Boolean
2312          UnusedConstraint = s.SparseA(i).Count = 0
              Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
2313          s.GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
2314          If Not RHSCellRange Is Nothing Then
2315              RHSstring = RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
2316          Else
2317              RHSstring = ConvertToCurrentLocale(StripWorksheetNameAndDollars(RHSstring, s.sheet))
2319          End If
              Dim summary As String
2320          summary = IIf(UnusedConstraint, "", "") & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & _
                        RelationEnumToString(s.Relation(constraint)) & RHSstring & IIf(UnusedConstraint, "", "")
              'Cell Range for each constraint
2321          sheet.Cells(row, Column).value = summary
              'Finds the nearest name for the constraint
2322          sheet.Cells(row, Column + 1) = findName(s.sheet, LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False))
2323          row = row + 1
2324      Next i

          'Format the sensitivity table
2325      FormatSensitivityTable sheet, row, s.NumVars

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("SolverModelDiff", "WriteConstraintSensitivityTable") Then Resume
          RaiseError = True
          GoTo ExitSub
          
End Sub

Sub FormatSensitivityTable(sheet As Worksheet, row As Long, NumVars As Double)
'Formats the sensitivity table on the new page with borders and bold writing
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim ScreenStatus As Boolean
          ScreenStatus = Application.ScreenUpdating
2326      Application.ScreenUpdating = False

          Dim currentSheet As Worksheet
          GetActiveSheetIfMissing currentSheet
          
2328      sheet.Select
          ActiveWindow.DisplayGridlines = False

          Dim startRow As String
2329      startRow = 6
          
2330      sheet.Cells.EntireColumn.AutoFit
2331      Columns("A:A").ColumnWidth = 5
2332      With sheet.Range(Cells(2, 2), Cells(row, 8))
2333          .HorizontalAlignment = xlCenter
2334      End With
          
          'Create the borders for the constraint table
2335      sheet.Range(Cells(startRow, 2), Cells(startRow + NumVars, 8)).Select
2336      With Selection.Borders(xlEdgeLeft)
2337          .LineStyle = xlContinuous
2338          .ColorIndex = 0
2339          .TintAndShade = 0
2340          .Weight = xlMedium
2341      End With
2342      With Selection.Borders(xlEdgeTop)
2343          .LineStyle = xlContinuous
2344          .ColorIndex = 0
2345          .TintAndShade = 0
2346          .Weight = xlMedium
2347      End With
2348      With Selection.Borders(xlEdgeBottom)
2349          .LineStyle = xlContinuous
2350          .ColorIndex = 0
2351          .TintAndShade = 0
2352          .Weight = xlMedium
2353      End With
2354      With Selection.Borders(xlEdgeRight)
2355          .LineStyle = xlContinuous
2356          .ColorIndex = 0
2357          .TintAndShade = 0
2358          .Weight = xlMedium
2359      End With
2360      With Selection.Borders(xlInsideVertical)
2361          .LineStyle = xlContinuous
2362          .Weight = xlThin
2363      End With
          
          'Create the borders for the variable table
2364      sheet.Range(Cells(NumVars + startRow + 3, 2), Cells(row - 1, 8)).Select
2365      With Selection.Borders(xlEdgeLeft)
2366          .LineStyle = xlContinuous
2367          .ColorIndex = 0
2368          .TintAndShade = 0
2369          .Weight = xlMedium
2370      End With
2371      With Selection.Borders(xlEdgeTop)
2372          .LineStyle = xlContinuous
2373          .ColorIndex = 0
2374          .TintAndShade = 0
2375          .Weight = xlMedium
2376      End With
2377      With Selection.Borders(xlEdgeBottom)
2378          .LineStyle = xlContinuous
2379          .ColorIndex = 0
2380          .TintAndShade = 0
2381          .Weight = xlMedium
2382      End With
2383      With Selection.Borders(xlEdgeRight)
2384          .LineStyle = xlContinuous
2385          .ColorIndex = 0
2386          .TintAndShade = 0
2387          .Weight = xlMedium
2388      End With
2389      With Selection.Borders(xlInsideVertical)
2390          .LineStyle = xlContinuous
2391          .Weight = xlThin
2392      End With
          
          'Bold the constraint table headings and make them blue as well as put a border around them
2393      sheet.Range(Cells(NumVars + startRow + 3, 2), Cells(NumVars + startRow + 3, 8)).Select
2394      With Selection.Borders(xlEdgeBottom)
2395          .LineStyle = xlContinuous
2396          .ColorIndex = 0
2397          .TintAndShade = 0
2398          .Weight = xlMedium
2399      End With
2400      With Selection.Font
2401          .Bold = True
2402          .ThemeColor = xlThemeColorLight2
2403      End With
          
          'Bold the variable table headings and make them blue as well as put a border around them
2404      sheet.Range(Cells(startRow, 2), Cells(startRow, 8)).Select
2405      With Selection.Borders(xlEdgeBottom)
2406          .LineStyle = xlContinuous
2407          .ColorIndex = 0
2408          .TintAndShade = 0
2409          .Weight = xlMedium
2410      End With
2411      With Selection.Font
2412          .Bold = True
2413          .ThemeColor = xlThemeColorLight2
2414      End With
          
          'Bold the headings
2415      With sheet.Range("A:A").Font
2416          .Bold = True
2417      End With
          
2419      currentSheet.Select

ExitSub:
          Application.ScreenUpdating = ScreenStatus
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("SolverModelDiff", "FormatSensitivityTable") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Function findName(searchSheet As Worksheet, cell As String) As String
'Finds the name of a constraint or variable by finding the nearest strings to the left and above the cell and putting these together
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim NotFoundStringLeft As Boolean, NotFoundStringTop As Boolean
          Dim row As Long, col As Long, i As Long, j As Long
          Dim CellValue As String, LHSName As String, AboveName As String

2421      row = searchSheet.Range(cell).row
2422      col = searchSheet.Range(cell).Column
2423      i = col - 1
2424      j = row - 1
          
2425      NotFoundStringLeft = True
2426      NotFoundStringTop = True
          'Loop through to the left and above the cell to find the first non-numeric cell
2427      While (NotFoundStringLeft And i > 0) Or (NotFoundStringTop And j > 0)
              'Find the nearest name to the left of the variable or constraint if one exists
2428          If i > 0 And NotFoundStringLeft Then
2429              CellValue = CStr(searchSheet.Cells(row, i).value)
                  'x = IsAmericanNumber(CellValue)
2430              If Not (IsNumeric(CellValue) Or (Left(CellValue, 1) = "=") Or (CellValue = "")) Then
2431                  NotFoundStringLeft = False
2432              End If
2433              i = i - 1
2434          End If
              'Find the nearest name above the variable or constraint if it exists
2435          If j > 0 And NotFoundStringTop Then
2436              CellValue = searchSheet.Cells(j, col)
2437              If Not (IsNumeric(CellValue) Or (Left(CellValue, 1) = "=") Or (CellValue = "")) Then
2438                  NotFoundStringTop = False
2439              End If
2440              j = j - 1
2441          End If
2442      Wend
2443      LHSName = CStr(searchSheet.Cells(row, i + 1).value)
2444      AboveName = CStr(searchSheet.Cells(j + 1, col).value)
          
          'Put the names together
2445      If AboveName = "" Then
2446          findName = LHSName
2447      ElseIf LHSName = "" Then
2448          findName = AboveName
2449      Else
2450          findName = LHSName & " " & AboveName
2451      End If

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("SolverModelDiff", "findName") Then Resume
          RaiseError = True
          GoTo ExitFunction

End Function

