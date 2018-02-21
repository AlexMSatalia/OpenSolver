VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAutoModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ObjectiveFunctionCellRefersTo As String
Public ObjSense As ObjectiveSenseType
Public DecVarsRange As Range
Public Constraints As Collection  ' CAutoModelConstraint

Public Function BuildModel(sheet As Worksheet, _
                           Optional MinimiseUserInteraction As Boolean = False, _
                           Optional SaveAfterBuilding As Boolean = False)
      Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim ShowModel As Boolean    ' Whether to highlight model on sheet after building
          Dim AskedToShow As Boolean  ' Whether we have already asked to show the model
          
3         FindObjective sheet
          
4         If Len(ObjectiveFunctionCellRefersTo) = 0 Then
5             If Not MinimiseUserInteraction Then
                  Dim frmAutoModel As FAutoModel
6                 Set frmAutoModel = New FAutoModel
                  
7                 Set frmAutoModel.sheet = sheet
8                 frmAutoModel.ObjectiveFunctionCellRefersTo = ObjectiveFunctionCellRefersTo
9                 frmAutoModel.ObjectiveSense = ObjSense
10                frmAutoModel.chkShow.value = SaveAfterBuilding
11                frmAutoModel.chkShow.Visible = SaveAfterBuilding
                  
12                frmAutoModel.Show
                  
13                If frmAutoModel.Tag = "Cancelled" Then
14                    Unload frmAutoModel
15                    GoTo ExitFunction
16                End If
                  
17                ObjectiveFunctionCellRefersTo = frmAutoModel.ObjectiveFunctionCellRefersTo
18                ObjSense = frmAutoModel.ObjectiveSense
19                ShowModel = frmAutoModel.chkShow.value
20                AskedToShow = True
                  
21                Unload frmAutoModel
22            End If
23        End If
          
24        If Not FindVarsAndCons(sheet) Then
25            If Not MinimiseUserInteraction Then MsgBox "Error while looking for variables and constraints"
26            GoTo ExitFunction
27        End If
          
28        If SaveAfterBuilding Then
29            SaveModel sheet
              
30            If MinimiseUserInteraction Then
31                ShowModel = True
32            ElseIf Not AskedToShow Then
33                If MsgBox("Automodel done! Show model?", vbYesNo, "OpenSolver - AutoModel") = vbYes Then ShowModel = True
34            End If
          
35            If ShowModel Then
36                OpenSolverVisualizer.ShowSolverModel sheet, HandleError:=True
37            Else
38                OpenSolverVisualizer.HideSolverModel sheet
39            End If
40        End If
          
41        BuildModel = True

ExitFunction:
42        If RaiseError Then RethrowError
43        Exit Function

ErrorHandler:
44        If Not ReportError("CAutoModel", "Build") Then Resume
45        RaiseError = True
46        GoTo ExitFunction
End Function

Private Sub FindObjective(sheet As Worksheet)
      ' Tries to find the objective function cell and sense by searching for likely keywords,
      ' then searching the area for appropriate calculations.
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim ObjSenseCell As Range
3         Set ObjSenseCell = Nothing
4         ObjSense = UnknownObjectiveSense
          
5         UpdateStatusBar "OpenSolver: Trying to determine objective sense...", True
6         Application.Cursor = xlWait

          Dim ObjKeyword As Variant
7         For Each ObjKeyword In StringArray("min", "minimise", "minimize", "max", "maximise", "maximize", _
                                             "min:", "minimise:", "minimize:", "max:", "maximise:", "maximize:")
8             Set ObjSenseCell = FindObjSense(ObjKeyword, sheet)
9             If Not ObjSenseCell Is Nothing Then
10                ObjSense = ObjectiveSenseStringToEnum(Replace(ObjKeyword, ":", ""))
11                Exit For
12            End If
13        Next ObjKeyword

          ' If we didn't find anything, give up here and report failure
14        If ObjSense = UnknownObjectiveSense Then GoTo ExitSub
          
15        UpdateStatusBar "OpenSolver: Found objective sense, looking for objective cell...", True

          ' Search for objective function cell
          Dim SearchFormula As Variant, RowOffsetVar As Variant, RowOffset As Long, ObjRange As Range
16        For Each SearchFormula In StringArray("sumproduct", "=") ' Look for sumproduct first, followed by any formula
17            For Each RowOffsetVar In Array(0, -1, 1)  ' Search current row, then above, then below
18                RowOffset = CLng(RowOffsetVar)
19                If ObjSenseCell.row + RowOffset > 0 Then
20                    Set ObjRange = FindObjCell(ObjSenseCell.row + RowOffset, SearchFormula, sheet)
21                    ObjectiveFunctionCellRefersTo = RangeToRefersTo(ObjRange)
22                    If Not ObjRange Is Nothing Then GoTo ExitSearchLoop
23                End If
24            Next RowOffsetVar
25        Next SearchFormula
          
ExitSearchLoop:
          
ExitSub:
26        Application.Cursor = xlDefault
27        Application.StatusBar = False
28        If RaiseError Then RethrowError
29        Exit Sub

ErrorHandler:
30        If Not ReportError("CAutoModel", "FindObjective") Then Resume
31        RaiseError = True
32        GoTo ExitSub
End Sub

Private Function FindObjSense(ByVal searchStr As String, sheet As Worksheet) As Range
      ' Run the right kind of search to find the objective sense (search values, don't match case)
1         Set FindObjSense = sheet.Cells.Find(What:=searchStr, After:=[a1], LookIn:=xlValues, _
                                              SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
End Function

Private Function FindObjCell(ByVal rowNum As Long, ByVal searchStr As String, sheet As Worksheet) As Range
      ' Run the right kind of search to find objective cell (look in specified row, search formulas, don't match case)
1         Set FindObjCell = sheet.Rows(rowNum).Find(What:=searchStr, LookIn:=xlFormulas, lookat:=xlPart, _
                                                    SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
End Function

' We have objective, now find all constraints.
Private Function FindVarsAndCons(sheet As Worksheet) As Boolean
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          ' Look for constraints and add them if they seem at all interesting (i.e. LHS or RHS has precedents)
3         UpdateStatusBar "OpenSolver:  Looking for constraints", True
4         On Error GoTo ConstraintErr

          Dim FoundLEQ As Range, FoundGEQ As Range, FoundEQ As Range
5         FindAllCells "<=", FoundLEQ, sheet
6         FindAllCells ">=", FoundGEQ, sheet
7         FindAllCells "=", FoundEQ, sheet
              
          ' Combine them as much as possible
          Dim AllCompOps As Range
8         Set AllCompOps = FoundEQ
9         Set AllCompOps = ProperUnion(AllCompOps, FoundLEQ)
10        Set AllCompOps = ProperUnion(AllCompOps, FoundGEQ)
         
          ' Now look for constraint cells
11        Set Constraints = New Collection
          
          Dim Area As Range
12        For Each Area In AllCompOps.Areas
              ' Determine the shape of the area
              Dim RowCount As Long, ColCount As Long
13            RowCount = Area.Rows.Count
14            ColCount = Area.Columns.Count
                 
              ' Depending on the shape, search differently
              Dim LHSs As Range, RHSs As Range
15            If ColCount = 1 Then
                  ' Vertical or singleton block of relations, search left and right for cells
16                Set LHSs = Area.Offset(0, -1)
17                Set RHSs = Area.Offset(0, 1)
18                If CheckPrecedentCells(LHSs, RHSs) Then
19                    AddRangeToConstraints LHSs, Area, RHSs, True, Constraints
20                    GoTo NextArea
21                End If
22            End If
23            If RowCount = 1 Then
                  ' Horizontal or singleton block of relations, search up and down for cells
24                Set LHSs = Area.Offset(-1, 0)
25                Set RHSs = Area.Offset(1, 0)
26                If CheckPrecedentCells(LHSs, RHSs) Then
27                    AddRangeToConstraints LHSs, Area, RHSs, False, Constraints
28                    GoTo NextArea
29                End If
30            End If
              ' If here, we have a block of relations (or a failed search)
              ' TODO - Handle this somehow, if it has an application
NextArea:
31        Next Area
          
          ' Use precedents of objective function and constraints to find the set of possible decision variables
32        UpdateStatusBar "OpenSolver: Searching for decision variables", True
          
          Dim DecRefCount As Dictionary
33        Set DecRefCount = New Dictionary
          
          ' Objective function precedents
34        UpdatePrecedentCount DecRefCount, GetRefersToRange(ObjectiveFunctionCellRefersTo)

          ' Constraint precedents
          Dim curConstraint As CAutoModelConstraint
35        For Each curConstraint In Constraints
36            UpdatePrecedentCount DecRefCount, curConstraint.LHS
37            UpdatePrecedentCount DecRefCount, curConstraint.RHS
38        Next
          
39        On Error GoTo ErrorHandler
40        UpdateStatusBar "OpenSolver: Selecting most likely decision variables", True
          ' If a cell has only been referenced once, we can't be sure it is a decision variable
          ' as constants are also referenced once, so take anything that is seen two or more times
          Dim addressKey As Variant
41        For Each addressKey In DecRefCount.Keys
42            If DecRefCount.Item(CStr(addressKey)) >= 2 Then
43                Set DecVarsRange = ProperUnion(DecVarsRange, sheet.Range(CStr(addressKey)))
44            End If
45        Next
              
          ' Look for type restrictions on decision variables
46        UpdateStatusBar "OpenSolver: Looking for variable type restrictions", True

          Dim CurDecVar As Range, PossibleType As String, VarTypeKeyword As Variant
47        For Each CurDecVar In DecVarsRange
              ' Look below it to see if there is type information
48            PossibleType = LCase(Trim(CurDecVar.Offset(1, 0).value))
49            For Each VarTypeKeyword In Array("integer", "int", "i", "binary", "bin", "b")  ' Keywords that indicate variable type
50                If PossibleType = VarTypeKeyword Then
51                    AddConstraintToModel Constraints, CurDecVar, RelationStringToEnum(VarTypeKeyword)
52                    Exit For
53                End If
54            Next
55        Next
          
          ' Combine adjacent constraints of the same type
56        UpdateStatusBar "OpenSolver: Rationalising constraints", True
57        RationaliseConstraints Constraints
          
          ' Finished!
58        FindVarsAndCons = True

ExitFunction:
59        Application.StatusBar = False
60        If RaiseError Then RethrowError
61        Exit Function

ErrorHandler:
62        If Not ReportError("CAutoModel", "FindVarsAndCons") Then Resume
63        RaiseError = True
64        GoTo ExitFunction
          
DecisionErr:
          ' Error occurred while trying to find decision variables
65        MsgBox "Error: an issue arose while finding decision variables." + vbNewLine + _
                 "Error number:" + str(Err.Number) + vbNewLine + _
                 "Error description: " + Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")"), _
                 vbExclamation Or vbOKOnly, "AutoModel"
66        FindVarsAndCons = False
67        GoTo ExitFunction
          
ConstraintErr:
          ' Error occurred while trying to find constraints
68        MsgBox "Error: an issue arose while finding constraints." + vbNewLine + _
                 "Error number:" + str(Err.Number) + vbNewLine + _
                 "Error description: " + Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")"), _
                 vbExclamation Or vbOKOnly, "AutoModel"
69        FindVarsAndCons = False
70        GoTo ExitFunction
End Function

' Increase precedent count by 1 for each precedent in the child cell
Private Sub UpdatePrecedentCount(ByRef PrecedentCount As Dictionary, ByRef ParentCell As Range)
          Dim RaiseError As Boolean
1         RaiseError = False
          
2         On Error Resume Next
          Dim ChildCell As Range
3         Set ChildCell = ParentCell.Precedents

4         On Error GoTo ErrorHandler
          Dim CurPrecedent As Range
5         If Not ChildCell Is Nothing Then
6             For Each CurPrecedent In ChildCell.Cells
7                 If PrecedentCount.Exists(CurPrecedent.Address) Then
8                     PrecedentCount.Item(CurPrecedent.Address) = PrecedentCount.Item(CurPrecedent.Address) + 1
9                 Else
10                    If Not CurPrecedent.HasFormula Then
11                        PrecedentCount.Add Item:=1, Key:=CurPrecedent.Address
12                    End If
13                End If
14            Next
15        End If

ExitSub:
16        If RaiseError Then RethrowError
17        Exit Sub

ErrorHandler:
18        If Not ReportError("CAutoModel", "UpdatePrecedentCount") Then Resume
19        RaiseError = True
20        GoTo ExitSub
End Sub

' Look for all cells in the sheet containing the search string (only in the value)
' Returns a range of these cells (may contain multiple areas)
Private Sub FindAllCells(ByVal searchStr As String, ByRef FoundCells As Range, ByRef sheet As Worksheet)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim LastCell As Range, FirstCell As Range
3         Set FoundCells = Nothing
          
          ' Find first cell that meets requirements
4         Set FirstCell = sheet.Cells.Find(What:=searchStr, After:=[a1], LookIn:=xlValues, _
                                           SearchOrder:=XlSearchOrder.xlByRows, _
                                           lookat:=XlLookAt.xlWhole, _
                                           SearchDirection:=XlSearchDirection.xlNext)
5         Set LastCell = FirstCell
6         If LastCell Is Nothing Then GoTo ExitSub ' If not even one, stop immediately
          
7         Do
8             Set FoundCells = ProperUnion(FoundCells, LastCell)
              ' Find next
9             Set LastCell = sheet.Cells.FindNext(LastCell)
              ' Loop until no more cells or we get back to the initial cell
10        Loop While (Not LastCell Is Nothing) And (FirstCell.Address <> LastCell.Address)

ExitSub:
11        If RaiseError Then RethrowError
12        Exit Sub

ErrorHandler:
13        If Not ReportError("CAutoModel", "FindAllCells") Then Resume
14        RaiseError = True
15        GoTo ExitSub
End Sub

' Determine if any of the LHS or RHS have a precedent
Private Function CheckPrecedentCells(ByRef LHSs As Range, ByRef RHSs As Range) As Boolean
          Dim CurCell As Range, PrecCells As Range
          Dim BothSides As Range
1         Set BothSides = Union(LHSs, RHSs)

2         For Each CurCell In BothSides.Cells
              ' If no precedents, error is thrown
3             Err.Clear
4             On Error Resume Next
5             Set PrecCells = CurCell.Precedents
6             If Err.Number = 0 Then
                  ' There is a precedent
7                 CheckPrecedentCells = True
8                 Exit Function
9             End If
10        Next
End Function

Private Sub AddRangeToConstraints(ByRef LHSs As Range, ByRef RelRange As Range, ByRef RHSs As Range, _
                                  IsVertical As Boolean, ByRef Constraints As Collection)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim CellCount As Long
3         CellCount = LHSs.Count
          
          Dim i As Long, LHSi As Range, RELi As Range, RHSi As Range
4         For i = 1 To CellCount
5             If IsVertical Then
6                 Set LHSi = LHSs(RowIndex:=i)
7                 Set RELi = RelRange(RowIndex:=i)
8                 Set RHSi = RHSs(RowIndex:=i)
9             Else
10                Set LHSi = LHSs(ColumnIndex:=i)
11                Set RELi = RelRange(ColumnIndex:=i)
12                Set RHSi = RHSs(ColumnIndex:=i)
13            End If

14            If Not TestKeyExists(Constraints, RELi.Address) Then
15                AddConstraintToModel Constraints, LHSi, RelationStringToEnum(RELi.value), RELi, RHSi
16            End If
17        Next i

ExitSub:
18        If RaiseError Then RethrowError
19        Exit Sub

ErrorHandler:
20        If Not ReportError("CAutoModel", "AddRangeToConstraints") Then Resume
21        RaiseError = True
22        GoTo ExitSub
End Sub

' Adds a single constraint, rather than a block
Private Sub AddConstraintToModel(constraintGroup As Collection, newLHS As Range, newType As RelationConsts, Optional newRelationCell As Range, Optional newRHS As Range)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim NewConstraint As New CAutoModelConstraint
3         NewConstraint.Init newLHS, newType, newRelationCell, newRHS
4         If NewConstraint.KeyCell Is Nothing Then
5             constraintGroup.Add NewConstraint
6         Else
7             constraintGroup.Add NewConstraint, NewConstraint.Key
8         End If

ExitSub:
9         If RaiseError Then RethrowError
10        Exit Sub

ErrorHandler:
11        If Not ReportError("CAutoModel", "AddConstraint") Then Resume
12        RaiseError = True
13        GoTo ExitSub
          
End Sub

' Group multiple individual constraints into 1 constraint if:
'   - They are next to each other
'   - They are of the same type
Private Sub RationaliseConstraints(ByRef Constraints As Collection)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim NewConstraints As Collection
3         Set NewConstraints = New Collection
          
          Dim CurrentRelation As RelationConsts
4         For CurrentRelation = RelationConsts.[_First] To RelationConsts.[_Last]
                
              Dim curCon As CAutoModelConstraint, UnionRange As Range
5             Set UnionRange = Nothing
6             For Each curCon In Constraints
7                 If curCon.RelationType = CurrentRelation Then
8                     Set UnionRange = ProperUnion(UnionRange, curCon.KeyCell)
9                 End If
10            Next curCon

11            If Not UnionRange Is Nothing Then
                  ' Now iterate through each area of the range - each represents a block
                  ' of constraints that are next to each other, with the same relation
                  Dim Area As Range
12                For Each Area In UnionRange.Areas
                      Dim LHSunion As Range, RHSunion As Range, RELunion As Range
13                    Set LHSunion = Nothing
14                    Set RHSunion = Nothing
15                    Set RELunion = Nothing

                      Dim CurCell As Range
16                    For Each CurCell In Area.Cells
17                        Set curCon = Constraints(CurCell.Address)
18                        Set LHSunion = ProperUnion(LHSunion, curCon.LHS)
19                        Set RHSunion = ProperUnion(RHSunion, curCon.RHS)
20                        Set RELunion = ProperUnion(RELunion, curCon.RelationCell)
21                    Next
22                    AddConstraintToModel NewConstraints, LHSunion, CurrentRelation, RELunion, RHSunion
23                Next Area
24            End If
25        Next CurrentRelation
          
          ' Update old constraints
26        Set Constraints = NewConstraints

ExitSub:
27        If RaiseError Then RethrowError
28        Exit Sub

ErrorHandler:
29        If Not ReportError("CAutoModel", "RationaliseConstraints") Then Resume
30        RaiseError = True
31        GoTo ExitSub
End Sub

Private Sub SaveModel(sheet As Worksheet)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' Check int/bin constraints not set on non-decision variables before we start saving things
          Dim curCon As CAutoModelConstraint, ShowWarning As Boolean
3         ShowWarning = False
4         For Each curCon In Constraints
5             If Not RelationHasRHS(curCon.RelationType) Then
6                 If Not SetDifference(curCon.LHS, DecVarsRange) Is Nothing Then
7                     If MsgBox("This model has specified that a non-decision cell must take an integer/binary value. " & _
                                "This is a valid model in some versions of Solver, but not one that OpenSolver can solve. " & _
                                "Do you wish to continue with saving this model?", _
                                vbQuestion + vbYesNo, "OpenSolver - Warning") = vbYes Then
8                         Exit For
9                     Else
10                        GoTo ExitSub
11                    End If
12                End If
13            End If
14        Next curCon

          ' Now save model
15        SetDecisionVariables DecVarsRange, sheet
16        SetObjectiveFunctionCellRefersTo ObjectiveFunctionCellRefersTo, sheet
17        SetObjectiveSense ObjSense, sheet
18        SetNumConstraints Constraints.Count, sheet
          Dim c As Long
19        For c = 1 To Constraints.Count
20            Set curCon = Constraints(c)
21            With curCon
22                UpdateConstraint c, .LHS, .RelationType, .RHS, "", sheet
23            End With
24        Next c

ExitSub:
25        If RaiseError Then RethrowError
26        Exit Sub

ErrorHandler:
27        If Not ReportError("CAutoModel", "SaveModel") Then Resume
28        RaiseError = True
29        GoTo ExitSub
End Sub
