Attribute VB_Name = "OpenSolverAutoModel"
Option Explicit

Public Function RunAutoModel(Optional MinimiseUserInteraction As Boolean = False, Optional ByRef InputModel As CModel) As Boolean
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    If Not CheckWorksheetAvailable Then GoTo ExitFunction
    Dim model As CModel, AskedToShow As Boolean, ShowModel As Boolean, DoBuild As Boolean
    If InputModel Is Nothing Then
        Set model = New CModel
        DoBuild = True
    Else
        Set model = InputModel
        DoBuild = False
    End If
    ShowModel = False
    AskedToShow = False
    
    Dim ObjectiveSense As ObjectiveSenseType, ObjectiveFunctionCell As Range
    FindObjective ActiveSheet, ObjectiveSense, ObjectiveFunctionCell
    Set model.ObjectiveFunctionCell = ObjectiveFunctionCell
    model.ObjectiveSense = ObjectiveSense
    
    If model.ObjectiveFunctionCell Is Nothing Then
        If Not MinimiseUserInteraction Then
            Dim frmAutoModel As FAutoModel
            Set frmAutoModel = New FAutoModel
            Set frmAutoModel.ObjectiveCell = model.ObjectiveFunctionCell
            frmAutoModel.ObjectiveSense = model.ObjectiveSense
            frmAutoModel.chkShow.value = DoBuild
            frmAutoModel.chkShow.Visible = DoBuild
            
            frmAutoModel.Show
            
            If frmAutoModel.Tag = "Cancelled" Then
                Unload frmAutoModel
                GoTo ExitFunction
            End If
            
            Set model.ObjectiveFunctionCell = frmAutoModel.ObjectiveCell
            model.ObjectiveSense = frmAutoModel.ObjectiveSense
            ShowModel = frmAutoModel.chkShow.value
            AskedToShow = True
            Unload frmAutoModel
        End If
    End If
    
    Dim SearchSheet As Worksheet
    If Not model.ObjectiveFunctionCell Is Nothing Then
        Set SearchSheet = model.ObjectiveFunctionCell.Parent
    Else
        Set SearchSheet = ActiveSheet
    End If
    
    Dim DecisionVariables As Range, Constraints As Collection
    If Not FindVarsAndCons(SearchSheet, model.ObjectiveFunctionCell, DecisionVariables, Constraints, True) Then
        If Not MinimiseUserInteraction Then MsgBox "Error while looking for variables and constraints"
        RunAutoModel = False
        GoTo ExitFunction
    End If
    Set model.DecisionVariables = DecisionVariables
    Set model.Constraints = Constraints
    
    model.NonNegativityAssumption = True
    
    If DoBuild Then
        model.BuildModel
        
        If MinimiseUserInteraction Then
            ShowModel = True
        ElseIf Not AskedToShow Then
            If MsgBox("Automodel done! Show model?", vbYesNo, "OpenSolver - AutoModel") = vbYes Then ShowModel = True
        End If
    
        If ShowModel Then
            OpenSolverVisualizer.ShowSolverModel
        Else
            OpenSolverVisualizer.HideSolverModel
        End If
    End If
    
    RunAutoModel = True

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverAutoModel", "RunAutoModel") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

' Tries to find the objective function cell and sense by searching for likely keywords,
' then searching the area for appropriate calculations.
Public Sub FindObjective(ByRef s As Worksheet, ByRef ObjectiveSense As ObjectiveSenseType, ByRef ObjectiveFunctionCell As Range)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim ObjSenseCell As Range
3586      Set ObjSenseCell = Nothing
3587      ObjectiveSense = UnknownObjectiveSense
          
3588      UpdateStatusBar "OpenSolver: Trying to determine objective sense...", True
3589      Application.Cursor = xlWait

          Dim ObjKeyword As Variant
          For Each ObjKeyword In StringArray("min", "minimise", "minimize", "max", "maximise", "maximize")
              FindObjSense s, ObjKeyword, ObjSenseCell
              If Not (ObjSenseCell Is Nothing) Then
                  ObjectiveSense = ObjectiveSenseStringToEnum(ObjKeyword)
                  Exit For
              End If
          Next ObjKeyword

          ' If we didn't find anything, give up here and report failure
3614      If ObjectiveSense = UnknownObjectiveSense Then GoTo ExitSub
          
3621      UpdateStatusBar "OpenSolver: Found objective sense, looking for objective cell...", True

          ' Search for objective function cell
          Dim SearchFormula As Variant, RowOffsetVar As Variant, RowOffset As Long
          For Each SearchFormula In StringArray("sumproduct", "=") ' Look for sumproduct first, followed by any formula
              For Each RowOffsetVar In Array(0, -1, 1)  ' Search current row, then above, then below
                  RowOffset = CLng(RowOffsetVar)
                  If ObjSenseCell.row + RowOffset > 0 Then
                      FindObjCell s, ObjSenseCell.row + RowOffset, SearchFormula, ObjectiveFunctionCell
                      If Not (ObjectiveFunctionCell Is Nothing) Then GoTo ExitSub
                  End If
              Next RowOffsetVar
          Next SearchFormula

ExitSub:
3648      Application.Cursor = xlDefault
3649      Application.StatusBar = False
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("OpenSolverAutoModel", "FindObjective") Then Resume
          RaiseError = True
          GoTo ExitSub
          
End Sub

' Run the right kind of search to find the objective sense (search values, don't match case)
Private Sub FindObjSense(ByRef s As Worksheet, ByVal searchStr As String, ByRef result As Range)
3650      Set result = s.Cells.Find(What:=searchStr, After:=[a1], LookIn:=xlValues, _
                                    SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
End Sub

' Run the right kind of search to find objective cell (look in specified row, search formulas, don't match case)
Private Sub FindObjCell(ByRef s As Worksheet, ByVal rowNum As Long, ByVal searchStr As String, ByRef result As Range)
3651      Set result = s.Rows(rowNum).Find(What:=searchStr, LookIn:=xlFormulas, lookat:=xlPart, _
                                           SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
End Sub

' We have objective, now find all constraints.
Public Function FindVarsAndCons(ByRef SearchSheet As Worksheet, ByRef ObjectiveFunctionCell As Range, ByRef DecisionVariables As Range, ByRef Constraints As Collection, IsFirstTime As Boolean) As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          ' Clear existing solution, if requested
3653      If IsFirstTime Then
3654          Set DecisionVariables = Nothing
3655          Set Constraints = New Collection
3656      End If
          
          ' Look for constraints and add them if they seem at all interesting (i.e. LHS or RHS has precedents)
3657      UpdateStatusBar "OpenSolver:  Looking for constraints", True
3658      On Error GoTo ConstraintErr

          Dim FoundLEQ As Range, FoundGEQ As Range, FoundEQ As Range
          FindAllCells "<=", FoundLEQ, SearchSheet
          FindAllCells ">=", FoundGEQ, SearchSheet
          FindAllCells "=", FoundEQ, SearchSheet
              
          ' Combine them as much as possible
          Dim AllCompOps As Range
3668      Set AllCompOps = FoundEQ
3669      Set AllCompOps = ProperUnion(AllCompOps, FoundLEQ)
3670      Set AllCompOps = ProperUnion(AllCompOps, FoundGEQ)
         
          ' Now look for constraint cells
          Dim Area As Range
3671      For Each Area In AllCompOps.Areas
              ' Determine the shape of the area
              Dim RowCount As Long, ColCount As Long
3672          RowCount = Area.Rows.Count
3673          ColCount = Area.Columns.Count
                 
              ' Depending on the shape, search differently
              Dim LHSs As Range, RHSs As Range
3674          If ColCount = 1 Then
                  ' Vertical or singleton block of relations, search left and right for cells
3675              Set LHSs = Area.Offset(0, -1)
3676              Set RHSs = Area.Offset(0, 1)
3677              If CheckPrecedentCells(LHSs, RHSs) Then
3678                  AddRangeToConstraints LHSs, Area, RHSs, True, Constraints
                      GoTo NextArea
3679              End If
              End If
3680          If RowCount = 1 Then
                  ' Horizontal or singleton block of relations, search up and down for cells
3681              Set LHSs = Area.Offset(-1, 0)
3682              Set RHSs = Area.Offset(1, 0)
3683              If CheckPrecedentCells(LHSs, RHSs) Then
3684                  AddRangeToConstraints LHSs, Area, RHSs, False, Constraints
                      GoTo NextArea
3685              End If
3698          End If
              ' If here, we have a block of relations (or a failed search)
              ' TODO - Handle this somehow, if it has an application
NextArea:
3700      Next Area
    
          ' Use precedents of objective function and constraints to find the set of possible decision variables
3702      UpdateStatusBar "OpenSolver: Searching for decision variables", True
          
          Dim DecRefCount As Dictionary
          Set DecRefCount = New Dictionary
          
          ' Objective function precedents
          UpdatePrecedentCount DecRefCount, ObjectiveFunctionCell

          ' Constraint precedents
          Dim curConstraint As CConstraint
3719      For Each curConstraint In Constraints
3723          UpdatePrecedentCount DecRefCount, curConstraint.LHS
              UpdatePrecedentCount DecRefCount, curConstraint.RHS
3752      Next
          
          On Error GoTo ErrorHandler
3753      UpdateStatusBar "OpenSolver: Selecting most likely decision variables", True
          ' If a cell has only been referenced once, we can't be sure it is a decision variable
          ' as constants are also referenced once, so take anything that is seen two or more times
          Dim addressKey As Variant
3754      For Each addressKey In DecRefCount.Keys
3755          If DecRefCount.Item(CStr(addressKey)) >= 2 Then
3756              Set DecisionVariables = ProperUnion(DecisionVariables, ActiveSheet.Range(CStr(addressKey)))
3757          End If
3758      Next
              
          ' Look for type restrictions on decision variables
3759      UpdateStatusBar "OpenSolver: Looking for variable type restrictions", True

          Dim CurDecVar As Range, PossibleType As String, VarTypeKeyword As Variant
3760      For Each CurDecVar In DecisionVariables
              ' Look below it to see if there is type information
3761          PossibleType = LCase(Trim(CurDecVar.Offset(1, 0).value))
              For Each VarTypeKeyword In Array("integer", "int", "i", "binary", "bin", "b")  ' Keywords that indicate variable type
                  If PossibleType = VarTypeKeyword Then
3763                  AddConstraintToModel Constraints, CurDecVar, RelationStringToEnum(VarTypeKeyword)
                      Exit For
                  End If
              Next
3781      Next
          
          ' Combine adjacent constraints of the same type
3782      UpdateStatusBar "OpenSolver: Rationalising constraints", True
          RationaliseConstraints Constraints
          
          ' Finished!
3783      FindVarsAndCons = True

ExitFunction:
          Application.StatusBar = False
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverAutoModel", "FindVarsAndCons") Then Resume
          RaiseError = True
          GoTo ExitFunction
          
DecisionErr:
          ' Error occurred while trying to find decision variables
3786      MsgBox "Error: an issue arose while finding decision variables." + vbNewLine + _
                 "Error number:" + str(Err.Number) + vbNewLine + _
                 "Error description: " + Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")"), _
                 vbExclamation Or vbOKOnly, "AutoModel"
3787      FindVarsAndCons = False
3788      GoTo ExitFunction
          
ConstraintErr:
          ' Error occurred while trying to find constraints
3790      MsgBox "Error: an issue arose while finding constraints." + vbNewLine + _
                 "Error number:" + str(Err.Number) + vbNewLine + _
                 "Error description: " + Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")"), _
                 vbExclamation Or vbOKOnly, "AutoModel"
3791      FindVarsAndCons = False
3792      GoTo ExitFunction
End Function

' Increase precedent count by 1 for each precedent in the child cell
Sub UpdatePrecedentCount(ByRef PrecedentCount As Dictionary, ByRef ParentCell As Range)
    Dim RaiseError As Boolean
    RaiseError = False
    
    On Error Resume Next
    Dim ChildCell As Range
    Set ChildCell = ParentCell.Precedents

    On Error GoTo ErrorHandler
    Dim CurPrecedent As Range
    If Not ChildCell Is Nothing Then
        For Each CurPrecedent In ChildCell.Cells
            If PrecedentCount.Exists(CurPrecedent.Address) Then
                PrecedentCount.Item(CurPrecedent.Address) = PrecedentCount.Item(CurPrecedent.Address) + 1
            Else
                If Not CurPrecedent.HasFormula Then
                    PrecedentCount.Add Item:=1, Key:=CurPrecedent.Address
                End If
            End If
        Next
    End If

ExitSub:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Sub

ErrorHandler:
    If Not ReportError("OpenSolverAutoModel", "UpdatePrecedentCount") Then Resume
    RaiseError = True
    GoTo ExitSub
End Sub

' Look for all cells in the sheet containing the search string (only in the value)
' Returns a range of these cells (may contain multiple areas)
Private Sub FindAllCells(ByVal searchStr As String, ByRef FoundCells As Range, ByRef sheet As Worksheet)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim LastCell As Range, FirstCell As Range
3794      Set FoundCells = Nothing
          
          ' Find first cell that meets requirements
3795      Set FirstCell = sheet.Cells.Find(What:=searchStr, After:=[a1], LookIn:=xlValues, _
                                           SearchOrder:=XlSearchOrder.xlByRows, _
                                           lookat:=XlLookAt.xlWhole, _
                                           SearchDirection:=XlSearchDirection.xlNext)
3796      Set LastCell = FirstCell
3797      If LastCell Is Nothing Then GoTo ExitSub ' If not even one, stop immediately
          
3798      Do
3802          Set FoundCells = ProperUnion(FoundCells, LastCell)
3803          ' Find next
3804          Set LastCell = sheet.Cells.FindNext(LastCell)
              ' Loop until no more cells or we get back to the initial cell
3805      Loop While (Not LastCell Is Nothing) And (FirstCell.Address <> LastCell.Address)

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("OpenSolverAutoModel", "FindAllCells") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

' Determine if any of the LHS or RHS have a precedent
Function CheckPrecedentCells(ByRef LHSs As Range, ByRef RHSs As Range) As Boolean
          Dim CurCell As Range, PrecCells As Range
          Dim BothSides As Range
3806      Set BothSides = Union(LHSs, RHSs)

3807      For Each CurCell In BothSides.Cells
              ' If no precedents, error is thrown
3808          Err.Clear
3809          On Error Resume Next
3810          Set PrecCells = CurCell.Precedents
3811          If Err.Number = 0 Then
                  ' There is a precedent
3813              CheckPrecedentCells = True
3814              Exit Function
3815          End If
3816      Next
End Function

Sub AddRangeToConstraints(ByRef LHSs As Range, ByRef RelRange As Range, ByRef RHSs As Range, _
                          IsVertical As Boolean, ByRef Constraints As Collection)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim CellCount As Long
3817      CellCount = LHSs.Count
          
          Dim i As Long
          Dim LHSi As Range, RELi As Range, RHSi As Range
          Dim NewConstraint As CConstraint
          
3818      For i = 1 To CellCount
3819          If IsVertical Then
3820              Set LHSi = LHSs(RowIndex:=i)
3821              Set RELi = RelRange(RowIndex:=i)
3822              Set RHSi = RHSs(RowIndex:=i)
3823          Else
3824              Set LHSi = LHSs(ColumnIndex:=i)
3825              Set RELi = RelRange(ColumnIndex:=i)
3826              Set RHSi = RHSs(ColumnIndex:=i)
3827          End If

3828          If Not TestKeyExists(Constraints, RELi.Address) Then
3829              AddConstraintToModel Constraints, LHSi, RelationStringToEnum(RELi.value), RELi, RHSi
3838          End If
3839      Next i

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("OpenSolverAutoModel", "AddRangeToConstraints") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

' Adds a single constraint, rather than a block
Sub AddConstraintToModel(constraintGroup As Collection, newLHS As Range, newType As RelationConsts, Optional newRelationCell As Range, Optional newRHS As Range, Optional newRHSstring As String)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim NewConstraint As New CConstraint
3975      NewConstraint.Init newLHS, newType, newRelationCell, newRHS, newRHSstring
          If NewConstraint.KeyCell Is Nothing Then
              constraintGroup.Add NewConstraint
          Else
3982          constraintGroup.Add NewConstraint, NewConstraint.Key
          End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("OpenSolverAutoModel", "AddConstraint") Then Resume
          RaiseError = True
          GoTo ExitSub
          
End Sub

' Group multiple individual constraints into 1 constraint if:
'   - They are next to each other
'   - They are of the same type
Public Sub RationaliseConstraints(ByRef Constraints As Collection)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim NewConstraints As Collection
          Set NewConstraints = New Collection
          
          Dim CurrentRelation As RelationConsts
3848      For CurrentRelation = RelationConsts.[_First] To RelationConsts.[_Last]
                
              Dim curCon As CConstraint, UnionRange As Range
3871          Set UnionRange = Nothing
3855          For Each curCon In Constraints
3856              If curCon.RelationType = CurrentRelation Then
3857                  Set UnionRange = ProperUnion(UnionRange, curCon.KeyCell)
3867              End If
3869          Next curCon

              If Not UnionRange Is Nothing Then
                  ' Now iterate through each area of the range - each represents a block
                  ' of constraints that are next to each other, with the same relation
                  Dim Area As Range
3881              For Each Area In UnionRange.Areas
                      Dim LHSunion As Range, RHSunion As Range, RELunion As Range
3883                  Set LHSunion = Nothing
3884                  Set RHSunion = Nothing
3885                  Set RELunion = Nothing

                      Dim CurCell As Range
3886                  For Each CurCell In Area.Cells
                          Set curCon = Constraints(CurCell.Address)
3887                      Set LHSunion = ProperUnion(LHSunion, curCon.LHS)
3888                      Set RHSunion = ProperUnion(RHSunion, curCon.RHS)
3889                      Set RELunion = ProperUnion(RELunion, curCon.RelationCell)
3891                  Next
                      AddConstraintToModel NewConstraints, LHSunion, CurrentRelation, RELunion, RHSunion
3902              Next Area
3903          End If
3904      Next CurrentRelation
          
          ' Update old constraints
3905      Set Constraints = NewConstraints

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("OpenSolverAutoModel", "RationaliseConstraints") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub
