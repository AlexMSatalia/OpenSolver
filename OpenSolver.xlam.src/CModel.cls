VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ObjectiveSense As ObjectiveSenseType
Public ObjectiveFunctionCell As Range
Public ObjectiveTarget As Double

Public DecisionVariables As Range

Public Constraints As Collection

Public Duals As Range   ' The range where we store the duals; specific to OpenSolver

Public NonNegativityAssumption As Boolean

' Tries to find the objective function cell and sense by searching for likely keywords,
' then searching the area for appropriate calculations.
Public Sub FindObjective(ByRef s As Worksheet)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim ObjSenseCell As Range
3586      Set ObjSenseCell = Nothing
3587      ObjectiveSense = UnknownObjectiveSense
          
3588      UpdateStatusBar "OpenSolver: Trying to determine objective sense...", True
3589      Application.Cursor = xlWait

          Dim ObjKeywords() As Variant, ObjSenseMin() As Variant, i As Long
          ObjKeywords = Array("min", "minimise", "minimize", "max", "maximise", "maximize")
          ObjSenseMin = Array(True, True, True, False, False, False)  ' True if corresponding keyword indicates minimisation objective
          
          For i = 0 To UBound(ObjKeywords)
              FindObjSense s, ObjKeywords(i), ObjSenseCell
              If Not (ObjSenseCell Is Nothing) Then
                  ObjectiveSense = IIf(ObjSenseMin(i), MinimiseObjective, MaximiseObjective)
                  Exit For
              End If
          Next

3614      If ObjectiveSense = UnknownObjectiveSense Then
              ' Didn't find anything, give up here and report failure
              GoTo ExitSub
3620      End If
          
3621      UpdateStatusBar "OpenSolver: Found objective sense, looking for objective cell...", True

          ' Search for objective function cell
          Dim SearchFormulae() As Variant, RowOffsets() As Variant, SearchFormula As Variant
          SearchFormulae = Array("sumproduct", "=")  ' Look for sumproduct first, followed by any formula
          RowOffsets = Array(0, -1, 1)  ' Search current row, then above, then below
          
          For Each SearchFormula In SearchFormulae
              For i = 0 To UBound(RowOffsets)
                  If ObjSenseCell.row + i > 0 Then
                      FindObjCell s, ObjSenseCell.row + i, SearchFormula, ObjectiveFunctionCell
                      If Not (ObjectiveFunctionCell Is Nothing) Then GoTo ExitSub
                  End If
              Next
          Next

ExitSub:
3648      Application.Cursor = xlDefault
3649      Application.StatusBar = False
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("CModel", "FindObjective") Then Resume
          RaiseError = True
          GoTo ExitSub
          
End Sub

' Run the right kind of search to find the objective sense (search values, don't match case)
Private Sub FindObjSense(ByRef s As Worksheet, ByVal searchStr As String, ByRef result As Range)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

3650      Set result = s.Cells.Find(What:=searchStr, After:=[a1], LookIn:=xlValues, _
                                    SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("CModel", "FindObjSense") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

' Run the right kind of search to find objective cell (look in specified row, search formulas, don't match case)
Private Sub FindObjCell(ByRef s As Worksheet, ByVal rowNum As Long, ByVal searchStr As String, ByRef result As Range)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

3651      Set result = s.Rows(rowNum).Find(What:=searchStr, LookIn:=xlFormulas, lookat:=xlPart, _
                                           SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("CModel", "FindObjCell") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

' We have objective, now find all constraints.
Public Function FindVarsAndCons(IsFirstTime As Boolean) As Boolean
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
          
          ' Find all the "relation operator" cells on the sheet that the objective function is on.
          Dim FoundLEQ As Range, FoundGEQ As Range, FoundEQ As Range, SearchSheet As Worksheet
          If Not ObjectiveFunctionCell Is Nothing Then
              Set SearchSheet = ObjectiveFunctionCell.Parent
          Else
              Set SearchSheet = ActiveSheet
          End If

          FindAllCells "<=", FoundLEQ, SearchSheet
          FindAllCells ">=", FoundGEQ, SearchSheet
          FindAllCells "=", FoundEQ, SearchSheet
              
          ' Combine them as much as possible
          Dim AllCompOps As Range
3668      Set AllCompOps = FoundEQ
3669      Set AllCompOps = ProperUnion(AllCompOps, FoundLEQ)
3670      Set AllCompOps = ProperUnion(AllCompOps, FoundGEQ)
         
          ' Now look for constraint cells
          Dim AreaIndex As Long
3671      For AreaIndex = 1 To AllCompOps.Areas.Count
              ' Determine the shape of the area
              Dim RowCount As Long, ColCount As Long
3672          RowCount = AllCompOps.Areas(AreaIndex).Rows.Count
3673          ColCount = AllCompOps.Areas(AreaIndex).Columns.Count
                 
              ' Depending on the shape, search differently
              Dim LHSs As Range, RHSs As Range
3674          If ColCount = 1 Then
                  ' Vertical or singleton block of relations, search left and right for cells
3675              Set LHSs = AllCompOps.Areas(AreaIndex).Offset(0, -1)
3676              Set RHSs = AllCompOps.Areas(AreaIndex).Offset(0, 1)
3677              If CheckPrecedentCells(LHSs, RHSs) Then
3678                  AddRangeToConstraints LHSs, AllCompOps.Areas(AreaIndex), RHSs, True, Constraints
                      GoTo NextArea
3679              End If
              End If
3680          If RowCount = 1 Then
                  ' Horizontal or singleton block of relations, search up and down for cells
3681              Set LHSs = AllCompOps.Areas(AreaIndex).Offset(-1, 0)
3682              Set RHSs = AllCompOps.Areas(AreaIndex).Offset(1, 0)
3683              If CheckPrecedentCells(LHSs, RHSs) Then
3684                  AddRangeToConstraints LHSs, AllCompOps.Areas(AreaIndex), RHSs, False, Constraints
                      GoTo NextArea
3685              End If
3698          End If
              ' If here, we have a block of relations (or a failed search)
              ' TODO - Handle this somehow, if it has an application
NextArea:
3700      Next AreaIndex
    
          ' Use precedents of objective function and constraints to find the set of possible decision variables
3702      UpdateStatusBar "OpenSolver: Searching for decision variables", True
          
          Dim DecRefCount As New Dictionary, PrecedentCells As Range
          
          ' Objective function precedents
          If Not ObjectiveFunctionCell Is Nothing Then
3703          On Error Resume Next
3704          Set PrecedentCells = Nothing
3705          Set PrecedentCells = ObjectiveFunctionCell.Precedents
3706          UpdatePrecedentCount DecRefCount, PrecedentCells
          End If

          ' Constraint precedents
3718      On Error Resume Next
          Dim curConstraint As CConstraint
3719      For Each curConstraint In Constraints
3720          On Error Resume Next
3721          Set PrecedentCells = Nothing
3722          Set PrecedentCells = curConstraint.LHS.Precedents
3723          UpdatePrecedentCount DecRefCount, PrecedentCells

3737          On Error Resume Next
3738          Set PrecedentCells = Nothing
3739          Set PrecedentCells = curConstraint.RHS.Precedents
3740          UpdatePrecedentCount DecRefCount, PrecedentCells
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
          Dim CurDecVar As Range, PossibleType As String, i As Long
          Dim VarTypeKeywords() As Variant, VarTypes() As Variant
          VarTypeKeywords = Array("integer", "int", "i", "binary", "bin", "b")  ' Keywords that indicate variable type
          VarTypes = Array("int", "int", "int", "bin", "bin", "bin")  ' Variable type for each keyword
          
3760      For Each CurDecVar In DecisionVariables
              ' Look below it to see if there is type information
3761          PossibleType = LCase(Trim(CurDecVar.Offset(1, 0).value))
              For i = 0 To UBound(VarTypeKeywords)
                  If PossibleType = VarTypeKeywords(i) Then
3763                  Set curConstraint = New CConstraint
3764                  With curConstraint
3765                      Set .LHS = CurDecVar
3766                      Set .RHS = Nothing
3767                      Set .Relation = Nothing
3768                      .ConstraintType = VarTypes(i)
3769                  End With
3770                  Constraints.Add curConstraint
                      Exit For
                  End If
              Next
3781      Next
          
          ' Combine adjacent constraints of the same type
3782      UpdateStatusBar "OpenSolver: Rationalising constraints", True
          RationaliseConstraints
          
          ' Finished!
3783      FindVarsAndCons = True

ExitFunction:
          Application.StatusBar = False
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CModel", "FindVarsAndCons") Then Resume
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
Sub UpdatePrecedentCount(ByRef PrecedentCount As Dictionary, ByRef ChildCell As Range)
    Dim RaiseError As Boolean
    RaiseError = False
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
    If Not ReportError("CModel", "UpdatePrecedentCount") Then Resume
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
3799          If FoundCells Is Nothing Then
3800              Set FoundCells = Range(LastCell.Address)
3801          Else
3802              Set FoundCells = Union(FoundCells, LastCell)
3803          End If
              ' Find next
3804          Set LastCell = sheet.Cells.FindNext(LastCell)
              ' Loop until no more cells or we get back to the initial cell
3805      Loop While (Not LastCell Is Nothing) And (FirstCell.Address <> LastCell.Address)

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("CModel", "FindAllCells") Then Resume
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
                  ' Constraint not already added - so add it!
3829              Set NewConstraint = New CConstraint
3830              With NewConstraint
3831                  Set .LHS = LHSi
3832                  Set .RHS = RHSi
3833                  Set .Relation = RELi
3834                  .ConstraintType = RELi.value
3835              End With
3836              Constraints.Add NewConstraint
3838          End If
3839      Next i

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("CModel", "AddRangeToConstraints") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

' Group multiple individual constraints into 1 constraint if:
'   - They are next to each other
'   - They are of the same type
Public Sub RationaliseConstraints()
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim Relations() As Variant, AddressedByLHS() As Variant
3841      Relations = Array("=", "<=", ">=", "int", "bin")
          AddressedByLHS = Array(False, False, False, True, True)  ' True for relations that only have a LHS
          
          Dim NewConstraints As New Collection
          
          Dim NumConstraints As Long
3846      NumConstraints = Constraints.Count
                    
          Dim conNum() As Variant
3847      ReDim conNum(NumConstraints + 1) As Variant
          
          Dim i As Long, j As Long, CurCell As Range, conKeys As Collection
3848      For i = 0 To UBound(Relations)
3849          Set conKeys = New Collection
                
              Dim c As Long, k As Long
              'Initialise the indices of the constraints to be empty
3850          For k = 1 To NumConstraints
3851              conNum(k) = Empty
3852          Next k
                
3853          c = 1
3854          k = 1
                
              Dim curCon As CConstraint
3855          For Each curCon In Constraints
3856              If curCon.ConstraintType = Relations(i) Then
                      Set CurCell = IIf(AddressedByLHS(i), curCon.LHS, curCon.Relation)
3857                  conKeys.Add CurCell.Address, CurCell.Address
3859                  conNum(k) = c
3860                  k = k + 1
3867              End If
3868              c = c + 1
3869          Next

3870          If conKeys.Count > 0 Then
3872              k = 1
3873              Do While conNum(k) <> ""
3874                  k = k + 1
3875              Loop
                
                  ' Build up a range that is the union of all the constraints
                  Dim UnionRange As Range
3871              Set UnionRange = Nothing
3876              For j = 1 To (k - 1)
                      Set CurCell = IIf(AddressedByLHS(i), Constraints(conNum(j)).LHS, Constraints(conNum(j)).Relation)
3877                  Set UnionRange = ProperUnion(UnionRange, CurCell)
3879              Next j

                  ' Now iterate through each area of the range - each represents a block
                  ' of constraints that are next to each other, with the same relation
                  Dim SavedNum As Long, num As Long, AreaIndex As Long
3880              SavedNum = 1
3881              For AreaIndex = 1 To UnionRange.Areas.Count
                      Dim LHSunion As Range, RHSunion As Range, RELunion As Range
3883                  Set LHSunion = Nothing
3884                  Set RHSunion = Nothing
3885                  Set RELunion = Nothing

3882                  num = SavedNum
3886                  For Each CurCell In UnionRange.Areas(AreaIndex).Cells
3887                      Set LHSunion = ProperUnion(LHSunion, Constraints(conNum(num)).LHS)
3888                      Set RHSunion = ProperUnion(RHSunion, Constraints(conNum(num)).RHS)
3889                      Set RELunion = ProperUnion(RELunion, Constraints(conNum(num)).Relation)
3890                      num = num + 1
3891                  Next
3892                  SavedNum = num

                      Dim NewConstraint As CConstraint
3893                  Set NewConstraint = New CConstraint
3894                  With NewConstraint
3895                      Set .LHS = LHSunion
3896                      Set .RHS = RHSunion
3897                      Set .Relation = RELunion
3898                      .ConstraintType = Relations(i)
3899                  End With

                      Set CurCell = IIf(AddressedByLHS(i), NewConstraint.LHS, NewConstraint.Relation)
3900                  NewConstraints.Add NewConstraint, CurCell.Address
3902              Next
3903          End If
3904      Next i
          
          ' Update old constraints
3905      Set Constraints = NewConstraints

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("CModel", "RationaliseConstraints") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Function BuildModel() As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          ' Check int/bin constraints not set on non-decision variables before we start saving things
          Dim c As Long, ShowWarning As Boolean
          ShowWarning = False
3924      For c = 1 To Constraints.Count
3925          If Constraints(c).ConstraintType = "bin" Or Constraints(c).ConstraintType = "int" Then
3927              If Not SetDifference(Constraints(c).LHS, DecisionVariables) Is Nothing Then
                      If MsgBox("This model has specified that a non-decision cell must take an integer/binary value. " & _
                                "This is a valid model, but not one that OpenSolver can solve. " & _
                                "Do you wish to continue with saving this model?", _
                                vbQuestion + vbYesNo, "OpenSolver - Warning") = vbYes Then
                          GoTo ContinueSaving
                      Else
                          BuildModel = False
                          GoTo ExitFunction
                      End If
                  End If
3934          End If
3935      Next c

ContinueSaving:
3906      SetDecisionVariables DecisionVariables
3911      SetObjectiveFunctionCell ObjectiveFunctionCell
          
          ' Objective type
3916      SetObjectiveSense ObjectiveSense
3918      If ObjectiveSense = TargetObjective Then
3922          SetObjectiveTargetValue ObjectiveTarget
3923      End If
          
          ' Constraints
3936      SetNumConstraints Constraints.Count
3937      For c = 1 To Constraints.Count
              With Constraints(c)
                  UpdateConstraint c, .LHS, RelationStringToEnum(.ConstraintType), .RHS, .RHSstring
              End With
3968      Next c
          
3969      SetDuals Duals
3974      SetNonNegativity NonNegativityAssumption

          BuildModel = True

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CModel", "BuildModel") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Adds a single constraint, rather than a block
Public Sub AddConstraint(newLHS As Range, newRHS As Range, newRHSstring As String, newType As String)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim NewConstraint As New CConstraint
3975      With NewConstraint
3976          Set .LHS = newLHS
3977          Set .Relation = Nothing
3978          Set .RHS = newRHS
3979          .RHSstring = newRHSstring
3980          .ConstraintType = newType
3981      End With
          
3982      Constraints.Add NewConstraint

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("CModel", "AddConstraint") Then Resume
          RaiseError = True
          GoTo ExitSub
          
End Sub

Public Sub PopulateConstraintListBox(ByRef lst As Object, showNameRange As Boolean)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim CS As String, c As Long
          Dim oldLI As Long
          
3984      If showNameRange Then SearchRangeName_DestroyCache
          
3985      oldLI = lst.ListIndex
3986      lst.Clear
3987      lst.AddItem "<Add new constraint>"
3988      For c = 1 To Constraints.Count
3989          With Constraints(c)
3990              CS = GetDisplayAddress(.LHS, showNameRange) + " " + .ConstraintType
3991              If Not .RHS Is Nothing Then
                      ' This is not a type restriction, so it has a RHS
3992                  CS = CS + " " + GetDisplayAddress(.RHS, showNameRange)
3993              ElseIf .RHSstring <> "" Then
                      Dim DisplayString As String
                      DisplayString = ConvertToCurrentLocale(.RHSstring)
3994                  If left(DisplayString, 1) = "=" Then DisplayString = Mid(DisplayString, 2)
3995                  CS = CS + " " + RemoveActiveSheetNameFromString(DisplayString)
3999              End If
4000          End With
4001          lst.AddItem CS
4002      Next c
4003      If oldLI >= 0 And oldLI <= lst.ListCount - 1 Then
4004          lst.ListIndex = oldLI
4005      Else
4006          lst.ListIndex = lst.ListCount - 1
4007      End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("CModel", "PopulateConstraintListBox") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Public Sub LoadFromSheet()
' TODO: Merge with BuildModelFromSolverData
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

4009      Set Constraints = New Collection

          Dim sheetName As String
4010      sheetName = EscapeSheetName(ActiveWorkbook.ActiveSheet)

          Set DecisionVariables = GetDecisionVariablesWithDefault()
          
4012      NonNegativityAssumption = GetNonNegativity()
           
4014      ObjectiveSense = GetObjectiveSense()
4016      ObjectiveTarget = GetObjectiveTargetValue()
4018      Set ObjectiveFunctionCell = GetObjectiveFunctionCell()
          
          Dim constraint As Long, NumConstraints As Long
4020      NumConstraints = GetNumConstraints()
4021      For constraint = 1 To NumConstraints
              
              Dim rLHS As Range, valRHS As Double, rRHS As Range, RHSRefersToFormula As Boolean, sRefersToRHS As String
              
4022          On Error Resume Next
              Set rLHS = GetConstraintLhs(constraint)
              If Err.Number <> 0 Then GoTo SkipCon
              On Error GoTo ErrorHandler
              
              Dim rel As Long, relString As String
4024          rel = GetConstraintRel(constraint)
              relString = RelationEnumToString(rel)
                            
4025          Select Case rel
              Case RelationINT, RelationBIN, RelationAllDiff
4026              AddConstraint rLHS, Nothing, "", relString
4029          Case RelationLE, RelationEQ, RelationGE
4030              On Error Resume Next
                  Set rRHS = GetConstraintRhs(constraint, sRefersToRHS, valRHS, RHSRefersToFormula)
4031              If Err.Number <> 0 Then GoTo SkipCon
                  On Error GoTo ErrorHandler

4033              AddConstraint rLHS, rRHS, sRefersToRHS, relString
              Case Else
                  GoTo SkipCon
4042          End Select
SkipCon:
4043      Next constraint
          
          Set Duals = GetDuals()
          
ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("CModel", "LoadFromSheet") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub
