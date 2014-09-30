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

'--------------------------------------------------------------------
' FindObjective
' Try and find the objective function cell and sense by searching
' for likely keywords, then search the area for appropriate
' calculations.
'
' Written by:       IRD
'--------------------------------------------------------------------
Public Function FindObjective(ByRef s As Worksheet) As String

          '----------------------------------------------------------------
          ' Determine objective sense
          ' Supported keywords:   min, minimise, minimize,
          '                       max, maximise, maximize
          Dim ObjSenseCell As Range
3586      Set ObjSenseCell = Nothing
3587      ObjectiveSense = UnknownObjectiveSense
          
3588      Application.StatusBar = "OpenSolver: Trying to determine objective sense..."
3589      Application.Cursor = xlWait
          
          ' MIN
3590      FindObjSense s, "min", ObjSenseCell
3591      If Not (ObjSenseCell Is Nothing) Then
3592          ObjectiveSense = MinimiseObjective: GoTo StopSearchingForModelType
3593      End If
3594      FindObjSense s, "minimise", ObjSenseCell
3595      If Not (ObjSenseCell Is Nothing) Then
3596          ObjectiveSense = MinimiseObjective: GoTo StopSearchingForModelType
3597      End If
3598      FindObjSense s, "minimize", ObjSenseCell
3599      If Not (ObjSenseCell Is Nothing) Then
3600          ObjectiveSense = MinimiseObjective: GoTo StopSearchingForModelType
3601      End If
          
          ' MAX
3602      FindObjSense s, "max", ObjSenseCell
3603      If Not (ObjSenseCell Is Nothing) Then
3604          ObjectiveSense = MaximiseObjective: GoTo StopSearchingForModelType
3605      End If
3606      FindObjSense s, "maximise", ObjSenseCell
3607      If Not (ObjSenseCell Is Nothing) Then
3608          ObjectiveSense = MaximiseObjective: GoTo StopSearchingForModelType
3609      End If
3610      FindObjSense s, "maximize", ObjSenseCell
3611      If Not (ObjSenseCell Is Nothing) Then
3612          ObjectiveSense = MaximiseObjective: GoTo StopSearchingForModelType
3613      End If
          
StopSearchingForModelType:

          ' Success?
3614      If ObjectiveSense = UnknownObjectiveSense Then
              ' Didn't find anything, give up here and report failure
3615          FindObjective = "NoSense"
3616          Application.Cursor = xlDefault
              'Application.StatusBar = "OpenSolver: Couldn't determine objective sense!"
3617          Application.StatusBar = False
3618          Exit Function
3619      Else
              ' At least got the sense, can continue
3620      End If
          
          '----------------------------------------------------------------
          ' Search for objective function cell
          ' - Start looking for a sumproduct in the row that model type
          '   was in, as well as row above and below
          ' - Otherwise find the first cell with a sum product in it
          '   looking across columns, then down rows
          ' - Can't find, give up
3621      Application.StatusBar = "OpenSolver: Found objective sense, looking for objective cell..."
          Dim RowLBound As Long, RowUBound As Long
          
          
          ' Prevent out of range errors
3622      If ObjSenseCell.row - 1 <= 0 Then
3623          RowLBound = 1
3624      Else
3625          RowLBound = ObjSenseCell.row - 1
3626      End If
3627      RowUBound = ObjSenseCell.row + 1
          
          ' SUMPRODUCT SEARCH FIRST
          ' SAME ROW AS MODEL TYPE
3628      FindObjCell s, ObjSenseCell.row, "sumproduct", ObjectiveFunctionCell
3629      If Not (ObjectiveFunctionCell Is Nothing) Then GoTo FoundObjCell
          ' ROW ABOVE
3630      FindObjCell s, RowLBound, "sumproduct", ObjectiveFunctionCell
3631      If Not (ObjectiveFunctionCell Is Nothing) Then GoTo FoundObjCell
          ' ROW BELOW
3632      FindObjCell s, RowUBound, "sumproduct", ObjectiveFunctionCell
3633      If Not (ObjectiveFunctionCell Is Nothing) Then GoTo FoundObjCell
          
          ' NEXT SEARCH FOR ANY FORMULA
          ' SAME ROW AS MODEL TYPE
3634      FindObjCell s, ObjSenseCell.row, "=", ObjectiveFunctionCell
3635      If Not (ObjectiveFunctionCell Is Nothing) Then GoTo FoundObjCell
          ' ROW ABOVE
3636      FindObjCell s, RowLBound, "=", ObjectiveFunctionCell
3637      If Not (ObjectiveFunctionCell Is Nothing) Then GoTo FoundObjCell
          ' ROW BELOW
3638      FindObjCell s, RowUBound, "=", ObjectiveFunctionCell
3639      If Not (ObjectiveFunctionCell Is Nothing) Then GoTo FoundObjCell
         
FoundObjCell:

3640      If ObjectiveFunctionCell Is Nothing Then
              ' Couldn't find it
3641          FindObjective = "SenseNoCell"
3642          Application.Cursor = xlDefault
              'Application.StatusBar = "OpenSolver: Determined objective sense, couldn't find obj. cell!"
3643          Application.StatusBar = False
3644          Exit Function
3645      Else
              ' Found something!
3646      End If
              
3647      FindObjective = "OK"
3648      Application.Cursor = xlDefault
          'Application.StatusBar = "OpenSolver: Found objective cell and sense."
3649      Application.StatusBar = False
          
End Function


'--------------------------------------------------------------------
' FindObjSense
' Helper function for FindObjective. Contains the settings for the
' right kind of search to find the objective sense (values, don't match case)
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub FindObjSense(ByRef s As Worksheet, ByVal searchStr As String, ByRef result As Range)
3650      Set result = s.Cells.Find( _
                      What:=searchStr, After:=[a1], _
                      LookIn:=xlValues, SearchOrder:=xlByRows, _
                      SearchDirection:=xlPrevious, MatchCase:=False)
End Sub


'--------------------------------------------------------------------
' FindObjCell
' Helper function for FindObjective. Contains the settings for the
' right kind of search to find objective cell (specified row, formulas, don't match case)
'
' Written by:       JWD
'--------------------------------------------------------------------
Private Sub FindObjCell(ByRef s As Worksheet, ByVal rowNum As Long, ByVal searchStr As String, ByRef result As Range)
      ' Mac cannot handle SearchFormat parameter
#If Mac Then
3651            Set result = _
                    s.Rows(rowNum).Find(What:=searchStr, _
                        LookIn:=xlFormulas, lookat:=xlPart, _
                        SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                        MatchCase:=False)
#Else
3652            Set result = _
                    s.Rows(rowNum).Find(What:=searchStr, _
                        LookIn:=xlFormulas, lookat:=xlPart, _
                        SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                        MatchCase:=False, SearchFormat:=False)
#End If
End Sub




'--------------------------------------------------------------------
' FindVarsAndCons
' We have objective, now find all constraints.
' We can build a collection containing all the precedents of the
' constraints and objective. All cells which are referenced
' more than once (i.e. objective and >= 1 constraint,
' or >= 2 constraints) are assumed to be decision variables.
'
' Written by:       IRD
'--------------------------------------------------------------------
Public Function FindVarsAndCons(IsFirstTime As Boolean) As Boolean

          Dim PrecedentCells As Range, CurPrecedent As Range
          
          '----------------------------------------------------------------
          ' Clear existing solution, if requested
3653      If IsFirstTime Then
3654          Set DecisionVariables = Nothing
3655          Set Constraints = New Collection
3656      End If
          
          '----------------------------------------------------------------
          ' Phase One
          '   - Look for constraints (i.e. <=, >=, =) and add them if they
          '     seem at all interesting (i.e. LHS or RHS has precedents)
3657      Application.StatusBar = "OpenSolver:  Looking for constraints"
3658      On Error GoTo ConstraintErr
          
              '------------------------------------------------------------
              ' Stage One - Find all the "relational operator" cells on the
              ' sheet that the objective function is on.

              ' Find the cells
              Dim FoundLEQ As Range, FoundGEQ As Range, FoundEQ As Range
3659          If Not ObjectiveFunctionCell Is Nothing Then
3660              FindAllCells "<=", FoundLEQ, ObjectiveFunctionCell.Parent
3661              FindAllCells ">=", FoundGEQ, ObjectiveFunctionCell.Parent
3662              FindAllCells "=", FoundEQ, ObjectiveFunctionCell.Parent
3663          Else
3664              FindAllCells "<=", FoundLEQ, ActiveSheet
3665              FindAllCells ">=", FoundGEQ, ActiveSheet
3666              FindAllCells "=", FoundEQ, ActiveSheet
3667          End If
              
              ' Combine them as much as possible
              Dim AllCompOps As Range
3668          Set AllCompOps = FoundEQ
3669          Set AllCompOps = ProperUnion(AllCompOps, FoundLEQ)
3670          Set AllCompOps = ProperUnion(AllCompOps, FoundGEQ)
          
              '------------------------------------------------------------
              ' Stage Two - Now we have the blocks of constraints. For
              ' blocks that have more rows than columns, we can look left
              ' and right. For more columns than rows, up and down. For
              ' single cells, look in all directions
              Dim AreaIndex As Long
3671          For AreaIndex = 1 To AllCompOps.Areas.Count
                  ''Debug.Print "FindVarsAndCons: Con area" + Str(AreaIndex)
                  
                  ' Determine the shape of the area
                  Dim RowCount As Long, ColCount As Long
3672              RowCount = AllCompOps.Areas(AreaIndex).Rows.Count
3673              ColCount = AllCompOps.Areas(AreaIndex).Columns.Count
                  
                  ' Depending on the shape, set the LHS and RHS
                  ' NB: in the case of row of constraints, LHS refers to
                  '     the cells above the row, RHS the cells below.
                  Dim LHSs As Range, RHSs As Range
3674              If ColCount = 1 And RowCount > 1 Then
                      ' Vertical
                      ''Debug.Print "FindVarsAndCons: Vertical", AreaIndex
3675                  Set LHSs = AllCompOps.Areas(AreaIndex).Offset(0, -1)
3676                  Set RHSs = AllCompOps.Areas(AreaIndex).Offset(0, 1)
3677                  If CheckPrecedentCells(LHSs, RHSs) Then
                          ' Its ok, add them
3678                      AddRangeToConstraints LHSs, AllCompOps.Areas(AreaIndex), RHSs, True, Constraints
                          ''Debug.Print "FindVarsAndCons: Added", AreaIndex
3679                  End If
                      
3680              ElseIf ColCount > 1 And RowCount = 1 Then
                      ' Horizontal
                      ''Debug.Print "FindVarsAndCons: Horizontal", AreaIndex
3681                  Set LHSs = AllCompOps.Areas(AreaIndex).Offset(-1, 0)
3682                  Set RHSs = AllCompOps.Areas(AreaIndex).Offset(1, 0)
3683                  If CheckPrecedentCells(LHSs, RHSs) Then
                          ' Its ok, add them
3684                      AddRangeToConstraints LHSs, AllCompOps.Areas(AreaIndex), RHSs, False, Constraints
                          ''Debug.Print "FindVarsAndCons: Added", AreaIndex
3685                  End If
                      
3686              ElseIf ColCount = 1 And RowCount = 1 Then
                      ' Singleton
                      ''Debug.Print "FindVarsAndCons: Singleton", AreaIndex
                      ' Look Left-Right
3687                  Set LHSs = AllCompOps.Areas(AreaIndex).Offset(0, -1)
3688                  Set RHSs = AllCompOps.Areas(AreaIndex).Offset(0, 1)
3689                  If CheckPrecedentCells(LHSs, RHSs) Then
                          ' Its ok, add them
3690                      AddRangeToConstraints LHSs, AllCompOps.Areas(AreaIndex), RHSs, True, Constraints
                          ''Debug.Print "FindVarsAndCons: Added LR", AreaIndex
3691                  Else
                          ' Didn't find anything interesting LR, look UD
3692                      Set LHSs = AllCompOps.Areas(AreaIndex).Offset(-1, 0)
3693                      Set RHSs = AllCompOps.Areas(AreaIndex).Offset(1, 0)
3694                      If CheckPrecedentCells(LHSs, RHSs) Then
                              ' Its ok, add them
3695                          AddRangeToConstraints LHSs, AllCompOps.Areas(AreaIndex), RHSs, False, Constraints
                              ''Debug.Print "FindVarsAndCons: Added UD", AreaIndex
3696                      End If
3697                  End If
3698              Else
                      ' Block
                      ' TODO - Handle this somehow, if it has an application
                      ''Debug.Print "FindVarsAndCons: Block TODO", AreaIndex
3699              End If
                  
3700          Next AreaIndex

3701          RationaliseConstraints
              
              
          '----------------------------------------------------------------
          ' Phase Two
          '       - use precedents of objective function and constraints
          '         to find the set of possible decision variables
3702      Application.StatusBar = "OpenSolver: Searching for decision variables"
          
          ' NB: The first collection has cell addresses as keys, and reference
          '     counts as values.
          '     The second collection has cell addresses as key and value.
          '     We need this because, as far as I know, you can't iterate
          '     through the keys of a collection.
          Dim DecRefCount As New Collection, DecRefAddress As New Collection
          
          
              '------------------------------------------------------------
              ' Stage One - Objective function precedents
3703          On Error Resume Next
              ' Find the precedents of the objective function
3704          Set PrecedentCells = Nothing
3705          Set PrecedentCells = ObjectiveFunctionCell.Precedents
3706          If PrecedentCells Is Nothing Then
                  ' Objective function must be constant, this is not necessarily an issue
                  ''Debug.Print "FindVarsAndCons: Constant objective, skipping stage one of phase two"
3707          Else
                  ' Go through each precedent of the objective function, and set the
                  ' reference counter for it to 1
3708              On Error GoTo DecisionErr
3709              For Each CurPrecedent In PrecedentCells.Cells
3710                  If Not CurPrecedent.HasFormula Then
3711                      If Not TestKeyExists(DecRefCount, CurPrecedent.Address) Then
3712                          DecRefCount.Add 1, CurPrecedent.Address
3713                          DecRefAddress.Add CurPrecedent.Address, CurPrecedent.Address
                              ''Debug.Print "obj", CurPrecedent.Address, 1
3714                      End If
3715                  End If
3716              Next
3717          End If
              
              ''Debug.Print "FindVarsAndCons: End of Phase One Stage One"


              '------------------------------------------------------------
              ' Stage Two - Constraint precedents
3718          On Error Resume Next
              Dim curConstraint As CConstraint, RC As Long
3719          For Each curConstraint In Constraints
                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  ' Look at LHS
3720              On Error Resume Next
3721              Set PrecedentCells = Nothing
3722              Set PrecedentCells = curConstraint.LHS.Precedents
3723              If Not PrecedentCells Is Nothing Then
3724                  For Each CurPrecedent In PrecedentCells.Cells
3725                      If TestKeyExists(DecRefCount, CurPrecedent.Address) Then
                              ' Exists, increment
                              ' Annoyingly, collections are read-only
                              ' I'm not sure how slow this is, but a way to
                              ' optimise would be to only do this if RC = 1, because
                              ' once RC hits two we will pick it anyway
3726                          RC = DecRefCount(CurPrecedent.Address)
3727                          DecRefCount.Remove CurPrecedent.Address
3728                          DecRefCount.Add RC + 1, CurPrecedent.Address
                              ''Debug.Print "con", CurPrecedent.Address, RC + 1
3729                      Else
                              ' Doesn't exist yet, start it at 1
3730                          If Not CurPrecedent.HasFormula Then
3731                              DecRefCount.Add 1, CurPrecedent.Address
3732                              DecRefAddress.Add CurPrecedent.Address, CurPrecedent.Address
                                  ''Debug.Print "connew", CurPrecedent.Address, 1
3733                          End If
3734                      End If
3735                  Next
3736              End If
                  
                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  ' Look at RHS
3737              On Error Resume Next
3738              Set PrecedentCells = Nothing
3739              Set PrecedentCells = curConstraint.RHS.Precedents
3740              If Not PrecedentCells Is Nothing Then
3741                  For Each CurPrecedent In PrecedentCells.Cells
3742                      If TestKeyExists(DecRefCount, CurPrecedent.Address) Then
                              ' Exists, increment
3743                          DecRefCount(CurPrecedent.Address) = DecRefCount(CurPrecedent.Address) + 1
3744                      Else
                              ' Doesn't exist yet, start it at 1
3745                          If Not CurPrecedent.HasFormula Then
3746                              DecRefCount.Add 1, CurPrecedent.Address
3747                              DecRefAddress.Add CurPrecedent.Address, CurPrecedent.Address
3748                          End If
3749                      End If
3750                  Next
3751              End If
              
3752          Next
          
          
          
          '----------------------------------------------------------------
          ' Phase Three
          '   If a decision variable has only been referenced once, we
          '   really can't be sure it is a decision variable - constants
          '   are also referenced once. So take anything that is two or
          '   or higher
3753      Application.StatusBar = "OpenSolver: Selecting most likely decision variables"
              
              Dim addressKey As Variant
              
3754          For Each addressKey In DecRefAddress
3755              If DecRefCount(addressKey) >= 2 Then
3756                  Set DecisionVariables = ProperUnion(DecisionVariables, ActiveSheet.Range(CStr(addressKey)))
3757              End If
3758          Next
              
          '----------------------------------------------------------------
          ' Phase Four
          '   - Look for type restrictions on decision variables
3759      Application.StatusBar = "OpenSolver: Looking for variable type restrictions"
              Dim CurDecVar As Range, PossibleType As String
          
3760          For Each CurDecVar In DecisionVariables
                  ' Look below it to see if there is type information
3761              PossibleType = LCase(Trim(CurDecVar.Offset(1, 0).value))
3762              If PossibleType = "integer" Or PossibleType = "int" Or PossibleType = "i" Then
                      ' This variable is integer
3763                  Set curConstraint = New CConstraint
3764                  With curConstraint
3765                      Set .LHS = CurDecVar
3766                      Set .RHS = Nothing
3767                      Set .Relation = Nothing
3768                      .ConstraintType = "int"
3769                  End With
3770                  Constraints.Add curConstraint ', curConstraint.GetKey
3771              ElseIf PossibleType = "binary" Or PossibleType = "bin" Or PossibleType = "b" Then
                      ' This variable is BINARY
3772                  Set curConstraint = New CConstraint
3773                  With curConstraint
3774                      Set .LHS = CurDecVar
3775                      Set .RHS = Nothing
3776                      Set .Relation = Nothing
3777                      .ConstraintType = "bin"
3778                  End With
3779                  Constraints.Add curConstraint ', curConstraint.GetKey
3780              End If
3781          Next
          
          
          '----------------------------------------------------------------
          ' Finished!
3782      Application.StatusBar = "OpenSolver: Rationalising constraints"
          'RationaliseConstraints ' Make constraints prettier ' BROKEN IN 1.3 COS USES KEYS
3783      FindVarsAndCons = True
          'Application.StatusBar = "OpenSolver: Finished finding decision variables and constraints"
3784      Application.StatusBar = False
3785      Exit Function
          
DecisionErr:
          ' Error occurred while trying to find decision variables
3786      MsgBox "Error: an issue arose while finding decision variables." + vbNewLine + _
                 "Error number:" + str(Err.Number) + vbNewLine + _
                 "Error description: " + Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")"), _
                 vbExclamation Or vbOKOnly, "AutoModel"
3787      FindVarsAndCons = False
3788      Application.StatusBar = False
3789      Exit Function
          
ConstraintErr:
          ' Error occurred while trying to find constraints
3790      MsgBox "Error: an issue arose while finding constraints." + vbNewLine + _
                 "Error number:" + str(Err.Number) + vbNewLine + _
                 "Error description: " + Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")"), _
                 vbExclamation Or vbOKOnly, "AutoModel"
3791      FindVarsAndCons = False
3792      Application.StatusBar = False
3793      Exit Function
End Function


'--------------------------------------------------------------------
' FindAllCells
' Returns a range (possibly with multiple areas) containing all cells
' in the sheet that contain the search string.
' Only looks in the value - hopefully avoiding conditionals that may
' be in formulae
'
' Inputs:
'   SearchStr       String to look for for
'   FoundCells      [REF] The range to store the results in. Will be
'                   Nothing if no cells found
'   sheet           Worksheet to look on
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub FindAllCells(ByVal searchStr As String, ByRef FoundCells As Range, _
                            ByRef sheet As Worksheet)

          ' Initialise
          Dim LastCell As Range, FirstCell As Range
3794      Set FoundCells = Nothing
          
          ' Find first cell that meets requirements
3795      Set LastCell = sheet.Cells.Find(What:=searchStr, _
                                  After:=[a1], _
                                  LookIn:=xlValues, _
                                  SearchOrder:=XlSearchOrder.xlByRows, _
                                  lookat:=XlLookAt.xlWhole, _
                                  SearchDirection:=XlSearchDirection.xlNext _
                                  )
3796      Set FirstCell = LastCell
3797      If LastCell Is Nothing Then Exit Sub ' If not even one, stop immediately
          
3798      Do
              ' Add it...
3799          If FoundCells Is Nothing Then
3800              Set FoundCells = Range(LastCell.Address)
3801          Else
3802              Set FoundCells = Union(FoundCells, LastCell)
3803          End If
              ' ...and find next
3804          Set LastCell = sheet.Cells.FindNext(LastCell)
              ' Loop until we don't find another cell, or until we get back to the
              ' initial cell we found
3805      Loop While (Not LastCell Is Nothing) And (FirstCell.Address <> LastCell.Address)
          
End Sub


'--------------------------------------------------------------------
' CheckPrecedentCells
' Currently just tries to determine if any of the LHS or RHS have
' a precedent. This is a cheap way of checking whether these
' constraints are 'active'.
'
' Inputs:
'   LHSs          [REF] LHS range
'   RHSs          [REF] RHS range
'
' Output:
'   A boolean - True means a precedent found for any cell in
'   either range.
'
' Written by:       IRD
'--------------------------------------------------------------------
Function CheckPrecedentCells(ByRef LHSs As Range, ByRef RHSs As Range) As Boolean
          Dim CurCell As Range, PrecCells As Range
          Dim BothSides As Range
3806      Set BothSides = Union(LHSs, RHSs)

3807      For Each CurCell In BothSides.Cells
              ' Check if cell has a precedent
              ' If no precedents, error is thrown
3808          Err.Clear
3809          On Error Resume Next
3810          Set PrecCells = CurCell.Precedents
3811          If Err.Number = 1004 Then
                  ' No precedents for this cell
3812          ElseIf Err.Number = 0 Then
                  ' There is a precedent
3813              CheckPrecedentCells = True
3814              Exit Function
3815          End If
3816      Next
          
End Function


'--------------------------------------------------------------------
' AddRangeToConstraints
' The LHS/Rel/RHS set is to be added to contraints.
' Problem: Doesn't group constraints, even if it could
' Short term fix: Add one-by-one
' Long term fix: Some ugly code, perhaps
'
' Inputs:
'   LHSs          [REF] LHS range
'   RelRange      [REF] The relation operator range
'   RHSs          [REF] RHS range
'   IsVertical    Boolean to tell shape
'   Constraints   [REF] The master constraint collection
'
' Written by:       IRD
'--------------------------------------------------------------------
Sub AddRangeToConstraints(ByRef LHSs As Range, ByRef RelRange As Range, _
                            ByRef RHSs As Range, IsVertical As Boolean, _
                            ByRef Constraints As Collection)
          Dim CellCount As Long
3817      CellCount = LHSs.Count
          
          Dim i As Long
          Dim LHSi As Range, RELi As Range, RHSi As Range
          Dim NewConstraint As CConstraint
          
          ' For each cell
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
              ' Check if added
3828          If Not TestKeyExists(Constraints, RELi.Address) Then
                  ' Constraint not already added - so add it!
3829              Set NewConstraint = New CConstraint
3830              With NewConstraint
3831                  Set .LHS = LHSi
3832                  Set .RHS = RHSi
3833                  Set .Relation = RELi
3834                  .ConstraintType = RELi.value
                      ''Debug.Print "Con added:", .LHS.Address, .ConstraintType, .RHS.Address
3835              End With
3836              Constraints.Add NewConstraint ', NewConstraint.GetKey
3837          Else
                  ' Added already
                  'Debug.Print "Con already added", LHSi.Address, RELi.Value, RHSi.Address
3838          End If
3839      Next i
      '          RationaliseConstraints
End Sub


'--------------------------------------------------------------------
' RationaliseConstraints
' Group multiple individual constraints into 1 constraint if:
'   - They are next to each other
'   - They are of the same type
'
'
' Written by:       MJVM & IRD
'--------------------------------------------------------------------
Public Sub RationaliseConstraints()
3840      On Error GoTo 0
          Dim Relations(1 To 5) As String
3841      Relations(1) = "="
3842      Relations(2) = "<="
3843      Relations(3) = ">="
3844      Relations(4) = "int"
3845      Relations(5) = "bin"
          
          Dim NewConstraints As New Collection
          Dim NewConstraint As CConstraint
          Dim conKeys As Collection
          Dim curCon As CConstraint
          Dim UnionRange As Range
          Dim AreaIndex As Long
          Dim LHSunion As Range, RHSunion As Range, RELunion As Range
          Dim CurCell As Range
          Dim NumConstraints As Long
          Dim conNum() As Variant
          
3846      NumConstraints = Constraints.Count
                    
3847      ReDim conNum(NumConstraints + 1) As Variant
          
          Dim i As Long, j As Long
3848      For i = 1 To 5
              ' Find all constraints of type Relations(i)...
3849          Set conKeys = New Collection
                
                Dim c As Long, k As Long
                
                'Initialise the indices of the constraints to be empty
3850            For k = 1 To NumConstraints
3851                conNum(k) = Empty
3852            Next k
                
3853            c = 1
3854            k = 1
                
3855          For Each curCon In Constraints
                  ' ... and add a reference to them in a collection and save the index of the constraint
3856              If curCon.ConstraintType = Relations(i) Then
3857                  If 1 <= i And i <= 3 Then
3858                        conKeys.Add curCon.Relation.Address, curCon.Relation.Address
3859                        conNum(k) = c
3860                        k = k + 1
3861                  End If
3862                  If 4 <= i And i <= 5 Then
3863                        conKeys.Add curCon.LHS.Address, curCon.LHS.Address
3864                        conNum(k) = c
3865                        k = k + 1
3866                  End If
3867              End If
3868              c = c + 1
3869          Next

3870          If conKeys.Count > 0 Then
3871              Set UnionRange = Nothing
                  'Find the number of constraints for relations(i)
3872              k = 1
                  
3873              Do While conNum(k) <> ""
3874                  k = k + 1
3875              Loop
                
                ' Build up a range that is the union of all the constraints
3876              For j = 1 To (k - 1)
3877                  If 1 <= i And i <= 3 Then Set UnionRange = ProperUnion(UnionRange, Constraints(conNum(j)).Relation)
3878                  If 4 <= i And i <= 5 Then Set UnionRange = ProperUnion(UnionRange, Constraints(conNum(j)).LHS)
3879              Next j
                  ' Now iterate through each area of the range - each represents a block
                  ' of constraints that are next to each other, with the same relation
                  
                  Dim SavedNum As Long, num As Long
                  
3880              SavedNum = 1
                  
3881              For AreaIndex = 1 To UnionRange.Areas.Count
                      'Debug.Print "Relation " + Relations(i) + ", area" + str(AreaIndex)
3882                  num = SavedNum
3883                  Set LHSunion = Nothing
3884                  Set RHSunion = Nothing
3885                  Set RELunion = Nothing

3886                  For Each CurCell In UnionRange.Areas(AreaIndex).Cells
3887                      Set LHSunion = ProperUnion(LHSunion, Constraints(conNum(num)).LHS)
3888                      Set RHSunion = ProperUnion(RHSunion, Constraints(conNum(num)).RHS)
3889                      Set RELunion = ProperUnion(RELunion, Constraints(conNum(num)).Relation)
3890                      num = num + 1
3891                  Next
3892                  SavedNum = num

3893                  Set NewConstraint = New CConstraint
3894                  With NewConstraint
3895                      Set .LHS = LHSunion
3896                      Set .RHS = RHSunion
3897                      Set .Relation = RELunion
3898                      .ConstraintType = Relations(i)
                          ''Debug.Print "SuperCon added:", .LHS.Address, .ConstraintType, .RHS.Address
3899                  End With
3900                  If 1 <= i And i <= 3 Then NewConstraints.Add NewConstraint, NewConstraint.Relation.Address
3901                  If 4 <= i And i <= 5 Then NewConstraints.Add NewConstraint, NewConstraint.LHS.Address
3902              Next
3903          End If
3904      Next i
          
          ' Update old constraints
3905      Set Constraints = NewConstraints
End Sub




'--------------------------------------------------------------------
' BuildModel
' Turn this object into an on-sheet model
'
' Written by:       IRD
'--------------------------------------------------------------------
Sub BuildModel()
          ' Decision variables
3906      If DecisionVariables Is Nothing Then
3907          DeleteSolverNameOnSheet "adj"
3908      Else
3909          SetSolverNamedRangeOnSheet "adj", DecisionVariables
3910      End If
          
          ' Objective function cell
3911      If ObjectiveFunctionCell Is Nothing Then
3912          DeleteSolverNameOnSheet "opt"
3913      Else
3914          SetSolverNamedRangeOnSheet "opt", ObjectiveFunctionCell
3915      End If
          
          ' Objective type
3916      If ObjectiveSense = MinimiseObjective Then
3917          SetSolverNameOnSheet "typ", "=2"
3918      ElseIf ObjectiveSense = MaximiseObjective Then
3919          SetSolverNameOnSheet "typ", "=1"
3920      ElseIf ObjectiveSense = TargetObjective Then
3921          SetSolverNameOnSheet "typ", "=3"
3922          SetSolverNameOnSheet "val", "=" + Trim(str(ObjectiveTarget)) ' AJM Changed from CStr to Str as we must store this using US numbers, and we must strip any leading space
3923      End If
          
          ' Check int/bin constraints not set on non-decision variables
          Dim c As Long, intersection As Range
3924      For c = 1 To Constraints.Count
3925          If Constraints(c).ConstraintType = "bin" _
              Or Constraints(c).ConstraintType = "int" Then
3926              Set intersection = Intersect(Constraints(c).LHS, DecisionVariables)
3927              If intersection Is Nothing Then
3928                  MsgBox "Warning: an int/bin constraint has been set on cells that are not decision variables. OpenSolver does not support this yet. An error message will appear if you have selected to show the Model."
3929              Else
3930                  If intersection.Count <> Constraints(c).LHS.Count Then
3931                      MsgBox "Warning: an int/bin constraint has been set on cells that are not decision variables. OpenSolver does not support this yet. An error message will appear if you have selected to show the Model."
3932                  End If
3933              End If
3934          End If
3935      Next c
          
          ' Constraints
3936      SetSolverNameOnSheet "num", "=" + CStr(Constraints.Count)
3937      For c = 1 To Constraints.Count
3938          If Constraints(c).ConstraintType = "bin" Then
3939              SetSolverNamedRangeOnSheet "lhs" & CStr(c), Constraints(c).LHS
3940              SetSolverNameOnSheet "rel" & CStr(c), "=5"
3941              SetSolverNameOnSheet "rhs" & CStr(c), "=binary"
3942          ElseIf Constraints(c).ConstraintType = "int" Then
3943              SetSolverNamedRangeOnSheet "lhs" & CStr(c), Constraints(c).LHS
3944              SetSolverNameOnSheet "rel" & CStr(c), "=4"
3945              SetSolverNameOnSheet "rhs" & CStr(c), "=integer"
3946          ElseIf Constraints(c).ConstraintType = "alldiff" Then
3947              SetSolverNamedRangeOnSheet "lhs" & CStr(c), Constraints(c).LHS
3948              SetSolverNameOnSheet "rel" & CStr(c), "=6"
3949              SetSolverNameOnSheet "rhs" & CStr(c), "=alldiff"
3950          Else
3951              SetSolverNamedRangeOnSheet "lhs" & CStr(c), Constraints(c).LHS
3952              If Constraints(c).RHS Is Nothing Then
3953                  If Mid(Constraints(c).RHSstring, 1, 1) = "=" Then
3954                      SetSolverNameOnSheet "rhs" & CStr(c), Constraints(c).RHSstring
3955                  Else
3956                      SetSolverNameOnSheet "rhs" & CStr(c), "=" & Constraints(c).RHSstring
3957                  End If
3958              Else
3959                  SetSolverNamedRangeOnSheet "rhs" & CStr(c), Constraints(c).RHS
3960              End If
3961              If Constraints(c).ConstraintType = ">=" Then SetSolverNameOnSheet "rel" & CStr(c), "=3"
3962              If Constraints(c).ConstraintType = ">" Then SetSolverNameOnSheet "rel" & CStr(c), "=3"
3963              If Constraints(c).ConstraintType = "<=" Then SetSolverNameOnSheet "rel" & CStr(c), "=1"
3964              If Constraints(c).ConstraintType = "<" Then SetSolverNameOnSheet "rel" & CStr(c), "=1"
3965              If Constraints(c).ConstraintType = "=" Then SetSolverNameOnSheet "rel" & CStr(c), "=2"
3966              If Constraints(c).ConstraintType = "'=" Then SetSolverNameOnSheet "rel" & CStr(c), "=2"
3967          End If
3968      Next c
          
3969      If Duals Is Nothing Then
3970          DeleteNameOnSheet "OpenSolver_Duals"
3971      Else
3972          SetNamedRangeOnSheet "OpenSolver_Duals", Duals
3973      End If
          
          ' Set the default Solver properties
          ' These values must be given in English, not using the local language (eg with . and not ,)
3974      SetSolverNameOnSheet "neg", IIf(NonNegativityAssumption, "=1", "=2")
          ' All of these have been initialised if any were missing when the dialog was first shown
          'SetSolverNameOnSheet "tim", "=100"
          'SetSolverNameOnSheet "itr", "=100"
          'SetSolverNameOnSheet "pre", "=0.000001"
          'SetSolverNameOnSheet "tol", "=0.05"
          'SetSolverNameOnSheet "cvg", "=0.0001"
          'SetSolverNameOnSheet "sho", "=2"
End Sub


'--------------------------------------------------------------------
' AddConstraint
' Add a constraint to the collection that is NOT an automatically
' found one (hence no Relation range is expected, and will be keyed
' on the LHS)
'
' Written by:       IRD
'--------------------------------------------------------------------
Public Sub AddConstraint(newLHS As Range, newRHS As Range, newRHSstring As String, newType As String)
          
          Dim NewConstraint As New CConstraint
          
3975      With NewConstraint
3976          Set .LHS = newLHS
3977          Set .Relation = Nothing
3978          Set .RHS = newRHS
3979          .RHSstring = newRHSstring
3980          .ConstraintType = newType
3981      End With
          
3982      Constraints.Add NewConstraint ', NewConstraint.GetKey
          
End Sub


'--------------------------------------------------------------------
' PopulateConstraintListBox
' Adds constraints to a listbox on a form
'
' Written by:       IRD
'--------------------------------------------------------------------
Public Sub PopulateConstraintListBox(ByRef lst As Object)
          Dim CS As String, c As Long
          Dim oldLI As Long
          
          Dim showNameRange As Boolean
3983      showNameRange = frmModel.chkNameRange.value
3984      If showNameRange Then SearchRangeName_DestroyCache
          
3985      oldLI = lst.ListIndex
3986      lst.Clear
3987      lst.AddItem "<Add new constraint>"
3988      For c = 1 To Constraints.Count
3989          With Constraints(c)
3990              CS = GetDisplayAddress(.LHS, showNameRange) + " " + .ConstraintType
3991              If Not (.RHS Is Nothing) Then
                      ' This is not a type restriction, so it has a RHS
3992                  CS = CS + " " + GetDisplayAddress(.RHS, showNameRange)
3993              ElseIf .RHS Is Nothing And .RHSstring <> "" Then
3994                  If Mid(.RHSstring, 1, 1) = "=" Then
3995                      CS = CS + " " + RemoveActiveSheetNameFromString(Mid(.RHSstring, 2, Len(.RHSstring)))
3996                  Else
3997                      CS = CS + " " + RemoveActiveSheetNameFromString(.RHSstring)
3998                  End If
3999              End If
4000          End With
4001          lst.AddItem CS
4002      Next c
4003      If oldLI >= 0 And oldLI <= lst.ListCount - 1 Then
4004          lst.ListIndex = oldLI
4005      Else
4006          lst.ListIndex = lst.ListCount - 1
4007      End If
End Sub


'--------------------------------------------------------------------
' LoadFromSheet
' Loads the Solver-style model
' Borrowed a lot from COpenSolver.BuildModelFromSolverData
'
' Written by:       IRD
'--------------------------------------------------------------------
Public Sub LoadFromSheet()

4008      On Error Resume Next ' TODO: This should not run under a blank RESUME NEXT
4009      Set Constraints = New Collection


          Dim sheetName As String
4010      sheetName = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!"   ' Must include quotes to handle sheet names with spaces

          ' ADJUSTABLE CELLS
4011      Set DecisionVariables = Range(sheetName & "solver_adj")
          
          ' NON-NEGATIVITY
          Dim s As String
4012      NonNegativityAssumption = False
4013      If GetNameValueIfExists(ActiveWorkbook, sheetName & "solver_neg", s) Then NonNegativityAssumption = (s = "1")
           
          ' OBJECTIVE SENSE (AND TARGET)
          Dim temp As Long
4014      ObjectiveSense = UnknownObjectiveSense
4015      If GetNamedIntegerIfExists(ActiveWorkbook, sheetName & "solver_typ", temp) Then ObjectiveSense = temp
4016      ObjectiveTarget = 0
4017      GetNamedNumericValueIfExists ActiveWorkbook, sheetName & "solver_val", ObjectiveTarget
          
          ' OBJECTIVE CELL
4018      Set ObjectiveFunctionCell = Nothing
4019      Set ObjectiveFunctionCell = ActiveWorkbook.Names(sheetName & "solver_opt").RefersToRange
          
          ' CONSTRAINTS
          Dim constraint As Long, NumConstraints As Long
          
4020      NumConstraints = Val(Mid(Names(sheetName & "solver_num"), 2))
4021      For constraint = 1 To NumConstraints
              ' Note: Solver enforces the following requirements.
              ' The LHS mut be a range with one or more cells
              ' The RHS can be either:
              '   A single-cell range (=A4)
              '   A multi-cell range of the same size as the LHS (=A4:B5)
              '   A single constant value (eg =2)
              '   A formula returning a single value (eg =sin(A4))
              
              Dim isRangeLHS As Boolean, valLHS As Double, rLHS As Range, LHSRefersToError As Boolean, LHSRefersToFormula As Boolean, sRefersToLHS As String, LHSisMissing As Boolean
              Dim isRangeRHS As Boolean, valRHS As Double, rRHS As Range, RHSRefersToError As Boolean, RHSRefersToFormula As Boolean, sRefersToRHS As String, RHSisMissing As Boolean
              
4022          GetNameAsValueOrRange ActiveWorkbook, sheetName & "solver_lhs" & constraint, LHSisMissing, isRangeLHS, rLHS, LHSRefersToFormula, LHSRefersToError, sRefersToLHS, valLHS
4023          If (LHSisMissing) Or (Not isRangeLHS) Then GoTo SkipCon
              
              Dim rel As Long
4024          rel = Val(Mid(Names(sheetName & "solver_rel" & constraint), 2))  ' Either <=, = or >=, or int or bin (or 'alldif' in Excel 2010)
              
4025          If rel = RelationINT Or rel = RelationBIN Or rel = RelationAllDiff Then
                  ' Make the LHS variables integer or binary
4026              If rel = RelationINT Then AddConstraint rLHS, Nothing, "", "int"
4027              If rel = RelationBIN Then AddConstraint rLHS, Nothing, "", "bin"
4028              If rel = RelationAllDiff Then AddConstraint rLHS, Nothing, "", "alldiff"
4029          Else
                  ' Constraint is a full equation with a RHS
4030              GetNameAsValueOrRange ActiveWorkbook, sheetName & "solver_rhs" & constraint, RHSisMissing, isRangeRHS, rRHS, RHSRefersToFormula, RHSRefersToError, sRefersToRHS, valRHS
4031              If RHSisMissing Then GoTo SkipCon
                  
4032              If rel <> RelationLE And rel <> RelationEQ And rel <> RelationGE Then GoTo SkipCon

4033              If Not isRangeRHS Then
4034                  If rel = RelationEQ Then AddConstraint rLHS, Nothing, Mid(Names(sheetName & "solver_rhs" & constraint), 2), "="
4035                  If rel = RelationLE Then AddConstraint rLHS, Nothing, Mid(Names(sheetName & "solver_rhs" & constraint), 2), "<="
4036                  If rel = RelationGE Then AddConstraint rLHS, Nothing, Mid(Names(sheetName & "solver_rhs" & constraint), 2), ">="
4037              Else
4038                  If rel = RelationEQ Then AddConstraint rLHS, rRHS, "", "="
4039                  If rel = RelationLE Then AddConstraint rLHS, rRHS, "", "<="
4040                  If rel = RelationGE Then AddConstraint rLHS, rRHS, "", ">="
4041              End If
4042          End If
SkipCon:
4043      Next constraint
          
          ' DUALS
4044      Set Duals = Range(sheetName & "OpenSolver_Duals")
          
ExitSub:
End Sub

