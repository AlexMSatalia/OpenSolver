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
36740     Set ObjSenseCell = Nothing
36750     ObjectiveSense = UnknownObjectiveSense
          
36760     Application.StatusBar = "OpenSolver: Trying to determine objective sense..."
36770     Application.Cursor = xlWait
          
          ' MIN
36780     FindObjSense s, "min", ObjSenseCell
36790     If Not (ObjSenseCell Is Nothing) Then
36800         ObjectiveSense = MinimiseObjective: GoTo StopSearchingForModelType
36810     End If
36820     FindObjSense s, "minimise", ObjSenseCell
36830     If Not (ObjSenseCell Is Nothing) Then
36840         ObjectiveSense = MinimiseObjective: GoTo StopSearchingForModelType
36850     End If
36860     FindObjSense s, "minimize", ObjSenseCell
36870     If Not (ObjSenseCell Is Nothing) Then
36880         ObjectiveSense = MinimiseObjective: GoTo StopSearchingForModelType
36890     End If
          
          ' MAX
36900     FindObjSense s, "max", ObjSenseCell
36910     If Not (ObjSenseCell Is Nothing) Then
36920         ObjectiveSense = MaximiseObjective: GoTo StopSearchingForModelType
36930     End If
36940     FindObjSense s, "maximise", ObjSenseCell
36950     If Not (ObjSenseCell Is Nothing) Then
36960         ObjectiveSense = MaximiseObjective: GoTo StopSearchingForModelType
36970     End If
36980     FindObjSense s, "maximize", ObjSenseCell
36990     If Not (ObjSenseCell Is Nothing) Then
37000         ObjectiveSense = MaximiseObjective: GoTo StopSearchingForModelType
37010     End If
          
StopSearchingForModelType:

          ' Success?
37020     If ObjectiveSense = UnknownObjectiveSense Then
              ' Didn't find anything, give up here and report failure
37030         FindObjective = "NoSense"
37040         Application.Cursor = xlDefault
              'Application.StatusBar = "OpenSolver: Couldn't determine objective sense!"
37050         Application.StatusBar = False
37060         Exit Function
37070     Else
              ' At least got the sense, can continue
37080     End If
          
          '----------------------------------------------------------------
          ' Search for objective function cell
          ' - Start looking for a sumproduct in the row that model type
          '   was in, as well as row above and below
          ' - Otherwise find the first cell with a sum product in it
          '   looking across columns, then down rows
          ' - Can't find, give up
37090     Application.StatusBar = "OpenSolver: Found objective sense, looking for objective cell..."
          Dim RowLBound As Long, RowUBound As Long
          
          
          ' Prevent out of range errors
37100     If ObjSenseCell.row - 1 <= 0 Then
37110         RowLBound = 1
37120     Else
37130         RowLBound = ObjSenseCell.row - 1
37140     End If
37150     RowUBound = ObjSenseCell.row + 1
          
          ' SUMPRODUCT SEARCH FIRST
          ' SAME ROW AS MODEL TYPE
37160     Set ObjectiveFunctionCell = _
              s.Rows(ObjSenseCell.row).Find(What:="sumproduct", _
                  LookIn:=xlFormulas, lookat:=xlPart, _
                  SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                  MatchCase:=False, SearchFormat:=False)
37170     If Not (ObjectiveFunctionCell Is Nothing) Then GoTo FoundObjCell
          ' ROW ABOVE
37180     Set ObjectiveFunctionCell = _
              s.Rows(RowLBound).Find(What:="sumproduct", _
                  LookIn:=xlFormulas, lookat:=xlPart, _
                  SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                  MatchCase:=False, SearchFormat:=False)
37190     If Not (ObjectiveFunctionCell Is Nothing) Then GoTo FoundObjCell
          ' ROW BELOW
37200     Set ObjectiveFunctionCell = _
              s.Rows(RowUBound).Find(What:="sumproduct", _
                  LookIn:=xlFormulas, lookat:=xlPart, _
                  SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                  MatchCase:=False, SearchFormat:=False)
37210     If Not (ObjectiveFunctionCell Is Nothing) Then GoTo FoundObjCell
          
          ' NEXT SEARCH FOR ANY FORMULA
          ' SAME ROW AS MODEL TYPE
37220     Set ObjectiveFunctionCell = _
              s.Rows(ObjSenseCell.row).Find(What:="=", _
                  LookIn:=xlFormulas, lookat:=xlPart, _
                  SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                  MatchCase:=False, SearchFormat:=False)
37230     If Not (ObjectiveFunctionCell Is Nothing) Then GoTo FoundObjCell
          ' ROW ABOVE
37240     Set ObjectiveFunctionCell = _
              s.Rows(RowLBound).Find(What:="=", _
                  LookIn:=xlFormulas, lookat:=xlPart, _
                  SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                  MatchCase:=False, SearchFormat:=False)
37250     If Not (ObjectiveFunctionCell Is Nothing) Then GoTo FoundObjCell
          ' ROW BELOW
37260     Set ObjectiveFunctionCell = _
              s.Rows(RowUBound).Find(What:="=", _
                  LookIn:=xlFormulas, lookat:=xlPart, _
                  SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                  MatchCase:=False, SearchFormat:=False)
37270     If Not (ObjectiveFunctionCell Is Nothing) Then GoTo FoundObjCell

         
FoundObjCell:

37280     If ObjectiveFunctionCell Is Nothing Then
              ' Couldn't find it
37290         FindObjective = "SenseNoCell"
37300         Application.Cursor = xlDefault
              'Application.StatusBar = "OpenSolver: Determined objective sense, couldn't find obj. cell!"
37310         Application.StatusBar = False
37320         Exit Function
37330     Else
              ' Found something!
37340     End If
              
37350     FindObjective = "OK"
37360     Application.Cursor = xlDefault
          'Application.StatusBar = "OpenSolver: Found objective cell and sense."
37370     Application.StatusBar = False
          
End Function


'--------------------------------------------------------------------
' FindObjSense
' Helper function for FindObjective. Contains the settings for the
' right kind of search (values, don't match case)
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub FindObjSense(ByRef s As Worksheet, ByVal searchStr As String, ByRef result As Range)
37380     Set result = s.Cells.Find( _
                      What:=searchStr, After:=[a1], _
                      LookIn:=xlValues, SearchOrder:=xlByRows, _
                      SearchDirection:=xlPrevious, MatchCase:=False)
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
37390     If IsFirstTime Then
37400         Set DecisionVariables = Nothing
37410         Set Constraints = New Collection
37420     End If
          
          '----------------------------------------------------------------
          ' Phase One
          '   - Look for constraints (i.e. <=, >=, =) and add them if they
          '     seem at all interesting (i.e. LHS or RHS has precedents)
37430     Application.StatusBar = "OpenSolver:  Looking for constraints"
37440     On Error GoTo ConstraintErr
          
              '------------------------------------------------------------
              ' Stage One - Find all the "relational operator" cells on the
              ' sheet that the objective function is on.

              ' Find the cells
              Dim FoundLEQ As Range, FoundGEQ As Range, FoundEQ As Range
37450         If Not ObjectiveFunctionCell Is Nothing Then
37460             FindAllCells "<=", FoundLEQ, ObjectiveFunctionCell.Parent
37470             FindAllCells ">=", FoundGEQ, ObjectiveFunctionCell.Parent
37480             FindAllCells "=", FoundEQ, ObjectiveFunctionCell.Parent
37490         Else
37500             FindAllCells "<=", FoundLEQ, ActiveSheet
37510             FindAllCells ">=", FoundGEQ, ActiveSheet
37520             FindAllCells "=", FoundEQ, ActiveSheet
37530         End If
              
              ' Combine them as much as possible
              Dim AllCompOps As Range
37540         Set AllCompOps = FoundEQ
37550         Set AllCompOps = ProperUnion(AllCompOps, FoundLEQ)
37560         Set AllCompOps = ProperUnion(AllCompOps, FoundGEQ)
          
              '------------------------------------------------------------
              ' Stage Two - Now we have the blocks of constraints. For
              ' blocks that have more rows than columns, we can look left
              ' and right. For more columns than rows, up and down. For
              ' single cells, look in all directions
              Dim AreaIndex As Long
37570         For AreaIndex = 1 To AllCompOps.Areas.Count
                  ''Debug.Print "FindVarsAndCons: Con area" + Str(AreaIndex)
                  
                  ' Determine the shape of the area
                  Dim RowCount As Long, ColCount As Long
37580             RowCount = AllCompOps.Areas(AreaIndex).Rows.Count
37590             ColCount = AllCompOps.Areas(AreaIndex).Columns.Count
                  
                  ' Depending on the shape, set the LHS and RHS
                  ' NB: in the case of row of constraints, LHS refers to
                  '     the cells above the row, RHS the cells below.
                  Dim LHSs As Range, RHSs As Range
37600             If ColCount = 1 And RowCount > 1 Then
                      ' Vertical
                      ''Debug.Print "FindVarsAndCons: Vertical", AreaIndex
37610                 Set LHSs = AllCompOps.Areas(AreaIndex).Offset(0, -1)
37620                 Set RHSs = AllCompOps.Areas(AreaIndex).Offset(0, 1)
37630                 If CheckPrecedentCells(LHSs, RHSs) Then
                          ' Its ok, add them
37640                     AddRangeToConstraints LHSs, AllCompOps.Areas(AreaIndex), RHSs, True, Constraints
                          ''Debug.Print "FindVarsAndCons: Added", AreaIndex
37650                 End If
                      
37660             ElseIf ColCount > 1 And RowCount = 1 Then
                      ' Horizontal
                      ''Debug.Print "FindVarsAndCons: Horizontal", AreaIndex
37670                 Set LHSs = AllCompOps.Areas(AreaIndex).Offset(-1, 0)
37680                 Set RHSs = AllCompOps.Areas(AreaIndex).Offset(1, 0)
37690                 If CheckPrecedentCells(LHSs, RHSs) Then
                          ' Its ok, add them
37700                     AddRangeToConstraints LHSs, AllCompOps.Areas(AreaIndex), RHSs, False, Constraints
                          ''Debug.Print "FindVarsAndCons: Added", AreaIndex
37710                 End If
                      
37720             ElseIf ColCount = 1 And RowCount = 1 Then
                      ' Singleton
                      ''Debug.Print "FindVarsAndCons: Singleton", AreaIndex
                      ' Look Left-Right
37730                 Set LHSs = AllCompOps.Areas(AreaIndex).Offset(0, -1)
37740                 Set RHSs = AllCompOps.Areas(AreaIndex).Offset(0, 1)
37750                 If CheckPrecedentCells(LHSs, RHSs) Then
                          ' Its ok, add them
37760                     AddRangeToConstraints LHSs, AllCompOps.Areas(AreaIndex), RHSs, True, Constraints
                          ''Debug.Print "FindVarsAndCons: Added LR", AreaIndex
37770                 Else
                          ' Didn't find anything interesting LR, look UD
37780                     Set LHSs = AllCompOps.Areas(AreaIndex).Offset(-1, 0)
37790                     Set RHSs = AllCompOps.Areas(AreaIndex).Offset(1, 0)
37800                     If CheckPrecedentCells(LHSs, RHSs) Then
                              ' Its ok, add them
37810                         AddRangeToConstraints LHSs, AllCompOps.Areas(AreaIndex), RHSs, False, Constraints
                              ''Debug.Print "FindVarsAndCons: Added UD", AreaIndex
37820                     End If
37830                 End If
37840             Else
                      ' Block
                      ' TODO - Handle this somehow, if it has an application
                      ''Debug.Print "FindVarsAndCons: Block TODO", AreaIndex
37850             End If
                  
37860         Next AreaIndex

37870         RationaliseConstraints
              
              
          '----------------------------------------------------------------
          ' Phase Two
          '       - use precedents of objective function and constraints
          '         to find the set of possible decision variables
37880     Application.StatusBar = "OpenSolver: Searching for decision variables"
          
          ' NB: The first collection has cell addresses as keys, and reference
          '     counts as values.
          '     The second collection has cell addresses as key and value.
          '     We need this because, as far as I know, you can't iterate
          '     through the keys of a collection.
          Dim DecRefCount As New Collection, DecRefAddress As New Collection
          
          
              '------------------------------------------------------------
              ' Stage One - Objective function precedents
37890         On Error Resume Next
              ' Find the precedents of the objective function
37900         Set PrecedentCells = Nothing
37910         Set PrecedentCells = ObjectiveFunctionCell.Precedents
37920         If PrecedentCells Is Nothing Then
                  ' Objective function must be constant, this is not necessarily an issue
                  ''Debug.Print "FindVarsAndCons: Constant objective, skipping stage one of phase two"
37930         Else
                  ' Go through each precedent of the objective function, and set the
                  ' reference counter for it to 1
37940             On Error GoTo DecisionErr
37950             For Each CurPrecedent In PrecedentCells.Cells
37960                 If Not CurPrecedent.HasFormula Then
37970                     If Not TestKeyExists(DecRefCount, CurPrecedent.Address) Then
37980                         DecRefCount.Add 1, CurPrecedent.Address
37990                         DecRefAddress.Add CurPrecedent.Address, CurPrecedent.Address
                              ''Debug.Print "obj", CurPrecedent.Address, 1
38000                     End If
38010                 End If
38020             Next
38030         End If
              
              ''Debug.Print "FindVarsAndCons: End of Phase One Stage One"


              '------------------------------------------------------------
              ' Stage Two - Constraint precedents
38040         On Error Resume Next
              Dim curConstraint As CConstraint, RC As Integer
38050         For Each curConstraint In Constraints
                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  ' Look at LHS
38060             On Error Resume Next
38070             Set PrecedentCells = Nothing
38080             Set PrecedentCells = curConstraint.LHS.Precedents
38090             If Not PrecedentCells Is Nothing Then
38100                 For Each CurPrecedent In PrecedentCells.Cells
38110                     If TestKeyExists(DecRefCount, CurPrecedent.Address) Then
                              ' Exists, increment
                              ' Annoyingly, collections are read-only
                              ' I'm not sure how slow this is, but a way to
                              ' optimise would be to only do this if RC = 1, because
                              ' once RC hits two we will pick it anyway
38120                         RC = DecRefCount(CurPrecedent.Address)
38130                         DecRefCount.Remove CurPrecedent.Address
38140                         DecRefCount.Add RC + 1, CurPrecedent.Address
                              ''Debug.Print "con", CurPrecedent.Address, RC + 1
38150                     Else
                              ' Doesn't exist yet, start it at 1
38160                         If Not CurPrecedent.HasFormula Then
38170                             DecRefCount.Add 1, CurPrecedent.Address
38180                             DecRefAddress.Add CurPrecedent.Address, CurPrecedent.Address
                                  ''Debug.Print "connew", CurPrecedent.Address, 1
38190                         End If
38200                     End If
38210                 Next
38220             End If
                  
                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  ' Look at RHS
38230             On Error Resume Next
38240             Set PrecedentCells = Nothing
38250             Set PrecedentCells = curConstraint.RHS.Precedents
38260             If Not PrecedentCells Is Nothing Then
38270                 For Each CurPrecedent In PrecedentCells.Cells
38280                     If TestKeyExists(DecRefCount, CurPrecedent.Address) Then
                              ' Exists, increment
38290                         DecRefCount(CurPrecedent.Address) = DecRefCount(CurPrecedent.Address) + 1
38300                     Else
                              ' Doesn't exist yet, start it at 1
38310                         If Not CurPrecedent.HasFormula Then
38320                             DecRefCount.Add 1, CurPrecedent.Address
38330                             DecRefAddress.Add CurPrecedent.Address, CurPrecedent.Address
38340                         End If
38350                     End If
38360                 Next
38370             End If
              
38380         Next
          
          
          
          '----------------------------------------------------------------
          ' Phase Three
          '   If a decision variable has only been referenced once, we
          '   really can't be sure it is a decision variable - constants
          '   are also referenced once. So take anything that is two or
          '   or higher
38390     Application.StatusBar = "OpenSolver: Selecting most likely decision variables"
              
              Dim addressKey As Variant
              
38400         For Each addressKey In DecRefAddress
38410             If DecRefCount(addressKey) >= 2 Then
38420                 Set DecisionVariables = ProperUnion(DecisionVariables, ActiveSheet.Range(CStr(addressKey)))
38430             End If
38440         Next
              
          '----------------------------------------------------------------
          ' Phase Four
          '   - Look for type restrictions on decision variables
38450     Application.StatusBar = "OpenSolver: Looking for variable type restrictions"
              Dim CurDecVar As Range, PossibleType As String
          
38460         For Each CurDecVar In DecisionVariables
                  ' Look below it to see if there is type information
38470             PossibleType = LCase(Trim(CurDecVar.Offset(1, 0).value))
38480             If PossibleType = "integer" Or PossibleType = "int" Or PossibleType = "i" Then
                      ' This variable is integer
38490                 Set curConstraint = New CConstraint
38500                 With curConstraint
38510                     Set .LHS = CurDecVar
38520                     Set .RHS = Nothing
38530                     Set .Relation = Nothing
38540                     .ConstraintType = "int"
38550                 End With
38560                 Constraints.Add curConstraint ', curConstraint.GetKey
38570             ElseIf PossibleType = "binary" Or PossibleType = "bin" Or PossibleType = "b" Then
                      ' This variable is BINARY
38580                 Set curConstraint = New CConstraint
38590                 With curConstraint
38600                     Set .LHS = CurDecVar
38610                     Set .RHS = Nothing
38620                     Set .Relation = Nothing
38630                     .ConstraintType = "bin"
38640                 End With
38650                 Constraints.Add curConstraint ', curConstraint.GetKey
38660             End If
38670         Next
          
          
          '----------------------------------------------------------------
          ' Finished!
38680     Application.StatusBar = "OpenSolver: Rationalising constraints"
          'RationaliseConstraints ' Make constraints prettier ' BROKEN IN 1.3 COS USES KEYS
38690     FindVarsAndCons = True
          'Application.StatusBar = "OpenSolver: Finished finding decision variables and constraints"
38700     Application.StatusBar = False
38710     Exit Function
          
DecisionErr:
          ' Error occurred while trying to find decision variables
38720     MsgBox "Error: an issue arose while finding decision variables." + vbNewLine + _
                 "Error number:" + str(Err.Number) + vbNewLine + _
                 "Error description: " + Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")"), _
                 vbExclamation Or vbOKOnly, "AutoModel"
38730     FindVarsAndCons = False
38740     Application.StatusBar = False
38750     Exit Function
          
ConstraintErr:
          ' Error occurred while trying to find constraints
38760     MsgBox "Error: an issue arose while finding constraints." + vbNewLine + _
                 "Error number:" + str(Err.Number) + vbNewLine + _
                 "Error description: " + Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")"), _
                 vbExclamation Or vbOKOnly, "AutoModel"
38770     FindVarsAndCons = False
38780     Application.StatusBar = False
38790     Exit Function
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
38800     Set FoundCells = Nothing
          
          ' Find first cell that meets requirements
38810     Set LastCell = sheet.Cells.Find(What:=searchStr, _
                                  After:=[a1], _
                                  LookIn:=xlValues, _
                                  SearchOrder:=XlSearchOrder.xlByRows, _
                                  lookat:=XlLookAt.xlWhole, _
                                  SearchDirection:=XlSearchDirection.xlNext _
                                  )
38820     Set FirstCell = LastCell
38830     If LastCell Is Nothing Then Exit Sub ' If not even one, stop immediately
          
38840     Do
              ' Add it...
38850         If FoundCells Is Nothing Then
38860             Set FoundCells = Range(LastCell.Address)
38870         Else
38880             Set FoundCells = Union(FoundCells, LastCell)
38890         End If
              ' ...and find next
38900         Set LastCell = sheet.Cells.FindNext(LastCell)
              ' Loop until we don't find another cell, or until we get back to the
              ' initial cell we found
38910     Loop While (Not LastCell Is Nothing) And (FirstCell.Address <> LastCell.Address)
          
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
38920     Set BothSides = Union(LHSs, RHSs)

38930     For Each CurCell In BothSides.Cells
              ' Check if cell has a precedent
              ' If no precedents, error is thrown
38940         Err.Clear
38950         On Error Resume Next
38960         Set PrecCells = CurCell.Precedents
38970         If Err.Number = 1004 Then
                  ' No precedents for this cell
38980         ElseIf Err.Number = 0 Then
                  ' There is a precedent
38990             CheckPrecedentCells = True
39000             Exit Function
39010         End If
39020     Next
          
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
39030     CellCount = LHSs.Count
          
          Dim i As Long
          Dim LHSi As Range, RELi As Range, RHSi As Range
          Dim NewConstraint As CConstraint
          
          ' For each cell
39040     For i = 1 To CellCount
39050         If IsVertical Then
39060             Set LHSi = LHSs(RowIndex:=i)
39070             Set RELi = RelRange(RowIndex:=i)
39080             Set RHSi = RHSs(RowIndex:=i)
39090         Else
39100             Set LHSi = LHSs(ColumnIndex:=i)
39110             Set RELi = RelRange(ColumnIndex:=i)
39120             Set RHSi = RHSs(ColumnIndex:=i)
39130         End If
              ' Check if added
39140         If Not TestKeyExists(Constraints, RELi.Address) Then
                  ' Constraint not already added - so add it!
39150             Set NewConstraint = New CConstraint
39160             With NewConstraint
39170                 Set .LHS = LHSi
39180                 Set .RHS = RHSi
39190                 Set .Relation = RELi
39200                 .ConstraintType = RELi.value
                      ''Debug.Print "Con added:", .LHS.Address, .ConstraintType, .RHS.Address
39210             End With
39220             Constraints.Add NewConstraint ', NewConstraint.GetKey
39230         Else
                  ' Added already
                  'Debug.Print "Con already added", LHSi.Address, RELi.Value, RHSi.Address
39240         End If
39250     Next i
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
39260     On Error GoTo 0
          Dim Relations(1 To 5) As String
39270     Relations(1) = "="
39280     Relations(2) = "<="
39290     Relations(3) = ">="
39300     Relations(4) = "int"
39310     Relations(5) = "bin"
          
          Dim NewConstraints As New Collection
          Dim NewConstraint As CConstraint
          Dim conKeys As Collection
          Dim curCon As CConstraint
          Dim UnionRange As Range
          Dim AreaIndex As Long
          Dim LHSunion As Range, RHSunion As Range, RELunion As Range
          Dim CurCell As Range
          Dim NumConstraints As Integer
          Dim conNum() As Variant
          
39320     NumConstraints = Constraints.Count
                    
39330     ReDim conNum(NumConstraints + 1) As Variant
          
          Dim i As Integer, j As Long
39340     For i = 1 To 5
              ' Find all constraints of type Relations(i)...
39350         Set conKeys = New Collection
                
                Dim c As Long, k As Integer
                
                'Initialise the indices of the constraints to be empty
39360           For k = 1 To NumConstraints
39370               conNum(k) = Empty
39380           Next k
                
39390           c = 1
39400           k = 1
                
39410         For Each curCon In Constraints
                  ' ... and add a reference to them in a collection and save the index of the constraint
39420             If curCon.ConstraintType = Relations(i) Then
39430                 If 1 <= i And i <= 3 Then
39440                       conKeys.Add curCon.Relation.Address, curCon.Relation.Address
39450                       conNum(k) = c
39460                       k = k + 1
39470                 End If
39480                 If 4 <= i And i <= 5 Then
39490                       conKeys.Add curCon.LHS.Address, curCon.LHS.Address
39500                       conNum(k) = c
39510                       k = k + 1
39520                 End If
39530             End If
39540             c = c + 1
39550         Next

39560         If conKeys.Count > 0 Then
39570             Set UnionRange = Nothing
                  'Find the number of constraints for relations(i)
39580             k = 1
                  
39590             Do While conNum(k) <> ""
39600                 k = k + 1
39610             Loop
                
                ' Build up a range that is the union of all the constraints
39620             For j = 1 To (k - 1)
39630                 If 1 <= i And i <= 3 Then Set UnionRange = ProperUnion(UnionRange, Constraints(conNum(j)).Relation)
39640                 If 4 <= i And i <= 5 Then Set UnionRange = ProperUnion(UnionRange, Constraints(conNum(j)).LHS)
39650             Next j
                  ' Now iterate through each area of the range - each represents a block
                  ' of constraints that are next to each other, with the same relation
                  
                  Dim SavedNum As Long, num As Long
                  
39660             SavedNum = 1
                  
39670             For AreaIndex = 1 To UnionRange.Areas.Count
                      'Debug.Print "Relation " + Relations(i) + ", area" + str(AreaIndex)
39680                 num = SavedNum
39690                 Set LHSunion = Nothing
39700                 Set RHSunion = Nothing
39710                 Set RELunion = Nothing

39720                 For Each CurCell In UnionRange.Areas(AreaIndex).Cells
39730                     Set LHSunion = ProperUnion(LHSunion, Constraints(conNum(num)).LHS)
39740                     Set RHSunion = ProperUnion(RHSunion, Constraints(conNum(num)).RHS)
39750                     Set RELunion = ProperUnion(RELunion, Constraints(conNum(num)).Relation)
39760                     num = num + 1
39770                 Next
39780                 SavedNum = num

39790                 Set NewConstraint = New CConstraint
39800                 With NewConstraint
39810                     Set .LHS = LHSunion
39820                     Set .RHS = RHSunion
39830                     Set .Relation = RELunion
39840                     .ConstraintType = Relations(i)
                          ''Debug.Print "SuperCon added:", .LHS.Address, .ConstraintType, .RHS.Address
39850                 End With
39860                 If 1 <= i And i <= 3 Then NewConstraints.Add NewConstraint, NewConstraint.Relation.Address
39870                 If 4 <= i And i <= 5 Then NewConstraints.Add NewConstraint, NewConstraint.LHS.Address
39880             Next
39890         End If
39900     Next i
          
          ' Update old constraints
39910     Set Constraints = NewConstraints
End Sub




'--------------------------------------------------------------------
' BuildModel
' Turn this object into an on-sheet model
'
' Written by:       IRD
'--------------------------------------------------------------------
Sub BuildModel()
          ' Decision variables
39920     If DecisionVariables Is Nothing Then
39930         DeleteSolverNameOnSheet "adj"
39940     Else
39950         SetSolverNamedRangeOnSheet "adj", DecisionVariables
39960     End If
          
          ' Objective function cell
39970     If ObjectiveFunctionCell Is Nothing Then
39980         DeleteSolverNameOnSheet "opt"
39990     Else
40000         SetSolverNamedRangeOnSheet "opt", ObjectiveFunctionCell
40010     End If
          
          ' Objective type
40020     If ObjectiveSense = MinimiseObjective Then
40030         SetSolverNameOnSheet "typ", "=2"
40040     ElseIf ObjectiveSense = MaximiseObjective Then
40050         SetSolverNameOnSheet "typ", "=1"
40060     ElseIf ObjectiveSense = TargetObjective Then
40070         SetSolverNameOnSheet "typ", "=3"
40080         SetSolverNameOnSheet "val", "=" + Trim(str(ObjectiveTarget)) ' AJM Changed from CStr to Str as we must store this using US numbers, and we must strip any leading space
40090     End If
          
          ' Check int/bin constraints not set on non-decision variables
          Dim c As Long, intersection As Range
40100     For c = 1 To Constraints.Count
40110         If Constraints(c).ConstraintType = "bin" _
              Or Constraints(c).ConstraintType = "int" Then
40120             Set intersection = Intersect(Constraints(c).LHS, DecisionVariables)
40130             If intersection Is Nothing Then
40140                 MsgBox "Warning: an int/bin constraint has been set on cells that are not decision variables. OpenSolver does not support this yet. An error message will appear if you have selected to show the Model."
40150             Else
40160                 If intersection.Count <> Constraints(c).LHS.Count Then
40170                     MsgBox "Warning: an int/bin constraint has been set on cells that are not decision variables. OpenSolver does not support this yet. An error message will appear if you have selected to show the Model."
40180                 End If
40190             End If
40200         End If
40210     Next c
          
          ' Constraints
40220     SetSolverNameOnSheet "num", "=" + CStr(Constraints.Count)
40230     For c = 1 To Constraints.Count
40240         If Constraints(c).ConstraintType = "bin" Then
40250             SetSolverNamedRangeOnSheet "lhs" & CStr(c), Constraints(c).LHS
40260             SetSolverNameOnSheet "rel" & CStr(c), "=5"
40270             SetSolverNameOnSheet "rhs" & CStr(c), "=binary"
40280         ElseIf Constraints(c).ConstraintType = "int" Then
40290             SetSolverNamedRangeOnSheet "lhs" & CStr(c), Constraints(c).LHS
40300             SetSolverNameOnSheet "rel" & CStr(c), "=4"
40310             SetSolverNameOnSheet "rhs" & CStr(c), "=integer"
40320         ElseIf Constraints(c).ConstraintType = "alldiff" Then
40330             SetSolverNamedRangeOnSheet "lhs" & CStr(c), Constraints(c).LHS
40340             SetSolverNameOnSheet "rel" & CStr(c), "=6"
40350             SetSolverNameOnSheet "rhs" & CStr(c), "=alldiff"
40360         Else
40370             SetSolverNamedRangeOnSheet "lhs" & CStr(c), Constraints(c).LHS
40380             If Constraints(c).RHS Is Nothing Then
40390                 If Mid(Constraints(c).RHSstring, 1, 1) = "=" Then
40400                     SetSolverNameOnSheet "rhs" & CStr(c), Constraints(c).RHSstring
40410                 Else
40420                     SetSolverNameOnSheet "rhs" & CStr(c), "=" & Constraints(c).RHSstring
40430                 End If
40440             Else
40450                 SetSolverNamedRangeOnSheet "rhs" & CStr(c), Constraints(c).RHS
40460             End If
40470             If Constraints(c).ConstraintType = ">=" Then SetSolverNameOnSheet "rel" & CStr(c), "=3"
40480             If Constraints(c).ConstraintType = ">" Then SetSolverNameOnSheet "rel" & CStr(c), "=3"
40490             If Constraints(c).ConstraintType = "<=" Then SetSolverNameOnSheet "rel" & CStr(c), "=1"
40500             If Constraints(c).ConstraintType = "<" Then SetSolverNameOnSheet "rel" & CStr(c), "=1"
40510             If Constraints(c).ConstraintType = "=" Then SetSolverNameOnSheet "rel" & CStr(c), "=2"
40520             If Constraints(c).ConstraintType = "'=" Then SetSolverNameOnSheet "rel" & CStr(c), "=2"
40530         End If
40540     Next c
          
40550     If Duals Is Nothing Then
40560         DeleteNameOnSheet "OpenSolver_Duals"
40570     Else
40580         SetNamedRangeOnSheet "OpenSolver_Duals", Duals
40590     End If
          
          ' Set the default Solver properties
          ' These values must be given in English, not using the local language (eg with . and not ,)
40600     SetSolverNameOnSheet "neg", IIf(NonNegativityAssumption, "=1", "=2")
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
          
40610     With NewConstraint
40620         Set .LHS = newLHS
40630         Set .Relation = Nothing
40640         Set .RHS = newRHS
40650         .RHSstring = newRHSstring
40660         .ConstraintType = newType
40670     End With
          
40680     Constraints.Add NewConstraint ', NewConstraint.GetKey
          
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
40690     showNameRange = frmModel.chkNameRange.value
40700     If showNameRange Then SearchRangeName_DestroyCache
          
40710     oldLI = lst.ListIndex
40720     lst.Clear
40730     lst.AddItem "<Add new constraint>"
40740     For c = 1 To Constraints.Count
40750         With Constraints(c)
40760             CS = GetDisplayAddress(.LHS, showNameRange) + " " + .ConstraintType
40770             If Not (.RHS Is Nothing) Then
                      ' This is not a type restriction, so it has a RHS
40780                 CS = CS + " " + GetDisplayAddress(.RHS, showNameRange)
40790             ElseIf .RHS Is Nothing And .RHSstring <> "" Then
40800                 If Mid(.RHSstring, 1, 1) = "=" Then
40810                     CS = CS + " " + RemoveActiveSheetNameFromString(Mid(.RHSstring, 2, Len(.RHSstring)))
40820                 Else
40830                     CS = CS + " " + RemoveActiveSheetNameFromString(.RHSstring)
40840                 End If
40850             End If
40860         End With
40870         lst.AddItem CS
40880     Next c
40890     If oldLI >= 0 And oldLI <= lst.ListCount - 1 Then
40900         lst.ListIndex = oldLI
40910     Else
40920         lst.ListIndex = lst.ListCount - 1
40930     End If
End Sub


'--------------------------------------------------------------------
' LoadFromSheet
' Loads the Solver-style model
' Borrowed a lot from COpenSolver.BuildModelFromSolverData
'
' Written by:       IRD
'--------------------------------------------------------------------
Public Sub LoadFromSheet()

40940     On Error Resume Next ' TODO: This should not run under a blank RESUME NEXT
40950     Set Constraints = New Collection


          Dim sheetName As String
40960     sheetName = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!"   ' Must include quotes to handle sheet names with spaces

          ' ADJUSTABLE CELLS
40970     Set DecisionVariables = Range(sheetName & "solver_adj")
          
          ' NON-NEGATIVITY
          Dim s As String
40980     NonNegativityAssumption = False
40990     If GetNameValueIfExists(ActiveWorkbook, sheetName & "solver_neg", s) Then NonNegativityAssumption = (s = "1")
           
          ' OBJECTIVE SENSE (AND TARGET)
          Dim temp As Integer
41000     ObjectiveSense = UnknownObjectiveSense
41010     If GetNamedIntegerIfExists(ActiveWorkbook, sheetName & "solver_typ", temp) Then ObjectiveSense = temp
41020     ObjectiveTarget = 0
41030     GetNamedNumericValueIfExists ActiveWorkbook, sheetName & "solver_val", ObjectiveTarget
          
          ' OBJECTIVE CELL
41040     Set ObjectiveFunctionCell = Nothing
41050     Set ObjectiveFunctionCell = ActiveWorkbook.Names(sheetName & "solver_opt").RefersToRange
          
          ' CONSTRAINTS
          Dim constraint As Long, NumConstraints As Long
          
41060     NumConstraints = Val(Mid(Names(sheetName & "solver_num"), 2))
41070     For constraint = 1 To NumConstraints
              ' Note: Solver enforces the following requirements.
              ' The LHS mut be a range with one or more cells
              ' The RHS can be either:
              '   A single-cell range (=A4)
              '   A multi-cell range of the same size as the LHS (=A4:B5)
              '   A single constant value (eg =2)
              '   A formula returning a single value (eg =sin(A4))
              
              Dim isRangeLHS As Boolean, valLHS As Double, rLHS As Range, LHSRefersToError As Boolean, LHSRefersToFormula As Boolean, sRefersToLHS As String, LHSisMissing As Boolean
              Dim isRangeRHS As Boolean, valRHS As Double, rRHS As Range, RHSRefersToError As Boolean, RHSRefersToFormula As Boolean, sRefersToRHS As String, RHSisMissing As Boolean
              
41080         GetNameAsValueOrRange ActiveWorkbook, sheetName & "solver_lhs" & constraint, LHSisMissing, isRangeLHS, rLHS, LHSRefersToFormula, LHSRefersToError, sRefersToLHS, valLHS
41090         If (LHSisMissing) Or (Not isRangeLHS) Then GoTo SkipCon
              
              Dim rel As Integer
41100         rel = Val(Mid(Names(sheetName & "solver_rel" & constraint), 2))  ' Either <=, = or >=, or int or bin (or 'alldif' in Excel 2010)
              
41110         If rel = RelationINT Or rel = RelationBIN Or rel = RelationAllDiff Then
                  ' Make the LHS variables integer or binary
41120             If rel = RelationINT Then AddConstraint rLHS, Nothing, "", "int"
41130             If rel = RelationBIN Then AddConstraint rLHS, Nothing, "", "bin"
41140             If rel = RelationAllDiff Then AddConstraint rLHS, Nothing, "", "alldiff"
41150         Else
                  ' Constraint is a full equation with a RHS
41160             GetNameAsValueOrRange ActiveWorkbook, sheetName & "solver_rhs" & constraint, RHSisMissing, isRangeRHS, rRHS, RHSRefersToFormula, RHSRefersToError, sRefersToRHS, valRHS
41170             If RHSisMissing Then GoTo SkipCon
                  
41180             If rel <> RelationLE And rel <> RelationEQ And rel <> RelationGE Then GoTo SkipCon

41190             If Not isRangeRHS Then
41200                 If rel = RelationEQ Then AddConstraint rLHS, Nothing, Mid(Names(sheetName & "solver_rhs" & constraint), 2), "="
41210                 If rel = RelationLE Then AddConstraint rLHS, Nothing, Mid(Names(sheetName & "solver_rhs" & constraint), 2), "<="
41220                 If rel = RelationGE Then AddConstraint rLHS, Nothing, Mid(Names(sheetName & "solver_rhs" & constraint), 2), ">="
41230             Else
41240                 If rel = RelationEQ Then AddConstraint rLHS, rRHS, "", "="
41250                 If rel = RelationLE Then AddConstraint rLHS, rRHS, "", "<="
41260                 If rel = RelationGE Then AddConstraint rLHS, rRHS, "", ">="
41270             End If
41280         End If
SkipCon:
41290     Next constraint
          
          ' DUALS
41300     Set Duals = Range(sheetName & "OpenSolver_Duals")
          
ExitSub:
End Sub

