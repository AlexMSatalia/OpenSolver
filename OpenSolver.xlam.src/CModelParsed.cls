VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CModelParsed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Takes a Solver model and
' 1. Validates it according to the same rules Solver uses
' 2. Converts it to a form that will be used by the output model builders
'==============================================================================
Option Explicit


Public Formulae As Collection
Public VarTypeMap As Collection
Public AdjCellNameMap As Collection
Public LHSKeys As Collection
Public RHSKeys As Collection
Public RELs As Collection

Public lngMaxDepth As Long

Dim AdjDependents As Range

Private s As COpenSolver

'==============================================================================
Public Function ProcessSolverModel(OpenSolver As COpenSolver) As Boolean
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
3         Set s = OpenSolver
4         Set AdjDependents = s.AdjustableCells.Dependents
          
5         Set SolverModelParsed.SheetNameMap = New Collection
6         Set SolverModelParsed.SheetNameMapReverse = New Collection

          ' =========================================
          ' Conduct parsed-specific processing
7         ProcessVarTypes
8         ProcessSolverModel = ParseFormulae()

ExitFunction:
9         If RaiseError Then RethrowError
10        Exit Function

ErrorHandler:
11        If Not ReportError("CModelParsed", "ProcessSolverModel") Then Resume
12        RaiseError = True
13        GoTo ExitFunction
End Function

Private Sub ProcessVarTypes()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' Get var types of constraints
3         Set VarTypeMap = New Collection
4         Set AdjCellNameMap = New Collection
          Dim curVarType As Long, lngArea As Long, cellName As String, c As Range
5         For lngArea = 1 To s.AdjustableCells.Areas.Count
6             For Each c In s.AdjustableCells.Areas(lngArea)
7                 cellName = ConvertCellToStandardName(c)
8                 If Not TestKeyExists(AdjCellNameMap, cellName) Then
9                     AdjCellNameMap.Add c, cellName
10                    curVarType = VarContinuous
11                    If Not (s.IntegerCellsRange Is Nothing) Then
12                        If Not (Intersect(c, s.IntegerCellsRange) Is Nothing) Then curVarType = VarInteger
13                    End If
14                    If Not (s.BinaryCellsRange Is Nothing) Then
15                        If Not (Intersect(c, s.BinaryCellsRange) Is Nothing) Then curVarType = VarBinary
16                    End If
17                    VarTypeMap.Add curVarType, cellName
18                End If
19            Next
20        Next

ExitSub:
21        If RaiseError Then RethrowError
22        Exit Sub

ErrorHandler:
23        If Not ReportError("CModelParsed", "ProcessVarTypes") Then Resume
24        RaiseError = True
25        GoTo ExitSub
End Sub

Private Function ParseFormulae() As Boolean
      ' Parses all constraints and objective formulae to create formulae that depend only on variable cells and constants
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          
          ' Create a collection for the spreadsheet's formulae (DAG)
3         Set Formulae = New Collection
          ' Remember the greatest depth of the DAG
4         lngMaxDepth = 0

          '==========================================================================
          ' STEP 1. Parse objective cell
          Dim objresult As String
          
5         If Not s.ObjRange Is Nothing Then
6             objresult = ExpandFormula(s.ObjRange.Formula, s.ObjRange, 0)
              
7             If objresult Like "*Error*" Then
8                 RaiseUserError "The objective contains an error"
9             End If
10            CleanFormulae 1, Formulae.Count
11        End If
          
          '==========================================================================
          ' STEP 2. Parse constraints
12        Set LHSKeys = New Collection
13        Set RHSKeys = New Collection
14        Set RELs = New Collection

          Dim i As Long
          Dim Count As Long
15        Count = 1
16        For i = 1 To s.NumConstraints
              Dim lngFormulaeCountBefore As Long
              Dim cRow As Long, cCol As Long, cRowCount As Long, cColCount As Long
              Dim varLHSFormulae As Variant, rngLHS As Range, curLHSFormulae As String
17            Set rngLHS = s.LHSRange(i)
              
              Dim rel As RelationConsts
18            rel = s.Relation(i)
19            If Not RelationHasRHS(rel) Then GoTo NextCons

              Dim strRHS As String, rngRHS As Range
20            Set rngRHS = s.RHSRange(i)
21            strRHS = s.RHSFormula(i)
              
22            lngFormulaeCountBefore = Formulae.Count
              
              ' Possibilities
              ' LHS is one cell, RHS is one cell/formula/value
              ' LHS is a range (any shape), RHS is one cell/formula/value
              ' LHS is a range (any shape), RHS is a range of same size, not necessarily same shape
              ' Because mismatched shapes are cruel, we will just throw an error if we find them
              ' OpenSolver1 just threw a general error.
              
23            cRowCount = rngLHS.Rows.Count
24            cColCount = rngLHS.Columns.Count
25            varLHSFormulae = rngLHS.Formula
                      
              Dim singleRngLHS As Range
26            For cRow = 1 To cRowCount
27                For cCol = 1 To cColCount
28                    Count = Count + 1
                      
                      ' Get LHS
29                    If cRowCount <> 1 Or cColCount <> 1 Then
30                        curLHSFormulae = varLHSFormulae(cRow, cCol)
31                        Set singleRngLHS = rngLHS(cRow, cCol)
32                    Else
33                        curLHSFormulae = varLHSFormulae
34                        Set singleRngLHS = rngLHS
35                    End If
                      
                      ' Get RHS
                      Dim rngRHSCell As Range, strRHSCell As String
36                    Select Case s.RHSType(i)
                      ' 1.1 RHS is a single cell
                      Case SingleCellRange
37                        Set rngRHSCell = rngRHS
38                        strRHSCell = rngRHSCell.Formula
                      ' 1.2 RHS is multiple cells
39                    Case MultiCellRange
40                        Set rngRHSCell = rngRHS.Cells(cRow, cCol)
41                        strRHSCell = rngRHSCell.Formula
                      ' 1.3 RHS is value or formula
42                    Case constant, Formula
43                        strRHSCell = "=" & strRHS
44                        Set rngRHSCell = Nothing
45                    End Select

                      ' If LHS is constant, then check if RHS is constant
46                    If CellIsConstant(singleRngLHS) Then
                          Dim isConst As Boolean
47                        isConst = False
                           
48                        Select Case s.RHSType(i)
                          Case constant
49                            isConst = True
50                        Case Formula
51                            isConst = StringFormulaIsConstant(strRHSCell)
52                        Case SingleCellRange, MultiCellRange
53                            isConst = CellIsConstant(rngRHSCell)
54                        End Select
                          
                          ' If both sides are constant we need to see if the constraint is valid
55                        If isConst Then
                              Dim isValid As Boolean, LHSValue As Double, RHSValue As Double
56                            isValid = True
                              
57                            LHSValue = singleRngLHS.Value2
58                            If Len(strRHSCell) = 0 Then
59                                RHSValue = 0
60                            Else
61                                RHSValue = s.sheet.Evaluate(strRHSCell)
62                            End If
                              
                              ' Test LHS and RHS in the constraint
63                            Select Case rel
                              Case RelationGE
64                                If LHSValue < RHSValue Then
65                                    isValid = False
66                                End If
67                            Case RelationLE
68                                If LHSValue > RHSValue Then
69                                    isValid = False
70                                End If
71                            Case RelationEQ
72                                If LHSValue <> RHSValue Then
73                                    isValid = False
74                                End If
75                            End Select
                              
                              ' If constraint is not satisfied, the model is infeasible
76                            If Not isValid Then
77                                s.SolveStatus = OpenSolverResult.Infeasible
78                                s.SolveStatusString = "Infeasible"
79                                s.SolveStatusComment = _
                                      "The model contains an Excel constraint which does not depend on the decision variables and is not satisfied." & vbCrLf & _
                                      "Constraint specifies: " & _
                                      "LHS=" & singleRngLHS.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "=" & LHSValue & _
                                      "   " & RelationEnumToString(rel) & "   " & _
                                      "RHS=" & "'" & strRHSCell & "'" & "=" & RHSValue
80                                ParseFormulae = False
81                                GoTo ExitFunction
82                            Else
                                  ' Otherwise the constraint is trivially satisfied, we can ignore it
83                                GoTo NextCol
84                            End If
85                        End If
86                    End If
                      
                      ' Parse LHS
87                    LHSKeys.Add ExpandFormula(curLHSFormulae, rngLHS(cRow, cCol), i)
                      
                      ' Parse RHS
88                    RHSKeys.Add ExpandFormula(strRHSCell, rngRHSCell, i)
                      
                      ' Determine appropriate relation
89                    RELs.Add rel
NextCol:
90                Next cCol
91            Next cRow
              
              ' Clean Formulae
92            CleanFormulae lngFormulaeCountBefore + 1, Formulae.Count, i
              
NextCons:
93        Next i
          
          '==========================================================================
          ' STEP 3. Make all possible substitutions in formula tree if required by solver
94        If DoBackSubstitution(s.Solver) Then
          
              Dim lngCurDepth As Long
              Dim lngIndex As Long, objCurNode As CFormula, strParentAdr As Variant
              
95            For lngCurDepth = lngMaxDepth To 0 Step -1
96                For lngIndex = 1 To Formulae.Count
97                    If Formulae(lngIndex).lngDepth = lngCurDepth Then
98                        Set objCurNode = Formulae(lngIndex)
                          
99                        For Each strParentAdr In objCurNode.Parents
100                           With Formulae(strParentAdr)
101                               .strFormulaParsed = Replace(.strFormulaParsed, objCurNode.strAddress, "(" & objCurNode.strFormulaParsed & ")")
102                           End With
103                       Next
104                   End If
105               Next lngIndex
106           Next lngCurDepth
107       End If
108       ParseFormulae = True

ExitFunction:
109       Application.StatusBar = False
110       If RaiseError Then RethrowError
111       Exit Function

ErrorHandler:
112       If Not ReportError("CModelParsed", "ParseFormulae") Then Resume
113       RaiseError = True
114       GoTo ExitFunction
End Function

Function DoBackSubstitution(Solver As ISolver) As Boolean
1         DoBackSubstitution = False
End Function

' Checks whether a cell is constant i.e. not a decision variable and does not depend on a decision variable
Function CellIsConstant(cell As Range) As Boolean
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' Check if cell is a variable cell
3         If cell.Parent.Name = s.AdjustableCells.Parent.Name Then
4             If Not Intersect(cell, s.AdjustableCells) Is Nothing Then
5                 CellIsConstant = False
6                 GoTo ExitFunction
7             End If
8         End If
          
          ' Check if dependent on decision variable
9         If cell.HasFormula Then
              
              ' Try and get cell precedents. If the precedents are in another sheet we get stuck
              ' TODO: We could try this: http://www.mrexcel.com/forum/excel-questions/371863-test-if-cell-has-dependents-precedents.html#post1851288
              Dim CellPrecedents As Range
10            On Error Resume Next
11            Set CellPrecedents = cell.Precedents
12            If Err.Number = 1004 Then
13                CellIsConstant = True
14                GoTo ExitFunction
15            End If
16            On Error GoTo ErrorHandler
              
17            If cell.Precedents.Parent.Name = s.AdjustableCells.Parent.Name Then
18                If Not Intersect(cell.Precedents, s.AdjustableCells) Is Nothing Then
19                    CellIsConstant = False
20                    GoTo ExitFunction
21                End If
22            End If
23        End If
24        CellIsConstant = True

ExitFunction:
25        If RaiseError Then RethrowError
26        Exit Function

ErrorHandler:
27        If Not ReportError("CModelParsed", "CellIsConstant") Then Resume
28        RaiseError = True
29        GoTo ExitFunction
End Function

' Checks whether a string formula depends on variable cells or is simply constant.
' If constant, we can evaluate it directly
Function StringFormulaIsConstant(strFormula As String) As Boolean
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim tksFormula As Tokens
3         Set tksFormula = ParseFormula(strFormula)
          
          Dim i As Long, c As Range, tkn As Token, CheckForColon As Boolean
          Dim LookingForEndOfRef As Boolean, strSheetPrefix As String, BuildingRef As String, rngBuildingRef As Range
          
          ' We loop through tokens looking for cell references.
          ' If we find any, we test them to see if they depend on the variable cells
4         For i = 1 To tksFormula.Count
5             Set tkn = tksFormula.Item(i)
              
6             If CheckForColon = True And tkn.Text <> ":" Then
                  ' We don't hit colon right after the reference, we abandon
                  ' search for another reference
7                 LookingForEndOfRef = False
8                 CheckForColon = False
9                 If Len(strSheetPrefix) > 0 Then BuildingRef = strSheetPrefix + "!" + BuildingRef
10                Set rngBuildingRef = s.sheet.Range(BuildingRef)
                  
                  ' Cell reference found, check if constant
11                If rngBuildingRef.Count > 1 Then
12                    For Each c In rngBuildingRef
13                        If Not CellIsConstant(c) Then GoTo NotConstant
14                    Next c
15                Else
16                    If Not CellIsConstant(rngBuildingRef) Then GoTo NotConstant
17                End If
                  
                  ' Reset the sheet we are tracking
18                strSheetPrefix = vbNullString
19            End If
          
              ' Decide what to insert based on token type
20            Select Case tkn.TokenType
                  Case TokenType.ErrorText
                      ' TODO: Can't handle that, throw error
21                    RaiseGeneralError "An error token was found while parsing the formulae."
                      
22                Case TokenType.Reference
                      ' Are we trying to complete a range?
23                    If LookingForEndOfRef Then
                          ' We were, so finish building it...
24                        BuildingRef = BuildingRef + tkn.Text
                          ' ... and stop trying to build it more
25                        LookingForEndOfRef = False
26                        CheckForColon = False
                          
                          ' Make sure this range is on the right sheet
27                        If Len(strSheetPrefix) > 0 Then BuildingRef = strSheetPrefix + "!" + BuildingRef
                          
                          ' Cell reference found, check if constant
28                        Set rngBuildingRef = s.sheet.Range(BuildingRef)
29                        For Each c In rngBuildingRef
30                            If Not Intersect(c.Precedents, s.AdjustableCells) Is Nothing Then GoTo NotConstant
31                        Next c

                          ' Reset the sheet we are tracking
32                        strSheetPrefix = vbNullString
                          
33                    Else
                          ' This is a new range, look for : and another cell ref
34                        BuildingRef = tkn.Text
35                        CheckForColon = True
36                        LookingForEndOfRef = True
37                    End If
                    
38                Case TokenType.RangeOperator
                      ' Colon. The range we were building is about to be completed
                      ' and should come up after this
39                    BuildingRef = BuildingRef + ":"
40                    CheckForColon = False
                      
41                Case TokenType.ReferenceQualifier
42                    strSheetPrefix = "'" + tkn.Text + "'"
                      
43                Case TokenType.ExternalReferenceOperator
                      ' We'll insert the ! ourselves manually later

44            End Select
45        Next i

          ' Are we still looking for another reference?
46        If LookingForEndOfRef = True Then
              ' It never came - probably a solitary cell reference at the end
              ' of a formula. Check if we need a new node, add token to parsed string
47            If Len(strSheetPrefix) > 0 Then BuildingRef = strSheetPrefix + "!" + BuildingRef
48            Set rngBuildingRef = s.sheet.Range(BuildingRef)
              
              ' Cell reference found, check if constant
49            If rngBuildingRef.Count > 1 Then
50                For Each c In rngBuildingRef
51                    If Not Intersect(c.Precedents, s.AdjustableCells) Is Nothing Then GoTo NotConstant
52                Next c
53            Else
54                If Not Intersect(rngBuildingRef.Precedents, s.AdjustableCells) Is Nothing Then GoTo NotConstant
55            End If
              
56            strSheetPrefix = vbNullString
57        End If
58        StringFormulaIsConstant = True
          
ExitFunction:
59        If RaiseError Then RethrowError
60        Exit Function

ErrorHandler:
61        If Not ReportError("CModelParsed", "StringFormulaIsConstant") Then Resume
62        RaiseError = True
63        GoTo ExitFunction

NotConstant:
64        StringFormulaIsConstant = False
65        GoTo ExitFunction
End Function

Function ExpandFormula(ByVal strFormula As String, _
                       ByRef rngSourceCell As Range, Optional ConstraintNumber As Long) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          '==========================================================================
          ' STEP 0. Misc Setup
          ' The parsed string, this is what we store at the node
          Dim strParsed As String
          ' The current node we are considering
          Dim objCurNode As CFormula
          ' The sheet the next reference token is on
          Dim strSheetPrefix As String
          ' The sheet the adjustable cells are on
          Dim strAdjCellSheet As String
3         strAdjCellSheet = s.AdjustableCells.Parent.Name
          ' Avoid unnecessary evaluations of AddNodeIfNew by checking if whole ranges
          ' are decision variables
          Dim rngBuildingRef As Range
          
          ' The following are used for converting SumProduct to AMPL
          Dim FunctionName As String
          Dim FunctionCount As Long
          Dim start As Long
          Dim output As String, node As String
          Dim Count As Long
4         FunctionName = vbNullString
5         FunctionCount = 0
          
          '==========================================================================
          ' STEP 1. Setup root node for search
          Dim strRootAddress As String
6         If Not (rngSourceCell Is Nothing) Then
7             strRootAddress = ConvertCellToStandardName(rngSourceCell)
              ' If root node exists already, don't need to do this again
8             If TestKeyExists(Formulae, strRootAddress) Then
9                 ExpandFormula = strRootAddress
10                GoTo ExitFunction
11            End If
              ' Check that the source cell (e.g. objective, LHS, RHS) is not an
              ' adjustable cell itself
12            If rngSourceCell.Parent.Name = strAdjCellSheet Then
13                If Not (Intersect(rngSourceCell, s.AdjustableCells) Is Nothing) Then
                      ' Its an adjustable cell, just return standard name and stop
14                    ExpandFormula = strRootAddress
15                    GoTo ExitFunction
16                End If
17            End If
18        End If
          ' It doesn't exist, it isn't an adjustable cell
          ' Is it even a formula?
19        If Len(strFormula) = 0 Then
20            ExpandFormula = "0": GoTo ExitFunction
21        End If
22        If Left(strFormula, 1) <> "=" Then
23            If IsAmericanNumber(strFormula) Then
24                ExpandFormula = strFormula
25            Else
26                ExpandFormula = "'" + strFormula + "'"
27            End If
28            GoTo ExitFunction
29        End If
          ' Store the last index of the list-representation of the DAG
          Dim lngBaseIndex As Long
30        lngBaseIndex = Formulae.Count
          ' Create the root node and add it
31        Set objCurNode = New CFormula
32        If Not (rngSourceCell Is Nothing) Then
33            objCurNode.strAddress = strRootAddress
34            objCurNode.initialValue = CDbl(rngSourceCell.Value2)
35            Set objCurNode.rngAddress = rngSourceCell
36            ExpandFormula = strRootAddress
37        Else
              ' Its a constant or formula - return the parsed formula directly
38            objCurNode.strAddress = "_" + CStr(Formulae.Count)
39            Set objCurNode.rngAddress = Nothing
40        End If
41        objCurNode.strFormula = strFormula  ' Could also get from rngSourceCell
42        objCurNode.lngDepth = 0             ' Root node depth is 0
43        objCurNode.boolIsRoot = True
44        Formulae.Add objCurNode, objCurNode.strAddress ' Add it to the DAG
          ' Start at the root node
          Dim lngIndex As Long
45        lngIndex = lngBaseIndex + 1
          
          ' tracking variables
          Dim TotalCount As Long, curProgress As Long
          
          '==========================================================================
          ' STEP 2. Keep processing nodes until nothing interesting left
46        Do While lngIndex <= Formulae.Count
47            UpdateStatusBar "OpenSolver: Constraint progress: " & ConstraintNumber & "/" & s.NumConstraints & ". " & _
                              "Processing formula " & lngIndex & "/" & Formulae.Count & ". "
              
48            Set objCurNode = Formulae(lngIndex)
49            strParsed = vbNullString
50            strSheetPrefix = vbNullString
          
              '======================================================================
              ' STEP 2.C. Tokenise if we can't evaluate
51            If objCurNode.boolCanEval Then GoTo skipTokenising
              
              ' Tokenize the formula
              Dim tksFormula As Tokens
52            Set tksFormula = ParseFormula(objCurNode.strFormula)
          
              ' If we find a reference to a cell name, it could be the beginning
              ' of a multi-cell range. So always assume thats going to happen, and
              ' look for another reference (LookingForEndOfRef) after a colon :
              ' (CheckForColon). If a colon isn't there, stop looking.
              Dim LookingForEndOfRef As Boolean, CheckForColon As Boolean
53            LookingForEndOfRef = False
54            CheckForColon = False
              
              ' The candidate multicell range we are building
              Dim BuildingRef As String
55            BuildingRef = vbNullString
                      
              ' Take a walk through the tokens
              Dim i As Long, c As Range, tkn As Token
56            For i = 1 To tksFormula.Count
57                Set tkn = tksFormula.Item(i)
                  
                  ' CheckForColon means we hit a cell reference, and want to see if
                  ' its actually referring to multiple cells
58                If CheckForColon = True And tkn.Text <> ":" Then
                      ' We don't hit colon right after the reference, we abandon
                      ' search for another reference
59                    LookingForEndOfRef = False
60                    CheckForColon = False
61                    If Len(strSheetPrefix) > 0 Then BuildingRef = strSheetPrefix + "!" + BuildingRef
62                    On Error Resume Next
63                    Set rngBuildingRef = s.sheet.Range(BuildingRef)
64                    If Err.Number <> 0 Then
65                        On Error GoTo ErrorHandler
66                        RaiseUserError "The model contains references to cells on other sheets, " & _
                                         "which is not currently supported by the parsing solvers."
67                    End If
68                    On Error GoTo ErrorHandler
69                    If rngBuildingRef.Count > 1 Then
70                        strParsed = strParsed + "["
71                        For Each c In rngBuildingRef
72                            strParsed = strParsed + AddNodeIfNew(objCurNode, c) + ","
73                        Next c
74                        strParsed = Left(strParsed, Len(strParsed) - 1) + "]"
75                    Else
76                        strParsed = strParsed + AddNodeIfNew(objCurNode, rngBuildingRef)
77                    End If
                      ' Reset the sheet we are tracking
78                    strSheetPrefix = vbNullString
79                End If
          
                  ' Decide what to insert based on token type
80                Select Case tkn.TokenType
                      Case TokenType.Text
                          ' Output with quotes, tokenizer turns "=""test""" -> test
81                        strParsed = strParsed + "'" + tkn.Text + "'"
                              
82                    Case TokenType.Number
                          ' TODO: Scientific notation for small/large numbers?
83                        strParsed = strParsed + StrExNoPlus(Val(tkn.Text))
                          
84                    Case TokenType.Bool
85                        strParsed = strParsed + IIf(tkn.Text = "TRUE", "True", "False")
                          
86                    Case TokenType.ErrorText
                          ' TODO: Can't handle that, throw error
87                        RaiseGeneralError "An error token was found while parsing the formulae."
                          
88                    Case TokenType.Reference
                          ' Are we trying to complete a range?
89                        If LookingForEndOfRef Then
                              ' We were, so finish building it...
90                            BuildingRef = BuildingRef + tkn.Text
                              ' ... and stop trying to build it more
91                            LookingForEndOfRef = False
92                            CheckForColon = False
                              
                              ' Make sure this range is on the right sheet
93                            If Len(strSheetPrefix) > 0 Then BuildingRef = strSheetPrefix + "!" + BuildingRef
                              
94                            Set rngBuildingRef = s.sheet.Range(BuildingRef)
                              
                              'Replace preceding comma with + if summing a second argument
95                            If FunctionName = "SUM" And FunctionCount > 0 Then
96                                strParsed = Left(strParsed, Len(strParsed) - 1) + "+"
97                            End If
                              
                              'Replace preceding comma with + if summing a second argument
98                            If FunctionName = "PRODUCT" And FunctionCount > 0 Then
99                                strParsed = Left(strParsed, Len(strParsed) - 1) + "*"
100                           End If

                              'If we are starting a SUMIF, we need to get the next argument now
101                           If FunctionName = "SUMIF" And FunctionCount = 0 Then
                                  Dim NewToken As Token, IfOperator As String
102                               Set NewToken = tksFormula.Item(i + 1)
103                               If NewToken.Text <> "," Then
104                                   MsgBox "Misformed SUMIF"
105                               End If
                                  
106                               Set NewToken = tksFormula.Item(i + 2)
107                               IfOperator = "="
108                               If NewToken.TokenType = Text Then
109                                   Select Case Left(NewToken.Text, 2)
                                          Case "<=", ">=", "<>"
110                                           IfOperator = vbNullString
111                                   End Select
112                                   Select Case Left(NewToken.Text, 1)
                                          Case "<", ">", "="
113                                           IfOperator = vbNullString
114                                   End Select
115                               End If
116                               If NewToken.TokenType = TokenType.Reference Then
117                                   IfOperator = IfOperator + AddNodeIfNew(objCurNode, s.sheet.Range(NewToken.Text))
118                               Else
119                                   IfOperator = IfOperator + CStr(NewToken.Text)
120                               End If
121                               i = i + 2
122                           End If
                              
                              ' If we are building up a function, save the start of the function so we don't overwrite earlier work
123                           If FunctionCount = 0 Then
124                               start = Len(strParsed)
125                           End If
                              
126                           Count = 0
127                           TotalCount = rngBuildingRef.Count
128                           curProgress = 0
129                           For Each c In rngBuildingRef
                                  ' Update status
130                               curProgress = curProgress + 1
131                               UpdateStatusBar "OpenSolver: Constraint progress: " & ConstraintNumber & "/" & s.NumConstraints & ". " & _
                                                  "Processing formula " & lngIndex & "/" & Formulae.Count & ". " & _
                                                  "Current formula progress: " & curProgress & "/" & TotalCount
                                  
132                               node = AddNodeIfNew(objCurNode, c)
                                  
133                               If FunctionName = "SUMPRODUCT" And Not FunctionCount = 0 Then
134                                   Count = Count + 1
135                                   If FunctionCount > 1 And Count < rngBuildingRef.Count Then
                                          Dim StartIndex As Long
136                                       StartIndex = FindPosition(strParsed, "+", Count)
137                                       strParsed = Left(strParsed, StartIndex - 1) & Replace(strParsed, "+", "*" & node & "+", StartIndex, 1)
138                                   Else
139                                       strParsed = Replace(strParsed, ",", "*" & node & "+", 1, 1)
140                                   End If
141                               ElseIf FunctionName = "SUMIF" Then
142                                   If FunctionCount = 0 Then
143                                       strParsed = strParsed + "if(" + node + IfOperator + ","
144                                   Else
                                          Dim Position As Long
145                                       Count = Count + 1
146                                       Position = FindPosition(strParsed, ",", 2 * Count, start)
147                                       strParsed = Left(strParsed, Position) + node + ",0)+" + Mid(strParsed, Position + 1)
148                                   End If
149                               ElseIf FunctionName = "SUM" Then
150                                   strParsed = strParsed + node + "+"
151                               ElseIf FunctionName = "PRODUCT" Then
152                                   strParsed = strParsed + node + "*"
153                               Else
                                      ' Check if we need a new node, add token to parsed string
154                                   strParsed = strParsed + node + ","
155                               End If
156                           Next c
157                           FunctionCount = FunctionCount + 1
                              ' Reset the sheet we are tracking
158                           strSheetPrefix = vbNullString
                              ' Kill the extra char added by last cell in range
159                           strParsed = Left(strParsed, Len(strParsed) - 1)
                              
160                       Else
                              ' This is a new range, look for : and another cell ref
161                           BuildingRef = tkn.Text
162                           CheckForColon = True
163                           LookingForEndOfRef = True
164                       End If
                          
165                   Case TokenType.whitespace
                          ' Do nothing
                          
166                   Case TokenType.UnaryOperator
167                       If tkn.Text = "+" Then
168                           strParsed = strParsed + "+"
169                       ElseIf tkn.Text = "-" Then
170                           strParsed = strParsed + "-"
171                       Else
172                           MsgBox "Unary Operator that isn't + or -! Develop code to handle this: " + tkn.Text
173                       End If
                          
174                   Case TokenType.ArithmeticOperator
175                       strParsed = strParsed + tkn.Text
                          
176                   Case TokenType.ComparisonOperator
177                       If tkn.Text = "=" Then
178                           strParsed = strParsed + ConvertFormula(tkn.Text)
179                       Else 'TODO (maybe): <>?
180                           strParsed = strParsed + tkn.Text
181                       End If
                          
182                   Case TokenType.TextOperator
                          ' Text concatenation &
183                       strParsed = strParsed + " + "
                          
184                   Case TokenType.RangeOperator
                          ' Colon. The range we were building is about to be completed
                          ' and should come up after this
185                       BuildingRef = BuildingRef + ":"
186                       CheckForColon = False
                          
187                   Case TokenType.ReferenceQualifier
188                       strSheetPrefix = "'" + tkn.Text + "'"
                          
189                   Case TokenType.ExternalReferenceOperator
                          ' We'll insert the ! ourselves manually later
                          
190                   Case TokenType.PostfixOperator
                          ' Only percentage sign it seems
191                       strParsed = strParsed + "/100.0"
                          
192                   Case TokenType.FunctionOpen
193                       FunctionName = tkn.Text
194                       output = tkn.Text
195                       Select Case output
                          Case "SUMPRODUCT"
196                           output = "("
197                       Case "SUM", "PRODUCT"
198                           output = "("
199                       Case Else
                              ' Delegate to solver specific conversion code
200                           output = ConvertFormula(output)
201                       End Select
202                       strParsed = strParsed + output
                          
203                   Case TokenType.ParameterSeparator
204                       strParsed = strParsed + tkn.Text
                          
205                   Case TokenType.FunctionClose
206                       strParsed = strParsed + ")"
207                       FunctionCount = 0
                          
208                   Case TokenType.SubExpressionOpen
209                       strParsed = strParsed + tkn.Text
                      
210                   Case TokenType.SubExpressionClose
211                       strParsed = strParsed + tkn.Text
                          
212                   Case Else
                          ' TODO: table things
213               End Select
214           Next i

              ' Are we still looking for another reference?
215           If LookingForEndOfRef = True Then
                  ' It never came - probably a solitary cell reference at the end
                  ' of a formula. Check if we need a new node, add token to parsed string
216               If Len(strSheetPrefix) > 0 Then BuildingRef = strSheetPrefix + "!" + BuildingRef
217               Set rngBuildingRef = s.sheet.Range(BuildingRef)
218               If rngBuildingRef.Count > 1 Then
219                   strParsed = strParsed + "["
220                   For Each c In rngBuildingRef
221                       strParsed = strParsed + AddNodeIfNew(objCurNode, c) + ","
222                   Next c
223                   strParsed = Left(strParsed, Len(strParsed) - 1) + "]"
224               Else
225                   strParsed = strParsed + AddNodeIfNew(objCurNode, rngBuildingRef)
226               End If
227               strSheetPrefix = vbNullString
228           End If
              
              
              '======================================================================
              ' STEP 2.C. Store parsed formula in the DAG
skipTokenising:
229           If objCurNode.boolIsConstant = False Then objCurNode.strFormulaParsed = strParsed
230           If Left(objCurNode.strAddress, 1) = "_" Then
                  ' Root node is actually a formula or constant, so return the parsed formula
231               ExpandFormula = strParsed
232           End If
233           If objCurNode.boolIsRoot And objCurNode.boolIsConstant Then
                  ' Was just a simple constant, no need to create anything in Formulae
234               ExpandFormula = objCurNode.strFormulaParsed
235               Formulae.Remove lngIndex
236               GoTo ExitFunction
237           End If
              
              '======================================================================
              ' On to the next node
238           lngIndex = lngIndex + 1
239       Loop
          
ExitFunction:
240       Application.StatusBar = False
241       If RaiseError Then RethrowError
242       Exit Function

ErrorHandler:
243       If Not ReportError("CModelParsed", "ExpandFormula") Then Resume
244       RaiseError = True
245       GoTo ExitFunction
End Function

Function FindPosition(Word As String, FindString As String, Count As Long, Optional start As Long = 1) As Long
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim It As Long
3         It = 1
          
4         FindPosition = start
          
          ' Loop until count reached
5         While It < Count
6             FindPosition = InStr(FindPosition + 1, Word, FindString)
7             It = It + 1
8         Wend

ExitFunction:
9         If RaiseError Then RethrowError
10        Exit Function

ErrorHandler:
11        If Not ReportError("CModelParsed", "FindPosition") Then Resume
12        RaiseError = True
13        GoTo ExitFunction
End Function

Sub CleanFormulae(ByVal lngFStart As Long, ByVal lngFEnd As Long, Optional ConstraintNumber As Long = 0)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim objCurNode As CFormula, lngIndex As Long, varReturn As Variant

          ' Any AdjCellUnknown can now be decided one way or the other
          Dim lngCurDepth As Long, strChildAddr As Variant
3         For lngCurDepth = lngMaxDepth To 0 Step -1
4             For lngIndex = lngFStart To lngFEnd
5                 UpdateStatusBar "OpenSolver: Constraint progress: " & ConstraintNumber & "/" & s.NumConstraints & ". " & _
                                  "Cleaning decision variables in formula " & lngIndex & "/" & Formulae.Count & ". "
                  
6                 Set objCurNode = Formulae(lngIndex)
7                 If objCurNode.lngDepth = lngCurDepth Then
8                     If objCurNode.intRefsAdjCell = AdjCellUnknown And Not objCurNode.boolIsRoot Then
                          ' Look through children - if any are adj cell depedendant,
                          ' this is too. Otherwise, its not!
                          
9                         objCurNode.intRefsAdjCell = AdjCellIndependent
10                        For Each strChildAddr In objCurNode.Children
11                            If Formulae(strChildAddr).intRefsAdjCell = AdjCellDependent Then
12                                objCurNode.intRefsAdjCell = AdjCellDependent
13                            End If
14                        Next
                          ' If its AdjCellIndependent, eval it!
15                        If objCurNode.intRefsAdjCell = AdjCellIndependent Then
16                            objCurNode.boolCanEval = True
                              ' NOTE: Application.Evaluate has a 255 character limit
                              ' See also: http://dutchgemini.wordpress.com/2009/08/07/error-2015-using-application-evaluate-in-excel-vba/
17                            varReturn = s.sheet.Evaluate(objCurNode.strFormula)
18                            If (VBA.VarType(varReturn) = vbError) Then
                                  ' Fall back to calculating the cell
19                                objCurNode.rngAddress.Calculate
20                                objCurNode.strFormulaParsed = CStr(objCurNode.rngAddress.value)
21                                objCurNode.boolEvaledWithCalculate = True
22                            Else
23                                objCurNode.strFormulaParsed = varReturn
24                            End If
25                            objCurNode.boolIsConstant = True
26                        End If
27                    End If
28                End If
29            Next lngIndex
30        Next lngCurDepth
          
          ' Pull up constants where possible
          Dim strParentAdr As Variant
31        For lngCurDepth = lngMaxDepth To 0 Step -1
32            For lngIndex = lngFStart To lngFEnd
33                UpdateStatusBar "OpenSolver: Constraint progress: " & ConstraintNumber & "/" & s.NumConstraints & ". " & _
                                  "Cleaning constants in formula " & lngIndex & "/" & Formulae.Count & ". "
                  
34                If Formulae(lngIndex).lngDepth = lngCurDepth Then
35                    If Formulae(lngIndex).boolIsConstant Then
                          ' Constant - so can pull up into its parents
36                        Set objCurNode = Formulae(lngIndex)
                          
37                        For Each strParentAdr In objCurNode.Parents
38                            With Formulae(strParentAdr)
39                                .strFormulaParsed = Replace(.strFormulaParsed, objCurNode.strAddress, objCurNode.strFormulaParsed)
40                            End With
41                        Next
                          
42                        objCurNode.boolCanBeRemoved = Not objCurNode.boolIsRoot
43                    End If
44                End If
45            Next lngIndex
46        Next lngCurDepth
          
          ' Remove any unneeded nodes
47        For lngIndex = lngFEnd To lngFStart Step -1
48            UpdateStatusBar "OpenSolver: Constraint progress: " & ConstraintNumber & "/" & s.NumConstraints & ". " & _
                              "Removing unnecessary formulae: " & lngIndex & "/" & Formulae.Count & ". "
          
49            If Formulae(lngIndex).boolCanBeRemoved Then Formulae.Remove lngIndex
50        Next lngIndex

ExitSub:
51        Application.StatusBar = False
52        If RaiseError Then RethrowError
53        Exit Sub

ErrorHandler:
54        If Not ReportError("CModelParsed", "CleanFormulae") Then Resume
55        RaiseError = True
56        GoTo ExitSub
End Sub

Sub PUSHDOWN(ByRef objNode As CFormula, ByVal lngNewDepth As Long)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If lngNewDepth > lngMaxDepth Then lngMaxDepth = lngNewDepth
4         If lngNewDepth > objNode.lngDepth Then objNode.lngDepth = lngNewDepth
          'Dim child As Long
          'For child = 1 To objNode.Children.Count
          '    PUSHDOWN Formulae(objNode.Children(child)), objNode.lngDepth + 1, Formulae, lngMaxDepth
          'Next
          Dim child As Variant
5         For Each child In objNode.Children
6             If TestKeyExists(Formulae, CStr(child)) Then PUSHDOWN Formulae(child), objNode.lngDepth + 1
7         Next

ExitSub:
8         If RaiseError Then RethrowError
9         Exit Sub

ErrorHandler:
10        If Not ReportError("CModelParsed", "PUSHDOWN") Then Resume
11        RaiseError = True
12        GoTo ExitSub
End Sub

Function AddNodeIfNew(ByRef objCurNode As CFormula, ByRef c As Range) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim cell As Range
          
          ' 1.1 Get the standard name of this cell
          Dim strStdName As String
3         strStdName = ConvertCellToStandardName(c)
          
          ' Check if cell is an adjustable cell
4         If c.Parent.Name = s.sheet.Name Then
5             If Not Intersect(c, s.AdjustableCells) Is Nothing Then
6                 AddNodeIfNew = strStdName
7                 GoTo finishedNode
8             End If
9         End If
          '==========================================================================
          ' 0. Is Simple Cell?
          ' 0.1 Get the standard name of this cell
          Dim strCFormula As String
10        strCFormula = c.Formula
          ' 0.2 Is blank cell (TODO: breaks some IF statements)
11        If Len(strCFormula) = 0 Then
12            AddNodeIfNew = "0"
13            GoTo finishedNode
14        End If
          ' 0.3 Is constant?
15        If Left(strCFormula, 1) <> "=" Then
              ' Number?
16            If IsAmericanNumber(strCFormula) Then
17                AddNodeIfNew = strCFormula
18            Else
                  ' Not a number, assume text
19                AddNodeIfNew = """" + strCFormula + """"
20            End If
21            GoTo finishedNode
22        End If
                  
          '==========================================================================
          ' 1. Is Existing Node?
          ' 1.2 Does the node exist yet?
23        If TestKeyExists(Formulae, strStdName) Then
              ' Node already exists
              Dim objExistNode As CFormula
24            Set objExistNode = Formulae(strStdName)
              ' Has it been evaluated?
25            If objExistNode.boolIsConstant Then
                  ' Its a constant, so we can just fold it straight in
26                AddNodeIfNew = objExistNode.strFormulaParsed
27                GoTo finishedNode
28            End If
              ' Is it dependent on a decision variable?
29            If objExistNode.intRefsAdjCell = AdjCellDependent Then
                  ' That means this cell is also dependent on adj. cells, which
                  ' we may not have known
30                objCurNode.intRefsAdjCell = AdjCellDependent
31            End If
              ' Add this node to the current nodes Children list
32            If Not TestKeyExists(objCurNode.Children, strStdName) Then
33                objCurNode.Children.Add strStdName, strStdName
34            End If
              ' Tell it about its new parent
35            If Not TestKeyExists(objExistNode.Parents, objCurNode.strAddress) Then
36                objExistNode.Parents.Add objCurNode.strAddress, objCurNode.strAddress
37            End If
              ' Ensure it is at the correct depth
38            PUSHDOWN objExistNode, objCurNode.lngDepth + 1
              ' Return the standardised name of this existing node
39            AddNodeIfNew = strStdName
40            GoTo finishedNode
41        End If
          
          '==========================================================================
          ' 2. Is Adjustable Cell?
          Dim varReturn As Variant
          Dim objNewNode As New CFormula
42        Set objNewNode.rngAddress = c
43        objNewNode.strAddress = strStdName
          ' Is on same sheet?
44        If objNewNode.GetSheet = s.AdjustableCells.Parent.Name Then
              ' This formula is on the same sheet as the adjacent cells
              ' Thus we can safely check if this formula's cell is depedent on
              ' a decision variable
45            If objNewNode.IsDependentOn(AdjDependents) Then
                  ' This formula depends on the value of the adjustable cells
                  ' This means we can NOT evaluate it
46                objNewNode.boolCanEval = False
47                objNewNode.boolIsConstant = False
48                objNewNode.intRefsAdjCell = AdjCellDependent
49            Else
                  ' This formula is NOT dependent on an adjustable cell, so we
                  ' can just evaluate its formula (TODO: or take its value?)
                  ' We know its not a simple constant, or we would of got
                  ' it earlier
                  ' NOTE: Application.Evaluate has a 255 character limit
                  
50                varReturn = s.sheet.Evaluate(strCFormula)
51                If (VBA.VarType(varReturn) = vbError) Then
                      ' Fall back to calculating the cell
52                    c.Calculate
53                    AddNodeIfNew = StrExNoPlus(c.value)
54                    GoTo finishedNode
55                ElseIf IsArray(varReturn) Then
56                    AddNodeIfNew = StrExNoPlus(CDbl(varReturn(1)))
57                    GoTo finishedNode
58                Else
59                    On Error Resume Next
60                    AddNodeIfNew = StrExNoPlus(varReturn)
61                    If Err.Number <> 0 Then
62                        On Error GoTo ErrorHandler
63                        RaiseUserError "The constant in cell " & c.Address & " has value `" & varReturn & _
                                         "` which is not numeric. " & _
                                         "Make sure this constant isn't being included in a formula unnecessarily."
64                    End If
65                    On Error GoTo ErrorHandler
66                    GoTo finishedNode
67                End If
68            End If
69        Else
              ' This formula is on a different sheet to the adjacent cells.
              ' This means we can't check if its depedent on an adjacent cell,
              ' because Range. Dependents only returns the cells for the same
              ' sheet as Range.Parent
              'objNewNode.boolCanEval = False
              'objNewNode.boolIsConstant = False
              'objNewNode.intRefsAdjCell = AdjCellUnknown
70            varReturn = s.sheet.Evaluate(strCFormula)
71                If (VBA.VarType(varReturn) = vbError) Then
                      ' Fall back to calculating the cell
72                    c.Calculate
73                    AddNodeIfNew = StrExNoPlus(c.value)
74                    GoTo finishedNode
75                Else
76                    AddNodeIfNew = StrExNoPlus(CDbl(varReturn))
77                    GoTo finishedNode
78                End If
79        End If
          
          '==========================================================================
          ' 3. Create new node
80        With objNewNode
81            .strFormula = strCFormula
82            If c.Value2 = vbNullString Then
83                .initialValue = 0
84            Else
85                .initialValue = CDbl(c.Value2)
86            End If
87            .lngDepth = objCurNode.lngDepth + 1
88            If .lngDepth > lngMaxDepth Then lngMaxDepth = .lngDepth
89            .Parents.Add objCurNode.strAddress, objCurNode.strAddress
90        End With
          ' Add this new node to the child list of current node
91        objCurNode.Children.Add strStdName, strStdName
          ' Add new node to the DAG
92        Formulae.Add objNewNode, strStdName
          ' Return the standardised name of this new node
93        AddNodeIfNew = strStdName

finishedNode:
          ' Done!
ExitFunction:
94        Application.StatusBar = False
95        If RaiseError Then RethrowError
96        Exit Function

ErrorHandler:
97        If Not ReportError("CModelParsed", "AddNodeIfNew") Then Resume
98        RaiseError = True
99        GoTo ExitFunction
End Function

Function ConvertFormula(tokenText As String) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim FileSolver As ISolverFile
3         Set FileSolver = s.Solver

4         Select Case FileSolver.FileType
          Case AMPL
5             ConvertFormula = ConvertFormula_AMPL(tokenText)
6         Case NL
7             ConvertFormula = ConvertFormula_NL(tokenText)
8         Case Else
9             ConvertFormula = tokenText + "("
10        End Select

ExitFunction:
11        If RaiseError Then RethrowError
12        Exit Function

ErrorHandler:
13        If Not ReportError("CModelParsed", "ConvertFormula") Then Resume
14        RaiseError = True
15        GoTo ExitFunction
End Function

Function ConvertFormula_AMPL(tokenText As String) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         tokenText = LCase(tokenText)
4         Select Case tokenText
          Case "="
5             ConvertFormula_AMPL = "=="
6         Case "radians"
7             ConvertFormula_AMPL = "pi/180*("
8         Case "pi"
9             ConvertFormula_AMPL = "(pi"
10        Case "sumif"
11            RaiseUserError "The operator " & UCase(tokenText) & " is not yet supported by OpenSolver when using the NEOS solvers."
12            ConvertFormula_AMPL = tokenText + "("
13        Case Else
14            ConvertFormula_AMPL = tokenText + "("
15        End Select

ExitFunction:
16        If RaiseError Then RethrowError
17        Exit Function

ErrorHandler:
18        If Not ReportError("CModelParsed", "ConvertFormula_AMPL") Then Resume
19        RaiseError = True
20        GoTo ExitFunction
End Function

Function ConvertFormula_NL(tokenText As String) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         Select Case tokenText
          ' RADIANS
          Case "RADIANS"
4             ConvertFormula_NL = "4*ATAN(1)/180*("
5         Case "SUMIF"
6             ConvertFormula_NL = "("
7         Case "PI"
8             ConvertFormula_NL = "(4*ATAN(1)"
9         Case Else
10            ConvertFormula_NL = tokenText + "("
11        End Select

ExitFunction:
12        If RaiseError Then RethrowError
13        Exit Function

ErrorHandler:
14        If Not ReportError("CModelParsed", "ConvertFormula_NL") Then Resume
15        RaiseError = True
16        GoTo ExitFunction
End Function


