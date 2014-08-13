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

'------------------------------------------------------------------------------
' MODEL DEFINITION
Public SolverModelSheet As Worksheet

Public AdjustableCells As Range
Public AdjustableCellCount As Long
Private AdjCellNames() As String
Public VarLowerBounds As Collection
    
Public ObjectiveSense As ObjectiveSenseType
Public ObjectiveTargetValue As Double
Public ObjectiveCell As Range
Public ObjectiveCellConstant As Double

Public ConstraintCount As Long
Public IntegerCells As Range
Public BinaryCells As Range
Private LHSType() As SolverInputType
Private RHSType() As SolverInputType

Public AssumeNonNegative As Boolean

Public Solver As String

Public Formulae As Collection
Public VarTypeMap As Collection
Public AdjCellNameMap As Collection
Public LHSKeys As Collection
Public RHSKeys As Collection
Public Rels As Collection

Public lngMaxDepth As Long

'------------------------------------------------------------------------------
' USEFUL GLOBALS
Public book As Workbook
Public strNameRoot As String

'==============================================================================
Public Sub Setup(ByRef BookToUse As Workbook, ByRef SheetToUse As Worksheet)

    Set book = BookToUse
    Set SolverModelSheet = SheetToUse
    strNameRoot = "'" & SolverModelSheet.Name & "'" & "!"
    
End Sub

'==============================================================================
Public Function ProcessSolverModel(s As COpenSolver) As Boolean
    Solver = s.Solver
    
    ' =========================================
    ' Process adjustable cells from COpenSolver
    Set AdjustableCells = s.AdjustableCells
    AdjustableCellCount = AdjustableCells.Count
    
    ' Get variable names
    ReDim AdjCellNames(1 To AdjustableCellCount) As String
    Dim i As Long
    For i = 1 To AdjustableCellCount
        AdjCellNames(i) = s.VarNamesP(i)
    Next i
    
    ' Get variable lower bounds
    Set VarLowerBounds = New Collection
    Dim c As Range
    For Each c In AdjustableCells
        If TestKeyExists(s.VarLowerBounds, c.Address) Then
            VarLowerBounds.Add s.VarLowerBounds(c.Address), c.Address(RowAbsolute:=False, ColumnAbsolute:=False)
        End If
    Next c
    
    ' =========================================
    ' Process objective from COpenSolver
    ObjectiveSense = s.ObjectiveSense
    ObjectiveTargetValue = s.ObjectiveTargetValue
    Set ObjectiveCell = s.ObjRange
    If Not ObjectiveCell Is Nothing Then
        ObjectiveCellConstant = ObjectiveCell.Value2
    End If
    
    ' =========================================
    ' Process constraints from COpenSolver
    ConstraintCount = s.NumConstraints
    Set IntegerCells = s.IntegerCellsRange
    Set BinaryCells = s.BinaryCellsRange
    
    If ConstraintCount > 0 Then
        ReDim LHSType(1 To ConstraintCount) As SolverInputType
        ReDim RHSType(1 To ConstraintCount) As SolverInputType
        For i = 1 To ConstraintCount
            LHSType(i) = s.LHSTypeP(i)
            RHSType(i) = s.RHSTypeP(i)
        Next i
    End If
    
    ' =========================================
    ' Process model options from COpenSolver
    AssumeNonNegative = s.AssumeNonNegativeVars
    
    ' =========================================
    ' Conduct parsed-specific processing
    ProcessVarTypes
    ProcessSolverModel = ParseFormulae(s)
End Function

Private Sub ProcessVarTypes()
' Get var types of constraints
    
    Set VarTypeMap = New Collection
    Set AdjCellNameMap = New Collection
    Dim curVarType As Long, lngArea As Long, cellName As String, c As Range
    For lngArea = 1 To AdjustableCells.Areas.Count
        For Each c In AdjustableCells.Areas(lngArea)
            cellName = ConvertCellToStandardName(c)
            If Not TestKeyExists(AdjCellNameMap, cellName) Then
                AdjCellNameMap.Add c, cellName
                curVarType = VarContinuous
                If Not (IntegerCells Is Nothing) Then
                    If Not (Intersect(c, IntegerCells) Is Nothing) Then curVarType = VarInteger
                End If
                If Not (BinaryCells Is Nothing) Then
                    If Not (Intersect(c, BinaryCells) Is Nothing) Then curVarType = VarBinary
                End If
                VarTypeMap.Add curVarType, cellName
            End If
        Next
    Next
End Sub

Private Function ParseFormulae(s As COpenSolver) As Boolean
' Parses all constraints and objective formulae to create formulae that depend only on variable cells and constants

    ' Create a collection for the spreadsheet's formulae (DAG)
    Set Formulae = New Collection
    ' Remember the greatest depth of the DAG
    lngMaxDepth = 0

    '==========================================================================
    ' STEP 1. Parse objective cell
    Dim objresult As String
    
    If Not ObjectiveCell Is Nothing Then
        objresult = ExpandFormula(ObjectiveCell.Formula, ObjectiveCell, 0)
        
        If objresult Like "*Error*" Then
            Err.Raise OpenSolver_BuildError, Description:="The objective contains an error"
        End If
        CleanFormulae 1, Formulae.Count
    End If
    
    '==========================================================================
    ' STEP 2. Parse constraints
    Set LHSKeys = New Collection
    Set RHSKeys = New Collection
    Set Rels = New Collection

    Dim i As Long
    Dim Count As Long
    Count = 1
    For i = 1 To ConstraintCount
        Dim nameLHSi As Name, nameRELi As Name, nameRHSi As Name
        Set nameLHSi = Names(strNameRoot + "solver_lhs" + CStr(i))
        Set nameRELi = Names(strNameRoot + "solver_rel" + CStr(i))
        Set nameRHSi = Names(strNameRoot + "solver_rhs" + CStr(i))
        
        Dim lngFormulaeCountBefore As Long
        Dim cRow As Long, cCol As Long, cRowCount As Long, cColCount As Long
        Dim varLHSFormulae As Variant, rngLHS As Range, curLHSFormulae As String
        
        If nameRELi.value = "=5" Or nameRELi.value = "=4" Then
            GoTo NextCons
        End If
        
        lngFormulaeCountBefore = Formulae.Count
        
        ' Possibilities
        ' LHS is one cell, RHS is one cell/formula/value
        ' LHS is a range (any shape), RHS is one cell/formula/value
        ' LHS is a range (any shape), RHS is a range of same size, not necessarily same shape
        ' Because mismatched shapes are cruel, we will just throw an error if we find them
        ' OpenSolver1 just threw a general error.
        
        cRowCount = nameLHSi.RefersToRange.Rows.Count
        cColCount = nameLHSi.RefersToRange.Columns.Count
        Set rngLHS = nameLHSi.RefersToRange
        varLHSFormulae = rngLHS.Formula
                
        Dim singleRngLHS As Range
        For cRow = 1 To cRowCount
            For cCol = 1 To cColCount
                Count = Count + 1
                
                ' Get LHS
                If cRowCount <> 1 Or cColCount <> 1 Then
                    curLHSFormulae = varLHSFormulae(cRow, cCol)
                    Set singleRngLHS = rngLHS(cRow, cCol)
                Else
                    curLHSFormulae = varLHSFormulae
                    Set singleRngLHS = rngLHS
                End If
                
                ' Get RHS
                Dim strRHSFormula As String, rngRHSCell As Range
                Select Case GetRHSType(i)
                ' 1.1 RHS is a single cell
                Case SingleCellRange
                    strRHSFormula = nameRHSi.RefersToRange.Formula
                    Set rngRHSCell = nameRHSi.RefersToRange
                ' 1.2 RHS is multiple cells
                Case MultiCellRange
                    strRHSFormula = nameRHSi.RefersToRange.Cells(cRow, cCol).Formula
                    Set rngRHSCell = nameRHSi.RefersToRange.Cells(cRow, cCol)
                ' 1.3 RHS is value or formula
                Case constant, Formula
                    strRHSFormula = nameRHSi.value
                    Set rngRHSCell = Nothing
                End Select

                ' If LHS is constant, then check if RHS is constant
                If CellIsConstant(singleRngLHS) Then
                    Dim isConst As Boolean
                    isConst = False
                     
                    Select Case GetRHSType(i)
                    Case constant
                        isConst = True
                    Case Formula
                        isConst = StringFormulaIsConstant(strRHSFormula)
                    Case SingleCellRange, MultiCellRange
                        isConst = CellIsConstant(rngRHSCell)
                    End Select
                    
                    ' If both sides are constant we need to see if the constraint is valid
                    If isConst Then
                        Dim isValid As Boolean, LHSValue As Double, RHSValue As Double
                        isValid = True
                        
                        LHSValue = singleRngLHS.Value2
                        RHSValue = Application.Evaluate(strRHSFormula)
                        
                        ' Test LHS and RHS in the constraint
                        Select Case ConvertRelationToEnum(nameRELi.value)
                        Case RelationGE
                            If LHSValue < RHSValue Then
                                isValid = False
                            End If
                        Case RelationLE
                            If LHSValue > RHSValue Then
                                isValid = False
                            End If
                        Case RelationEQ
                            If LHSValue <> RHSValue Then
                                isValid = False
                            End If
                        End Select
                        
                        ' If constraint is not satisfied, the model is infeasible
                        If Not isValid Then
                            s.SolveStatus = OpenSolverResult.Infeasible
                            s.SolveStatusString = "Infeasible"
                            s.SolveStatusComment = _
                                "The model contains an Excel constraint which does not depend on the decision variables and is not satisfied." & vbCrLf & _
                                "Constraint specifies: " & _
                                "LHS=" & singleRngLHS.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "=" & LHSValue & _
                                "   " & SolverRelationAsString(ConvertRelationToEnum(nameRELi.value)) & "   " & _
                                "RHS=" & "'" & strRHSFormula & "'" & "=" & RHSValue
                            ParseFormulae = False
                            Exit Function
                        Else
                            ' Otherwise the constraint is trivially satisfied, we can ignore it
                            GoTo NextCol
                        End If
                    End If
                End If
                
                ' Parse LHS
                LHSKeys.Add ExpandFormula(curLHSFormulae, rngLHS(cRow, cCol), i)
                
                ' Parse RHS
                RHSKeys.Add ExpandFormula(strRHSFormula, rngRHSCell, i)
                
                ' Determine appropriate relation
                Rels.Add ConvertRelationToEnum(nameRELi.value)
NextCol:
            Next cCol
        Next cRow
        
        ' Clean Formulae
        CleanFormulae lngFormulaeCountBefore + 1, Formulae.Count
        
NextCons:
    Next i
    Application.StatusBar = False
    
    '==========================================================================
    ' STEP 3. Make all possible substitutions in formula tree if required by solver
    If DoBackSubstitution(Solver) Then
    
        Dim lngCurDepth As Long
        Dim lngIndex As Long, objCurNode As CFormula, strParentAdr As Variant
        
        For lngCurDepth = lngMaxDepth To 0 Step -1
            For lngIndex = 1 To Formulae.Count
                If Formulae(lngIndex).lngDepth = lngCurDepth Then
                    Set objCurNode = Formulae(lngIndex)
                    
                    For Each strParentAdr In objCurNode.Parents
                        With Formulae(strParentAdr)
                            .strFormulaParsed = Replace(.strFormulaParsed, objCurNode.strAddress, "(" & objCurNode.strFormulaParsed & ")")
                        End With
                    Next
                End If
            Next lngIndex
        Next lngCurDepth
    End If
    ParseFormulae = True
End Function
'==============================================================================
Public Function GetLHSType(lngConNum As Long) As SolverInputType
    GetLHSType = LHSType(lngConNum)
End Function

Public Function GetRHSType(lngConNum As Long) As SolverInputType
    GetRHSType = RHSType(lngConNum)
End Function

Public Function GetAdjCellName(lngVar As Long) As String
    GetAdjCellName = AdjCellNames(lngVar)
End Function

' Checks whether a cell is constant i.e. not a decision variable and does not depend on a decision variable
Function CellIsConstant(cell As Range) As Boolean
    ' Check if cell is a variable cell
    If Not Intersect(cell, AdjustableCells) Is Nothing Then
        CellIsConstant = False
        Exit Function
    End If
    
    ' Check if dependent on decision variable
    If cell.HasFormula Then
        If Not Intersect(cell.Precedents, AdjustableCells) Is Nothing Then
            CellIsConstant = False
            Exit Function
        End If
    End If
    CellIsConstant = True
End Function

' Checks whether a string formula depends on variable cells or is simply constant.
' If constant, we can evaluate it directly
Function StringFormulaIsConstant(strFormula As String) As Boolean
    Dim tksFormula As Tokens
    Set tksFormula = ParseFormula(strFormula)
    
    Dim i As Long, c As Range, tkn As Token, CheckForColon As Boolean
    Dim LookingForEndOfRef As Boolean, strSheetPrefix As String, BuildingRef As String, rngBuildingRef As Range
    
    ' We loop through tokens looking for cell references.
    ' If we find any, we test them to see if they depend on the variable cells
    For i = 1 To tksFormula.Count
        Set tkn = tksFormula.Item(i)
        
        If CheckForColon = True And tkn.Text <> ":" Then
            ' We don't hit colon right after the reference, we abandon
            ' search for another reference
            LookingForEndOfRef = False
            CheckForColon = False
            If strSheetPrefix <> "" Then BuildingRef = strSheetPrefix + "!" + BuildingRef
            Set rngBuildingRef = Range(BuildingRef)
            
            ' Cell reference found, check if constant
            If rngBuildingRef.Count > 1 Then
                For Each c In rngBuildingRef
                    If Not CellIsConstant(c) Then GoTo NotConstant
                Next c
            Else
                If Not CellIsConstant(rngBuildingRef) Then GoTo NotConstant
            End If
            
            ' Reset the sheet we are tracking
            strSheetPrefix = ""
        End If
    
        ' Decide what to insert based on token type
        Select Case tkn.TokenType
            Case TokenType.ErrorText
                ' TODO: Can't handle that, throw error
                Err.Raise OpenSolver_BuildError, "Parsing model formulae", "An error token was found while parsing the formulae."
                
            Case TokenType.Reference
                ' Are we trying to complete a range?
                If LookingForEndOfRef Then
                    ' We were, so finish building it...
                    BuildingRef = BuildingRef + tkn.Text
                    ' ... and stop trying to build it more
                    LookingForEndOfRef = False
                    CheckForColon = False
                    
                    ' Make sure this range is on the right sheet
                    If strSheetPrefix <> "" Then BuildingRef = strSheetPrefix + "!" + BuildingRef
                    
                    ' Cell reference found, check if constant
                    Set rngBuildingRef = Range(BuildingRef)
                    For Each c In rngBuildingRef
                        If Not Intersect(c.Precedents, AdjustableCells) Is Nothing Then GoTo NotConstant
                    Next c

                    ' Reset the sheet we are tracking
                    strSheetPrefix = ""
                    
                Else
                    ' This is a new range, look for : and another cell ref
                    BuildingRef = tkn.Text
                    CheckForColon = True
                    LookingForEndOfRef = True
                End If
              
            Case TokenType.RangeOperator
                ' Colon. The range we were building is about to be completed
                ' and should come up after this
                BuildingRef = BuildingRef + ":"
                CheckForColon = False
                
            Case TokenType.ReferenceQualifier
                strSheetPrefix = "'" + tkn.Text + "'"
                
            Case TokenType.ExternalReferenceOperator
                ' We'll insert the ! ourselves manually later

        End Select
    Next i

    ' Are we still looking for another reference?
    If LookingForEndOfRef = True Then
        ' It never came - probably a solitary cell reference at the end
        ' of a formula. Check if we need a new node, add token to parsed string
        If strSheetPrefix <> "" Then BuildingRef = strSheetPrefix + "!" + BuildingRef
        Set rngBuildingRef = Range(BuildingRef)
        
        ' Cell reference found, check if constant
        If rngBuildingRef.Count > 1 Then
            For Each c In rngBuildingRef
                If Not Intersect(c.Precedents, AdjustableCells) Is Nothing Then GoTo NotConstant
            Next c
        Else
            If Not Intersect(rngBuildingRef.Precedents, AdjustableCells) Is Nothing Then GoTo NotConstant
        End If
        
        strSheetPrefix = ""
    End If
    StringFormulaIsConstant = True
    Exit Function
    
NotConstant:
    StringFormulaIsConstant = False
End Function

Function ExpandFormula(ByVal strFormula As String, _
                       ByRef rngSourceCell As Range, Optional ConstraintNumber As Long) As String
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
    strAdjCellSheet = AdjustableCells.Parent.Name
    ' Avoid unnecessary evaluations of AddNodeIfNew by checking if whole ranges
    ' are decision variables
    Dim rngBuildingRef As Range
    
    ' The following are used for converting SumProduct to AMPL
    Dim FunctionName As String
    Dim FunctionCount As Long
    Dim start As Long
    Dim output As String, node As String
    Dim Count As Long
    FunctionName = ""
    FunctionCount = 0
    
    '==========================================================================
    ' STEP 1. Setup root node for search
    Dim strRootAddress As String
    If Not (rngSourceCell Is Nothing) Then
        strRootAddress = ConvertCellToStandardName(rngSourceCell)
        ' If root node exists already, don't need to do this again
        If TestKeyExists(Formulae, strRootAddress) Then
            ExpandFormula = strRootAddress
            Exit Function
        End If
        ' Check that the source cell (e.g. objective, LHS, RHS) is not an
        ' adjustable cell itself
        If rngSourceCell.Parent.Name = strAdjCellSheet Then
            If Not (Intersect(rngSourceCell, AdjustableCells) Is Nothing) Then
                ' Its an adjustable cell, just return standard name and stop
                ExpandFormula = strRootAddress
                Exit Function
            End If
        End If
    End If
    ' It doesn't exist, it isn't an adjustable cell
    ' Is it even a formula?
    If Len(strFormula) = 0 Then
        ExpandFormula = "0": Exit Function
    End If
    If left(strFormula, 1) <> "=" Then
        If IsAmericanNumber(strFormula) Then
            ExpandFormula = strFormula
        Else
            ExpandFormula = "'" + strFormula + "'"
        End If
        Exit Function
    End If
    ' Store the last index of the list-representation of the DAG
    Dim lngBaseIndex As Long
    lngBaseIndex = Formulae.Count
    ' Create the root node and add it
    Set objCurNode = New CFormula
    If Not (rngSourceCell Is Nothing) Then
        objCurNode.strAddress = strRootAddress
        objCurNode.initialValue = CDbl(rngSourceCell.Value2)
        Set objCurNode.rngAddress = rngSourceCell
        ExpandFormula = strRootAddress
    Else
        ' Its a constant or formula - return the parsed formula directly
        objCurNode.strAddress = "_" + CStr(Formulae.Count)
        Set objCurNode.rngAddress = Nothing
    End If
    objCurNode.strFormula = strFormula  ' Could also get from rngSourceCell
    objCurNode.lngDepth = 0             ' Root node depth is 0
    objCurNode.boolIsRoot = True
    Formulae.Add objCurNode, objCurNode.strAddress ' Add it to the DAG
    ' Start at the root node
    Dim lngIndex As Long
    lngIndex = lngBaseIndex + 1
    
    ' tracking variables
    Dim TotalCount As Long, curProgress As Long
    
    '==========================================================================
    ' STEP 2. Keep processing nodes until nothing interesting left
    Do While lngIndex <= Formulae.Count
        'DoEvents
        'Debug.Print lngIndex, Formulae.Count
        'If lngIndex = 96 Then
        '    'Debug.Print 96
        'End If
        Set objCurNode = Formulae(lngIndex)
        strParsed = ""
        strSheetPrefix = ""
    
        '======================================================================
        ' STEP 2.C. Tokenise if we can't evaluate
        If objCurNode.boolCanEval Then GoTo skipTokenising
        
        ' Tokenize the formula
        Dim tksFormula As Tokens
        Set tksFormula = ParseFormula(objCurNode.strFormula)
    
        ' If we find a reference to a cell name, it could be the beginning
        ' of a multi-cell range. So always assume thats going to happen, and
        ' look for another reference (LookingForEndOfRef) after a colon :
        ' (CheckForColon). If a colon isn't there, stop looking.
        Dim LookingForEndOfRef As Boolean, CheckForColon As Boolean
        LookingForEndOfRef = False
        CheckForColon = False
        
        ' The candidate multicell range we are building
        Dim BuildingRef As String
        BuildingRef = ""
                
        ' Take a walk through the tokens
        Dim i As Long, c As Range, tkn As Token
        For i = 1 To tksFormula.Count
            Set tkn = tksFormula.Item(i)
            
            ' CheckForColon means we hit a cell reference, and want to see if
            ' its actually referring to multiple cells
            If CheckForColon = True And tkn.Text <> ":" Then
                ' We don't hit colon right after the reference, we abandon
                ' search for another reference
                LookingForEndOfRef = False
                CheckForColon = False
                If strSheetPrefix <> "" Then BuildingRef = strSheetPrefix + "!" + BuildingRef
                Set rngBuildingRef = Range(BuildingRef)
                If rngBuildingRef.Count > 1 Then
                    strParsed = strParsed + "["
                    For Each c In rngBuildingRef
                        strParsed = strParsed + AddNodeIfNew(objCurNode, c) + ","
                    Next c
                    strParsed = left(strParsed, Len(strParsed) - 1) + "]"
                Else
                    strParsed = strParsed + AddNodeIfNew(objCurNode, rngBuildingRef)
                End If
                ' Reset the sheet we are tracking
                strSheetPrefix = ""
            End If
    
            ' Decide what to insert based on token type
            Select Case tkn.TokenType
                Case TokenType.Text
                    ' Output with quotes, tokenizer turns "=""test""" -> test
                    strParsed = strParsed + "'" + tkn.Text + "'"
                        
                Case TokenType.Number
                    ' TODO: Scientific notation for small/large numbers?
                    strParsed = strParsed + tkn.Text
                    
                Case TokenType.Bool
                    strParsed = strParsed + IIf(tkn.Text = "TRUE", "True", "False")
                    
                Case TokenType.ErrorText
                    ' TODO: Can't handle that, throw error
                    Err.Raise OpenSolver_BuildError, "Parsing model formulae", "An error token was found while parsing the formulae."
                    
                Case TokenType.Reference
                    ' Are we trying to complete a range?
                    If LookingForEndOfRef Then
                        ' We were, so finish building it...
                        BuildingRef = BuildingRef + tkn.Text
                        ' ... and stop trying to build it more
                        LookingForEndOfRef = False
                        CheckForColon = False
                        
                        ' Make sure this range is on the right sheet
                        If strSheetPrefix <> "" Then BuildingRef = strSheetPrefix + "!" + BuildingRef
                        
                        Set rngBuildingRef = Range(BuildingRef)
                        
                        'Replace preceding comma with + if summing a second argument
                        If FunctionName = "SUM" And FunctionCount > 0 Then
                            strParsed = left(strParsed, Len(strParsed) - 1) + "+"
                        End If
                        
                        ' If we are building up a function, save the start of the function so we don't overwrite earlier work
                        If FunctionCount = 0 Then
                            start = Len(strParsed)
                        End If
                        
                        Count = 0
                        TotalCount = rngBuildingRef.Count
                        curProgress = 0
                        For Each c In rngBuildingRef
                            curProgress = curProgress + 1
                            If curProgress Mod 10 = 0 Then _
                                Application.StatusBar = "OpenSolver: Constraint progress: " & ConstraintNumber & "/" & ConstraintCount & ". " & _
                                                        "Processing formula " & lngIndex & "/" & Formulae.Count & ". " & _
                                                        "Current formula progress: " & curProgress & "/" & TotalCount
                            node = AddNodeIfNew(objCurNode, c)
                            
                            If FunctionName = "SUMPRODUCT" And Not FunctionCount = 0 Then
                                Count = Count + 1
                                If FunctionCount > 1 And Count < rngBuildingRef.Count Then
                                    strParsed = Replace(strParsed, "+", "*" & node & "+", FindPosition(strParsed, "+", Count), 1)
                                Else
                                    strParsed = Replace(strParsed, ",", "*" & node & "+", 1, 1)
                                End If
                            ElseIf FunctionName = "SUM" Then
                                strParsed = strParsed + node + "+"
                            Else
                                ' Check if we need a new node, add token to parsed string
                                strParsed = strParsed + node + ","
                            End If
                        Next c
                        FunctionCount = FunctionCount + 1
                        ' Reset the sheet we are tracking
                        strSheetPrefix = ""
                        ' Kill the extra char added by last cell in range
                        strParsed = left(strParsed, Len(strParsed) - 1)
                        
                    Else
                        ' This is a new range, look for : and another cell ref
                        BuildingRef = tkn.Text
                        CheckForColon = True
                        LookingForEndOfRef = True
                    End If
                    
                Case TokenType.whitespace
                    ' Do nothing
                    
                Case TokenType.UnaryOperator
                    If tkn.Text = "+" Then
                        strParsed = strParsed + "+"
                    ElseIf tkn.Text = "-" Then
                        strParsed = strParsed + "-"
                    Else
                        MsgBox "Unary Operator that isn't + or -! Develop code to handle this: " + tkn.Text
                    End If
                    
                Case TokenType.ArithmeticOperator
                    strParsed = strParsed + tkn.Text
                    
                Case TokenType.ComparisonOperator
                    If tkn.Text = "=" Then
                        strParsed = strParsed + "=="
                    Else 'TODO (maybe): <>?
                        strParsed = strParsed + tkn.Text
                    End If
                    
                Case TokenType.TextOperator
                    ' Text concatenation &
                    strParsed = strParsed + " + "
                    
                Case TokenType.RangeOperator
                    ' Colon. The range we were building is about to be completed
                    ' and should come up after this
                    BuildingRef = BuildingRef + ":"
                    CheckForColon = False
                    
                Case TokenType.ReferenceQualifier
                    strSheetPrefix = "'" + tkn.Text + "'"
                    
                Case TokenType.ExternalReferenceOperator
                    ' We'll insert the ! ourselves manually later
                    
                Case TokenType.PostfixOperator
                    ' Only percentage sign it seems
                    strParsed = strParsed + "/100.0"
                    
                Case TokenType.FunctionOpen
                    FunctionName = tkn.Text
                    output = tkn.Text
                    Select Case output
                    Case "SUMPRODUCT"
                        output = "("
                    Case "SUM"
                        output = "("
                    Case Else
                        ' Delegate to solver specific conversion code
                        output = ConvertFormula(Solver, output)
                    End Select
                    strParsed = strParsed + output
                    
                Case TokenType.ParameterSeparator
                    strParsed = strParsed + tkn.Text
                    
                Case TokenType.FunctionClose
                    strParsed = strParsed + ")"
                    FunctionCount = 0
                    
                Case TokenType.SubExpressionOpen
                    strParsed = strParsed + tkn.Text
                
                Case TokenType.SubExpressionClose
                    strParsed = strParsed + tkn.Text
                    
                Case Else
                    ' TODO: table things
            End Select
        Next i

        ' Are we still looking for another reference?
        If LookingForEndOfRef = True Then
            ' It never came - probably a solitary cell reference at the end
            ' of a formula. Check if we need a new node, add token to parsed string
            If strSheetPrefix <> "" Then BuildingRef = strSheetPrefix + "!" + BuildingRef
            Set rngBuildingRef = Range(BuildingRef)
            If rngBuildingRef.Count > 1 Then
                strParsed = strParsed + "["
                For Each c In rngBuildingRef
                    strParsed = strParsed + AddNodeIfNew(objCurNode, c) + ","
                Next c
                strParsed = left(strParsed, Len(strParsed) - 1) + "]"
            Else
                strParsed = strParsed + AddNodeIfNew(objCurNode, rngBuildingRef)
            End If
            strSheetPrefix = ""
        End If
        
        
        '======================================================================
        ' STEP 2.C. Store parsed formula in the DAG
skipTokenising:
        If objCurNode.boolIsConstant = False Then objCurNode.strFormulaParsed = strParsed
        If left(objCurNode.strAddress, 1) = "_" Then
            ' Root node is actually a formula or constant, so return the parsed formula
            ExpandFormula = strParsed
        End If
        If objCurNode.boolIsRoot And objCurNode.boolIsConstant Then
            ' Was just a simple constant, no need to create anything in Formulae
            ExpandFormula = objCurNode.strFormulaParsed
            Formulae.Remove lngIndex
            Exit Function
        End If
        
        '======================================================================
        ' On to the next node
        lngIndex = lngIndex + 1
    Loop
    
    Exit Function
End Function

Function FindPosition(Word As String, FindString As String, Count As Long, Optional start As Long = 1) As Long
    Dim It As Long
    
    ' Define iterator
    It = 1
    
    ' Define initial position
    FindPosition = start
    
    ' Loop until count reached
    While It < Count
        FindPosition = FindPosition + InStr(1, Word, FindString)
        It = It + 1
    Wend
End Function

Sub CleanFormulae(ByVal lngFStart As Long, ByVal lngFEnd As Long)

    Dim objCurNode As CFormula, lngIndex As Long, varReturn As Variant

    ' Any AdjCellUnknown can now be decided one way or the other
    Dim lngCurDepth As Long, strChildAddr As Variant
    For lngCurDepth = lngMaxDepth To 0 Step -1
        For lngIndex = lngFStart To lngFEnd
            Set objCurNode = Formulae(lngIndex)
            If objCurNode.lngDepth = lngCurDepth Then
                If objCurNode.intRefsAdjCell = AdjCellUnknown And Not objCurNode.boolIsRoot Then
                    ' Look through children - if any are adj cell depedendant,
                    ' this is too. Otherwise, its not!
                    
                    objCurNode.intRefsAdjCell = AdjCellIndependent
                    For Each strChildAddr In objCurNode.Children
                        If Formulae(strChildAddr).intRefsAdjCell = AdjCellDependent Then
                            objCurNode.intRefsAdjCell = AdjCellDependent
                        End If
                    Next
                    ' If its AdjCellIndependent, eval it!
                    If objCurNode.intRefsAdjCell = AdjCellIndependent Then
                        objCurNode.boolCanEval = True
                        ' NOTE: Application.Evaluate has a 255 character limit
                        ' See also: http://dutchgemini.wordpress.com/2009/08/07/error-2015-using-application-evaluate-in-excel-vba/
                        varReturn = Application.Evaluate(objCurNode.strFormula)
                        If (VBA.VarType(varReturn) = vbError) Then
                            ' Fall back to calculating the cell
                            objCurNode.rngAddress.Calculate
                            objCurNode.strFormulaParsed = CStr(objCurNode.rngAddress.value)
                            objCurNode.boolEvaledWithCalculate = True
                        Else
                            objCurNode.strFormulaParsed = varReturn
                        End If
                        objCurNode.boolIsConstant = True
                    End If
                End If
            End If
        Next lngIndex
    Next lngCurDepth
    
    ' Pull up constants where possible
    Dim strParentAdr As Variant
    For lngCurDepth = lngMaxDepth To 0 Step -1
        For lngIndex = lngFStart To lngFEnd
            If Formulae(lngIndex).lngDepth = lngCurDepth Then
                If Formulae(lngIndex).boolIsConstant Then
                    ' Constant - so can pull up into its parents
                    Set objCurNode = Formulae(lngIndex)
                    
                    For Each strParentAdr In objCurNode.Parents
                        With Formulae(strParentAdr)
                            .strFormulaParsed = Replace(.strFormulaParsed, objCurNode.strAddress, objCurNode.strFormulaParsed)
                        End With
                    Next
                    
                    objCurNode.boolCanBeRemoved = Not objCurNode.boolIsRoot
                End If
            End If
        Next lngIndex
    Next lngCurDepth
    
    ' Remove any unneeded nodes
    For lngIndex = lngFEnd To lngFStart Step -1
        If Formulae(lngIndex).boolCanBeRemoved Then Formulae.Remove lngIndex
    Next lngIndex
End Sub

Sub PUSHDOWN(ByRef objNode As CFormula, ByVal lngNewDepth As Long)
    If lngNewDepth > lngMaxDepth Then lngMaxDepth = lngNewDepth
    If lngNewDepth > objNode.lngDepth Then objNode.lngDepth = lngNewDepth
    'Dim child As Long
    'For child = 1 To objNode.Children.Count
    '    PUSHDOWN Formulae(objNode.Children(child)), objNode.lngDepth + 1, Formulae, lngMaxDepth
    'Next
    Dim child As Variant
    For Each child In objNode.Children
        If TestKeyExists(Formulae, CStr(child)) Then PUSHDOWN Formulae(child), objNode.lngDepth + 1
    Next
End Sub

Function AddNodeIfNew(ByRef objCurNode As CFormula, ByRef c As Range) As String
    Dim cell As Range
    
    ' 1.1 Get the standard name of this cell
    Dim strStdName As String
    strStdName = ConvertCellToStandardName(c)
    
    ' Check if cell is an adjustable cell
    If c.Parent.Name = SolverModelSheet.Name Then
        If Not Intersect(c, AdjustableCells) Is Nothing Then
            AddNodeIfNew = strStdName
            GoTo finishedNode
        End If
    End If
    '==========================================================================
    ' 0. Is Simple Cell?
    ' 0.1 Get the standard name of this cell
    Dim strCFormula As String
    strCFormula = c.Formula
    ' 0.2 Is blank cell (TODO: breaks some IF statements)
    If Len(strCFormula) = 0 Then
        AddNodeIfNew = "0"
        GoTo finishedNode
    End If
    ' 0.3 Is constant?
    If left(strCFormula, 1) <> "=" Then
        ' Number?
        If IsAmericanNumber(strCFormula) Then
            AddNodeIfNew = strCFormula
        Else
            ' Not a number, assume text
            AddNodeIfNew = "'" + strCFormula + "'"
        End If
        GoTo finishedNode
    End If
            
    '==========================================================================
    ' 1. Is Existing Node?
    ' 1.2 Does the node exist yet?
    If TestKeyExists(Formulae, strStdName) Then
        ' Node already exists
        Dim objExistNode As CFormula
        Set objExistNode = Formulae(strStdName)
        ' Has it been evaluated?
        If objExistNode.boolIsConstant Then
            ' Its a constant, so we can just fold it straight in
            AddNodeIfNew = objExistNode.strFormulaParsed
            GoTo finishedNode
        End If
        ' Is it dependent on a decision variable?
        If objExistNode.intRefsAdjCell = AdjCellDependent Then
            ' That means this cell is also dependent on adj. cells, which
            ' we may not have known
            objCurNode.intRefsAdjCell = AdjCellDependent
        End If
        ' Add this node to the current nodes Children list
        If Not TestKeyExists(objCurNode.Children, strStdName) Then
            objCurNode.Children.Add strStdName, strStdName
        End If
        ' Tell it about its new parent
        If Not TestKeyExists(objExistNode.Parents, objCurNode.strAddress) Then
            objExistNode.Parents.Add objCurNode.strAddress, objCurNode.strAddress
        End If
        ' Ensure it is at the correct depth
        PUSHDOWN objExistNode, objCurNode.lngDepth + 1
        ' Return the standardised name of this existing node
        AddNodeIfNew = strStdName
        GoTo finishedNode
    End If
    
    '==========================================================================
    ' 2. Is Adjustable Cell?
    Dim varReturn As Variant
    Dim objNewNode As New CFormula
    Set objNewNode.rngAddress = c
    objNewNode.strAddress = strStdName
    ' Is on same sheet?
    If objNewNode.GetSheet = AdjustableCells.Parent.Name Then
        ' This formula is on the same sheet as the adjacent cells
        ' Thus we can safely check if this formula's cell is depedent on
        ' a decision variable
        If objNewNode.IsDependentOn(AdjustableCells.Dependents) Then
            ' This formula depends on the value of the adjustable cells
            ' This means we can NOT evaluate it
            objNewNode.boolCanEval = False
            objNewNode.boolIsConstant = False
            objNewNode.intRefsAdjCell = AdjCellDependent
        Else
            ' This formula is NOT dependent on an adjustable cell, so we
            ' can just evaluate its formula (TODO: or take its value?)
            ' We know its not a simple constant, or we would of got
            ' it earlier
            ' NOTE: Application.Evaluate has a 255 character limit
            
            varReturn = Application.Evaluate(strCFormula)
            If (VBA.VarType(varReturn) = vbError) Then
                ' Fall back to calculating the cell
                c.Calculate
                AddNodeIfNew = CStr(c.value)
                GoTo finishedNode
            Else
                AddNodeIfNew = varReturn
                GoTo finishedNode
            End If
        End If
    Else
        ' This formula is on a different sheet to the adjacent cells.
        ' This means we can't check if its depedent on an adjacent cell,
        ' because Range. Dependents only returns the cells for the same
        ' sheet as Range.Parent
        'objNewNode.boolCanEval = False
        'objNewNode.boolIsConstant = False
        'objNewNode.intRefsAdjCell = AdjCellUnknown
        varReturn = Application.Evaluate(strCFormula)
            If (VBA.VarType(varReturn) = vbError) Then
                ' Fall back to calculating the cell
                c.Calculate
                AddNodeIfNew = CStr(c.value)
                GoTo finishedNode
            Else
                AddNodeIfNew = varReturn
                GoTo finishedNode
            End If
    End If
    
    '==========================================================================
    ' 3. Create new node
    With objNewNode
        .strFormula = strCFormula
        .initialValue = CDbl(c.Value2)
        .lngDepth = objCurNode.lngDepth + 1
        If .lngDepth > lngMaxDepth Then lngMaxDepth = .lngDepth
        .Parents.Add objCurNode.strAddress, objCurNode.strAddress
    End With
    ' Add this new node to the child list of current node
    objCurNode.Children.Add strStdName, strStdName
    ' Add new node to the DAG
    Formulae.Add objNewNode, strStdName
    ' Return the standardised name of this new node
    AddNodeIfNew = strStdName

finishedNode:
End Function

Function ConvertFormula(Solver As String, tokenText As String) As String
    Select Case Solver
    Case "PuLP"
        ConvertFormula = ConvertFormula_PuLP(tokenText)
    Case "NeosBon", "NeosCou"
        ConvertFormula = ConvertFormula_AMPL(tokenText)
    Case "Couenne", "Bonmin"
        ConvertFormula = ConvertFormula_NL(tokenText)
    Case Else
        ConvertFormula = tokenText + "("
    End Select
End Function

Function ConvertFormula_AMPL(tokenText As String) As String
   tokenText = LCase(tokenText)
    Select Case tokenText
    ' RADIANS
    Case "radians"
        ConvertFormula_AMPL = "pi/180*("
    Case "min", "max"
        If Solver = "NeosCou" Then
            Err.Raise OpenSolver_BuildError, "Parsing model formulae", "The operator " & UCase(tokenText) & " is not yet supported by Couenne."
        End If
        ConvertFormula_AMPL = tokenText + "("
    Case Else
        ConvertFormula_AMPL = tokenText + "("
    End Select
    Exit Function
    
    ' If function is broken jump to here
errorHandler:
    MsgBox tokenText & " not implemented for AMPL yet"
    ConvertFormula_AMPL = tokenText + "("
End Function

Function ConvertFormula_NL(tokenText As String) As String
    Select Case tokenText
    ' RADIANS
    Case "RADIANS"
        ConvertFormula_NL = "4*ATAN(1)/180*("
    Case "MIN", "MAX"
        If Solver = "Couenne" Then
            Err.Raise OpenSolver_BuildError, "Parsing model formulae", "The operator " & UCase(tokenText) & " is not yet supported by Couenne."
        End If
        ConvertFormula_NL = tokenText + "("
    Case Else
        ConvertFormula_NL = tokenText + "("
    End Select
    Exit Function
    
    ' If function is broken jump to here
errorHandler:
    MsgBox tokenText & " not implemented for .nl solvers yet"
    ConvertFormula_NL = tokenText + "("
End Function


