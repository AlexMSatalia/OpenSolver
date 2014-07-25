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
    
Public ObjectiveSense As ObjectiveSenseType
Public ObjectiveTargetValue As Double
Public ObjectiveCell As Range
Public ObjectiveCellConstant As Double

Public ConstraintCount As Long
Public IntegerCells As Range
Public BinaryCells As Range
Private LHSType() As SolverConstraintType
Private RELType() As RelationConsts
Private RHSType() As SolverConstraintType

Public AssumeNonNegative As Boolean

Public Solver As String

Public Formulae As Collection
Public VarTypeMap As Collection
Public AdjCellNameMap As Collection
Public LHSKeys As Collection
Public RHSKeys As Collection
Public Rels As Collection

Private lngMaxDepth As Long

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
Public Sub ProcessSolverModel(SolverType As String)
    Solver = SolverType
    ProcessAdjustableCells
    ProcessObjective
    ProcessConstraints
    ProcessVarTypes
    ParseFormulae
    ProcessModelOptions
End Sub

'==============================================================================
Private Sub ProcessAdjustableCells()
    
    ' Try to get the range
    On Error Resume Next
    Set AdjustableCells = Range(strNameRoot & NameAdjCells)
    ' Catch any error related to either the model being incomplete,
    ' or invalid entry for adjustable cells.
    If Err.Number <> 0 Then
        On Error GoTo 0 'Stop Resume Next
        RaiseOSErr OSErrAdjCellMissing
    End If
    On Error GoTo 0
    
    ' Range might have multiple areas that overlap
    ' This will cause issues later on, so we remove them
    Set AdjustableCells = RemoveRangeOverlap(AdjustableCells)
    
    ' Handle any merged ranges
    Dim BadMergedCell As Range
    If Not CheckRangeContainsNoAmbiguousMergedCells(AdjustableCells, BadMergedCell) Then
        RaiseOSErr OSErrAdjCellBadMerge, BadMergedCell
    End If
    
    ' Check we have any cells after all this
    AdjustableCellCount = AdjustableCells.Count
    If AdjustableCellCount < 1 Then RaiseOSErr OSErrAdjCellNone
    
    ' Get names for all the variables
    ReDim AdjCellNames(1 To AdjustableCellCount) As String
    Dim AdjCell As Range, i As Long
    i = 0
    For Each AdjCell In AdjustableCells
        i = i + 1
        AdjCellNames(i) = AdjCell.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) ' eg A1
    Next AdjCell
End Sub


'==============================================================================
Private Sub ProcessObjective()

    Dim errorPrefix As String
    errorPrefix = "OpenSolver Non-Linear Model Build"

    ' Get the objective sense
    Dim temp As Long
    ObjectiveSense = UnknownObjectiveSense
    If modNameHandlers.GetIntegerFromName(book, strNameRoot & NameObjSense, temp) Then ObjectiveSense = temp
    If ObjectiveSense = TargetObjective Then
        ' Target will be 0 if this fails
        modNameHandlers.GetValueFromName book, strNameRoot & NameObjTarget, ObjectiveTargetValue
    End If
    
    ' Get and check the objective function
    Dim objIsMissing As Boolean, objIsRange As Boolean, objIsFormula As Boolean
    Dim objIsError As Boolean, sRefersTo As String, objValue As Double
    GetValueOrRangeFromName book, strNameRoot & NameObjCell, objIsMissing, objIsRange, ObjectiveCell, _
                            objIsFormula, objIsError, sRefersTo, objValue
    
    ' Valdiate objective function:
    ' 1. Name must exist
    If objIsMissing Then Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="OpenSolver cannot find the objective ('solver_opt' is out of date). Please re-enter the objective, and try again."
    ' Objective is corrupted somehow
    If objIsError Then Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="The objective is marked #REF!, indicating this cell has been deleted. Please fix the objective, and try again."
    ' Objective has a value that is not a number
    If VarType(ObjectiveCell.value) <> vbDouble Then
        If VarType(ObjectiveCell.value) = vbError Then
            Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="The objective cell appears to contain an error. This could have occurred if there is a divide by zero error or if you have used the wrong function (eg #DIV/0! or #VALUE!). Please fix this, and try again."
        Else
            Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="The objective cell does not appear to contain a numeric value. Please fix this, and try again."
        End If
    End If
    
    ' Objective is calid.
    ' Store the value currently in the objective cell.
    ' NOTE: This assumes the decision variables have already been set to zero
    ObjectiveCell.Calculate
    ObjectiveCellConstant = ObjectiveCell.value

End Sub


'==============================================================================
Private Sub ProcessConstraints()
    
    ' Try and get the number of constraints
    If Not GetIntegerFromName(book, strNameRoot + NameConsCount, ConstraintCount) Then
        RaiseOSErr OSErrConsMissing
    End If
    
    If ConstraintCount < 1 Then Exit Sub
    
    ' Create an array to store the details of the LHS and RHS, which
    ' is useful for parsing and classic
    ReDim LHSType(1 To ConstraintCount) As SolverConstraintType
    ReDim RELType(1 To ConstraintCount) As RelationConsts
    ReDim RHSType(1 To ConstraintCount) As SolverConstraintType
    
    ' Iterate through all the constraints and check they are OK
    Dim i As Integer
    For i = 1 To ConstraintCount
        ' Rules for valid constraints:
        ' The LHS mut be a range with one or more cells
        ' The RHS can be either:
        '  - A single-cell range (=A4)
        '  - A multi-cell range of the same size as the LHS (=A4:B5)
        '  - A single constant value (eg =2)
        '  - A formula returning a single value (eg =sin(A4))
        
        ' LHS
        Dim LHSisMissing As Boolean, LHSisRange As Boolean, rLHS As Range
        Dim LHSisFormula As Boolean, LHSisError As Boolean
        Dim LHSRefersTo As String, LHSValue As Double
        GetValueOrRangeFromName book, strNameRoot & NameLHS & CStr(i), _
                LHSisMissing, LHSisRange, rLHS, LHSisFormula, _
                LHSisError, LHSRefersTo, LHSValue
        If LHSisMissing Then RaiseOSErr OSErrConsLHSMissing, CStr(i)
        If LHSisError Then RaiseOSErr OSErrConsLHSError, CStr(i)
        If Not LHSisRange Then RaiseOSErr OSErrConsLHSNotRange, CStr(i)
        If rLHS.Count = 1 Then
            LHSType(i) = ConIsSingleCell
        Else
            LHSType(i) = ConIsMultipleCell
        End If
        
        ' REL
        Dim rel As Long
        GetIntegerFromName book, strNameRoot & NameREL & CStr(i), rel
        
        If rel = RelationBIN Or rel = RelationINT Then
            ' Record that these variables are type constrained and check
            ' that they are adjustable cells
            RELType(i) = rel
            If Intersect(rLHS, AdjustableCells).Count = rLHS.Count Then
                If rel = RelationINT Then
                    Set IntegerCells = ProperUnion(IntegerCells, rLHS)
                Else 'if the relation is binary
                    Set BinaryCells = ProperUnion(BinaryCells, rLHS)
                End If
            Else
                RaiseOSErr OSErrConsIntBinOnNonAdjustable, CStr(i)
            End If
        Else
            ' Check it isn't something like 'alldiff' from Excel 2010
            If rel <> RelationLE And rel <> RelationEQ And rel <> RelationGE Then
                RaiseOSErr OSErrConsUnknownRel, CStr(i)
            End If
            RELType(i) = rel
                
            ' RHS
            Dim RHSisMissing As Boolean, RHSisRange As Boolean, rRHS As Range
            Dim RHSisFormula As Boolean, RHSisError As Boolean
            Dim RHSRefersTo As String, RHSValue As Double
            GetValueOrRangeFromName book, strNameRoot & NameRHS & CStr(i), _
                    RHSisMissing, RHSisRange, rRHS, RHSisFormula, _
                    RHSisError, RHSRefersTo, RHSValue
            If RHSisMissing Then RaiseOSErr OSErrConsRHSMissing, CStr(i)
            If RHSisError Then RaiseOSErr OSErrConsRHSError, CStr(i)
            
            ' Record what type the RHS is
            If RHSisRange Then
                If rRHS.Count = 1 Then RHSType(i) = ConIsSingleCell
                If rRHS.Count > 1 Then RHSType(i) = ConIsMultipleCell
            Else
                RHSType(i) = ConIsValueOrFormula
            End If
            
            ' If RHS is a range, need to have an understandable constraint system
            Dim LHSCount As Long, RHSCount As Long
            LHSCount = rLHS.Count
            RHSCount = 1
            If RHSisRange Then RHSCount = rRHS.Count
            If (LHSCount <> 1 And RHSCount <> 1) And (LHSCount <> RHSCount) Then
                RaiseOSErr OSErrConsBadCellCount, CStr(i)
            End If
        End If
    Next i
End Sub

Private Sub ProcessVarTypes()
' Get var types of constraints
    
    Set VarTypeMap = New Collection
    Set AdjCellNameMap = New Collection
    Dim curVarType As Integer, lngArea As Long, cellName As String, c As Range
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

Private Sub ParseFormulae()
' Parses all constraints and objective formulae to create formulae that depend only on variable cells and constants

    ' Create a collection for the spreadsheet's formulae (DAG)
    Set Formulae = New Collection
    ' Remember the greatest depth of the DAG
    lngMaxDepth = 0

    '==========================================================================
    ' STEP 1. Parse objective cell
    Dim objresult As String
    
    objresult = ExpandFormula(ObjectiveCell.Formula, ObjectiveCell)
    
    If objresult Like "*Error*" Then
        Err.Raise OpenSolver_BuildError, Description:="The objective contains an error"
    End If
    CleanFormulae 1, Formulae.Count
    
    '==========================================================================
    ' STEP 2. Parse constraints
    Set LHSKeys = New Collection
    Set RHSKeys = New Collection
    Set Rels = New Collection

    Dim i As Long
    Dim Count As Integer
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
                
        For cRow = 1 To cRowCount
            For cCol = 1 To cColCount
                Count = Count + 1
                
                If cRowCount <> 1 Or cColCount <> 1 Then
                    curLHSFormulae = varLHSFormulae(cRow, cCol)
                Else
                    curLHSFormulae = varLHSFormulae
                End If
                LHSKeys.Add ExpandFormula(curLHSFormulae, rngLHS(cRow, cCol))

                ' Determine appropriate relation
                Rels.Add ConvertRelationToEnum(nameRELi.value)
                    
                ' Parse RHS
                Dim strRHSFormula As String, rngRHSCell As Range
                ' 1.1 RHS is a single cell
                If GetRHSType(i) = ConIsSingleCell Then
                    strRHSFormula = nameRHSi.RefersToRange.Formula
                    Set rngRHSCell = nameRHSi.RefersToRange
                ' 1.2 RHS is multiple cells
                ElseIf GetRHSType(i) = ConIsMultipleCell Then
                    strRHSFormula = nameRHSi.RefersToRange.Cells(cRow, cCol).Formula
                    Set rngRHSCell = nameRHSi.RefersToRange.Cells(cRow, cCol)
                ' 1.3 RHS is value or formula
                ElseIf GetRHSType(i) = ConIsValueOrFormula Then
                    strRHSFormula = nameRHSi.value
                    Set rngRHSCell = Nothing
                End If
                
                RHSKeys.Add ExpandFormula(strRHSFormula, rngRHSCell)
            Next cCol
        Next cRow
        
        ' Clean Formulae
        CleanFormulae lngFormulaeCountBefore + 1, Formulae.Count
        
NextCons:
    Next i
    
    '==========================================================================
    ' STEP 3. Make all possible substitutions in formula tree
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

End Sub

'==============================================================================
Private Sub ProcessModelOptions()
    
    ' Assume Non-negative
    ' Excel adds a zero lower bound to any variable which does not have any
    ' other lower bound specified.
    Dim s As String
    AssumeNonNegative = False
    If GetStringFromName(book, strNameRoot & NameNonNegative, s) Then
        AssumeNonNegative = (s = "1")
    End If
    
    ' Excel 2007: Assume Linear Model checked.
    ' Excel 2010: Engine Type = Simplex.
    Dim AssumeLinearModel As Boolean ' Excel stores ...!solver_lin=1 if Assume Linear Model is turned on, 2 if off (NB: Not 0!)
    Dim SimplexEngineSelected As Boolean '2010 excel - stores the type of engine used
    
    ' Try the 2007 name - may be present if opening a 2007 file in 2010.
    AssumeLinearModel = False
    If GetStringFromName(book, strNameRoot & NameLinear, s) Then
        AssumeLinearModel = (s = "1")
    End If
    
    ' Try to get the 2010 optimisation engine
    SimplexEngineSelected = False
    If Application.Version = "14.0" And GetStringFromName(book, strNameRoot & NameEngine, s) Then
        SimplexEngineSelected = (s = "2")
    End If
        
    ' If we are in Excel 2010, and the user has set "Solve Relaxtion",
    ' then we report an error telling them to turn it off or to use
    ' the solve relaxation command
    If Application.Version = "14.0" And GetStringFromName(book, strNameRoot & NameRelaxation, s) Then
        If (s = "1") Then RaiseOSErr OSErrRelaxationSet
    End If
    
    ' If user has not enabled linearity, we need to ask them about it
    If (Not SimplexEngineSelected And Not AssumeLinearModel) Then
        Dim strLinearityMsg As String
        strLinearityMsg = "OpenSolver assumes the model is linear." + vbNewLine
        If Val(Application.Version) <= 12 Then ' Excel 2007 or earlier
            strLinearityMsg = strLinearityMsg + "However, Assume Linear Model is not turned on." + vbNewLine
        Else ' Excel 2010
            strLinearityMsg = strLinearityMsg + "However, the Solving Method is not set as Simplex." + vbNewLine
        End If
        strLinearityMsg = strLinearityMsg + "If you would like to correct this, click Yes." + vbNewLine
        strLinearityMsg = strLinearityMsg + "Click No to proceed assuming the model is linear."
        
        Dim msgResponse As VbMsgBoxResult
        'msgResponse = MsgBox(strLinearityMsg, vbYesNoCancel, "OpenSolver")
        
        If msgResponse = vbCancel Then
            RaiseOSErr OSErrUserCancel
        ElseIf msgResponse = vbYes Then
            If Val(Application.Version) <= 12 Then ' Excel 2007 or earlier
                SetSolverNameOnSheet "lin", "=1"
            Else ' Excel 2010
                SetSolverNameOnSheet "eng", "=2"
            End If
        ElseIf msgResponse = vbNo Then
            ' We'll just assume its linear and keep on going
        End If
    End If
    
End Sub

'==============================================================================
Public Function GetLHSType(lngConNum As Long) As SolverConstraintType
    GetLHSType = LHSType(lngConNum)
End Function
Public Function GetLHSRange(lngConNum As Long) As Range
    Set GetLHSRange = book.Names(strNameRoot + NameLHS + CStr(lngConNum)).RefersToRange
End Function

Public Function GetRELType(lngConNum As Long) As RelationConsts
    GetRELType = RELType(lngConNum)
End Function

Public Function GetRHSType(lngConNum As Long) As SolverConstraintType
    GetRHSType = RHSType(lngConNum)
End Function
Public Function GetRHSRange(lngConNum As Long) As Range
    Set GetRHSRange = book.Names(strNameRoot + NameRHS + CStr(lngConNum)).RefersToRange
End Function
Public Function GetRHSValue(lngConNum As Long) As String
    GetRHSValue = book.Names(strNameRoot + NameRHS + CStr(lngConNum)).RefersTo
End Function

Public Function GetAdjCellName(lngVar As Long) As String
    GetAdjCellName = AdjCellNames(lngVar)
End Function

Function ExpandFormula(ByVal strFormula As String, _
                       ByRef rngSourceCell As Range) As String
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
    Dim FunctionCount As Integer
    Dim start As Integer
    Dim output As String, node As String
    Dim Count As Integer
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
        Set objCurNode.rngAddress = rngSourceCell
        ExpandFormula = strRootAddress
    Else
        ' Its a constant or formula - return the parsed formula directly
        objCurNode.strAddress = "!" + CStr(Formulae.Count)
        Set objCurNode.rngAddress = Nothing
    End If
    objCurNode.strFormula = strFormula  ' Could also get from rngSourceCell
    objCurNode.lngDepth = 0             ' Root node depth is 0
    objCurNode.boolIsRoot = True
    Formulae.Add objCurNode, objCurNode.strAddress ' Add it to the DAG
    ' Start at the root node
    Dim lngIndex As Long
    lngIndex = lngBaseIndex + 1
    
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
        Set tksFormula = modTokeniser.ParseFormula(objCurNode.strFormula)
    
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
        Dim i As Integer, c As Range, tkn As Token
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
                    RaiseOSErr OSErrPulpTokenErrText
                    
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
                        For Each c In rngBuildingRef
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
        If left(objCurNode.strAddress, 1) = "!" Then
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

Function FindPosition(Word As String, FindString As String, Count As Integer, Optional start As Integer = 1) As Integer
    Dim It As Integer
    
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
    
    ' Check if cell is adjustable cell, otherwise use value
    For Each cell In AdjustableCells
        If ConvertCellToStandardName(cell) = strStdName Then
            AddNodeIfNew = strStdName
            GoTo finishedNode
        End If
    Next
    
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
    ' TODO: Unhandled yet
    Case "min", "max"
        GoTo errorHandler
    Case Else
        ConvertFormula_AMPL = tokenText + "("
    End Select
    Exit Function
    
errorHandler:
    MsgBox tokenText & " not implemented for AMPL yet"
    ConvertFormula_AMPL = tokenText + "("
End Function
