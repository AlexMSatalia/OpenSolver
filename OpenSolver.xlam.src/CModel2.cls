VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CModel2"
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
Public Sub ProcessSolverModel()
    ProcessAdjustableCells
    ProcessObjective
    ProcessConstraints
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
    
    ' Set all the adjustable cells to zero
    'AdjustableCells.value = 0
End Sub


'==============================================================================
Private Sub ProcessObjective()

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

