Attribute VB_Name = "OpenSolverModelValidation"
Option Explicit

Public Sub ValidateObjectiveFunctionCell(ObjectiveFunctionCell As Range)
    If Not ObjectiveFunctionCell Is Nothing Then
        ' Check objective is a single cell
        If ObjectiveFunctionCell.Count <> 1 Then Err.Raise OpenSolver_ModelError, Description:="The objective function cell must be a single cell."
    End If
End Sub
Public Sub ValidateObjectiveFunctionCellRefersTo(ObjectiveFunctionCellRefersTo As String)
    Dim ObjectiveFunctionCell As Range
    Set ObjectiveFunctionCell = GetRefersToRange(ObjectiveFunctionCellRefersTo)
    ValidateObjectiveFunctionCell ObjectiveFunctionCell
End Sub

Public Sub ValidateDecisionVariables(DecisionVariables As Range)
     ' We allow empty range to be set, it will just throw an error when solving
     'If DecisionVariables Is Nothing Then Err.Raise OpenSolver_ModelError, Description:="The adjustable cells must be specified"
End Sub
Public Sub ValidateDecisionVariablesRefersTo(DecisionVariablesRefersTo As String)
    Dim DecisionVariables As Range
    Set DecisionVariables = GetRefersToRange(DecisionVariablesRefersTo)
    ValidateDecisionVariables DecisionVariables
End Sub

Public Sub ValidateConstraint(LHSRange As Range, Relation As RelationConsts, RHSRange As Range, RHSFormula As String, sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    
    If LHSRange.Areas.Count > 1 Then
        Err.Raise OpenSolver_ModelError, Description:="Left-hand-side of constraint must have only one area."
    End If
    
    If RelationHasRHS(Relation) Then
        If Not RHSRange Is Nothing Then
            If RHSRange.Count > 1 And RHSRange.Count <> LHSRange.Count Then
                Err.Raise OpenSolver_ModelError, Description:="Right-hand-side of constraint has more than one cell, and does not match the number of cells on the left-hand-side."
            End If
        Else
            ' Try to convert it to a US locale string internally
            Dim internalRHS As String
            internalRHS = ConvertFromCurrentLocale(RHSFormula)

            ' Can we evaluate this function or constant?
            Dim varReturn As Variant
            varReturn = sheet.Evaluate(internalRHS) ' Must be worksheet.evaluate to get references to names local to the sheet
            If VBA.VarType(varReturn) = vbError Then
                ' We want to catch any error that arises from the *structure* of the formula, **not** the current values of its precedent cells
                Select Case CLng(varReturn)
                Case 2000  ' #NULL!
                    ' Allow
                Case 2007  ' #DIV/0!
                    ' Allow
                Case 2015  ' #VALUE!
                    ' Invalid formula structure
                    Err.Raise OpenSolver_ModelError, Description:="The formula or value for the RHS is not valid. Please check and try again."
                Case 2023  ' #REF!
                    ' Missing cell reference inside formula
                    Err.Raise OpenSolver_ModelError, Description:="The RHS is marked #REF!, indicating it has been deleted. Please fix and try again."
                Case 2029  ' #NAME?
                    ' Valid formula that evaluates to an error
                    ' Allow
                Case 2036  ' #NUM!
                    ' Allow
                Case 2042  ' #N/A
                    ' Allow
                End Select
            ElseIf VarType(varReturn) = vbArray + vbVariant Then  ' Formula evaluates to array
                Err.Raise OpenSolver_ModelError, Description:="The formula for the RHS is not valid because it evaluates to multiple cells. The result of the formula must be a single numeric value. Please fix this and try again."
            End If

            ' Convert any cell references to absolute
            If Left(internalRHS, 1) <> "=" Then internalRHS = "=" & internalRHS
            varReturn = Application.ConvertFormula(internalRHS, FromReferenceStyle:=xlA1, ToReferenceStyle:=xlA1, ToAbsolute:=xlAbsolute)

            If (VBA.VarType(varReturn) = vbError) Then
                ' Its valid, but couldn't convert to standard form, probably because not A1... just leave it
            Else
                ' Always comes back with a = at the start
                ' Unfortunately, return value will have wrong locale...
                ' But not much can be done with that?
                internalRHS = Mid(varReturn, 2, Len(varReturn))
            End If
            RHSFormula = internalRHS
        End If
        
    Else
        If Not RHSRange Is Nothing Or _
           (RHSFormula <> "" And RHSFormula <> "integer" And RHSFormula <> "binary" And RHSFormula <> "alldiff") Then
            Err.Raise OpenSolver_ModelError, Description:="No right-hand-side is permitted for this relation"
        End If
    End If
End Sub

Public Sub ValidateConstraintRefersTo(LHSRefersTo As String, Relation As RelationConsts, RHSRefersTo As String, sheet As Worksheet)
    Dim LHSRange As Range, RHSRange As Range, RHSFormula As String
    ConvertRefersToConstraint LHSRefersTo, RHSRefersTo, LHSRange, RHSRange, RHSFormula
    
    If LHSRange Is Nothing Then
        Err.Raise OpenSolver_ModelError, Description:="Left-hand-side of constraints must be a range."
    End If
    
    If RelationHasRHS(Relation) Then
        If Len(RHSRefersTo) = 0 Then
            Err.Raise OpenSolver_ModelError, Description:="Right-hand-side cannot be blank!"
        End If
    End If
    
    ValidateConstraint LHSRange, Relation, RHSRange, RHSFormula, sheet
End Sub

Public Sub ValidateDualsRefersTo(DualsRefersTo As String)
    Dim Duals As Range
    Set Duals = GetRefersToRange(DualsRefersTo)
    ValidateDuals Duals
End Sub
Public Sub ValidateDuals(Duals As Range)
    If Not Duals Is Nothing Then
        If Duals.Count <> 1 Then Err.Raise OpenSolver_ModelError, Description:="The dual range must be a single cell."
    End If
End Sub

Public Sub ConvertRefersToConstraint(LHSRefersTo As String, RHSRefersTo As String, LHSRange As Range, RHSRange As Range, RHSFormula As String)
    On Error Resume Next
    Set LHSRange = GetRefersToRange(LHSRefersTo)
    Set RHSRange = GetRefersToRange(RHSRefersTo)
    On Error GoTo 0
    If RHSRange Is Nothing Then RHSFormula = RHSRefersTo
End Sub

Sub ValidateSolverParameters(SolverParameters As Range)
    If SolverParameters Is Nothing Then Exit Sub
    If SolverParameters.Areas.Count > 1 Or SolverParameters.Columns.Count <> 2 Then
        Err.Raise OpenSolver_SolveError, Description:="The Extra Solver Parameters range must be a single two-column table of keys and values."
    End If
End Sub

Sub ValidateSolverParametersRefersTo(SolverParametersRefersTo As String)
    Dim SolverParameters As Range
    Set SolverParameters = GetRefersToRange(SolverParametersRefersTo)
    ValidateSolverParameters SolverParameters
End Sub

Sub ValidateQuickSolveParameters(QuickSolveParameters As Range, sheet As Worksheet)
    If QuickSolveParameters.Worksheet.Name <> sheet.Name Then
        Err.Raise OpenSolver_BuildError, Description:="The parameter cells need to be on the current worksheet."
    End If
End Sub
Sub ValidateQuickSolveParametersRefersTo(QuickSolveParametersRefersTo As String, sheet As Worksheet)
    Dim QuickSolveParameters As Range
    Set QuickSolveParameters = GetRefersToRange(QuickSolveParametersRefersTo)
    ValidateQuickSolveParameters QuickSolveParameters, sheet
End Sub
