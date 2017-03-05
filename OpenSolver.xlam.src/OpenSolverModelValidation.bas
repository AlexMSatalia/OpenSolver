Attribute VB_Name = "OpenSolverModelValidation"
Option Explicit

Public Sub ValidateObjectiveFunctionCell(ObjectiveFunctionCell As Range)
1         If Not ObjectiveFunctionCell Is Nothing Then
              ' Check objective is a single cell
2             If ObjectiveFunctionCell.Count <> 1 Then RaiseUserError "The objective function cell must be a single cell."
3         End If
End Sub
Public Sub ValidateObjectiveFunctionCellRefersTo(ObjectiveFunctionCellRefersTo As String)
          Dim ObjectiveFunctionCell As Range
1         Set ObjectiveFunctionCell = GetRefersToRange(ObjectiveFunctionCellRefersTo)
2         ValidateObjectiveFunctionCell ObjectiveFunctionCell
End Sub

Public Sub ValidateDecisionVariables(DecisionVariables As Range)
     ' We allow empty range to be set, it will just throw an error when solving
     'If DecisionVariables Is Nothing Then RaiseUserError "The adjustable cells must be specified"
End Sub
Public Sub ValidateDecisionVariablesRefersTo(DecisionVariablesRefersTo As String)
          Dim DecisionVariables As Range
1         Set DecisionVariables = GetRefersToRange(DecisionVariablesRefersTo)
2         ValidateDecisionVariables DecisionVariables
End Sub

Public Sub ValidateConstraint(LHSRange As Range, Relation As RelationConsts, RHSRange As Range, RHSFormula As String, sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
          
2         If LHSRange.Areas.Count > 1 Then
3             RaiseUserError "Left-hand-side of constraint must have only one area."
4         End If
          
5         If RelationHasRHS(Relation) Then
6             If Not RHSRange Is Nothing Then
7                 If RHSRange.Count > 1 And RHSRange.Count <> LHSRange.Count Then
8                     RaiseUserError "Right-hand-side of constraint has more than one cell, and does not match the number of cells on the left-hand-side."
9                 End If
10            Else
                  ' Try to convert it to a US locale string internally
                  Dim internalRHS As String
11                internalRHS = ConvertFromCurrentLocale(RHSFormula)

                  ' Can we evaluate this function or constant?
                  Dim varReturn As Variant
12                varReturn = sheet.Evaluate(AddEquals(internalRHS)) ' Must be worksheet.evaluate to get references to names local to the sheet
13                If VBA.VarType(varReturn) = vbError Then
                      ' We want to catch any error that arises from the *structure* of the formula, **not** the current values of its precedent cells
14                    Select Case CLng(varReturn)
                      Case 2000  ' #NULL!
                          ' Allow
15                    Case 2007  ' #DIV/0!
                          ' Allow
16                    Case 2015  ' #VALUE!
                          ' Invalid formula structure
17                        RaiseUserError "The formula or value for the RHS is not valid. Please check and try again."
18                    Case 2023  ' #REF!
                          ' Missing cell reference inside formula
19                        RaiseUserError "The RHS is marked #REF!, indicating it has been deleted. Please fix and try again."
20                    Case 2029  ' #NAME?
                          ' Valid formula that evaluates to an error
                          ' Allow
21                    Case 2036  ' #NUM!
                          ' Allow
22                    Case 2042  ' #N/A
                          ' Allow
23                    End Select
24                ElseIf VarType(varReturn) = vbArray + vbVariant Then  ' Formula evaluates to array
25                    RaiseUserError "The formula for the RHS is not valid because it evaluates to multiple cells. The result of the formula must be a single numeric value. Please fix this and try again."
26                End If

                  ' Convert any cell references to absolute
27                If Left(internalRHS, 1) <> "=" Then internalRHS = "=" & internalRHS
28                varReturn = Application.ConvertFormula(internalRHS, FromReferenceStyle:=xlA1, ToReferenceStyle:=xlA1, ToAbsolute:=xlAbsolute)

29                If (VBA.VarType(varReturn) = vbError) Then
                      ' Its valid, but couldn't convert to standard form, probably because not A1... just leave it
30                Else
                      ' Always comes back with a = at the start
                      ' Unfortunately, return value will have wrong locale...
                      ' But not much can be done with that?
31                    internalRHS = Mid(varReturn, 2, Len(varReturn))
32                End If
33                RHSFormula = internalRHS
34            End If
              
35        Else
36            If Not RHSRange Is Nothing Or _
                 (Len(RHSFormula) > 0 And RHSFormula <> "integer" And RHSFormula <> "binary" And RHSFormula <> "alldiff") Then
37                RaiseUserError "No right-hand-side is permitted for this relation"
38            End If
39        End If
End Sub

Public Sub ValidateConstraintRefersTo(LHSRefersTo As String, Relation As RelationConsts, RHSRefersTo As String, sheet As Worksheet)
          Dim LHSRange As Range, RHSRange As Range, RHSFormula As String
1         ConvertRefersToConstraint LHSRefersTo, RHSRefersTo, LHSRange, RHSRange, RHSFormula
          
2         If LHSRange Is Nothing Then
3             RaiseUserError "Left-hand-side of constraints must be a range."
4         End If
          
5         If RelationHasRHS(Relation) Then
6             If Len(RHSRefersTo) = 0 Then
7                 RaiseUserError "Right-hand-side cannot be blank!"
8             End If
9         End If
          
10        ValidateConstraint LHSRange, Relation, RHSRange, RHSFormula, sheet
End Sub

Public Sub ValidateDualsRefersTo(DualsRefersTo As String)
          Dim Duals As Range
1         Set Duals = GetRefersToRange(DualsRefersTo)
2         ValidateDuals Duals
End Sub
Public Sub ValidateDuals(Duals As Range)
1         If Not Duals Is Nothing Then
2             If Duals.Count <> 1 Then RaiseUserError "The dual range must be a single cell."
3         End If
End Sub

Public Sub ConvertRefersToConstraint(LHSRefersTo As String, RHSRefersTo As String, LHSRange As Range, RHSRange As Range, RHSFormula As String)
1         On Error Resume Next
2         Set LHSRange = GetRefersToRange(LHSRefersTo)
3         Set RHSRange = GetRefersToRange(RHSRefersTo)
4         On Error GoTo 0
5         If RHSRange Is Nothing Then RHSFormula = RHSRefersTo
End Sub

Sub ValidateSolverParameters(SolverParameters As Range)
1         If SolverParameters Is Nothing Then Exit Sub
2         If SolverParameters.Areas.Count > 1 Or SolverParameters.Columns.Count <> 2 Then
3             RaiseUserError "The Extra Solver Parameters range must be a single two-column table of keys and values.", _
                             "http://opensolver.org/using-opensolver/#extra-parameters"
4         End If
End Sub

Sub ValidateSolverParametersRefersTo(SolverParametersRefersTo As String)
          Dim SolverParameters As Range
1         Set SolverParameters = GetRefersToRange(SolverParametersRefersTo)
2         ValidateSolverParameters SolverParameters
End Sub

Sub ValidateQuickSolveParameters(QuickSolveParameters As Range, sheet As Worksheet)
1         If QuickSolveParameters.Worksheet.Name <> sheet.Name Then
2             RaiseUserError "The parameter cells need to be on the current worksheet."
3         End If
End Sub
Sub ValidateQuickSolveParametersRefersTo(QuickSolveParametersRefersTo As String, sheet As Worksheet)
          Dim QuickSolveParameters As Range
1         Set QuickSolveParameters = GetRefersToRange(QuickSolveParametersRefersTo)
2         ValidateQuickSolveParameters QuickSolveParameters, sheet
End Sub

Sub ValidateTolerance(Tolerance As Double)
1         If Tolerance < 0 Or Tolerance > 1 Then
2             RaiseUserError "Tolerance needs to be between 0 and 1."
3         End If
End Sub
Sub ValidateToleranceAsPercentage(Tolerance As Double)
1         If Tolerance < 0 Or Tolerance > 100 Then
2             RaiseUserError "Tolerance needs to be between 0 and 100."
3         End If
End Sub

Sub ValidateMaxTime(MaxTime As Double)
1         If MaxTime < 0 Then
2             RaiseUserError "Maximum solution time needs to be positive."
3         End If
End Sub

Sub ValidateMaxIterations(MaxIterations As Double)
1         If MaxIterations < 0 Then
2             RaiseUserError "Maximum number of iterations needs to be positive."
3         End If
End Sub

Sub ValidatePrecision(Precision As Double)
1         If Precision < 0 Then
2             RaiseUserError "Precision needs to be positive."
3         End If
End Sub
