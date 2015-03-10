Attribute VB_Name = "OpenSolverAPI"
Option Explicit

Public Function GetAvailableSolvers() As Variant()
    GetAvailableSolvers = Array("CBC", "Gurobi", "NeosCBC", "Bonmin", "Couenne", "NOMAD", "NeosBon", "NeosCou")
End Function

Public Function GetChosenSolver(Optional book As Workbook, Optional sheet As Worksheet) As String
    GetActiveBookAndSheetIfMissing book, sheet
    If Not GetNameValueIfExists(book, EscapeSheetName(sheet) & "OpenSolver_ChosenSolver", GetChosenSolver) Then
        GoTo SetDefault
    End If
    
    ' Check solver is an allowed solver
    On Error GoTo SetDefault
    WorksheetFunction.Match GetChosenSolver, GetAvailableSolvers, 0
    Exit Function
    
SetDefault:
    GetChosenSolver = GetAvailableSolvers()(0)
    SetChosenSolver GetChosenSolver, book, sheet
End Function

Public Sub SetChosenSolver(Solver As String, Optional book As Workbook, Optional sheet As Worksheet)
    ' Check that a valid solver has been specified
    On Error GoTo SolverNotAllowed
    WorksheetFunction.Match Solver, GetAvailableSolvers, 0
        
    SetNameOnSheet "OpenSolver_ChosenSolver", "=" & Solver, book, sheet
    Exit Sub
    
SolverNotAllowed:
    Err.Raise OpenSolver_ModelError, Description:="The specified solver (" & Solver & ") is not in the list of available solvers. " & _
                                                  "Please see the OpenSolverAPI module for the list of available solvers."
End Sub

Public Function GetDualsNewSheet(Optional book As Workbook, Optional sheet As Worksheet) As Boolean
    GetDualsNewSheet = GetNamedBooleanWithDefault("OpenSolver_DualsNewSheet", book, sheet, False)
End Function

Public Sub SetDualsNewSheet(DualsNewSheet As Boolean, Optional book As Workbook, Optional sheet As Worksheet)
    SetBooleanNameOnSheet "OpenSolver_DualsNewSheet", DualsNewSheet, book, sheet
End Sub

Public Function GetUpdateSensitivity(Optional book As Workbook, Optional sheet As Worksheet) As Boolean
    GetUpdateSensitivity = GetNamedBooleanWithDefault("OpenSolver_UpdateSensitivity", book, sheet, True)
End Function

Public Sub SetUpdateSensitivity(UpdateSensitivity As Boolean, Optional book As Workbook, Optional sheet As Worksheet)
    SetBooleanNameOnSheet "OpenSolver_UpdateSensitivity", UpdateSensitivity, book, sheet
End Sub

Public Function GetLinearityCheck(Optional book As Workbook, Optional sheet As Worksheet) As Boolean
    GetActiveBookAndSheetIfMissing book, sheet
    
    Dim value As Long
    If Not GetNamedIntegerIfExists(book, EscapeSheetName(sheet) & "OpenSolver_LinearityCheck", value) Then GoTo SetDefault
    If value <> 2 Then GoTo SetDefault
    GetLinearityCheck = False
    Exit Function
    
SetDefault:
    GetLinearityCheck = True
    SetLinearityCheck GetLinearityCheck, book, sheet
End Function

Public Sub SetLinearityCheck(LinearityCheck As Boolean, Optional book As Workbook, Optional sheet As Worksheet)
    If LinearityCheck Then
        DeleteNameOnSheet "OpenSolver_LinearityCheck", book, sheet
    Else
        SetNameOnSheet "OpenSolver_LinearityCheck", "=2", book, sheet
    End If
End Sub

Public Function GetDuals(Optional book As Workbook, Optional sheet As Worksheet) As Range
    If Not GetNamedRangeIfExistsOnSheet(sheet, "OpenSolver_Duals", GetDuals) Then Set GetDuals = Nothing
End Function

Public Sub SetDuals(Duals As Range, Optional book As Workbook, Optional sheet As Worksheet)
    SetNamedRangeIfExists "OpenSolver_Duals", Duals, book, sheet
End Sub

Public Function GetSolverParameters(Solver As String, Optional book As Workbook, Optional sheet As Worksheet) As Range
    GetActiveBookAndSheetIfMissing book, sheet
    If Not GetNamedRangeIfExistsOnSheet(sheet, "OpenSolver_" & Solver & "Parameters", GetSolverParameters) Then Set GetSolverParameters = Nothing
End Function

Public Sub SetSolverParameters(Solver As String, SolverParameters As Range, Optional book As Workbook, Optional sheet As Worksheet)
    SetNamedRangeIfExists "OpenSolver_" & Solver & "Parameters", SolverParameters, book, sheet
End Sub

Public Function GetQuickSolveParameters(Optional book As Workbook, Optional sheet As Worksheet) As Range
    If Not GetNamedRangeIfExistsOnSheet(sheet, "OpenSolverModelParameters", GetQuickSolveParameters) Then Set GetQuickSolveParameters = Nothing
End Function

Public Sub SetQuickSolveParameters(QuickSolveParameters As Range, Optional book As Workbook, Optional sheet As Worksheet)
    SetNamedRangeIfExists "OpenSolverModelParameters", QuickSolveParameters, book, sheet
End Sub

Public Function GetDecisionVariables(Optional book As Workbook, Optional sheet As Worksheet) As Range
    GetActiveBookAndSheetIfMissing book, sheet
              
    ' We check to see if a model exists by getting the adjustable cells. We check for a name first, as this may contain =Sheet1!$C$2:$E$2,Sheet1!#REF!
    Dim n As Name
    If Not NameExistsInWorkbook(book, EscapeSheetName(sheet) & "solver_adj", n) Then
        Err.Raise Number:=OpenSolver_ModelError, Description:="No Solver model with decision variables was found on sheet " & sheet.Name
    End If
    
    GetNamedRangeIfExistsOnSheet sheet, "solver_adj", GetDecisionVariables
    If GetDecisionVariables Is Nothing Then
        Err.Raise OpenSolver_ModelError, Description:="A model was found on the sheet " & sheet.Name & " but the decision variable cells (" & n & ") could not be interpreted. Please redefine the decision variable cells, and try again."
    End If
End Function

Public Function GetDecisionVariablesWithDefault(Optional book As Workbook, Optional sheet As Worksheet) As Range
    On Error GoTo SetDefault:
    Set GetDecisionVariablesWithDefault = GetDecisionVariables(book, sheet)
    Exit Function
    
SetDefault:
    Set GetDecisionVariablesWithDefault = Nothing
End Function

Public Function GetDecisionVariablesNoOverlap(Optional book As Workbook, Optional sheet As Worksheet) As Range
    Set GetDecisionVariablesNoOverlap = RemoveRangeOverlap(GetDecisionVariables(book, sheet))
End Function

Public Sub SetDecisionVariables(DecisionVariables As Range, Optional book As Workbook, Optional sheet As Worksheet)
    SetSolverNamedRangeIfExists "adj", DecisionVariables
End Sub
