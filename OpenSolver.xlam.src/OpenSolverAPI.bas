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
    GetActiveBookAndSheetIfMissing book, sheet
    
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
    GetActiveBookAndSheetIfMissing book, sheet
    GetDualsNewSheet = GetNamedBooleanWithDefault("OpenSolver_DualsNewSheet", book, sheet, False)
End Function

Public Sub SetDualsNewSheet(DualsNewSheet As Boolean, Optional book As Workbook, Optional sheet As Worksheet)
    GetActiveBookAndSheetIfMissing book, sheet
    SetBooleanNameOnSheet "OpenSolver_DualsNewSheet", DualsNewSheet, book, sheet
End Sub

Public Function GetUpdateSensitivity(Optional book As Workbook, Optional sheet As Worksheet) As Boolean
    GetActiveBookAndSheetIfMissing book, sheet
    GetUpdateSensitivity = GetNamedBooleanWithDefault("OpenSolver_UpdateSensitivity", book, sheet, True)
End Function

Public Sub SetUpdateSensitivity(UpdateSensitivity As Boolean, Optional book As Workbook, Optional sheet As Worksheet)
    GetActiveBookAndSheetIfMissing book, sheet
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
    GetActiveBookAndSheetIfMissing book, sheet
    If LinearityCheck Then
        DeleteNameOnSheet "OpenSolver_LinearityCheck", book, sheet
    Else
        SetNameOnSheet "OpenSolver_LinearityCheck", "=2", book, sheet
    End If
End Sub

Public Function GetDuals(Optional book As Workbook, Optional sheet As Worksheet) As Range
    GetActiveBookAndSheetIfMissing book, sheet
    If Not GetNamedRangeIfExistsOnSheet(sheet, "OpenSolver_Duals", GetDuals) Then Set GetDuals = Nothing
End Function

Public Sub SetDuals(Duals As Range, Optional book As Workbook, Optional sheet As Worksheet)
    GetActiveBookAndSheetIfMissing book, sheet
    SetNamedRangeIfExists "OpenSolver_Duals", Duals, book, sheet
End Sub
