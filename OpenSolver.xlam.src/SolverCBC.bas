Attribute VB_Name = "SolverCBC"
Option Explicit

Sub LaunchCommandLine_CBC()
' Open the CBC solver with our last model loaded.
' If we have a worksheet open with a model, then we pass the solver options (max runtime etc) from this model to CBC. Otherwise, we don't pass any options.
6347      On Error GoTo ErrorHandler

          Dim ModelFilePathName As String
6353      GetLPFilePath ModelFilePathName

          Dim SolverParametersString As String
            
          Dim sheet As Worksheet
          On Error GoTo NoSheet
6349      GetActiveSheetIfMissing sheet
          On Error GoTo ErrorHandler

          Dim Solver As ISolver
          Set Solver = CreateSolver("CBC")
          
          Dim SolverPath As String, errorString As String
6350      If Not SolverIsAvailable(Solver, SolverPath, errorString) Then
6351          Err.Raise OpenSolver_CBCError, Description:=errorString
6352      End If
          
          
          Dim SolveOptions As SolveOptionsType, SolverParameters As New Dictionary
          GetSolveOptions sheet, SolveOptions
          Set SolverParameters = New Dictionary
          PopulateSolverParameters Solver, sheet, SolverParameters, SolveOptions
          SolverParametersString = ParametersToFlags(SolverParameters)
             
NoSheet:
          Dim CBCRunString As String
6364      CBCRunString = " -directory " & MakePathSafe(Left(GetTempFolder, Len(GetTempFolder) - 1)) _
                           & " -import " & MakePathSafe(ModelFilePathName) _
                           & " " & SolverParametersString _
                           & " -" ' Force CBC to accept commands from the command line
6365      RunExternalCommand MakePathSafe(SolverPath) & CBCRunString, "", Normal, False

ExitSub:
          Exit Sub

ErrorHandler:
          ReportError "SolverCBC", "LaunchCommandLine_CBC", True
          GoTo ExitSub
End Sub
