Attribute VB_Name = "SolverCBC"
Option Explicit

Sub LaunchCommandLine_CBC()
' Open the CBC solver with our last model loaded.
' If we have a worksheet open with a model, then we pass the solver options (max runtime etc) from this model to CBC. Otherwise, we don't pass any options.
1         On Error GoTo ErrorHandler

          Dim ModelFilePathName As String
2         GetLPFilePath ModelFilePathName
3         If Not FileOrDirExists(ModelFilePathName) Then
4             RaiseUserError "There is no .lp file (" & ModelFilePathName & ") to open. Please solve the OpenSolver model and then try again."
5         End If

          Dim SolverParametersString As String
            
          Dim sheet As Worksheet
6         On Error GoTo NoSheet
7         GetActiveSheetIfMissing sheet
8         On Error GoTo ErrorHandler

          Dim Solver As ISolver
9         Set Solver = CreateSolver("CBC")
          
          Dim SolverPath As String, errorString As String
10        If Not SolverIsAvailable(Solver, SolverPath, errorString) Then
11            RaiseGeneralError errorString
12        End If
          
          Dim SolverParameters As New Dictionary
13        Set SolverParameters = GetSolverParametersDict(Solver, sheet)
14        SolverParametersString = ParametersToFlags(SolverParameters)
             
NoSheet:
          Dim CBCRunString As String
15        CBCRunString = " -directory " & MakePathSafe(Left(GetTempFolder, Len(GetTempFolder) - 1)) _
                           & " -import " & MakePathSafe(ModelFilePathName) _
                           & " " & SolverParametersString _
                           & " -" ' Force CBC to accept commands from the command line
16        ExecAsync MakePathSafe(SolverPath) & CBCRunString, GetTempFolder(), True

ExitSub:
17        Exit Sub

ErrorHandler:
18        ReportError "SolverCBC", "LaunchCommandLine_CBC", True
19        GoTo ExitSub
End Sub
