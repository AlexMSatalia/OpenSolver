VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSolverCbc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISolver
Implements ISolverLocal
Implements ISolverLocalExec
Implements ISolverLinear
Implements ISolverFile

Private Const pTitle = "COIN-OR CBC (Linear solver)"
Private Const pDesc = "The COIN Branch and Cut solver (CBC) is the default solver for OpenSolver and is an open-source mixed-integer program (MIP) solver written in C++. CBC is an active open-source project led by John Forrest at www.coin-or.org."
Private Const pLink = "http://www.coin-or.org/Cbc/cbcuserguide.html"
Private Const pName = "CBC"
Private Const pShortName = "CBC"

Private Const pExecName = "cbc" & ExecExtension
Private Const pScriptFile = "cbc" & ScriptExtension

Private Const pSensitivityAnalysisAvailable = True

Private Const pCostRangesFile = "costranges.txt"
Private Const pRHSRangesFile = "rhsranges.txt"
Private Const pFileType = OpenSolver_FileType.LP
Private Const pModelType = OpenSolver_ModelType.Diff

Private Const pIterationLimitName = "maxIterations"
Private Const pTimeLimitName = "seconds"
Private Const pToleranceName = "ratioGap"

' ISolver

Private Property Get ISolver_Name() As String
1         ISolver_Name = pName
End Property

Private Property Get ISolver_ShortName() As String
1         ISolver_ShortName = pShortName
End Property

Private Property Get ISolver_Title() As String
1         ISolver_Title = pTitle
End Property

Private Property Get ISolver_Desc() As String
1         ISolver_Desc = pDesc
End Property

Private Property Get ISolver_Link() As String
1         ISolver_Link = pLink
End Property

Private Property Get ISolver_IterationLimitName() As String
1         ISolver_IterationLimitName = pIterationLimitName
End Property

Private Property Get ISolver_PrecisionName() As String
End Property

Private Property Get ISolver_TimeLimitName() As String
1         ISolver_TimeLimitName = pTimeLimitName
End Property

Private Property Get ISolver_ToleranceName() As String
1         ISolver_ToleranceName = pToleranceName
End Property

Private Property Get ISolver_ModelType() As OpenSolver_ModelType
1         ISolver_ModelType = pModelType
End Property

Private Sub ISolver_CleanFiles()
1         CleanFiles
End Sub

Private Sub ISolver_CheckLog(s As COpenSolver)
1         CheckLog s
End Sub

' ISolverLocal

Private Property Get ISolverLocal_Version() As String
1         ISolverLocal_Version = Version()
End Property

Private Property Get ISolverLocal_Bitness() As String
1         SolverIsPresent Me, Bitness:=ISolverLocal_Bitness
End Property

' ISolverLocalExec

Private Property Get ISolverLocalExec_ExecName() As String
1         ISolverLocalExec_ExecName = pExecName
End Property

Private Function ISolverLocalExec_GetExecPath(Optional errorString As String, Optional Bitness As String) As String
1         ISolverLocalExec_GetExecPath = SolverExecPath(Me, errorString, Bitness)
End Function

Private Function ISolverLocalExec_CreateSolveCommand(s As COpenSolver) As String
1         ISolverLocalExec_CreateSolveCommand = CreateSolveCommand(s)
End Function

' ISolverLinear

Private Property Get ISolverLinear_SensitivityAnalysisAvailable() As Boolean
1         ISolverLinear_SensitivityAnalysisAvailable = pSensitivityAnalysisAvailable
End Property

' ISolverFile

Private Sub ISolverFile_ReadResults(s As COpenSolver, Optional solution As String)
1         ReadResults s
End Sub

Private Property Get ISolverFile_FileType() As OpenSolver_FileType
1         ISolverFile_FileType = pFileType
End Property

' Methods

Private Function Version() As String
      ' Get CBC version by running 'cbc -exit' at command line
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim SolverPath As String
3         If Not SolverIsPresent(Me, SolverPath) Then
4             Version = vbNullString
5             GoTo ExitFunction
6         End If
          
          Dim result As String
7         result = ExecCapture(MakePathSafe(SolverPath) & " -exit")
8         Version = Mid(result, InStr(result, "Version:") + 9, 5)

ExitFunction:
9         If RaiseError Then RethrowError
10        Exit Function

ErrorHandler:
11        If Not ReportError("CSolverCbc", "Version") Then Resume
12        RaiseError = True
13        GoTo ExitFunction
End Function

Private Sub CleanFiles()
1         DeleteFileAndVerify CostRangesFilePath()
2         DeleteFileAndVerify RHSRangesFilePath()
3         DeleteFileAndVerify ScriptFilePath()
End Sub

Function ScriptFilePath() As String
1         GetTempFilePath pScriptFile, ScriptFilePath
End Function

Function CostRangesFilePath() As String
1         GetTempFilePath pCostRangesFile, CostRangesFilePath
End Function

Function RHSRangesFilePath() As String
1         GetTempFilePath pRHSRangesFile, RHSRangesFilePath
End Function

Private Function CreateSolveCommand(s As COpenSolver) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
3         CreateSolveCommand = _
              MakePathSafe(ISolverLocalExec_GetExecPath()) & _
              " -directory " & MakePathSafe(Left(GetTempFolder, Len(GetTempFolder) - 1)) & _
              " -import " & MakePathSafe(GetModelFilePath(Me)) & _
              " " & ParametersToFlags(s.SolverParameters) & _
              " -solve" & _
              IIf(s.bGetDuals, " -printingOptions all", vbNullString) & _
              " -solution " & MakePathSafe(s.SolutionFilePathName) & _
              IIf(s.bGetDuals, " -printingOptions rhs  -solution " & pRHSRangesFile & _
                               " -printingOptions objective -solution " & pCostRangesFile, vbNullString)
              
4         CreateScriptFile ScriptFilePath, CreateSolveCommand

ExitFunction:
5         If RaiseError Then RethrowError
6         Exit Function

ErrorHandler:
7         If Not ReportError("CSolverCbc", "CreateSolveScript") Then Resume
8         RaiseError = True
9         GoTo ExitFunction
End Function

Private Sub ReadResults(s As COpenSolver)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
3         If Not FileOrDirExists(s.SolutionFilePathName) Then
4             RaiseGeneralError "The solver did not create a solution file. No new solution is available.", _
                                "http://opensolver.org/help/#cbcfails"
5         End If

6         Open s.SolutionFilePathName For Input As #1
          Dim Response As String
7         Line Input #1, Response  ' Optimal - objective value              22

8         s.SolutionWasLoaded = True
9         If Response Like "Optimal*" Then
10            s.SolveStatus = OpenSolverResult.Optimal
11            s.SolveStatusString = "Optimal"
              '
12        ElseIf Response Like "Infeasible*" Then
13            s.SolveStatus = OpenSolverResult.Infeasible
14            s.SolveStatusString = "No Feasible Solution"
              '
15        ElseIf Response Like "Integer infeasible*" Then
16            s.SolveStatus = OpenSolverResult.Infeasible
17            s.SolveStatusString = "No Feasible Integer Solution"
              '
18        ElseIf Response Like "Unbounded*" Then
19            s.SolveStatus = OpenSolverResult.Unbounded
20            s.SolveStatusString = "No Solution Found (Unbounded)"
21            s.SolutionWasLoaded = False
              '
22        ElseIf Response Like "Stopped on time *" Then ' Stopped on iterations or time
23            s.SolveStatus = OpenSolverResult.LimitedSubOptimal
24            s.SolveStatusString = "Stopped on Time Limit"
25            If Response Like "*(no integer solution - continuous used)*" Then
26                s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
27            End If
              '
28        ElseIf Response Like "Stopped on iterations*" Then ' Stopped on iterations or time
29            s.SolveStatus = OpenSolverResult.LimitedSubOptimal
30            s.SolveStatusString = "Stopped on Iteration Limit"
31            If Response Like "*(no integer solution - continuous used)*" Then
32                s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
33            End If
              '
34        ElseIf Response Like "Stopped on difficulties*" Then ' Stopped on iterations or time
35            s.SolveStatus = OpenSolverResult.LimitedSubOptimal
36            s.SolveStatusString = "Stopped on CBC difficulties"
37            If Response Like "*(no integer solution - continuous used)*" Then
38                s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
39            End If
              '
40        ElseIf Response Like "Stopped on ctrl-c*" Then ' Stopped on iterations or time
41            s.SolveStatus = OpenSolverResult.LimitedSubOptimal
42            s.SolveStatusString = "Stopped on Ctrl-C"
43            If Response Like "*(no integer solution - continuous used)*" Then
44                s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
45            End If
              '
46        ElseIf Response Like "Status unknown*" Then
47            RaiseGeneralError "The CBC solver did not solve the problem, suggesting there was an error " & _
                                "in the CBC input parameters. The response was: " & vbNewLine & _
                                Response & vbNewLine & _
                                "The CBC command line can be found at:" & vbNewLine & _
                                ScriptFilePath()
48        Else
49            RaiseGeneralError "The response from the CBC solver is not recognised. The response was: " & Response
50        End If
          
          ' Remove the double spaces from Response
51        Response = Replace(Response, "    ", " ")
52        Response = Replace(Response, "   ", " ")
53        Response = Replace(Response, "  ", " ")

54        If s.SolutionWasLoaded Then
              ' We read in whatever solution CBC returned
55            UpdateStatusBar "OpenSolver: Loading Solution... " & Response
              
              Dim m As CModelDiff
56            Set m = s.DiffModel
          
              Dim Line As String, SplitLine() As String, Index As Double, NameValue As String, value As Double, CBCConstraintIndex As Long, StartOffset As Long
57            If s.bGetDuals Then
                  ' Read in the Solution File
                  ' Line format: Index ConstraintName Value ConShadowPrice
                  
                  Dim j As Long, row As Long
58                CBCConstraintIndex = 0
                  
                  ' Throw away first constraint if it was from a seek objective model
59                If s.ObjectiveSense = TargetObjective Then
60                    Line Input #1, Line
61                    CBCConstraintIndex = CBCConstraintIndex + 1
62                End If

63                For row = 1 To s.NumRows
                      ' Only process rows that were in the model
64                    If s.SparseA(row).Count > 0 Then
65                        Line Input #1, Line
66                        SplitLine = SplitWithoutRepeats(Line, " ")

                          ' In the case of LpStatusInfeasible, we can get lines that start **. We strip the **
67                        StartOffset = 0
68                        If SplitLine(StartOffset) = "**" Then StartOffset = 1

                          ' Check the index of the row
69                        If CLng(Val(SplitLine(StartOffset))) <> CBCConstraintIndex Then
70                            RaiseGeneralError "While reading the CBC solution file, OpenSolver found an unexpected constraint row."
71                        End If

72                        NameValue = SplitLine(StartOffset + 1)
                          
73                        s.ConFinalValue(row) = Val(SplitLine(StartOffset + 2))
74                        value = Val(SplitLine(StartOffset + 3))
75                        If s.ObjectiveSense = MaximiseObjective Then value = -value
76                        s.ConShadowPrice(row) = value
77                        CBCConstraintIndex = CBCConstraintIndex + 1
78                    Else
79                        s.ConFinalValue(row) = s.RHS(row)
80                        s.ConShadowPrice(row) = 0
81                    End If
82                Next row
83                ReadSensitivityData s
84            End If
            
              ' Now we read in the decision variable values
              ' Line format: Index VariableName Value ReducedCost
85            j = 1
86            While Not EOF(1)
87                Line Input #1, Line
88                SplitLine = SplitWithoutRepeats(Line, " ")

                  ' In the case of LpStatusInfeasible, we can get lines that start **. We strip the **
89                StartOffset = 0
90                If SplitLine(StartOffset) = "**" Then StartOffset = 1
                  
91                Index = CLng(Val(SplitLine(StartOffset)))
92                s.VarCellName(j) = GetVarNameFromLPName(SplitLine(StartOffset + 1))
93                s.VarFinalValue(j) = Val(SplitLine(StartOffset + 2))
                 
94                If s.bGetDuals Then
95                    value = Val(SplitLine(StartOffset + 3))
96                    If s.ObjectiveSense = MaximiseObjective Then value = -value
97                    s.VarReducedCost(j) = value
98                End If
99                j = j + 1
100           Wend
101           s.SolutionWasLoaded = True

102       End If

ExitSub:
103       Close #1
104       If RaiseError Then RethrowError
105       Exit Sub

ErrorHandler:
106       If Not ReportError("CSolverCBC", "ReadResults") Then Resume
107       RaiseError = True
108       GoTo ExitSub
End Sub

Sub ReadSensitivityData(s As COpenSolver)
      'Reads the two files with the limits on the bounds of shadow prices and reduced costs
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim LineData() As String
          Dim Line As String, row As Long, j As Long
          
          'Find the ranges on the constraints

3         Open RHSRangesFilePath For Input As #2 ' supply path with filename
4         Line Input #2, Line 'Dont want first line
5         row = 1
          
          ' Throw away first constraint if it was from a seek objective model
6         If s.ObjectiveSense = TargetObjective Then
7             Line Input #2, Line
8         End If
          
9         While Not EOF(2)
10            If s.SparseA(row).Count > 0 Then
11                Line Input #2, Line
12                LineData() = Split(Line, ",")
13                s.ConIncrease(row) = Val(LineData(2))
14                s.ConDecrease(row) = Val(LineData(4))
15            Else
16                s.ConIncrease(row) = 0
17                s.ConDecrease(row) = 0
18            End If
19            row = row + 1
20        Wend
21        Close 2
          
22        j = 1
          'Find the ranges on the variables
23        Open CostRangesFilePath For Input As #2 ' supply path with filename
24        Line Input #2, Line 'Dont want first line
25        row = s.NumRows + 2
26        While Not EOF(2)
27            Line Input #2, Line
28            LineData() = Split(Line, ",")
29            If s.ObjectiveSense = MaximiseObjective Then
30                s.VarIncrease(j) = Val(LineData(4))
31                s.VarDecrease(j) = Val(LineData(2))
32            Else
33                s.VarIncrease(j) = Val(LineData(2))
34                s.VarDecrease(j) = Val(LineData(4))
35            End If
36            j = j + 1
37        Wend

ExitSub:
38        Close #2
39        If RaiseError Then RethrowError
40        Exit Sub

ErrorHandler:
41        If Not ReportError("CSolverCbc", "ReadSensitivityData") Then Resume
42        RaiseError = True
43        GoTo ExitSub
End Sub

Private Sub CheckLog(s As COpenSolver)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
3         If Not FileOrDirExists(s.LogFilePathName) Then
4             RaiseGeneralError "The solver did not create a log file. No new solution is available.", _
                                "http://opensolver.org/help/#cbcfails"
5         End If
          
          Dim message As String
6         Open s.LogFilePathName For Input As #1
7             message = LCase(Input$(LOF(1), 1))
8         Close #1
          
          Dim Key As Variant
9         For Each Key In s.SolverParameters.Keys
10            If InStr(message, LCase("parameter " & Key & " value remains")) > 0 Then
11                RaiseUserError _
                      "The value for the parameter '" & Key & "' supplied to CBC was invalid. " & _
                      "Please check the value you have specified, or consult the CBC documentation for more information.", _
                      "http://opensolver.org/using-opensolver/#extra-parameters"
12            End If
13            If InStr(message, LCase("No match for " & Key)) > 0 Then
14                RaiseUserError _
                      "The parameter '" & Key & "' was not recognised by CBC. " & _
                      "Please check the parameter name you have specified, or consult the CBC documentation for more information.", _
                      "http://opensolver.org/using-opensolver/#extra-parameters"
15            End If
16        Next Key

ExitSub:
17        If RaiseError Then RethrowError
18        Exit Sub

ErrorHandler:
19        If Not ReportError("CSolverCbc", "CheckLog") Then Resume
20        RaiseError = True
21        GoTo ExitSub
End Sub


