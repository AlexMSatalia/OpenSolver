VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSolverGurobi"
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

Private Const pTitle = "Gurobi (Linear solver)"
Private Const pDesc = "Gurobi is a solver for linear programming (LP), quadratic and quadratically constrained programming (QP and QCP), and mixed-integer programming (MILP, MIQP, and MIQCP). It requires the user to download and install a version of the Gurobi and to have GurobiOSRun.py in the OpenSolver directory."
Private Const pLink = "http://www.gurobi.com/resources/documentation"
Private Const pName = "Gurobi"
Private Const pShortName = "Gurobi"

Private Const pExecName = "gurobi_cl" & ExecExtension
Private Const pScriptFile = "gurobi_tmp" & ScriptExtension
Private Const pPythonScript = "gurobiOSRun.py"
Private Const pMipStartFile = "model.mst"

Private Const pSensitivityAnalysisAvailable = True

Private Const pStatusFile = "status.sol"
Private Const pSensitivityFile = "sensitivityData.sol"
Private Const pFileType = OpenSolver_FileType.LP
Private Const pModelType = OpenSolver_ModelType.Diff

Private Const pIterationLimitName = "IterationLimit"
Private Const pTimeLimitName = "TimeLimit"
Private Const pToleranceName = "MIPGap"

Private Enum GurobiResult
    Optimal = 2
    Infeasible = 3
    InfOrUnbound = 4
    Unbounded = 5
    SolveStoppedIter = 7
    SolveStoppedTime = 9
    SolveStoppedUser = 11
    Unsolved = 12
    SubOptimal = 13
End Enum

Private Const pGurobiError = "Gurobi Error: "


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
1         ISolverLocal_Bitness = Bitness()
End Property

' ISolverLocalExec

Private Property Get ISolverLocalExec_ExecName() As String
1         ISolverLocalExec_ExecName = pExecName
End Property

Private Function ISolverLocalExec_GetExecPath(Optional errorString As String, Optional Bitness As String) As String
1         IsPresent ISolverLocalExec_GetExecPath, errorString
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

Private Function PythonScriptPath() As String
1         PythonScriptPath = JoinPaths(SolverDir, pPythonScript)
End Function

Private Function ScriptFilePath() As String
1         GetTempFilePath pScriptFile, ScriptFilePath
End Function

Function SensitivityFilePath() As String
1         GetTempFilePath pSensitivityFile, SensitivityFilePath
End Function

Function StatusFilePath() As String
1         GetTempFilePath pStatusFile, StatusFilePath
End Function

Function ExecFilePath() As String
1         ExecFilePath = JoinPaths(GetBinDir(), pExecName)
End Function

Private Function PythonPath() As String
          ' So that we have direct control over the running process, we start python itself rather than using the gurobi script file.
    #If Mac Then
              ' We need to use the default system python (pre-installed on mac) and not any other version (e.g. a version from homebrew)
1             PythonPath = JoinPaths(GetRootDriveName(), "usr", "bin", "python")
    #Else
              ' Gurobi supplies a python install on windows in the %GUROBI_HOME% folder
2             PythonPath = JoinPaths(GetGurobiHomeDir(), "python27", "bin", "python" & ExecExtension)
    #End If
End Function

Function MipStartFilePath() As String
1         GetTempFilePath pMipStartFile, MipStartFilePath
End Function

Private Function GetGurobiHomeDir() As String
1         GetGurobiHomeDir = Environ("GUROBI_HOME")
End Function

Private Function GetBinDir() As String
    #If Mac Then
1             GetBinDir = JoinPaths(GetRootDriveName(), "usr", "local", "bin")
    #Else
2             GetBinDir = JoinPaths(GetGurobiHomeDir(), "bin")
    #End If
End Function

Function IsPresent(Optional SolverPath As String, Optional errorString As String) As Boolean
          ' Returns true if Gurobi is available and sets SolverPath as path to gurobi_cl
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
3         If Not FileOrDirExists(PythonScriptPath()) Then
4             errorString = "Unable to find OpenSolver Gurobi script ('" & pPythonScript & "'). Folders searched:" & _
                            vbNewLine & MakePathSafe(SolverDir)
    #If Win32 Then
5             ElseIf Not FileOrDirExists(GetGurobiHomeDir()) Then
6                 errorString = "No Gurobi installation was detected. " & _
                                "The value of %GUROBI_HOME% was " & IIf(Len(GetGurobiHomeDir()) = 0, "not set.", "invalid (" & MakePathSafe(GetGurobiHomeDir()) & ").")
    #End If
7         ElseIf Not FileOrDirExists(ExecFilePath()) Then
8             errorString = "No Gurobi installation was detected. The following file was not found: " & MakePathSafe(ExecFilePath())
9         ElseIf Not FileOrDirExists(PythonPath()) Then
10            errorString = "No Gurobi installation was detected. The following file was not found: " & MakePathSafe(PythonPath())
11        End If

12        If Len(errorString) > 0 Then
13            IsPresent = False
14        Else
15            SolverPath = PythonPath()
16            IsPresent = True
17        End If

ExitFunction:
18        If RaiseError Then RethrowError
19        Exit Function

ErrorHandler:
20        If Not ReportError("CSolverGurobi", "IsPresent") Then Resume
21        RaiseError = True
22        GoTo ExitFunction
End Function

Private Function Version() As String
          ' Get Gurobi version by running 'gurobi_cl -v' at command line
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If Not SolverIsPresent(Me) Then
4             Version = vbNullString
5             GoTo ExitFunction
6         End If
          
          Dim SolverPath As String
7         SolverPath = ExecFilePath()
          
          Dim result As String
8         result = ExecCapture(MakePathSafe(SolverPath) & " -v")
9         Version = Mid(result, 26, 5)

ExitFunction:
10        Close #1
11        If RaiseError Then RethrowError
12        Exit Function

ErrorHandler:
13        If Not ReportError("CSolverGurobi", "Version") Then Resume
14        RaiseError = True
15        GoTo ExitFunction
End Function

Private Function Bitness() As String
          ' Get Gurobi bitness by running 'gurobi_cl -v' at command line
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If Not SolverIsPresent(Me) Then
4             Bitness = vbNullString
5             GoTo ExitFunction
6         End If
          
          Dim SolverPath As String
7         SolverPath = ExecFilePath()

          Dim result As String
8         result = ExecCapture(MakePathSafe(SolverPath) & " -v")
9         Bitness = IIf(InStr(result, "64)") > 0, "64", "32")

ExitFunction:
10        If RaiseError Then RethrowError
11        Exit Function

ErrorHandler:
12        If Not ReportError("CSolverGurobi", "Bitness") Then Resume
13        RaiseError = True
14        GoTo ExitFunction
End Function

Private Sub CleanFiles()
1         DeleteFileAndVerify SensitivityFilePath()
2         DeleteFileAndVerify ScriptFilePath()
3         DeleteFileAndVerify StatusFilePath()
End Sub

Private Function CreateSolveCommand(s As COpenSolver) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
3         DeleteFileAndVerify MipStartFilePath()
4         If s.InitialSolutionIsValid Then
5             CreateMipStartFile s
6         End If

          ' We use the -u flag on python to force it to unbuffer its output
7         CreateSolveCommand = MakePathSafe(PythonPath()) & " -u " & _
                               MakePathSafe(PythonScriptPath()) & " " & _
                               MakePathSafe(GetModelFilePath(Me)) & " " & _
                               MakePathSafe(MipStartFilePath()) & " " & _
                               MakePathSafe(StatusFilePath()) & " " & _
                               MakePathSafe(s.SolutionFilePathName) & " " & _
                               MakePathSafe(SensitivityFilePath()) & " " & _
                               ParametersToKwargs(s.SolverParameters)
                               
8         CreateScriptFile ScriptFilePath, CreateSolveCommand

ExitFunction:
9         If RaiseError Then RethrowError
10        Exit Function

ErrorHandler:
11        If Not ReportError("CSolverGurobi", "CreateSolveScript") Then Resume
12        RaiseError = True
13        GoTo ExitFunction
End Function

Private Sub CreateMipStartFile(s As COpenSolver)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim FileNum As Long
3         FileNum = FreeFile()
4         Open MipStartFilePath() For Output As FileNum
5         Print #FileNum, "# MIP start for problem on sheet '"; s.sheet.Name; "'"
          
          Dim i As Long
6         For i = 1 To s.NumVars
7             Print #FileNum, GetLPNameFromVarName(s.VarName(i)); " "; StrEx(s.VarInitialValue(i))
8         Next i

ExitSub:
9         Close #FileNum
10        If RaiseError Then RethrowError
11        Exit Sub

ErrorHandler:
12        If Not ReportError("CSolverGurobi", "CreateMipStartFile") Then Resume
13        RaiseError = True
14        GoTo ExitSub
End Sub

Sub ReadResults(s As COpenSolver)
          
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If Not FileOrDirExists(StatusFilePath()) Then
4             RaiseGeneralError "The solver did not create a status file", _
                                LINK_NO_SOLUTION_FILE
5         End If
          
6         s.SolutionWasLoaded = True
          
7         Open StatusFilePath() For Input As #1 ' supply path with filename
          Dim Line As String, Index As Long
8         Line Input #1, Line
9         Close #1
          ' Check for python exception while running Gurobi
10        If Left(Line, Len(pGurobiError)) = pGurobiError Then
11            RaiseGeneralError Line
12        End If
          'Get the returned status code from gurobi.
          'List of return codes can be seen at - http://www.gurobi.com/documentation/7.0/refman/optimization_status_codes.html
13        If Line = GurobiResult.Optimal Then
14            s.SolveStatus = OpenSolverResult.Optimal
15            s.SolveStatusString = "Optimal"
16        ElseIf Line = GurobiResult.Infeasible Then
17            s.SolveStatus = OpenSolverResult.Infeasible
18            s.SolveStatusString = "No Feasible Solution"
19            s.SolutionWasLoaded = False
20        ElseIf Line = GurobiResult.InfOrUnbound Then
21            s.SolveStatus = OpenSolverResult.Unbounded
22            s.SolveStatusString = "No Solution Found (Infeasible or Unbounded)"
23            s.SolutionWasLoaded = False
24        ElseIf Line = GurobiResult.Unbounded Then
25            s.SolveStatus = OpenSolverResult.Unbounded
26            s.SolveStatusString = "No Solution Found (Unbounded)"
27            s.SolutionWasLoaded = False
28        ElseIf Line = GurobiResult.SolveStoppedTime Then
29            s.SolveStatus = OpenSolverResult.LimitedSubOptimal
30            s.SolveStatusString = "Stopped on Time Limit"
31        ElseIf Line = GurobiResult.SolveStoppedIter Then
32            s.SolveStatus = OpenSolverResult.LimitedSubOptimal
33            s.SolveStatusString = "Stopped on Iteration Limit"
34        ElseIf Line = GurobiResult.SolveStoppedUser Then
35            s.SolveStatus = OpenSolverResult.LimitedSubOptimal
36            s.SolveStatusString = "Stopped on Ctrl-C"
37        ElseIf Line = GurobiResult.Unsolved Then
38            s.SolveStatus = OpenSolverResult.LimitedSubOptimal
39            s.SolveStatusString = "Stopped on Gurobi Numerical difficulties"
40        ElseIf Line = GurobiResult.SubOptimal Then
41            s.SolveStatus = OpenSolverResult.LimitedSubOptimal
42            s.SolveStatusString = "Unable to satisfy optimality tolerances; a sub-optimal solution is available."
43        Else
44            RaiseGeneralError "The response from the Gurobi solver is not recognised. The response was: " & Line
45        End If
          
46        If s.SolutionWasLoaded Then

47            If Not FileOrDirExists(s.SolutionFilePathName) Then
48                RaiseGeneralError "The solver did not create a solution file. No new solution is available.", _
                                    LINK_NO_SOLUTION_FILE
49            End If
50            Open s.SolutionFilePathName() For Input As #1

51            UpdateStatusBar "OpenSolver: Loading Solution... " & s.SolveStatusString, True
              
              ' Loop over all comment lines
52            Do While Not EOF(1)
53                Line Input #1, Line
54                If Len(Line) = 0 Or Left(Line, 1) <> "#" Then
55                    Exit Do
56                End If
57            Loop
              
              Dim NumVar As Long, SplitLine() As String, i As Long
58            i = 1
59            Do While True
60                If Len(Line) > 0 Then
61                    SplitLine = SplitWithoutRepeats(Line, " ")
                      Dim ConFinalValue As String
62                    ConFinalValue = SplitLine(1)
                      
                      ' Check for an exponent that is too large, rougly >300
                      ' Only do it for -ve exponents, since e-30 ~= e-300
63                    Index = InStr(LCase(ConFinalValue), "e-")
64                    If Index > 0 Then
                          ' Trim the final digit if the exponent has 3 digits
65                        If Len(ConFinalValue) - Index - 1 > 2 Then
66                            ConFinalValue = Left(ConFinalValue, Len(ConFinalValue) - 1)
67                        End If
68                    End If
                      
69                    s.VarFinalValue(i) = Val(ConFinalValue)
70                    s.VarCellName(i) = GetVarNameFromLPName(SplitLine(0))
          
                      ' Save number of vars read
71                    NumVar = i
72                    i = i + 1
73                End If

74                If EOF(1) Then
75                    Exit Do
76                Else
77                    Line Input #1, Line
78                End If
79            Loop
              
80            If s.bGetDuals Then
                  Dim CostVector() As Double
81                CostVector = s.CostCoeffs.AsVector(s.NumVars)
              
82                Open SensitivityFilePath() For Input As 2
83                For i = 1 To NumVar
84                    Line Input #2, Line
85                    SplitLine = SplitWithoutRepeats(Line, ",")
86                    s.VarReducedCost(i) = Val(SplitLine(0))
87                    s.VarDecrease(i) = CostVector(i) - Val(SplitLine(1))
88                    s.VarIncrease(i) = Val(SplitLine(2)) - CostVector(i)
89                Next i
                  
                  ' Throw away first constraint if it was from a seek objective model
90                If s.ObjectiveSense = TargetObjective Then
91                    Line Input #2, Line
92                End If
                  
                  Dim row As Long
93                For row = 1 To s.NumRows
                      ' Skip any rows that weren't written
94                    If s.SparseA(row).Count > 0 Then
95                        Line Input #2, Line
96                        SplitLine = SplitWithoutRepeats(Line, ",")
97                        s.ConShadowPrice(row) = Val(SplitLine(0))
                          Dim RHSValue As Double
98                        RHSValue = Val(SplitLine(1))
99                        s.ConIncrease(row) = Val(SplitLine(4)) - RHSValue
100                       s.ConDecrease(row) = RHSValue - Val(SplitLine(3))
101                       s.ConFinalValue(row) = RHSValue - Val(SplitLine(2))
102                   Else
103                       s.ConFinalValue(row) = s.RHS(row)
104                       s.ConShadowPrice(row) = 0
105                       s.ConIncrease(row) = 0
106                       s.ConDecrease(row) = 0
107                   End If
108               Next row
109           End If
110       End If

ExitSub:
111       Application.StatusBar = False
112       Close #1
113       Close #2
114       If RaiseError Then RethrowError
115       Exit Sub

ErrorHandler:
116       If Not ReportError("CSolverGurobi", "ReadResults") Then Resume
117       RaiseError = True
118       GoTo ExitSub
End Sub

Private Sub CheckLog(s As COpenSolver)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If Not FileOrDirExists(s.LogFilePathName) Then
4             RaiseGeneralError "The solver did not create a log file. No new solution is available.", _
                                LINK_NO_SOLUTION_FILE
5         End If
          
          Dim message As String
6         Open s.LogFilePathName For Input As #1
7             message = LCase(Input$(LOF(1), 1))
8         Close #1

          Dim Key As Variant
9         For Each Key In s.SolverParameters.Keys
10            If InStr(message, LCase("No parameters matching '" & Key & "' found")) > 0 Then
11                RaiseUserError _
                      "The parameter '" & Key & "' was not recognised by Gurobi. " & _
                      "Please check the parameter name you have specified, or consult the Gurobi documentation for more information.", _
                      LINK_PARAMETER_DOCS
12            End If
13        Next Key

ExitSub:
14        Close #1
15        If RaiseError Then RethrowError
16        Exit Sub

ErrorHandler:
17        If Not ReportError("CSolverGurobi", "CheckLog") Then Resume
18        RaiseError = True
19        GoTo ExitSub
End Sub
