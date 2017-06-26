Attribute VB_Name = "SolverCommon"
Option Explicit

Public Const LogFileName = "log1.tmp"
Public Const SolutionFileName = "model.sol"

Public Const SolverDirName As String = "Solvers"
Public Const SolverDirMac As String = "osx"
Public Const SolverDirWin32 As String = "win32"
Public Const SolverDirWin64 As String = "win64"

Public LastUsedSolver As String

Public Function SolverDir() As String
          Dim SolverDirBase As String
    #If Mac And MAC_OFFICE_VERSION >= 15 Then
              ' On Mac 2016, we need to access the solvers from a folder that has execute permissions in the sandbox
1             SolverDirBase = "/Library/OpenSolver"
2             If Not FileOrDirExists(SolverDirBase) Then
3                 RaiseUserError "Unable to find the solvers in `" & SolverDirBase & "`. Make sure you have run the `OpenSolver Solvers.pkg` installer in the `Solvers/osx` folder where you unzipped OpenSolver.", "http://opensolver.org/installing-opensolver/"
4             End If
    #Else
5             SolverDirBase = ThisWorkbook.Path
    #End If
6         SolverDir = JoinPaths(SolverDirBase, SolverDirName)
End Function

Sub SolveModel(s As COpenSolver, ShouldSolveRelaxation As Boolean, ShouldMinimiseUserInteraction As Boolean)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
3         Application.EnableCancelKey = xlErrorHandler
4         Application.Cursor = xlWait

5         s.SolveStatus = OpenSolverResult.Unsolved
6         s.SolveStatusString = "Unsolved"
7         s.SolveStatusComment = vbNullString
8         s.SolutionWasLoaded = False
          
          ' Track whether to show messages
9         s.MinimiseUserInteraction = ShouldMinimiseUserInteraction
10        s.SolveRelaxation = ShouldSolveRelaxation

          Dim oldCalculationMode As Long, oldIterationMode As Boolean
11        oldIterationMode = Application.Iteration
12        oldCalculationMode = Application.Calculation
13        Application.Calculation = xlCalculationManual
          
          Dim ScreenStatus As Boolean
14        ScreenStatus = Application.ScreenUpdating
15        Application.ScreenUpdating = False
          
16        If s.ModelStatus <> Built Then
17            RaiseGeneralError "The model cannot be solved as it has not yet been built."
18        End If

          'Check that solver is available
          Dim errorString As String
19        If Not SolverIsAvailable(s.Solver, errorString:=errorString) Then
20            RaiseGeneralError errorString
21        End If

          Dim LogFilePathName As String
22        If GetLogFilePath(LogFilePathName) Then DeleteFileAndVerify LogFilePathName
23        s.LogFilePathName = LogFilePathName
          
          ' Clean Solver specific files
24        s.Solver.CleanFiles
          
          'Check that we can write to all cells
25        TestCellsForWriting s.AdjustableCells

26        If TypeOf s.Solver Is ISolverLocalLib Then
              Dim LocalLibSolver As ISolverLocalLib
27            Set LocalLibSolver = s.Solver
28            LocalLibSolver.Solve s
29        ElseIf TypeOf s.Solver Is ISolverFile Then
              Dim LinearSolver As ISolverLinear
              Dim FileSolver As ISolverFile
30            Set FileSolver = s.Solver
              
              'Delete any existing solution file
              Dim SolutionFilePathName As String
31            If GetSolutionFilePath(SolutionFilePathName) Then DeleteFileAndVerify SolutionFilePathName
32            s.SolutionFilePathName = SolutionFilePathName
              
              ' Check if we need to request duals from the solver
33            If TypeOf s.Solver Is ISolverLinear Then
34                Set LinearSolver = s.Solver
35                Set s.rConstraintList = GetDuals(s.sheet)
36                s.DualsOnSameSheet = (Not s.rConstraintList Is Nothing)
37                s.DualsOnNewSheet = GetDualsOnSheet(s.sheet)
38                s.bGetDuals = ((s.NumDiscreteVars = 0) Or s.SolveRelaxation) And _
                                (s.DualsOnNewSheet Or s.DualsOnSameSheet) And LinearSolver.SensitivityAnalysisAvailable
39            End If
              
              ' Set up arrays to hold solution values (to avoid dynamically resizing them later)
40            s.PrepareForSolution
              
              ' Write file
              Dim SolverCommand As String
41            UpdateStatusBar "OpenSolver: Writing Model to disk... " & s.NumVars & " vars, " & s.NumRows & " rows.", True
42            SolverCommand = WriteModelFile(s)
              
              ' Check if anything detected while writing file
43            If s.SolveStatus <> OpenSolverResult.Unsolved Then GoTo ExitSub
              
              ' Run solver and read results
              Dim solution As String
44            solution = RunSolver(s, SolverCommand)
45            FileSolver.ReadResults s, solution
              
46            s.LoadResultsToSheet
              
47            If TypeOf s.Solver Is ISolverLinear Then
                  ' Get sensitivity results
48                If s.bGetDuals And s.SolveStatus = OpenSolverResult.Optimal Then
                      'write the duals on the same sheet if the user has picked this option
49                    If s.DualsOnSameSheet Then WriteConstraintListToSheet s.rConstraintList, s
                      'If the user wants a new sheet with the sensitivity data then call the functions that write this
50                    If s.DualsOnNewSheet Then
                          Dim newSheet As Worksheet, currentSheet As Worksheet
                          ' Save sheet selection
51                        GetActiveSheetIfMissing currentSheet
                          
52                        Set newSheet = MakeNewSheet(s.sheet.Name & " Sensitivity", GetUpdateSensitivity(s.sheet))
                          
                          ' Restore old sheet selection
53                        currentSheet.Select
54                        WriteConstraintSensitivityTable newSheet, s
55                    End If
56                ElseIf Not s.bGetDuals And (s.DualsOnNewSheet Or s.DualsOnSameSheet) And LinearSolver.SensitivityAnalysisAvailable Then
57                    RaiseUserError _
                          "Could not get sensitivity analysis due to binary and/or integer constraints." & vbNewLine & vbNewLine & _
                          "Turn off sensitivity in the model dialogue or reformulate your model without these constraints." & vbNewLine & vbNewLine & _
                          "The " & s.Solver.ShortName & " solution has been returned to the sheet." & vbNewLine
58                End If
59            End If
60        End If

ExitSub:
61        Application.Cursor = xlDefault
62        Application.StatusBar = False ' Resume normal status bar behaviour
63        Application.ScreenUpdating = ScreenStatus
64        Application.Calculation = oldCalculationMode
65        Application.Iteration = oldIterationMode
66        Application.Calculate
67        If RaiseError Then RethrowError
68        Exit Sub

ErrorHandler:
69        If Not ReportError("SolverCommon", "SolveModel") Then Resume
70        RaiseError = True
71        GoTo ExitSub

End Sub

Function CreateSolver(SolverShortName As String) As ISolver
1         Select Case LCase(SolverShortName)
          Case "cbc":         Set CreateSolver = New CSolverCbc
2         Case "gurobi":      Set CreateSolver = New CSolverGurobi
3         Case "neoscbc":     Set CreateSolver = New CSolverNeosCbc
4         Case "bonmin":      Set CreateSolver = New CSolverBonmin
5         Case "couenne":     Set CreateSolver = New CSolverCouenne
6         Case "nomad":       Set CreateSolver = New CSolverNomad
7         Case "neosbon":     Set CreateSolver = New CSolverNeosBon
8         Case "neoscou":     Set CreateSolver = New CSolverNeosCou
          Case "solveengine": Set CreateSolver = New CSolverSolveEngine
9         Case Else: RaiseGeneralError "The specified solver ('" & SolverShortName & "') was not recognised."
10        End Select
End Function

Function GetLogFilePath(ByRef Path As String) As Boolean
1         GetLogFilePath = GetTempFilePath(LogFileName, Path)
End Function

Function GetSolutionFilePath(ByRef Path As String) As Boolean
1         GetSolutionFilePath = GetTempFilePath(SolutionFileName, Path)
End Function

Function IterationLimitAvailable(Solver As ISolver) As Boolean
1         IterationLimitAvailable = (Len(Solver.IterationLimitName) <> 0)
End Function

Function PrecisionAvailable(Solver As ISolver) As Boolean
1         PrecisionAvailable = (Len(Solver.PrecisionName) <> 0)
End Function

Function TimeLimitAvailable(Solver As ISolver) As Boolean
1         TimeLimitAvailable = (Len(Solver.TimeLimitName) <> 0)
End Function

Function ToleranceAvailable(Solver As ISolver) As Boolean
1         ToleranceAvailable = (Len(Solver.ToleranceName) <> 0)
End Function

Function SolverIsPresent(Solver As ISolver, Optional SolverPath As String, Optional errorString As String, Optional Bitness As String) As Boolean
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
3         If Not SolverDirIsPresent Then
4             errorString = "Could not find the Solvers folder in the folder containing OpenSolver.xlam, " & _
                            "indicating OpenSolver has not been properly installed. Make sure you have " & _
                            "unzipped all files from the downloaded zip file to the same place."
5             SolverIsPresent = False
6             Exit Function
7         End If

8         If TypeOf Solver Is ISolverLocalExec Then
              Dim LocalExecSolver As ISolverLocalExec
9             Set LocalExecSolver = Solver
              
10            SolverPath = LocalExecSolver.GetExecPath(errorString, Bitness)
11            If Len(SolverPath) = 0 Then
12                SolverIsPresent = False
13            Else
14                SolverIsPresent = True
            #If Mac Then
                      ' Make sure solver is executable on Mac
15                    Exec "chmod +x " & MakePathSafe(SolverPath)
            #End If
16            End If
17        ElseIf TypeOf Solver Is ISolverLocalLib Then
              Dim LocalLibSolver As ISolverLocalLib
18            Set LocalLibSolver = Solver
19            SolverPath = LocalLibSolver.GetLibPath(errorString, Bitness)
20            SolverIsPresent = (Len(SolverPath) > 0)
21        ElseIf TypeOf Solver Is ISolverNeos Then
        #If Mac Then
22                SolverPath = NeosClientScriptPath()
23                If FileOrDirExists(SolverPath) Then
24                    SolverIsPresent = True
25                Else
26                    errorString = "Unable to find the NeosClient.py file at " & MakeSpacesNonBreaking(MakePathSafe(SolverPath))
27                    SolverPath = ""
28                End If
        #Else
29                SolverIsPresent = True
        #End If
          ElseIf TypeOf Solver Is CSolverSolveEngine Then
              SolverIsPresent = True
              
30        Else
31            SolverIsPresent = False
32        End If

ExitFunction:
33        If RaiseError Then RethrowError
34        Exit Function

ErrorHandler:
35        If Not ReportError("SolverCommon", "SolverIsPresent") Then Resume
36        RaiseError = True
37        GoTo ExitFunction
End Function

Function SolverIsAvailable(Solver As ISolver, Optional SolverPath As String, Optional errorString As String) As Boolean
1         If Not SolverIsPresent(Solver, SolverPath, errorString) Then Exit Function
          
2         If TypeOf Solver Is ISolverLocal Then
              Dim LocalSolver As ISolverLocal
3             Set LocalSolver = Solver
              
4             On Error GoTo ErrorHandlerLocal
5             If Len(LocalSolver.Version) <> 0 Then
6                 SolverIsAvailable = True
7             Else
ErrorHandlerLocal:
8                 SolverIsAvailable = False
9                 errorString = "Unable to access the " & DisplayName(Solver) & " solver at " & MakeSpacesNonBreaking(MakePathSafe(SolverPath))
10            End If
11        ElseIf TypeOf Solver Is ISolverNeos Then
12            SolverIsAvailable = True
          ElseIf TypeOf Solver Is CSolverSolveEngine Then
              SolverIsAvailable = True
          
13        Else
14            SolverIsAvailable = False
15        End If
End Function

Function AboutLocalSolver(LocalSolver As ISolverLocal) As String
          Dim SolverPath As String, errorString As String, Solver As ISolver
1         Set Solver = LocalSolver
2         If Not SolverIsPresent(Solver, SolverPath, errorString) Then
3             AboutLocalSolver = errorString
4         Else
              Dim LibVersion As String
5             If TypeOf Solver Is ISolverLocalLib Then
                  Dim LocalLibSolver As ISolverLocalLib
6                 Set LocalLibSolver = Solver
7                 LibVersion = "using " & LocalLibSolver.LibName & " v" & LocalLibSolver.LibVersion & " "
8             ElseIf TypeOf Solver Is CSolverGurobi Then
                  Dim GurobiSolver As CSolverGurobi
9                 Set GurobiSolver = Solver
10                SolverPath = GurobiSolver.ExecFilePath()
11            End If
          
12            AboutLocalSolver = DisplayName(Solver) & " " & _
                                 "v" & LocalSolver.Version & " " & _
                                 "(" & LocalSolver.Bitness & "-bit) " & _
                                 LibVersion & _
                                 "at " & MakeSpacesNonBreaking(MakePathSafe(SolverPath))
13        End If
End Function

Sub RunLocalSolver(s As COpenSolver, ExternalCommand As String)

          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         UpdateStatusBar "OpenSolver: Solving " & IIf(s.SolveRelaxation, "Relaxed ", vbNullString) & "Model... " & _
                          s.NumVars & " vars, " & _
                          s.NumDiscreteVars & " int vars " & "(" & s.NumBinVars & " bin), " & _
                          s.NumRows & " rows, " & _
                          s.SolverParameters.Item(s.Solver.TimeLimitName) & "s time limit, " & _
                          "limit of " & s.SolverParameters.Item(s.Solver.IterationLimitName) & " iterations, " & _
                          s.SolverParameters.Item(s.Solver.ToleranceName) * 100 & "% tolerance.", True
                                
          Dim exeResult As Long
4         ExecCapture ExternalCommand, s.LogFilePathName, GetTempFolder(), s.ShowIterationResults And Not s.MinimiseUserInteraction, exeResult
          
          ' Check log for any errors which can offer more descriptive messages than exeresult <> 0
5         s.Solver.CheckLog s
6         If exeResult <> 0 Then
7             RaiseGeneralError "The " & DisplayName(s.Solver) & " solver did not complete, but aborted with the error code " & exeResult & "." & vbCrLf & vbCrLf & "The last log file can be viewed under the OpenSolver menu and may give you more information on what caused this error.", _
                                "http://opensolver.org/help/#cbccrashes"
8         End If

ExitSub:
9         If RaiseError Then RethrowError
10        Exit Sub

ErrorHandler:
11        If Not ReportError("SolverCommon", "RunLocalSolver") Then Resume
12        RaiseError = True
13        GoTo ExitSub
End Sub

Function LibDir(Optional Bitness As String) As String
    #If Mac Then
1             LibDir = SolverDirMac
2             Bitness = "64"
    #ElseIf Win64 Then
3             LibDir = SolverDirWin64
4             Bitness = "64"
    #Else
5             LibDir = SolverDirWin32
6             Bitness = "32"
    #End If
7         LibDir = JoinPaths(SolverDir, LibDir)
End Function

Function SolverLibPath(LocalLibSolver As ISolverLocalLib, Optional errorString As String, Optional Bitness As String) As String
1         If Not GetExistingFilePathName(LibDir(Bitness), LocalLibSolver.LibBinary, SolverLibPath) Then
2             SolverLibPath = vbNullString
              Dim Solver As ISolver
3             Set Solver = LocalLibSolver
4             errorString = "Unable to find " & DisplayName(Solver) & " ('" & LocalLibSolver.LibBinary & "'). Folders searched:" & _
                            vbNewLine & MakePathSafe(LibDir())
5         End If
End Function

Function SolverExecPath(LocalExecSolver As ISolverLocalExec, Optional errorString As String, Optional Bitness As String) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim Solver As ISolver
3         Set Solver = LocalExecSolver

          Dim SearchPath As String
4         SearchPath = SolverDir
5         errorString = vbNullString
6         Bitness = vbNullString
          
    #If Mac Then
7             If GetExistingFilePathName(JoinPaths(SearchPath, SolverDirMac), LocalExecSolver.ExecName, SolverExecPath) Then
8                 Bitness = "64"
9             Else
10                errorString = errorString & vbNewLine & MakePathSafe(JoinPaths(SearchPath, SolverDirMac))
11                SolverExecPath = ""
12            End If
    #Else
              ' Look for the 64 bit version
13            If SystemIs64Bit And GetExistingFilePathName(JoinPaths(SearchPath, SolverDirWin64), LocalExecSolver.ExecName, SolverExecPath) Then
14                Bitness = "64"
              ' Look for the 32 bit version
15            Else
16                errorString = vbNewLine & MakePathSafe(JoinPaths(SearchPath, SolverDirWin64))
17                If GetExistingFilePathName(JoinPaths(SearchPath, SolverDirWin32), LocalExecSolver.ExecName, SolverExecPath) Then
18                    Bitness = "32"
19                    If SystemIs64Bit Then
20                        errorString = "Unable to find 64-bit " & DisplayName(Solver) & " in the Solvers folder. A 32-bit version will be used instead."
21                    End If
22                Else
23                    errorString = errorString & vbNewLine & MakePathSafe(JoinPaths(SearchPath, SolverDirWin32))
24                End If
25            End If
    #End If
              
26        If Len(Bitness) = 0 Then
              ' Failed to find a valid solver
27            SolverExecPath = vbNullString
28            errorString = "Unable to find " & DisplayName(Solver) & " ('" & LocalExecSolver.ExecName & "'). Folders searched:" & errorString
29        End If

ExitFunction:
30        If RaiseError Then RethrowError
31        Exit Function

ErrorHandler:
32        If Not ReportError("SolverCommon", "SolverExecPath") Then Resume
33        RaiseError = True
34        GoTo ExitFunction
End Function

Function SolverLinearity(Solver As ISolver) As OpenSolver_SolverType
1         If TypeOf Solver Is ISolverLinear Then
2             SolverLinearity = Linear
3         Else
4             SolverLinearity = NonLinear
5         End If
End Function

Function SensitivityAnalysisAvailable(Solver As ISolver) As Boolean
1         SensitivityAnalysisAvailable = False
2         If TypeOf Solver Is ISolverLinear Then
              Dim LinearSolver As ISolverLinear
3             Set LinearSolver = Solver
4             If LinearSolver.SensitivityAnalysisAvailable Then
5                 SensitivityAnalysisAvailable = True
6             End If
7         End If
End Function

Function SolverUsesUpperBounds(SolverShortName As String) As Boolean
1         Select Case LCase(SolverShortName)
          Case "nomad"
2             SolverUsesUpperBounds = True
3         Case Else
4             SolverUsesUpperBounds = False
5         End Select
End Function

Function WriteModelFile(s As COpenSolver) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim ModelFilePathName As String
3         ModelFilePathName = GetModelFilePath(s.Solver)
4         If FileOrDirExists(ModelFilePathName) Then DeleteFileAndVerify ModelFilePathName
          
          Dim FileSolver As ISolverFile
5         Set FileSolver = s.Solver
          
6         If s.Solver.ModelType = Diff Then
7             Select Case FileSolver.FileType
              Case AMPL
8                 WriteAMPLFile_Diff s, ModelFilePathName
9             Case LP
10                WriteLPFile_Diff s, ModelFilePathName
11            End Select
12        ElseIf s.Solver.ModelType = Parsed Then
13            Select Case FileSolver.FileType
              Case AMPL
14                WriteAMPLFile_Parsed s, ModelFilePathName
15            Case NL
16                WriteNLFile_Parsed s, ModelFilePathName
17            End Select
18        End If
          
19        If TypeOf s.Solver Is ISolverLocalExec Then
              ' Create a script to run the solver
              Dim LocalExecSolver As ISolverLocalExec
20            Set LocalExecSolver = s.Solver
21            WriteModelFile = LocalExecSolver.CreateSolveCommand(s)
22        ElseIf TypeOf s.Solver Is ISolverNeos Or TypeOf s.Solver Is CSolverSolveEngine Then
              ' Load the model file back into a string
23            Open ModelFilePathName For Input As #1
24                WriteModelFile = Input$(LOF(1), 1)
25            Close #1
26        End If
          
ExitFunction:
27        Close #1
28        If RaiseError Then RethrowError
29        Exit Function

ErrorHandler:
30        If Not ReportError("SolverCommon", "WriteModelFile") Then Resume
31        RaiseError = True
32        GoTo ExitFunction

End Function

Function GetModelFilePath(FileSolver As ISolverFile) As String
1         Select Case FileSolver.FileType
          Case AMPL
2             GetAMPLFilePath GetModelFilePath
3         Case NL
4             GetNLModelFilePath GetModelFilePath
5         Case LP
6             GetLPFilePath GetModelFilePath
7         End Select
End Function

Function RunSolver(s As COpenSolver, SolverCommand As String) As String
1         If TypeOf s.Solver Is ISolverLocalExec Then
2             RunLocalSolver s, SolverCommand
3             RunSolver = vbNullString
4         ElseIf TypeOf s.Solver Is ISolverNeos Then
5             RunSolver = CallNEOS(s, SolverCommand)
          ElseIf TypeOf s.Solver Is CSolverSolveEngine Then
              RunSolver = CallSolveEngine(s, SolverCommand)
6         End If
End Function

Function DisplayName(Solver As ISolver) As String
1         DisplayName = Solver.Name
2         If TypeOf Solver Is ISolverNeos Then
3             DisplayName = DisplayName & " on NEOS"
4         End If
End Function
