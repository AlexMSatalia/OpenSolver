Attribute VB_Name = "SolverCommon"
Option Explicit

Public Const LogFileName = "log1.tmp"
Public Const SolutionFileName = "model.sol"

Public Const SolverDir As String = "Solvers"
Public Const SolverDirMac As String = "osx"
Public Const SolverDirWin32 As String = "win32"
Public Const SolverDirWin64 As String = "win64"

Public LastUsedSolver As String

Sub SolveModel(s As COpenSolver, ShouldSolveRelaxation As Boolean, ShouldMinimiseUserInteraction As Boolean)
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
    
    Application.EnableCancelKey = xlErrorHandler
    Application.Cursor = xlWait

    s.SolveStatus = OpenSolverResult.Unsolved
    s.SolveStatusString = "Unsolved"
    s.SolveStatusComment = ""
    s.SolutionWasLoaded = False
    
    ' Track whether to show messages
    s.MinimiseUserInteraction = ShouldMinimiseUserInteraction
    s.SolveRelaxation = ShouldSolveRelaxation

    Dim oldCalculationMode As Long, oldIterationMode As Boolean
    oldIterationMode = Application.Iteration
    oldCalculationMode = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    Dim ScreenStatus As Boolean
    ScreenStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    If s.ModelStatus <> Built Then
        Err.Raise Number:=OpenSolver_SolveError, Description:="The model cannot be solved as it has not yet been built."
    End If

    'Check that solver is available
    Dim errorString As String
    If Not SolverIsAvailable(s.Solver, errorString:=errorString) Then
        Err.Raise Number:=OpenSolver_SolveError, Description:=errorString
    End If

    Dim LogFilePathName As String
    If GetLogFilePath(LogFilePathName) Then DeleteFileAndVerify LogFilePathName
    s.LogFilePathName = LogFilePathName
    
    ' Clean Solver specific files
    s.Solver.CleanFiles
    
    'Check that we can write to all cells
    TestCellsForWriting s.AdjustableCells

    If TypeOf s.Solver Is ISolverLocalLib Then
        Dim LocalLibSolver As ISolverLocalLib
        Set LocalLibSolver = s.Solver
        LocalLibSolver.Solve s
    ElseIf TypeOf s.Solver Is ISolverFile Then
        Dim LinearSolver As ISolverLinear
        Dim FileSolver As ISolverFile
        Set FileSolver = s.Solver
        
        'Delete any existing solution file
        Dim SolutionFilePathName As String
        If GetSolutionFilePath(SolutionFilePathName) Then DeleteFileAndVerify SolutionFilePathName
        s.SolutionFilePathName = SolutionFilePathName
        
        ' Check if we need to request duals from the solver
        If TypeOf s.Solver Is ISolverLinear Then
            Set LinearSolver = s.Solver
            Set s.rConstraintList = GetDuals(s.sheet)
            s.DualsOnSameSheet = (Not s.rConstraintList Is Nothing)
            s.DualsOnNewSheet = GetDualsOnSheet(s.sheet)
            s.bGetDuals = ((s.IntegerCellsRange Is Nothing And s.BinaryCellsRange Is Nothing) Or s.SolveRelaxation) And _
                          (s.DualsOnNewSheet Or s.DualsOnSameSheet) And LinearSolver.SensitivityAnalysisAvailable
        End If
        
        ' Set up arrays to hold solution values (to avoid dynamically resizing them later)
        s.PrepareForSolution
        
        ' Write file
        Dim SolverCommand As String
        UpdateStatusBar "OpenSolver: Writing Model to disk... " & s.numVars & " vars, " & s.NumRows & " rows.", True
        SolverCommand = WriteModelFile(s)
        
        ' Check if anything detected while writing file
        If s.SolveStatus <> OpenSolverResult.Unsolved Then GoTo ExitSub
        
        ' Run solver and read results
        Dim solution As String
        solution = RunSolver(s, SolverCommand)
        FileSolver.ReadResults s, solution
        
        s.LoadResultsToSheet
        
        If TypeOf s.Solver Is ISolverLinear Then
            ' Perform a linearity check unless the user has requested otherwise
            If GetLinearityCheck() Then
                Dim fullLinearityCheckWasPerformed As Boolean
                QuickLinearityCheck fullLinearityCheckWasPerformed, s
                If fullLinearityCheckWasPerformed Then
                    s.SolveStatus = OpenSolverResult.AbortedThruUserAction
                    s.SolveStatusString = "No Solution Found"
                End If
            End If
            
            ' Get sensitivity results
            If s.bGetDuals And s.SolveStatus = OpenSolverResult.Optimal Then
                'write the duals on the same sheet if the user has picked this option
                If s.DualsOnSameSheet Then WriteConstraintListToSheet s.rConstraintList, s
                'If the user wants a new sheet with the sensitivity data then call the functions that write this
                If s.DualsOnNewSheet Then
                    Dim newSheet As Worksheet, currentSheet As Worksheet
                    ' Save sheet selection
                    GetActiveSheetIfMissing currentSheet
                    
                    Set newSheet = MakeNewSheet(s.sheet.Name & " Sensitivity", GetUpdateSensitivity(s.sheet))
                    
                    ' Restore old sheet selection
                    currentSheet.Select
                    WriteConstraintSensitivityTable newSheet, s
                End If
            ElseIf Not s.bGetDuals And (s.DualsOnNewSheet Or s.DualsOnSameSheet) And LinearSolver.SensitivityAnalysisAvailable Then
                Err.Raise Number:=OpenSolver_SolveError, Description:= _
                    "Could not get sensitivity analysis due to binary and/or integer constraints." & vbNewLine & vbNewLine & _
                    "Turn off sensitivity in the model dialogue or reformulate your model without these constraints." & vbNewLine & vbNewLine & _
                    "The " & s.Solver.ShortName & " solution has been returned to the sheet." & vbNewLine
            End If
        End If
    End If

ExitSub:
    Application.Cursor = xlDefault
    Application.StatusBar = False ' Resume normal status bar behaviour
    Application.ScreenUpdating = ScreenStatus
    Application.Calculation = oldCalculationMode
    Application.Iteration = oldIterationMode
    Application.Calculate
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Sub

ErrorHandler:
    If Not ReportError("SolverCommon", "SolveModel") Then Resume
    RaiseError = True
    GoTo ExitSub

End Sub

Function CreateSolver(SolverShortName As String) As ISolver
    Select Case SolverShortName
    Case "CBC":     Set CreateSolver = New CSolverCbc
    Case "Gurobi":  Set CreateSolver = New CSolverGurobi
    Case "NeosCBC": Set CreateSolver = New CSolverNeosCbc
    Case "Bonmin":  Set CreateSolver = New CSolverBonmin
    Case "Couenne": Set CreateSolver = New CSolverCouenne
    Case "NOMAD":   Set CreateSolver = New CSolverNomad
    Case "NeosBon": Set CreateSolver = New CSolverNeosBon
    Case "NeosCou": Set CreateSolver = New CSolverNeosCou
    Case Else: Err.Raise OpenSolver_ModelError, Description:="The specified solver ('" & SolverShortName & "') was not recognised."
    End Select
End Function

Function GetLogFilePath(ByRef Path As String) As Boolean
    GetLogFilePath = GetTempFilePath(LogFileName, Path)
End Function

Function GetSolutionFilePath(ByRef Path As String) As Boolean
    GetSolutionFilePath = GetTempFilePath(SolutionFileName, Path)
End Function

Function IterationLimitAvailable(Solver As ISolver) As Boolean
    IterationLimitAvailable = (Len(Solver.IterationLimitName) <> 0)
End Function

Function PrecisionAvailable(Solver As ISolver) As Boolean
    PrecisionAvailable = (Len(Solver.PrecisionName) <> 0)
End Function

Function TimeLimitAvailable(Solver As ISolver) As Boolean
    TimeLimitAvailable = (Len(Solver.TimeLimitName) <> 0)
End Function

Function ToleranceAvailable(Solver As ISolver) As Boolean
    ToleranceAvailable = (Len(Solver.ToleranceName) <> 0)
End Function

Function SolverIsPresent(Solver As ISolver, Optional SolverPath As String, Optional errorString As String, Optional Bitness As String) As Boolean
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    If TypeOf Solver Is ISolverLocalExec Then
        Dim LocalExecSolver As ISolverLocalExec
        Set LocalExecSolver = Solver
        
        SolverPath = LocalExecSolver.GetExecPath(errorString, Bitness)
        If SolverPath = "" Then
            SolverIsPresent = False
        Else
            SolverIsPresent = True
            #If Mac Then
                ' Make sure solver is executable on Mac
                Exec "chmod +x " & MakePathSafe(SolverPath)
            #End If
        End If
    ElseIf TypeOf Solver Is ISolverLocalLib Then
        Dim LocalLibSolver As ISolverLocalLib
        Set LocalLibSolver = Solver
        SolverPath = LocalLibSolver.GetLibPath(errorString, Bitness)
        SolverIsPresent = (SolverPath <> "")
    ElseIf TypeOf Solver Is ISolverNeos Then
        #If Mac Then
            SolverPath = NeosClientScriptPath()
            If FileOrDirExists(SolverPath) Then
                SolverIsPresent = True
            Else
                errorString = "Unable to find the NeosClient.py file at " & MakeSpacesNonBreaking(MakePathSafe(SolverPath))
                SolverPath = ""
            End If
        #Else
            SolverIsPresent = True
        #End If
    Else
        SolverIsPresent = False
    End If

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("SolverCommon", "SolverIsPresent") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

Function SolverIsAvailable(Solver As ISolver, Optional SolverPath As String, Optional errorString As String) As Boolean
    If Not SolverIsPresent(Solver, SolverPath, errorString) Then Exit Function
    
    If TypeOf Solver Is ISolverLocal Then
        Dim LocalSolver As ISolverLocal
        Set LocalSolver = Solver
        
        On Error GoTo ErrorHandlerLocal
        If Len(LocalSolver.Version) <> 0 Then
            SolverIsAvailable = True
        Else
ErrorHandlerLocal:
            SolverIsAvailable = False
            errorString = "Unable to access the " & DisplayName(Solver) & " solver at " & MakeSpacesNonBreaking(MakePathSafe(SolverPath))
        End If
    ElseIf TypeOf Solver Is ISolverNeos Then
        SolverIsAvailable = True
    Else
        SolverIsAvailable = False
    End If
End Function

Function AboutLocalSolver(LocalSolver As ISolverLocal) As String
    Dim SolverPath As String, errorString As String, Solver As ISolver
    Set Solver = LocalSolver
    If Not SolverIsPresent(Solver, SolverPath, errorString) Then
        AboutLocalSolver = errorString
    Else
        Dim LibVersion As String
        If TypeOf Solver Is ISolverLocalLib Then
            Dim LocalLibSolver As ISolverLocalLib
            Set LocalLibSolver = Solver
            LibVersion = "using " & LocalLibSolver.LibName & " v" & LocalLibSolver.LibVersion & " "
        ElseIf TypeOf Solver Is CSolverGurobi Then
            Dim GurobiSolver As CSolverGurobi
            Set GurobiSolver = Solver
            SolverPath = GurobiSolver.ExecFilePath()
        End If
    
        AboutLocalSolver = DisplayName(Solver) & " " & _
                           "v" & LocalSolver.Version & " " & _
                           "(" & LocalSolver.Bitness & "-bit) " & _
                           LibVersion & _
                           "at " & MakeSpacesNonBreaking(MakePathSafe(SolverPath))
    End If
End Function

Sub RunLocalSolver(s As COpenSolver, ExternalCommand As String)

    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    UpdateStatusBar "OpenSolver: Solving " & IIf(s.SolveRelaxation, "Relaxed ", "") & "Model... " & _
                    s.numVars & " vars, " & _
                    s.NumIntVars & " int vars " & "(" & s.NumBinVars & " bin), " & _
                    s.NumRows & " rows, " & _
                    s.SolverParameters.Item(s.Solver.TimeLimitName) & "s time limit, " & _
                    s.SolverParameters.Item(s.Solver.ToleranceName) * 100 & "% tolerance.", True
                          
    Dim exeResult As Long
    ExecCapture ExternalCommand, s.LogFilePathName, GetTempFolder(), s.ShowIterationResults And Not s.MinimiseUserInteraction, exeResult
    
    ' Check log for any errors which can offer more descriptive messages than exeresult <> 0
    s.Solver.CheckLog s
    If exeResult <> 0 Then
        Err.Raise Number:=OpenSolver_SolveError, Description:="The " & DisplayName(s.Solver) & " solver did not complete, but aborted with the error code " & exeResult & "." & vbCrLf & vbCrLf & "The last log file can be viewed under the OpenSolver menu and may give you more information on what caused this error."
    End If

ExitSub:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Sub

ErrorHandler:
    If Not ReportError("SolverCommon", "RunLocalSolver") Then Resume
    RaiseError = True
    GoTo ExitSub
End Sub

Function LibDir(Optional Bitness As String) As String
    LibDir = JoinPaths(ThisWorkbook.Path, SolverDir)
    #If Mac Then
        LibDir = JoinPaths(LibDir, SolverDirMac)
        Bitness = "64"
    #ElseIf Win64 Then
        LibDir = JoinPaths(LibDir, SolverDirWin64)
        Bitness = "64"
    #Else
        LibDir = JoinPaths(LibDir, SolverDirWin32)
        Bitness = "32"
    #End If
End Function

Function SolverLibPath(LocalLibSolver As ISolverLocalLib, Optional errorString As String, Optional Bitness As String) As String
    If Not GetExistingFilePathName(LibDir(Bitness), LocalLibSolver.LibBinary, SolverLibPath) Then
        SolverLibPath = ""
        Dim Solver As ISolver
        Set Solver = LocalLibSolver
        errorString = "Unable to find " & DisplayName(Solver) & " ('" & LocalLibSolver.LibBinary & "'). Folders searched:" & _
                      vbNewLine & MakePathSafe(LibDir())
    End If
End Function

Function SolverExecPath(LocalExecSolver As ISolverLocalExec, Optional errorString As String, Optional Bitness As String) As String
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
    
    Dim Solver As ISolver
    Set Solver = LocalExecSolver

    Dim SearchPath As String
    SearchPath = JoinPaths(ThisWorkbook.Path, SolverDir)
    errorString = ""
    Bitness = ""
    
    #If Mac Then
        If GetExistingFilePathName(JoinPaths(SearchPath, SolverDirMac), LocalExecSolver.ExecName, SolverExecPath) Then
            Bitness = "64"
        Else
            errorString = errorString & vbNewLine & MakePathSafe(JoinPaths(SearchPath, SolverDirMac))
            SolverExecPath = ""
        End If
    #Else
        ' Look for the 64 bit version
        If SystemIs64Bit And GetExistingFilePathName(JoinPaths(SearchPath, SolverDirWin64), LocalExecSolver.ExecName, SolverExecPath) Then
            Bitness = "64"
        ' Look for the 32 bit version
        Else
            errorString = vbNewLine & MakePathSafe(JoinPaths(SearchPath, SolverDirWin64))
            If GetExistingFilePathName(JoinPaths(SearchPath, SolverDirWin32), LocalExecSolver.ExecName, SolverExecPath) Then
                Bitness = "32"
                If SystemIs64Bit Then
                    errorString = "Unable to find 64-bit " & DisplayName(Solver) & " in the Solvers folder. A 32-bit version will be used instead."
                End If
            Else
                errorString = errorString & vbNewLine & MakePathSafe(JoinPaths(SearchPath, SolverDirWin32))
            End If
        End If
    #End If
        
    If Len(Bitness) = 0 Then
        ' Failed to find a valid solver
        SolverExecPath = ""
        errorString = "Unable to find " & DisplayName(Solver) & " ('" & LocalExecSolver.ExecName & "'). Folders searched:" & errorString
    End If

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("SolverCommon", "SolverExecPath") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

Function SolverLinearity(Solver As ISolver) As OpenSolver_SolverType
    If TypeOf Solver Is ISolverLinear Then
        SolverLinearity = Linear
    Else
        SolverLinearity = NonLinear
    End If
End Function

Function SensitivityAnalysisAvailable(Solver As ISolver) As Boolean
    SensitivityAnalysisAvailable = False
    If TypeOf Solver Is ISolverLinear Then
        Dim LinearSolver As ISolverLinear
        Set LinearSolver = Solver
        If LinearSolver.SensitivityAnalysisAvailable Then
            SensitivityAnalysisAvailable = True
        End If
    End If
End Function

Function SolverUsesUpperBounds(SolverShortName As String) As Boolean
    Select Case SolverShortName
    Case "NOMAD"
        SolverUsesUpperBounds = True
    Case Else
        SolverUsesUpperBounds = False
    End Select
End Function

Function WriteModelFile(s As COpenSolver) As String
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    Dim ModelFilePathName As String
    ModelFilePathName = GetModelFilePath(s.Solver)
    If FileOrDirExists(ModelFilePathName) Then DeleteFileAndVerify ModelFilePathName
    
    Dim FileSolver As ISolverFile
    Set FileSolver = s.Solver
    
    If s.Solver.ModelType = Diff Then
        Select Case FileSolver.FileType
        Case AMPL
            WriteAMPLFile_Diff s, ModelFilePathName
        Case LP
            WriteLPFile_Diff s, ModelFilePathName
        End Select
    ElseIf s.Solver.ModelType = Parsed Then
        Select Case FileSolver.FileType
        Case AMPL
            WriteAMPLFile_Parsed s, ModelFilePathName
        Case NL
            WriteNLFile_Parsed s, ModelFilePathName
        End Select
    End If
    
    If TypeOf s.Solver Is ISolverLocalExec Then
        ' Create a script to run the solver
        Dim LocalExecSolver As ISolverLocalExec
        Set LocalExecSolver = s.Solver
        WriteModelFile = LocalExecSolver.CreateSolveCommand(s)
    ElseIf TypeOf s.Solver Is ISolverNeos Then
        ' Load the model file back into a string
        Open ModelFilePathName For Input As #1
            WriteModelFile = Input$(LOF(1), 1)
        Close #1
    End If
    
ExitFunction:
    Close #1
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("SolverCommon", "WriteModelFile") Then Resume
    RaiseError = True
    GoTo ExitFunction

End Function

Function GetModelFilePath(FileSolver As ISolverFile) As String
    Select Case FileSolver.FileType
    Case AMPL
        GetAMPLFilePath GetModelFilePath
    Case NL
        GetNLModelFilePath GetModelFilePath
    Case LP
        GetLPFilePath GetModelFilePath
    End Select
End Function

Function RunSolver(s As COpenSolver, SolverCommand As String) As String
    If TypeOf s.Solver Is ISolverLocalExec Then
        RunLocalSolver s, SolverCommand
        RunSolver = ""
    ElseIf TypeOf s.Solver Is ISolverNeos Then
        RunSolver = CallNEOS(s, SolverCommand)
    End If
End Function

Function DisplayName(Solver As ISolver) As String
    DisplayName = Solver.Name
    If TypeOf Solver Is ISolverNeos Then
        DisplayName = DisplayName & " on NEOS"
    End If
End Function
