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

Private Const pSolutionFile = "modelsolution.txt"
Private Const pCostRangesFile = "costranges.txt"
Private Const pRHSRangesFile = "rhsranges.txt"
Private Const pFileType = OpenSolver_FileType.LP
Private Const pModelType = OpenSolver_ModelType.Diff

Private Const pIterationLimitName = "maxIterations"
Private Const pTimeLimitName = "seconds"
Private Const pToleranceName = "ratioGap"

' ISolver

Private Property Get ISolver_Name() As String
    ISolver_Name = pName
End Property

Private Property Get ISolver_ShortName() As String
    ISolver_ShortName = pShortName
End Property

Private Property Get ISolver_Title() As String
    ISolver_Title = pTitle
End Property

Private Property Get ISolver_Desc() As String
    ISolver_Desc = pDesc
End Property

Private Property Get ISolver_Link() As String
    ISolver_Link = pLink
End Property

Private Property Get ISolver_IterationLimitName() As String
    ISolver_IterationLimitName = pIterationLimitName
End Property

Private Property Get ISolver_PrecisionName() As String
End Property

Private Property Get ISolver_TimeLimitName() As String
    ISolver_TimeLimitName = pTimeLimitName
End Property

Private Property Get ISolver_ToleranceName() As String
    ISolver_ToleranceName = pToleranceName
End Property

Private Property Get ISolver_ModelType() As OpenSolver_ModelType
    ISolver_ModelType = pModelType
End Property

Private Sub ISolver_CleanFiles()
    CleanFiles
End Sub

Private Sub ISolver_CheckLog(s As COpenSolver)
    CheckLog s
End Sub

' ISolverLocal

Private Property Get ISolverLocal_Version() As String
    ISolverLocal_Version = Version()
End Property

Private Property Get ISolverLocal_Bitness() As String
    SolverIsPresent Me, Bitness:=ISolverLocal_Bitness
End Property

' ISolverLocalExec

Private Property Get ISolverLocalExec_ExecName() As String
    ISolverLocalExec_ExecName = pExecName
End Property

Private Function ISolverLocalExec_GetExecPath(Optional errorString As String, Optional Bitness As String) As String
    ISolverLocalExec_GetExecPath = SolverExecPath(Me, errorString, Bitness)
End Function

Private Function ISolverLocalExec_CreateSolveScript(ModelFilePathName As String, s As COpenSolver) As String
    ISolverLocalExec_CreateSolveScript = CreateSolveScript(ModelFilePathName, s)
End Function

' ISolverLinear

Private Property Get ISolverLinear_SensitivityAnalysisAvailable() As Boolean
    ISolverLinear_SensitivityAnalysisAvailable = pSensitivityAnalysisAvailable
End Property

' ISolverFile

Private Sub ISolverFile_ReadResults(s As COpenSolver, Optional solution As String)
    ReadResults s
End Sub

Private Property Get ISolverFile_FileType() As OpenSolver_FileType
    ISolverFile_FileType = pFileType
End Property

' Methods

Private Function Version() As String
' Get CBC version by running 'cbc -exit' at command line
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    Dim SolverPath As String
    If Not SolverIsPresent(Me, SolverPath) Then
        Version = ""
        GoTo ExitFunction
    End If
    
    Dim result As String
    result = ReadExternalCommandOutput(MakePathSafe(SolverPath) & " -exit")
    Version = Mid(result, InStr(result, "Version:") + 9, 5)

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("CSolverCbc", "Version") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

Private Sub CleanFiles()
    DeleteFileAndVerify SolutionFilePath()
    DeleteFileAndVerify CostRangesFilePath()
    DeleteFileAndVerify RHSRangesFilePath()
    DeleteFileAndVerify ScriptFilePath()
End Sub

Function ScriptFilePath() As String
    GetTempFilePath pScriptFile, ScriptFilePath
End Function

Function SolutionFilePath() As String
    GetTempFilePath pSolutionFile, SolutionFilePath
End Function

Function CostRangesFilePath() As String
    GetTempFilePath pCostRangesFile, CostRangesFilePath
End Function

Function RHSRangesFilePath() As String
    GetTempFilePath pRHSRangesFile, RHSRangesFilePath
End Function

Private Function CreateSolveScript(ModelFilePathName As String, s As COpenSolver) As String
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
    
    Dim CommandLineRunString As String, SolverParametersString As String
    SolverParametersString = ParametersToFlags(s.SolverParameters)
    
    CommandLineRunString = " -directory " & MakePathSafe(left(GetTempFolder, Len(GetTempFolder) - 1)) & _
                           " -import " & MakePathSafe(ModelFilePathName) & _
                           " " & SolverParametersString & _
                           " -solve " & _
                           IIf(s.bGetDuals, " -printingOptions all ", "") & _
                           " -solution " & MakePathSafe(SolutionFilePath) & _
                           IIf(s.bGetDuals, " -printingOptions rhs  -solution " & pRHSRangesFile & _
                                            " -printingOptions objective -solution " & pCostRangesFile, "")
                  
    Dim scriptFile As String, scriptFileContents As String
    scriptFile = ScriptFilePath()
    scriptFileContents = MakePathSafe(ISolverLocalExec_GetExecPath()) & CommandLineRunString
    CreateScriptFile scriptFile, scriptFileContents
    
    CreateSolveScript = scriptFile

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("CSolverCbc", "CreateSolveScript") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

Private Sub ReadResults(s As COpenSolver)
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    If Not FileOrDirExists(SolutionFilePath) Then
        Err.Raise Number:=OpenSolver_SolveError, Description:="The solver did not create a solution file. No new solution is available."
    End If

    Dim Response As String
    ' All parameters valid, now we can check the solution
    Open SolutionFilePath For Input As #1 ' supply path with filename
    Line Input #1, Response  ' Optimal - objective value              22

    s.SolutionWasLoaded = True
    If Response Like "Optimal*" Then
        s.SolveStatus = OpenSolverResult.Optimal
        s.SolveStatusString = "Optimal"
        '
    ElseIf Response Like "Infeasible*" Then
        s.SolveStatus = OpenSolverResult.Infeasible
        s.SolveStatusString = "No Feasible Solution"
        '
    ElseIf Response Like "Integer infeasible*" Then
        s.SolveStatus = OpenSolverResult.Infeasible
        s.SolveStatusString = "No Feasible Integer Solution"
        '
    ElseIf Response Like "Unbounded*" Then
        s.SolveStatus = OpenSolverResult.Unbounded
        s.SolveStatusString = "No Solution Found (Unbounded)"
        s.SolutionWasLoaded = False
        '
    ElseIf Response Like "Stopped on time *" Then ' Stopped on iterations or time
        s.SolveStatus = OpenSolverResult.LimitedSubOptimal
        s.SolveStatusString = "Stopped on Time Limit"
        If Response Like "*(no integer solution - continuous used)*" Then
            s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
        End If
        '
    ElseIf Response Like "Stopped on iterations*" Then ' Stopped on iterations or time
        s.SolveStatus = OpenSolverResult.LimitedSubOptimal
        s.SolveStatusString = "Stopped on Iteration Limit"
        If Response Like "*(no integer solution - continuous used)*" Then
            s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
        End If
        '
    ElseIf Response Like "Stopped on difficulties*" Then ' Stopped on iterations or time
        s.SolveStatus = OpenSolverResult.LimitedSubOptimal
        s.SolveStatusString = "Stopped on CBC difficulties"
        If Response Like "*(no integer solution - continuous used)*" Then
            s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
        End If
        '
    ElseIf Response Like "Stopped on ctrl-c*" Then ' Stopped on iterations or time
        s.SolveStatus = OpenSolverResult.LimitedSubOptimal
        s.SolveStatusString = "Stopped on Ctrl-C"
        If Response Like "*(no integer solution - continuous used)*" Then
            s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
        End If
        '
    ElseIf Response Like "Status unknown*" Then
        Err.Raise OpenSolver_CBCError, Description:="CBC solver did not solve the problem, suggesting there was an error in the CBC input parameters. The response was: " & vbCrLf _
         & Response _
         & vbCrLf & "The CBC command line can be found at:" _
         & vbCrLf & ScriptFilePath()
    Else
        Err.Raise OpenSolver_CBCError, Description:="The response from the CBC solver is not recognised. The response was: " & Response
    End If
    
    ' Remove the double spaces from Response
    Response = Replace(Response, "    ", " ")
    Response = Replace(Response, "   ", " ")
    Response = Replace(Response, "  ", " ")

    If s.SolutionWasLoaded Then
        ' We read in whatever solution CBC returned
        UpdateStatusBar "OpenSolver: Loading Solution... " & Response
    
        Dim Line As String, SplitLine() As String, Index As Double, NameValue As String, value As Double, CBCConstraintIndex As Long, StartOffset As Long
        If s.bGetDuals Then
            ' Read in the Solution File
            ' Line format: Index ConstraintName Value ShadowPrice
            
            Dim j As Long, row As Long
            CBCConstraintIndex = 0
            
            ' Throw away first constraint if it was from a seek objective model
            If s.ObjectiveSense = TargetObjective Then
                Line Input #1, Line
                CBCConstraintIndex = CBCConstraintIndex + 1
            End If

            j = 1
            For row = 1 To s.NumRows
                If s.GetSparseACount(row) = 0 Then
                    ' This constraint was not written to the model, as it had no coefficients. Just ignore it.
                    s.rConstraintList.Cells(row, 2).ClearContents
                Else
                    Line Input #1, Line
                    SplitLine = SplitWithoutRepeats(Line, " ")

                    ' In the case of LpStatusInfeasible, we can get lines that start **. We strip the **
                    StartOffset = 0
                    If SplitLine(StartOffset) = "**" Then StartOffset = 1

                    ' Check the index of the row
                    If CInt(SplitLine(StartOffset)) <> CBCConstraintIndex Then
                        Err.Raise OpenSolver_CBCError, Description:="While reading the CBC solution file, OpenSolver found an unexpected constraint row."
                    End If

                    NameValue = SplitLine(StartOffset + 1)
                    s.FinalValue(j) = Val(SplitLine(StartOffset + 2))
                    value = Val(SplitLine(StartOffset + 3))
                    If s.ObjectiveSense = MaximiseObjective Then value = -value
                    s.ShadowPrice(j) = value
                    CBCConstraintIndex = CBCConstraintIndex + 1
                    j = j + 1
                End If
            Next row
            ReadSensitivityData s
        End If
      
        ' Now we read in the decision variable values
        ' Line format: Index VariableName Value ReducedCost
        j = 1
        While Not EOF(1)
            Line Input #1, Line
            SplitLine = SplitWithoutRepeats(Line, " ")

            ' In the case of LpStatusInfeasible, we can get lines that start **. We strip the **
            StartOffset = 0
            If SplitLine(StartOffset) = "**" Then StartOffset = 1
            
            Index = CInt(SplitLine(StartOffset))
            NameValue = SplitLine(StartOffset + 1)
            If left(NameValue, 1) = "_" Then NameValue = Mid(NameValue, 2) ' Strip any _ character added to make a valid name
            s.VarCell(j) = NameValue
            s.FinalVarValue(j) = Val(SplitLine(StartOffset + 2))
           
            If s.bGetDuals Then
                value = Val(SplitLine(StartOffset + 3))
                If s.ObjectiveSense = MaximiseObjective Then value = -value
                s.ReducedCosts(j) = value
            End If
            j = j + 1
        Wend
        s.SolutionWasLoaded = True

    End If

ExitSub:
    Close #1
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Sub

ErrorHandler:
    If Not ReportError("CSolverCBC", "ReadResults") Then Resume
    RaiseError = True
    GoTo ExitSub
End Sub

Sub ReadSensitivityData(s As COpenSolver)
'Reads the two files with the limits on the bounds of shadow prices and reduced costs
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
    
    Dim LineData() As String
    Dim Line As String, row As Long, j As Long
    
    'Find the ranges on the constraints

    Open RHSRangesFilePath For Input As #2 ' supply path with filename
    Line Input #2, Line 'Dont want first line
    j = 1
    While Not EOF(2)
        Line Input #2, Line
        LineData() = Split(Line, ",")
        s.IncreaseCon(j) = Val(LineData(2))
        s.DecreaseCon(j) = Val(LineData(4))
        j = j + 1
    Wend
    Close 2
    
    j = 1
    'Find the ranges on the variables
    Open CostRangesFilePath For Input As #2 ' supply path with filename
    Line Input #2, Line 'Dont want first line
    row = s.NumRows + 2
    While Not EOF(2)
        Line Input #2, Line
        LineData() = Split(Line, ",")
        If s.ObjectiveSense = MaximiseObjective Then
            s.IncreaseVar(j) = Val(LineData(4))
            s.DecreaseVar(j) = Val(LineData(2))
        Else
            s.IncreaseVar(j) = Val(LineData(2))
            s.DecreaseVar(j) = Val(LineData(4))
        End If
        j = j + 1
    Wend

ExitSub:
    Close #2
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Sub

ErrorHandler:
    If Not ReportError("CSolverCbc", "ReadSensitivityData") Then Resume
    RaiseError = True
    GoTo ExitSub
End Sub

Private Sub CheckLog(s As COpenSolver)
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
    
    If Not FileOrDirExists(s.LogFilePathName) Then
        Err.Raise Number:=OpenSolver_SolveError, Description:="The solver did not create a log file. No new solution is available."
    End If
    
    Dim message As String
    Open s.LogFilePathName For Input As #1
        message = Input$(LOF(1), 1)
    Close #1
    
    Dim Key As Variant
    For Each Key In s.SolverParameters.Keys
        If InStrText(message, "parameter " & Key & " value remains") Then
            Err.Raise OpenSolver_SolveError, Description:= _
                "The value for the parameter '" & Key & "' supplied to CBC was invalid. " & _
                "Please check the value you have specified, or consult the CBC documentation for more information."
        End If
        If InStrText(message, "No match for " & Key) Then
            Err.Raise OpenSolver_SolveError, Description:= _
                "The parameter '" & Key & "' was not recognised by CBC. " & _
                "Please check the parameter name you have specified, or consult the CBC documentation for more information."
        End If
    Next Key

ExitSub:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Sub

ErrorHandler:
    If Not ReportError("CSolverCbc", "CheckLog") Then Resume
    RaiseError = True
    GoTo ExitSub
End Sub


