VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSolverNomad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISolver
Implements ISolverLocal
Implements ISolverLocalLib

Private Const pTitle = "NOMAD (Non-linear solver)"
Private Const pDesc = "Nomad (Nonsmooth Optimization by Mesh Adaptive Direct search) is a C++ implementation of the Mesh Adaptive Direct Search (Mads) algorithm that solves non-linear problems. It works by updating the values on the sheet and passing them to the C++ solver. Like many non-linear solvers NOMAD cannot guarantee optimality of its solutions."
Private Const pLink = "http://www.gerad.ca/nomad/Project/Home.html"
Private Const pType = OpenSolver_SolverType.NonLinear
Private Const pName = "NOMAD"
Private Const pShortName = "NOMAD"

Private Const pIterationLimitName = "MAX_BB_EVAL"
Private Const pPrecisionName = "EPSILON"
Private Const pTimeLimitName = "MAX_TIME"

Private Const pModelType = OpenSolver_ModelType.None

Private Const pLibName = "OpenSolverNomad.dll"

'NOMAD return status codes
Private Enum NomadResult
    UserCancelled = -3
    Optimal = 0
    ErrorOccured = 1
    SolveStoppedIter = 2
    SolveStoppedTime = 3
    Infeasible = 4
    SolveStoppedNoSolution = 10
End Enum

' Don't forget we need to chdir to the directory containing the lib before calling any of the functions
#If VBA7 Then
    Private Declare PtrSafe Function NomadMain Lib "OpenSolverNomad.dll" (ByVal SolveRelaxation As Boolean) As Long
    Private Declare PtrSafe Function NomadVersion Lib "OpenSolverNomad.dll" () As String
    Private Declare PtrSafe Function NomadLibVersion Lib "OpenSolverNomad.dll" Alias "NomadDLLVersion" () As String
#Else
    Private Declare Function NomadMain Lib "OpenSolverNomad.dll" (ByVal SolveRelaxation As Boolean) As Long
    Private Declare Function NomadVersion Lib "OpenSolverNomad.dll" () As String
    Private Declare Function NomadLibVersion Lib "OpenSolverNomad.dll" Alias "NomadDLLVersion" () As String
#End If

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

Private Property Get ISolver_SolverType() As OpenSolver_SolverType
    ISolver_SolverType = pType
End Property

Private Property Get ISolver_IterationLimitName() As String
    ISolver_IterationLimitName = pIterationLimitName
End Property

Private Property Get ISolver_PrecisionName() As String
    ISolver_PrecisionName = pPrecisionName
End Property

Private Property Get ISolver_TimeLimitName() As String
    ISolver_TimeLimitName = pTimeLimitName
End Property

Private Property Get ISolver_ToleranceName() As String
End Property

Private Property Get ISolver_ModelType() As OpenSolver_ModelType
    ISolver_ModelType = pModelType
End Property

Private Sub ISolver_CleanFiles()
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

' ISolverLocalLib

Private Function ISolverLocalLib_GetLibPath(Optional errorString As String, Optional Bitness As String) As String
    ISolverLocalLib_GetLibPath = GetPath(errorString, Bitness)
End Function

Private Property Get ISolverLocalLib_LibName() As String
    ISolverLocalLib_LibName = pLibName
End Property

Private Property Get ISolverLocalLib_LibVersion() As String
    ISolverLocalLib_LibVersion = LibVersion()
End Property

Private Sub ISolverLocalLib_Solve(s As COpenSolver)
    Solve s
End Sub

' Methods

Private Function GetPath(Optional errorString As String, Optional Bitness As String) As String
    #If Mac Then
          errorString = "NOMAD for OpenSolver is not currently supported on Mac"
          GetPath = ""
    #Else
          GetPath = SolverLibPath(Me, errorString, Bitness)
    #End If
End Function

Private Function Version() As String
    If Not SolverIsPresent(Me) Then
        Version = ""
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
       
    ' Set current dir for finding the DLL
    Dim currentDir As String
    currentDir = CurDir
    SetCurrentDirectory LibDir()
    
    ' Get version info from DLL
    ' Save to a new string first - modifying the string from the DLL can sometimes crash Excel
    Dim sNomadVersion As String
    sNomadVersion = NomadVersion()
    sNomadVersion = left(Replace(sNomadVersion, vbNullChar, ""), 5)
    
    SetCurrentDirectory currentDir
    
    Version = sNomadVersion
    Exit Function

ErrorHandler:
    Version = ""
End Function

Private Function LibVersion() As String
    If Not SolverIsPresent(Me) Then
        LibVersion = ""
        Exit Function
    End If
    
    Dim currentDir As String
    
    ' Set current dir for finding the DLL
    currentDir = CurDir
    SetCurrentDirectory LibDir()
    
    ' Get version info from DLL
    ' Save to a new string first - modifying the string from the DLL can sometimes crash Excel
    Dim sLibVersion As String
    sLibVersion = NomadLibVersion()
    sLibVersion = left(Replace(sLibVersion, vbNullChar, ""), 5)
    
    SetCurrentDirectory currentDir
    
    LibVersion = sLibVersion
End Function

Private Sub Solve(s As COpenSolver)
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    Dim ScreenStatus As Boolean
    ScreenStatus = Application.ScreenUpdating

    If Not GetShowSolverProgress() Then Application.ScreenUpdating = False

    ' Check precision is not 0
    Dim SolveOptions As SolveOptionsType
    GetSolveOptions s.sheet, SolveOptions
    
    If SolveOptions.Precision <= 0 Then
        Err.Raise Number:=OpenSolver_NomadError, Description:="The current level of precision (" & CStr(SolveOptions.Precision) & ") is invalid. Please set the precision to a small positive (non-zero) value and try again."
    End If
    
    Dim currentDir As String
    currentDir = CurDir
    
    SetCurrentDirectory LibDir()

    Set SolverNOMAD.OS = s
    
    ' We need to call NomadMain directly rather than use Application.Run .
    ' Using Application.Run causes the API calls inside the DLL to fail on 64 bit Office
    Dim NomadRetVal As Long
    NomadRetVal = NomadMain(s.SolveRelaxation)
    
    'Catch any errors that occured while Nomad was solving
    Select Case NomadRetVal
    Case NomadResult.ErrorOccured
        s.SolveStatus = OpenSolverResult.ErrorOccurred

        ' Check logs for more info and raise an error if we find anything specific
        CheckLog s
        
        Err.Raise Number:=OpenSolver_NomadError, Description:="There was an error while Nomad was solving. No solution has been loaded into the sheet."
    Case NomadResult.SolveStoppedIter
        s.SolveStatus = OpenSolverResult.LimitedSubOptimal
        s.SolveStatusString = "NOMAD reached the maximum number of iterations and returned the best feasible solution it found. " & _
                              "This solution is not guaranteed to be an optimal solution." & vbNewLine & vbNewLine & _
                              "You can increase the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
        s.SolutionWasLoaded = True
    Case NomadResult.SolveStoppedTime
        s.SolveStatusString = "NOMAD reached the maximum time and returned the best feasible solution it found. " & _
                              "This solution is not guaranteed to be an optimal solution." & vbNewLine & vbNewLine & _
                              "You can increase the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
        s.SolveStatus = OpenSolverResult.LimitedSubOptimal
        s.SolutionWasLoaded = True
    Case NomadResult.Infeasible
        s.SolveStatusString = "Nomad reached the maximum time or number of iterations without finding a feasible solution. " & _
                              "The best infeasible solution has been returned to the sheet." & vbNewLine & vbNewLine & _
                              "You can increase the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
        s.SolveStatus = OpenSolverResult.Infeasible
        s.SolutionWasLoaded = True
    Case NomadResult.SolveStoppedNoSolution
        s.SolveStatusString = "Nomad could not find a feasible solution. " & _
                              "The best infeasible solution has been returned to the sheet." & vbNewLine & vbNewLine & _
                              "Try resolving at a different start point or check whether your model is feasible or relax some of your constraints."
        s.SolveStatus = OpenSolverResult.Infeasible
        s.SolutionWasLoaded = True
    Case NomadResult.UserCancelled
        Err.Raise OpenSolver_UserCancelledError, "Running NOMAD", "Model solve cancelled by user."
    Case NomadResult.Optimal
        s.SolveStatus = OpenSolverResult.Optimal
        s.SolveStatusString = "Optimal"
    End Select
    
ExitSub:
    SetCurrentDirectory currentDir
    Application.Calculate
    Set SolverNOMAD.OS = Nothing
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Sub

ErrorHandler:
    If Not ReportError("CSolverNomad", "Solve") Then Resume
    RaiseError = True
    GoTo ExitSub

End Sub

Private Sub CheckLog(s As COpenSolver)
' If NOMAD encounters an error, we dump the exception to the log file. We can use this to deduce what went wrong
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
    
    If Not FileOrDirExists(s.LogFilePathName) Then
        Err.Raise Number:=OpenSolver_SolveError, Description:="The solver did not create a log file. No new solution is available."
    End If
    
    Dim message As String
    Open s.LogFilePathName For Input As #3
        message = Input$(LOF(3), 3)
    Close #3
    
    If Not InStrText(message, "NOMAD") Then GoTo ExitSub

    If InStrText(message, "invalid parameter: DIMENSION") Then
        Dim MaxSize As Long, Position As Long
        Position = InStrRev(message, " ")
        MaxSize = CInt(Mid(message, Position + 1, InStrRev(message, ")") - Position - 1))
        Err.Raise OpenSolver_NomadError, Description:="This model contains too many variables for NOMAD to solve. NOMAD is only capable of solving models with up to " & MaxSize & " variables."
    End If
    
    Dim Key As Variant
    For Each Key In s.SolverParameters.Keys()
        If InStrText(message, "invalid parameter: " & UCase(Key) & " - unknown") Then
            Err.Raise OpenSolver_NomadError, Description:="The parameter '" & UCase(Key) & "' was not understood by NOMAD. Check that you have specified a valid parameter name, or consult the NOMAD documentation for more information."
        End If
        If InStrText(message, "invalid parameter: " & UCase(Key)) Then
            Err.Raise OpenSolver_NomadError, Description:="The value of the parameter '" & UCase(Key) & "' supplied to NOMAD was invalid. Check that you have specified a valid value for this parameter, or consult the NOMAD documentation for more information."
        End If
    Next Key
        
    If InStrText(message, "invalid parameter") Then
        Err.Raise OpenSolver_NomadError, Description:="One of the parameters supplied to NOMAD was invalid. This usually happens if the precision is too large. Try adjusting the values in the Solve Options dialog box."
    End If

ExitSub:
    Close #3
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Sub

ErrorHandler:
    If Not ReportError("CSolverNomad", "CheckLog") Then Resume
    RaiseError = True
    GoTo ExitSub
End Sub

