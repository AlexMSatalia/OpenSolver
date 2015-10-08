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

Private Const pLibName = "OpenSolverNomad"
#If Mac Then
    Private Const pLibBinary = "OpenSolverNomad"
#Else
    Private Const pLibBinary = "OpenSolverNomad.dll"
#End If

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
    #If Mac Then
        ISolverLocal_Version = Version_Mac()
    #Else
        ISolverLocal_Version = Version_Win()
    #End If
End Property

Private Property Get ISolverLocal_Bitness() As String
    SolverIsPresent Me, Bitness:=ISolverLocal_Bitness
End Property

' ISolverLocalLib

Private Function ISolverLocalLib_GetLibPath(Optional errorString As String, Optional Bitness As String) As String
    ISolverLocalLib_GetLibPath = SolverLibPath(Me, errorString, Bitness)
End Function

Private Property Get ISolverLocalLib_LibName() As String
    ISolverLocalLib_LibName = pLibName
End Property

Private Property Get ISolverLocalLib_LibBinary() As String
    ISolverLocalLib_LibBinary = pLibBinary
End Property

Private Property Get ISolverLocalLib_LibVersion() As String
    #If Mac Then
        ISolverLocalLib_LibVersion = LibVersion_Mac()
    #Else
        ISolverLocalLib_LibVersion = LibVersion_Win()
    #End If
End Property

Private Sub ISolverLocalLib_Solve(s As COpenSolver)
    Solve s
End Sub

' Methods

Private Function Version_Win() As String
    If Not SolverIsPresent(Me) Then
        Version_Win = ""
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
    sNomadVersion = Left(Replace(sNomadVersion, vbNullChar, ""), 5)
    
    SetCurrentDirectory currentDir
    
    Version_Win = sNomadVersion
    Exit Function

ErrorHandler:
    Version_Win = ""
End Function

Private Function Version_Mac() As String
    Dim SolverPath As String
    If Not SolverIsPresent(Me, SolverPath) Then
        Version_Mac = ""
        Exit Function
    End If
    
    Version_Mac = ExecCapture(MakePathSafe(SolverPath) & " -nv")
End Function

Private Function LibVersion_Win() As String
    If Not SolverIsPresent(Me) Then
        LibVersion_Win = ""
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
    sLibVersion = Left(Replace(sLibVersion, vbNullChar, ""), 5)
    
    SetCurrentDirectory currentDir
    
    LibVersion_Win = sLibVersion
End Function

Private Function LibVersion_Mac() As String
    Dim SolverPath As String
    If Not SolverIsPresent(Me, SolverPath) Then
        LibVersion_Mac = ""
        Exit Function
    End If
    
    LibVersion_Mac = ExecCapture(MakePathSafe(SolverPath) & " -v")
End Function

Private Sub Solve(s As COpenSolver)
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    Dim ScreenStatus As Boolean
    ScreenStatus = Application.ScreenUpdating
    Application.ScreenUpdating = GetShowSolverProgress()

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
    
    #If Mac Then
        s.SolveStatus = Pending
        Dim LocalLibSolver As ISolverLocalLib
        Set LocalLibSolver = Me
        Shell LocalLibSolver.GetLibPath()
    #Else
        ' We need to call NomadMain directly rather than use Application.Run .
        ' Using Application.Run causes the API calls inside the DLL to fail on 64 bit Office
        Dim NomadRetVal As Long
        NomadRetVal = NomadMain(s.SolveRelaxation)
        NOMAD_LoadResult NomadRetVal
        s.SolveStatus = OS.SolveStatus
        s.SolveStatusComment = OS.SolveStatusComment
    #End If
    
ExitSub:
    #If Mac Then
    #Else
        SetCurrentDirectory currentDir
        Application.Calculate
        Set SolverNOMAD.OS = Nothing
    #End If
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Sub

ErrorHandler:
    If Not ReportError("CSolverNomad", "Solve") Then Resume
    RaiseError = True
    GoTo ExitSub

End Sub

