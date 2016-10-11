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

Private Property Get ISolver_SolverType() As OpenSolver_SolverType
1         ISolver_SolverType = pType
End Property

Private Property Get ISolver_IterationLimitName() As String
1         ISolver_IterationLimitName = pIterationLimitName
End Property

Private Property Get ISolver_PrecisionName() As String
1         ISolver_PrecisionName = pPrecisionName
End Property

Private Property Get ISolver_TimeLimitName() As String
1         ISolver_TimeLimitName = pTimeLimitName
End Property

Private Property Get ISolver_ToleranceName() As String
End Property

Private Property Get ISolver_ModelType() As OpenSolver_ModelType
1         ISolver_ModelType = pModelType
End Property

Private Sub ISolver_CleanFiles()
End Sub

Private Sub ISolver_CheckLog(s As COpenSolver)
1         CheckLog s
End Sub

' ISolverLocal

Private Property Get ISolverLocal_Version() As String
    #If Mac Then
1             ISolverLocal_Version = Version_Mac()
    #Else
2             ISolverLocal_Version = Version_Win()
    #End If
End Property

Private Property Get ISolverLocal_Bitness() As String
1         SolverIsPresent Me, Bitness:=ISolverLocal_Bitness
End Property

' ISolverLocalLib

Private Function ISolverLocalLib_GetLibPath(Optional errorString As String, Optional Bitness As String) As String
1         ISolverLocalLib_GetLibPath = SolverLibPath(Me, errorString, Bitness)
End Function

Private Property Get ISolverLocalLib_LibName() As String
1         ISolverLocalLib_LibName = pLibName
End Property

Private Property Get ISolverLocalLib_LibBinary() As String
1         ISolverLocalLib_LibBinary = pLibBinary
End Property

Private Property Get ISolverLocalLib_LibVersion() As String
    #If Mac Then
1             ISolverLocalLib_LibVersion = LibVersion_Mac()
    #Else
2             ISolverLocalLib_LibVersion = LibVersion_Win()
    #End If
End Property

Private Sub ISolverLocalLib_Solve(s As COpenSolver)
1         Solve s
End Sub

' Methods

Private Function Version_Win() As String
1         If Not SolverIsPresent(Me) Then
2             Version_Win = vbNullString
3             Exit Function
4         End If
          
5         On Error GoTo ErrorHandler
             
          ' Set current dir for finding the DLL
          Dim currentDir As String
6         currentDir = CurDir
7         SetCurrentDirectory LibDir()
          
          ' Get version info from DLL
          ' Save to a new string first - modifying the string from the DLL can sometimes crash Excel
          Dim sNomadVersion As String
8         sNomadVersion = NomadVersion()
9         sNomadVersion = Left(Replace(sNomadVersion, vbNullChar, vbNullString), 5)
          
10        SetCurrentDirectory currentDir
          
11        Version_Win = sNomadVersion
12        Exit Function

ErrorHandler:
13        Version_Win = vbNullString
End Function

Private Function Version_Mac() As String
          Dim SolverPath As String
1         If Not SolverIsPresent(Me, SolverPath) Then
2             Version_Mac = vbNullString
3             Exit Function
4         End If
          
5         Version_Mac = ExecCapture(MakePathSafe(SolverPath) & " -nv")
End Function

Private Function LibVersion_Win() As String
1         If Not SolverIsPresent(Me) Then
2             LibVersion_Win = vbNullString
3             Exit Function
4         End If
          
          Dim currentDir As String
          
          ' Set current dir for finding the DLL
5         currentDir = CurDir
6         SetCurrentDirectory LibDir()
          
          ' Get version info from DLL
          ' Save to a new string first - modifying the string from the DLL can sometimes crash Excel
          Dim sLibVersion As String
7         sLibVersion = NomadLibVersion()
8         sLibVersion = Left(Replace(sLibVersion, vbNullChar, vbNullString), 5)
          
9         SetCurrentDirectory currentDir
          
10        LibVersion_Win = sLibVersion
End Function

Private Function LibVersion_Mac() As String
          Dim SolverPath As String
1         If Not SolverIsPresent(Me, SolverPath) Then
2             LibVersion_Mac = vbNullString
3             Exit Function
4         End If
          
5         LibVersion_Mac = ExecCapture(MakePathSafe(SolverPath) & " -v")
End Function

Private Sub Solve(s As COpenSolver)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim ScreenStatus As Boolean
3         ScreenStatus = Application.ScreenUpdating
4         Application.ScreenUpdating = GetShowSolverProgress(s.sheet)
          
          Dim currentDir As String
5         currentDir = CurDir
          
6         SetCurrentDirectory LibDir()

7         Set SolverNOMAD.OS = s
          
    #If Mac Then
8             s.SolveStatus = Pending
              Dim LocalLibSolver As ISolverLocalLib
9             Set LocalLibSolver = Me
10            Shell LocalLibSolver.GetLibPath()
    #Else
              ' We need to call NomadMain directly rather than use Application.Run .
              ' Using Application.Run causes the API calls inside the DLL to fail on 64 bit Office
              Dim NomadRetVal As Long
11            NomadRetVal = NomadMain(s.SolveRelaxation)
12            GetNomadSolveResult NomadRetVal, s
    #End If
          
ExitSub:
    #If Mac Then
    #Else
13            SetCurrentDirectory currentDir
14            Application.Calculate
15            Set SolverNOMAD.OS = Nothing
    #End If
16        If RaiseError Then RethrowError
17        Exit Sub

ErrorHandler:
18        If Not ReportError("CSolverNomad", "Solve") Then Resume
19        RaiseError = True
20        GoTo ExitSub

End Sub

