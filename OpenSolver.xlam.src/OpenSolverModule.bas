Attribute VB_Name = "OpenSolverModule"
' OpenSolver
' Copyright Andrew Mason 2010
' http://www.OpenSolver.org
' This software is distributed under the terms of the GNU General Public License
'
' Where marked, portions of this code have been sourced from online sources with no explicit license given.
'
' This file is part of OpenSolver.
'
' OpenSolver is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' OpenSolver is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with OpenSolver.  If not, see <http://www.gnu.org/licenses/>.
'
'
' OpenSolver v0.8
'
' v0.2: Switched to Application.Calculation = manual; it is twice as fast. Looping thru the LHS and RHS ranges is an insignificant time when compared to the calculation time
'       Eg, Run time with no LHS and RHS range loopups is 4.2s, this goes to 4.4 or 4.5 when we loop thru the LHS and RHS ranges
'       Note: Reading cell values one by one is very slow.
'       Instead, see:
'          http://www.xtremevbtalk.com/showthread.php?t=296858
'          http://www.avdf.com/apr98/art_ot003.html
'          http://www.food-info.net/uk/e/e173.htm - very good info on writing fast code
'          http://blogs.msdn.com/excel/archive/2009/03/12/excel-vba-performance-coding-best-practices.aspx - fast coding
'          http://msdn.microsoft.com/en-us/library/aa730921.aspx - microsoft info on Excel 2007 what's new
'          http://support.microsoft.com/kb/153090/EN-US/ - pass an Excel array to VB
'          http://support.microsoft.com/kb/177991 - limitations when passing arrays to sheets

Option Explicit
Option Base 1

Public Const EPSILON = 0.000001
Public Const ZERO = 0.00000001

'Solution results, as reported by Excel Solver
' FROM http://msdn.microsoft.com/en-us/library/ff197237.aspx
' 0 Solver found a solution. All constraints and optimality conditions are satisfied.
' 1 Solver has converged to the current solution. All constraints are satisfied.
' 2 Solver cannot improve the current solution. All constraints are satisfied.
' 3 Stop chosen when the maximum iteration limit was reached.
' 4 The Objective Cell values do not converge.
' 5 Solver could not find a feasible solution.
' 6 Solver stopped at user's request.
' 7 The linearity conditions required by this LP Solver are not satisfied.
' 8 The problem is too large for Solver to handle.
' 9 Solver encountered an error value in a target or constraint cell.
' 10 Stop chosen when the maximum time limit was reached.
' 11 There is not enough memory available to solve the problem.
' 13 Error in model. Please verify that all cells and constraints are valid.
' 14 Solver found an integer solution within tolerance. All constraints are satisfied.
' 15 Stop chosen when the maximum number of feasible [integer] solutions was reached.
' 16 Stop chosen when the maximum number of feasible [integer] subproblems was reached.
' 17 Solver converged in probability to a global solution.
' 18 All variables must have both upper and lower bounds.
' 19 Variable bounds conflict in binary or alldifferent constraint.
' 20 Lower and upper bounds on variables allow no feasible solution.

' -----------------------------------------------------------------------------
' OpenSolver results, as given by OpenSolver.SolveStatus
' See also OpenSolver.SolveStatusString, which gives a slightly more detailed text summary
' and OpenSolver.SolveStatusComment, for any detailed comment on an infeasible problem
Enum OpenSolverResult
   AbortedThruUserAction = -3   ' Used to indicate that a non-linearity check was made (losing the solution)
   ErrorOccurred = -2    ' Added by us - used in RunOpenSolver to indicate an error occured and has been reported to the user
   Unsolved = -1        ' Added by us - indicates a model not yet solved
   Optimal = 0
   Unbounded = 4        ' "The Objective Cell values do not converge" => unbounded
   Infeasible = 5
   TimeLimitedSubOptimal = 10    ' CBC stopped before finding an optimal/feasible/integer solution because of CBC errors or time/iteration limits
   NotLinear = 7 ' Report non-linearity so that it can be picked up in silent mode
   ' ErrorInTargetOrConstraint = 9  We throw an error instead
   ' ErrorInModel = 13 We throw an error instead
   ' IntegerOptimal = 14 We just return Optimal
End Enum

' OpenSolver Solver Types
Enum OpenSolver_SolverType
    Unknown = -1
    Linear = 1
    NonLinear = 2
End Enum

'--------------------------------------------------------------------------------
' OpenSolver solve results as read directly from CBC, and given by OpenSolver.CBCSolveStatus
' See also OpenSolver.CBCSolveStatusString for a direct text equivalent
' and OpenSolver.CBCSolutionWasLoaded, which is true if any CBC result (suboptimal or infeasible) was loaded back into the sheet
Public Enum LinearSolveResult
    Unsolved = -1
    Optimal = 1
    Infeasible = 2
    Unbounded = 3
    SolveStopped = 4
    IntegerInfeasible = 5
End Enum

Public Const ModelStatus_Unitialized = 0
Public Const ModelStatus_Built = 1

' OpenSolver error numbers.
Public Const OpenSolver_BuildError = vbObjectError + 1000 ' An error occured while building the model
Public Const OpenSolver_SolveError = vbObjectError + 1001 ' An error occured while solving the model
Public Const OpenSolver_UserCancelledError = vbObjectError + 1002 ' The user cancelled the model build or the model solve
Public Const OpenSolver_CBCMissingError = vbObjectError + 1003  ' We cannot find the CBC.exe file
Public Const OpenSolver_CBCExecutionError = vbObjectError + 1004 ' Something went wrong trying to run CBC

Public Const OpenSolver_NoWorksheet = vbObjectError + 1010 ' There is no active workbook
Public Const OpenSolver_NoWorkbook = vbObjectError + 1011  ' There is no active worksheet

Public Const OpenSolver_NomadError = vbObjectError + 1012 ' An error occured while running Nomad non-linear solver

'Gurobi return status codes
Public Enum GurobiResult
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

' Solver's different types of constraints
Public Enum RelationConsts
    RelationLE = 1
    RelationEQ = 2
    RelationGE = 3
    RelationINT = 4
    RelationBIN = 5
    RelationAllDiff = 6
End Enum

Public Enum ObjectiveSenseType
   UnknownObjectiveSense = 0
   MaximiseObjective = 1
   MinimiseObjective = 2
   TargetObjective = 3   ' Seek a specific value
End Enum

Public Enum VariableType
   VarContinuous = 0
   VarInteger = 1
   VarBinary = 2
End Enum

Public Type SolveOptionsType
    maxTime As Double ' "MaxTime"=Max run time in seconds
    MaxIterations As Long ' "Iterations" = max number of branch and bound nodes?
    Precision As Double ' ???
    Tolerance As Double ' Tolerance, being allowable percentage gap. NB: Solver shows this as a percentage, but stores it as a value, eg 1% is stored as 0.01
    ' Convergence As Double   ' Convergence, being ??
    ShowIterationResults As Boolean   ' Excel stores ...!solver_sho=1 if Show Iteration Results is turned on, 2 if off (NB: Not 0!)
End Type

' This name is used to define a table of parameters that get changed between successive solves using QuickSolve
Const ParamRangeName As String = "OpenSolverModelParameters"


'CACHE for SearchRange - Saves defined names from user
Private SearchRangeNameCACHE As Collection  'by ASL 20130126

#If Mac Then
    Public Const ScriptExtension = ".sh"
    Public Const NBSP = 202 ' ascii char code for non-breaking space on Mac
#Else
    Public Const ScriptExtension = ".bat"
    Public Const NBSP = 160 ' ascii char code for non-breaking space on Windows
#End If

Public TempFolderPathCached As String

#If Mac Then
    ' Variables for caching error messages on Mac when control passes between Class and Module
    Public OpenSolver_ErrNumber As Long
    Public OpenSolver_ErrSource As String
    Public OpenSolver_ErrDescription As String
    
    ' Variable to save Drive name
    Public CachedDriveName As String
#End If

' TODO: These & other declarations, and type definitons, need to be updated for 64 bit systems; see:
'   http://msdn.microsoft.com/en-us/library/ee691831.aspx
'   http://technet.microsoft.com/en-us/library/ee833946.aspx
#If Win32 Then
    #If VBA7 Then
        Private Declare PtrSafe Function GetTempPath Lib "kernel32" _
        Alias "GetTempPathA" (ByVal nBufferLength As Long, _
        ByVal lpBuffer As String) As Long
    #Else
        Private Declare Function GetTempPath Lib "kernel32" _
        Alias "GetTempPathA" (ByVal nBufferLength As Long, _
        ByVal lpBuffer As String) As Long
    #End If
#End If

#If Win32 Then
    #If VBA7 Then
        Declare PtrSafe Function SetCurrentDirectoryA Lib "kernel32" _
        (ByVal lpPathName As String) As Long
    #Else
        Declare Function SetCurrentDirectoryA Lib "kernel32" _
        (ByVal lpPathName As String) As Long
    #End If
#End If

'***************** Code Start ******************
'This code was originally written by Terry Kreft.
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
'Code Courtesy of
'Terry Kreft
Private Const STARTF_USESHOWWINDOW& = &H1
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&


#If VBA7 Then

    Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Long
        cbReserved2 As Long
        lpReserved2 As Long
        hStdInput As LongPtr
        hStdOutput As LongPtr
        hStdError As LongPtr
    End Type
    
    Private Type PROCESS_INFORMATION
        hProcess As LongPtr
        hThread As LongPtr
        dwProcessID As Long
        dwThreadID As Long
    End Type
    
    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal _
    hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
    
    Private Declare PtrSafe Function CreateProcessA Lib "kernel32" (ByVal _
        lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
        lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
        ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
        ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
        lpStartupInfo As STARTUPINFO, lpProcessInformation As _
        PROCESS_INFORMATION) As Long
    
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal _
        hObject As LongPtr) As Long
#Else

    Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Long
        cbReserved2 As Long
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
    End Type
    
    Private Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessID As Long
        dwThreadID As Long
    End Type

    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
    hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
    lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
    lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As _
    PROCESS_INFORMATION) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal _
    hObject As Long) As Long
#End If
'***************** Code End Terry Kreft ****************
#If VBA7 Then
    Declare PtrSafe Function TerminateProcess Lib "kernel32.dll" (ByVal ApphProcess As LongPtr, _
    ByVal uExitCode As Long) As Long
#Else
    Declare Function TerminateProcess Lib "kernel32.dll" (ByVal ApphProcess As Long, _
    ByVal uExitCode As Long) As Long
#End If

' For ShowWindow
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2

'Code Courtesy of Dev Ashish
#If VBA7 Then
    Private Declare PtrSafe Function apiShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" _
        (ByVal hwnd As LongPtr, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) _
        As Long
#Else
    Private Declare Function apiShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" _
        (ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) _
        As Long
#End If

Public Const WIN_NORMAL = 1         'Open Normal
Public Const WIN_MAX = 2            'Open Maximized
Public Const WIN_MIN = 3            'Open Minimized

Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&

#If VBA7 Then
   Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As LongPtr, lpExitCode As Long) As Long
#Else
   Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
#End If

#If VBA7 Then
  Private Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As LongPtr) As Long
#Else
  Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
#End If

'=====================================================================

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

#If Mac Then
    Public Declare Function system Lib "libc.dylib" (ByVal cammand As String) As Long
#End If


'***************** Code Start ******************
'This code was originally written by Terry Kreft.
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
'Code Courtesy of
'Terry Kreft
' Modified by A Mason
Function RunExternalCommand(CommandString As String, Optional logPath As String, Optional WindowStyle As Long, Optional WaitForCompletion As Boolean, Optional userCancelled As Boolean, Optional exeResult As Long) As Boolean
#If Mac Then
          If WindowStyle = SW_HIDE Then
              Dim ret As Long
26            ret = system(CommandString & IIf(logPath <> "", " > " & logPath, ""))
27            If ret = 0 Then RunExternalCommand = True
          Else
              Dim CommandToRun As String
              ' Applescript escapes double quotes with a backslash
              CommandToRun = Replace(CommandString, """", "\""")
              ' Pipe through tee if we are logging
              If logPath <> "" Then
                  CommandToRun = CommandToRun & " | tee " & Replace(logPath, """", "\""")
              End If
          
              ' Applescript for opening a terminal to run our command
              ' 1. Create window if terminal not already open, then activate window
              ' 2. Run our shell command in the terminal, saving a reference to the open window
              ' 3. Loop until the window is no longer busy
              Dim script As String
              script = _
                  "tell application ""Terminal""" & vbNewLine & _
                  "    if not (exists window 1) then reopen" & vbNewLine & _
                  "    activate" & vbNewLine & _
                  "    set w to do script """ & CommandToRun & """ in window 1" & vbNewLine & _
                  "    repeat" & vbNewLine & _
                  "        delay 1" & vbNewLine & _
                  "        if not busy of w then exit repeat" & vbNewLine & _
                  "    end repeat" & vbNewLine & _
                  "end tell"
              MacScript (script)
          End If
              
#Else
      'TODO: Optional for Boolean doesn't seem to work IsMissing is always false and value is false?
      ' Returns true if successful completion, false if escape was pressed
28        RunExternalCommand = False
29        userCancelled = False
30        exeResult = -1
          Dim proc As PROCESS_INFORMATION
          Dim start As STARTUPINFO
          Dim ret As Long
          ' Initialize the STARTUPINFO structure:
31        With start
32            .cb = Len(start)
33        If Not isMissing(WindowStyle) Then
34            .dwFlags = STARTF_USESHOWWINDOW
35            .wShowWindow = WindowStyle
36        End If
37        End With

          If logPath <> "" Then
              If WindowStyle = SW_HIDE Then
                  logPath = " > " & logPath
              Else
                  Dim mteePath As String
                  GetExistingFilePathName JoinPaths(JoinPaths(ThisWorkbook.Path, "Utils"), "mtee"), "mtee.exe", mteePath
                  logPath = " | " & MakePathSafe(mteePath) & " " & logPath
              End If
          End If

          ' Start the shelled application:
38        ret& = CreateProcessA(0&, CommandString & logPath, 0&, 0&, 1&, _
                                NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
39        If ret& = 0 Then
41            Err.Raise Number:=OpenSolver_CBCExecutionError, Source:="OpenSolver", _
              Description:="Unable to run the external program: " & CommandString & ". " & vbCrLf & vbCrLf _
              & "Error " & Err.LastDllError & ": " & DLLErrorText(Err.LastDllError)
42        End If
43        If Not isMissing(WaitForCompletion) Then
44            If Not WaitForCompletion Then GoTo ExitSuccessfully
45        End If
          
          ' Wait for the shelled application to finish:
          ' Allow the user to cancel the run. Pressing ESC seems to be well detected with this loop structure
          ' if the new process is hidden; if it is just minimized, then Escape does not seem to be well detected.
          'TODO: Put up a modal dialog for long runs....
46        On Error GoTo errorHandler
47        Do
              ' ret& = WaitForSingleObject(proc.hProcess, INFINITE)
48            ret& = WaitForSingleObject(proc.hProcess, 50) ' Wait for up to 50 milliseconds
              ' Application.CheckAbort  ' We don't need this as the escape key already causes any error
49        Loop Until ret& <> 258

          ' Get the return code for the executable; http://msdn.microsoft.com/en-us/library/windows/desktop/ms683189%28v=vs.85%29.aspx
          Dim lExitCode As Long
50        If GetExitCodeProcess(proc.hProcess, lExitCode) = 0 Then GoTo DLLErrorHandler
51        If Not isMissing(exeResult) Then
52            exeResult = lExitCode
53        End If

ExitSuccessfully:
54        RunExternalCommand = True
          
ExitSub:
55        On Error Resume Next
56        ret& = CloseHandle(proc.hProcess)
57        Exit Function
          
errorHandler:
          Dim ErrorNumber As Long, ErrorDescription As String, ErrorSource As String
58        ErrorNumber = Err.Number
59        ErrorDescription = Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
60        ErrorSource = Err.Source
          
61        If Err.Number = 18 Then
              ' Firstly show the CBC
              ' m_dwSirenProcessID = proc.dwProcessID;
              ' hWnd = GetWindowHandle(m_dwSirenProcessID); enumerates windows, using GetWindowThreadProcessId
              ' ::ShowWindowAsync(hWnd,sw_WindowState);
              ' See http://www.everythingaccess.com/tutorials.asp?ID=Bring-an-external-application-window-to-the-foreground
              '     for an example of finding a given running application's window

              Dim f As UserForm
#If Mac Then
62            Set f = New MacUserFormInterrupt
#Else
63            Set f = New UserFormInterrupt
#End If
64            Application.Cursor = xlDefault
65            f.Show
              'If msgbox("You have pressed the Escape key. Do you wish to cancel?", _
                         vbQuestion + vbYesNo + vbDefaultButton1, _
                         "OpenSolver: User Interrupt Occured...") = vbNo Then
              Dim StopSolving As Boolean
66            StopSolving = f.Tag = vbCancel
67            Unload f
68            Application.Cursor = xlWait
69            If Not StopSolving Then
70                Resume 'continue on from where error occured
71            Else
                  ' Kill CBC (if it is still running?)
72                TerminateProcess proc.hProcess, 0   ' Give an exit code of 0?
73                userCancelled = True
74                Resume ExitSub
75            End If
76        End If
          
77        On Error Resume Next
78        ret& = CloseHandle(proc.hProcess)
79        Err.Raise ErrorNumber, "OpenSolver RunExternalCommand", ErrorDescription
80        Exit Function
DLLErrorHandler:
81        On Error Resume Next
82        ret& = CloseHandle(proc.hProcess)
83        Err.Raise Err.LastDllError, "OpenSolver OSSolverSync", DLLErrorText(Err.LastDllError) & IIf(Erl = 0, "", " (at line " & Erl & ")")
#End If
End Function
'***************** Code End Terry Kreft ****************

' From http://stackoverflow.com/questions/1439200/vba-shell-and-wait-with-exit-code
Public Function DLLErrorText(ByVal lLastDLLError As Long) As String
          Dim sBuff As String * 256
          Dim lCount As Long
          Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100, FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
          Const FORMAT_MESSAGE_FROM_HMODULE = &H800, FORMAT_MESSAGE_FROM_STRING = &H400
          Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000, FORMAT_MESSAGE_IGNORE_INSERTS = &H200
          Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

84        lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
85        If lCount Then
86            DLLErrorText = left$(sBuff, lCount - 2) ' Remove line feeds
87        End If

End Function


Function GetTempFolder() As String

88        If TempFolderPathCached = "" Then
#If Mac Then
89      TempFolderPathCached = MacScript("return (path to temporary items) as string")
#Else
          'Get Temp Folder
          ' See http://www.pcreview.co.uk/forums/thread-934893.php
          Dim fctRet As Long
90        TempFolderPathCached = String$(255, 0)
91        fctRet = GetTempPath(255, TempFolderPathCached)
92        If fctRet <> 0 Then
93            TempFolderPathCached = left(TempFolderPathCached, fctRet)
94            If right(TempFolderPathCached, 1) <> "\" Then TempFolderPathCached = TempFolderPathCached & "\"
95        Else
96            TempFolderPathCached = ""
97        End If
#End If
        '  NEW CODE 2013-01-22 - Andres Sommerhoff (ASL) - Country: Chile
        '  Use Environment Var to have the option to a different Temp path for Opensolver.
        '  To allow have independent configuration in different computers, Environment Var
        '  is used instead of saving the option in the excel.
        '  This also work as workaround to avoid problem with spaces in the temp path.
98        If Environ("OpenSolverTempPath") <> "" Then
99              TempFolderPathCached = Environ("OpenSolverTempPath")
100       End If
        '  ASL END NEW CODE

101       End If
102       GetTempFolder = TempFolderPathCached
103       Exit Function
End Function

Function GetTempFilePath(FileName As String) As String
104       GetTempFilePath = GetTempFolder & FileName
End Function

Function FileOrDirExists(pathName As String) As Boolean
           ' from http://www.vbaexpress.com/kb/getarticle.php?kb_id=559
           'Macro Purpose: Function returns TRUE if the specified file
           '               or folder exists, false if not.
           'PathName     : Supports Windows mapped drives or UNC
           '             : Supports Macintosh paths
           'File usage   : Provide full file path and extension
           'Folder usage : Provide full folder path
           '               Accepts with/without trailing "\" (Windows)
           '               Accepts with/without trailing ":" (Macintosh)
           
          Dim iTemp As Long
           
           'Ignore errors to allow for error evaluation
105       On Error Resume Next
106       iTemp = GetAttr(pathName)
           
           'Check if error exists and set response appropriately
107       Select Case Err.Number
          Case Is = 0
108           FileOrDirExists = True
109       Case Else
110           FileOrDirExists = False
111       End Select
           
           'Resume error checking
112       On Error GoTo 0
End Function

Function GetParamRangeName() As String
113       GetParamRangeName = ParamRangeName
End Function

Function JoinPaths(Path1 As String, Path2 As String) As String
114       JoinPaths = Path1
115       If right(" " & JoinPaths, 1) <> Application.PathSeparator Then JoinPaths = JoinPaths & Application.PathSeparator
116       JoinPaths = JoinPaths & Path2
End Function

'Function GetNameRefersTo(TheName As String) As String
    ' See http://www.cpearson.com/excel/DefinedNames.aspx
'    Dim s As String
'    Dim HasRef As Boolean
'    Dim r As Range
'    Dim NM As Name
'    Set NM = ThisWorkbook.Names(TheName)
'    On Error Resume Next
'    Set r = NM.RefersToRange
'    If Err.Number = 0 Then
'        HasRef = True
'    Else
'        HasRef = False
'    End If
'    On Error GoTo 0
'    If HasRef Then
'        s = r.Text
'    Else
'        s = NM.RefersTo
'        If StrComp(Mid(s, 2, 1), Chr(34), vbBinaryCompare) = 0 Then
'            ' text constant
'            s = Mid(s, 3, Len(s) - 3)
'        Else
'            ' numeric contant (AJM: or formula)
'            s = Mid(s, 2)
'        End If
'    End If
'    GetNameRefersTo = s
'End Function

Function GetNameValueIfExists(w As Workbook, theName As String, ByRef value As String) As Boolean
          ' See http://www.cpearson.com/excel/DefinedNames.aspx
          Dim s As String
          Dim HasRef As Boolean
          Dim r As Range
          Dim NM As Name
          
117       On Error Resume Next
118       Set NM = w.Names(theName)
119       If Err.Number <> 0 Then ' Name does not exist
120           value = ""
121           GetNameValueIfExists = False
122           Exit Function
123       End If
          
124       On Error Resume Next
125       Set r = NM.RefersToRange
126       If Err.Number = 0 Then
127           HasRef = True
128       Else
129           HasRef = False
130       End If
131       If HasRef = True Then
132           value = r.value
133       Else
134           s = NM.RefersTo
135           If StrComp(Mid(s, 2, 1), Chr(34), vbBinaryCompare) = 0 Then
                  ' text constant
136               value = Mid(s, 3, Len(s) - 3)
137           Else
                  ' numeric contant (AJM: or Formula)
138               value = Mid(s, 2)
139           End If
140       End If
141       GetNameValueIfExists = True
End Function

'Function NamedRangeExists(Name As String) As Boolean
'    Dim r As Range
'    On Error Resume Next
'    r = Names(Name).value
'    NamedRangeExists = (Err.Number = 0)
'End Function

Function NameExistsInWorkbook(book As Workbook, Name As String) As Boolean
          ' WARNING: If the name has a sheet prefix, eg Sheet1!OpenSolverCBCParameters, then this will NOT find the range
          ' if the range has been defined globally (which happens when the user defines a name if that name exists only once)
          Dim o As Object
142       On Error Resume Next
143       Set o = book.Names(Name)
144       NameExistsInWorkbook = (Err.Number = 0)
End Function

Function GetNameRefersToIfExists(book As Workbook, Name As String, RefersTo As String) As Boolean
          ' WARNING: If the name has a sheet prefix, eg Sheet1!OpenSolverCBCParameters, then this will NOT find the range
          ' if the range has been defined globally (which happens when the user defines a name if that name exists only once)
145       On Error Resume Next
146       RefersTo = book.Names(Name).RefersTo
147       GetNameRefersToIfExists = (Err.Number = 0)
End Function

Function GetNamedRangeIfExists(book As Workbook, Name As String, r As Range) As Boolean
          ' WARNING: If the name has a sheet prefix, eg Sheet1!OpenSolverCBCParameters, then this will NOT find the range
          ' if the range has been defined globally (which happens when the user defines a name if that name exists only once)
148       On Error Resume Next
149       Set r = book.Names(Name).RefersToRange
150       GetNamedRangeIfExists = (Err.Number = 0)
End Function

Function GetNamedRangeIfExistsOnSheet(sheet As Worksheet, Name As String, r As Range) As Boolean
          ' This finds a named range (either local or global) if it exists, and if it refers to the specified sheet.
          ' It will not find a globally defined name
          ' GetNamedRangeIfExistsOnSheet = False
151       On Error Resume Next
152       Set r = sheet.Range(Name)   ' This will return either a local or globally defined named range, that must refer to the specified sheet. OTherwise there is an error
153       GetNamedRangeIfExistsOnSheet = Err.Number = 0
          ' If r.Worksheet.Name <> Sheet.Name Then Exit Function
          ' GetNamedRangeIfExistsOnSheet = True
End Function

Function GetNamedNumericValueIfExists(book As Workbook, Name As String, value As Double) As Boolean
          ' Get a named range that must contain a double value or the form "=12.34" or "=12" etc, with no spaces
          Dim isRange As Boolean, r As Range, RefersToFormula As Boolean, RangeRefersToError As Boolean, RefersTo As String, isMissing As Boolean
154       GetNameAsValueOrRange book, Name, isMissing, isRange, r, RefersToFormula, RangeRefersToError, RefersTo, value
155       GetNamedNumericValueIfExists = Not isMissing And Not isRange And Not RefersToFormula And Not RangeRefersToError
End Function

Function GetNamedIntegerIfExists(book As Workbook, Name As String, IntegerValue As Long) As Boolean
          ' Get a named range that must contain an integer value
          Dim value As Double
156       If GetNamedNumericValueIfExists(book, Name, value) Then
157           IntegerValue = Int(value)
158           GetNamedIntegerIfExists = IntegerValue = value
159       Else
160           GetNamedIntegerIfExists = False
161       End If
End Function

Function GetNamedStringIfExists(book As Workbook, Name As String, value As String) As Boolean
          ' Get a named range that must contain a string value (probably with quotes)
162       If GetNameRefersToIfExists(book, Name, value) Then
163           If left(value, 2) = "=""" Then ' Remove delimiters and equals in: ="...."
164               value = Mid(value, 3, Len(value) - 3)
165           ElseIf left(value, 1) = "=" Then
166               value = Mid(value, 2)
167           End If
168           GetNamedStringIfExists = True
169       Else
170           GetNamedStringIfExists = False
171       End If
End Function

Sub GetNameAsValueOrRange(book As Workbook, theName As String, isMissing As Boolean, isRange As Boolean, r As Range, RefersToFormula As Boolean, RangeRefersToError As Boolean, RefersTo As String, value As Double)
          ' See http://www.cpearson.com/excel/DefinedNames.aspx, but see below for internationalisation problems with this code
172       RangeRefersToError = False
173       RefersToFormula = False
          ' Dim r As Range
          Dim NM As Name
174       On Error Resume Next
175       Set NM = book.Names(theName)
176       If Err.Number <> 0 Then
177           isMissing = True
178           Exit Sub
179       End If
180       isMissing = False
181       On Error Resume Next
182       Set r = NM.RefersToRange
183       If Err.Number = 0 Then
184           isRange = True
185       Else
186           isRange = False
187       End If
188       If Not isRange Then
              ' String will be of form: "=5", or "=Sheet1!#REF!" or "=Test4!$M$11/4+Test4!$A$3"
189           RefersTo = Mid(NM.RefersTo, 2)
190           If right(RefersTo, 6) = "!#REF!" Then
191               RangeRefersToError = True
192           Else
              ' If StrComp(Mid(S, 2, 1), Chr(34), vbBinaryCompare) = 0 Then
                  ' text constant
              '    S = Mid(S, 3, Len(S) - 3)
              'Else
                  ' numeric contant (or possibly a string? We ignore strings - Solver rejects them as invalid on entry)
                  ' The following Pearson code FAILS because "Value=RefersTo" applies regional settings, but Names are always stored as strings containing values in US settings (with no regionalisation)
                  ' value = RefersTo
                  ' If Err.Number = 13 Then
                  '    RefersToFormula = True
                  ' End If
                  
                  ' Test for a numeric constant, in US format
193               If IsAmericanNumber(RefersTo) Then
194                   value = Val(RefersTo)   ' Force a conversion to a number using Val which uses US settings (no regionalisation)
195               Else
196                   RefersToFormula = True
197               End If
198           End If
199       End If
End Sub

Function GetDisplayAddress(r As Range, Optional showRangeName As Boolean = False) As String
             ' Get a name to display for this range which includes a sheet name if this range is not on the active sheet
              Dim s As String
              Dim R2 As Range
              Dim Rname As Name
              Dim i As Long
          
              'Find if the range has a defined name
200           If r.Worksheet.Name = ActiveSheet.Name Then
201               GetDisplayAddress = r.Address
202               If showRangeName Then
203                   Set Rname = SearchRangeInVisibleNames(r)
204                   If Not Rname Is Nothing Then
205                       GetDisplayAddress = StripWorksheetNameAndDollars(Rname.Name, ActiveSheet)
206                   End If
207               End If
208               Exit Function
209           End If

              ' We first attempt converting without quoting the worksheet name
210           On Error GoTo Try2
211           Set R2 = r.Areas(1)
212           s = R2.Worksheet.Name & "!" & R2.Address
213           If showRangeName Then
214               Set Rname = SearchRangeInVisibleNames(R2)
215               If Not Rname Is Nothing Then
216                   s = R2.Worksheet.Name & "!" & StripWorksheetNameAndDollars(Rname.Name, R2.Worksheet)
217               End If
218           End If

              Dim pre As String
              ' Conversion must also work with multiple areas, eg: A1,B5 converts to Sheet1!A1,Sheet1!B5
219           For i = 2 To r.Areas.Count
220               Set R2 = r.Areas(i)
221               pre = R2.Worksheet.Name & "!" & R2.Address
222               If showRangeName Then
223                   Set Rname = SearchRangeInVisibleNames(R2)
224                   If Not Rname Is Nothing Then
225                       pre = R2.Worksheet.Name & "!" & StripWorksheetNameAndDollars(Rname.Name, R2.Worksheet)
226                   End If
227               End If
228               s = s & "," & pre
229           Next i
230           Set R2 = Range(s) ' Check it has worked!
231           GetDisplayAddress = s
232           Exit Function

Try2:
              ' We now try with quotes around the worksheet name
              ' TODO: This can probably be done more efficiently!
              ' Note that we need to double any single quotes in the name to double quotes in the process (2012.10.29)
233           On Error GoTo 0 ' Turn back on error handling; a failure now shoudl throw an error
234           Set R2 = r.Areas(1)
235           s = "'" & Replace(R2.Worksheet.Name, "'", "''") & "'!" & R2.Address
236           If showRangeName Then
237               Set Rname = SearchRangeInVisibleNames(R2)
238               If Not Rname Is Nothing Then
239                   s = "'" & Replace(R2.Worksheet.Name, "'", "''") & "'!" & StripWorksheetNameAndDollars(Rname.Name, R2.Worksheet)
240               End If
241           End If
              ' Conversion must also work with multiple areas, eg: A1,B5 converts to Sheet1!A1,Sheet1!B5
242           For i = 2 To r.Areas.Count
243               Set R2 = r.Areas(i)
244               pre = "'" & Replace(R2.Worksheet.Name, "'", "''") & "'!" & R2.Address
245               If showRangeName Then
246                   Set Rname = SearchRangeInVisibleNames(R2)
247                   If Not Rname Is Nothing Then
248                       pre = "'" & Replace(R2.Worksheet.Name, "'", "''") & "'!" & StripWorksheetNameAndDollars(Rname.Name, R2.Worksheet)
249                   End If
250               End If
251               s = s & "," & pre
252           Next i
253           Set R2 = Range(s) ' Check it has worked!
              'Show the proper sheet name without the doubled quotes
              's = Replace(s, "''", "'")
254           GetDisplayAddress = s
255           Exit Function
End Function

Function GetDisplayAddressInCurrentLocale(r As Range) As String
      ' Get a name to display for this range which includes a sheet name if this range is not on the active sheet
          Dim s As String, R2 As Range
256       If r.Worksheet.Name = ActiveSheet.Name Then
257           GetDisplayAddressInCurrentLocale = r.AddressLocal
258           Exit Function
259       End If
260       On Error GoTo Try2
          Dim i As Long
261       Set R2 = r.Areas(1)
262       s = R2.Worksheet.Name & "!" & R2.AddressLocal
          ' Conversion must also work with multiple areas, eg: A1,B5 converts to Sheet1!A1,Sheet1!B5
263       For i = 2 To r.Areas.Count
264          Set R2 = r.Areas(i)
265          s = s & Application.International(xlListSeparator) & R2.Worksheet.Name & "!" & R2.AddressLocal
266       Next i
267       Set R2 = Range(ConvertFromCurrentLocale(s)) ' Check it has worked!
268       GetDisplayAddressInCurrentLocale = s
269       Exit Function
Try2:
270       On Error GoTo 0 ' Turn back on error handling; a failure now should throw an error
271       Set R2 = r.Areas(1)
272       s = "'" & Replace(R2.Worksheet.Name, "'", "''") & "'!" & R2.Address ' NB: We jhave to double any single quotes when we quote the name
          ' Conversion must also work with multiple areas, eg: A1,B5 converts to Sheet1!A1,Sheet1!B5
273       For i = 2 To r.Areas.Count
274          Set R2 = r.Areas(i)
275          s = s & Application.International(xlListSeparator) & "'" & Replace(R2.Worksheet.Name, "'", "''") & "'!" & R2.AddressLocal
276       Next i
277       Set R2 = Range(ConvertFromCurrentLocale(s)) ' Check it has worked!
278       GetDisplayAddressInCurrentLocale = s
279       Exit Function
End Function

Function RemoveActiveSheetNameFromString(s As String) As String
          ' Try the active sheet name in quotes first
          Dim sheetName As String
280       sheetName = "'" & Replace(ActiveSheet.Name, "'", "''") & "'!" ' We double any single quotes when we quote the name
281       If InStr(s, sheetName) Then
282           RemoveActiveSheetNameFromString = Replace(s, sheetName, "")
283           Exit Function
284       End If
285       sheetName = ActiveSheet.Name & "!"
286       If InStr(s, sheetName) Then
287           RemoveActiveSheetNameFromString = Replace(s, sheetName, "")
288           Exit Function
289       End If
290       RemoveActiveSheetNameFromString = s
End Function

Function ConvertFromCurrentLocale(ByVal s As String) As String
          ' Convert a formula or a range from the current locale into US locale
          ' This will add a leading "=" if its not already there
          ' A blank string is returned if any errors occur
          ' This works by putting the expression into cell A1 on Sheet1 of the add-in!
          ' We turn off calculation & hide alerts as we don't want Excel popping up dialogs asking for references to other sheets
                   
          Dim oldCalculation As Long
291       oldCalculation = Application.Calculation
          Dim oldDisplayAlerts As Boolean
292       oldDisplayAlerts = Application.DisplayAlerts
293       On Error GoTo errorHandler
294       s = Trim(s)
          Dim equalsAdded As Boolean
295       If left(s, 1) <> "=" Then
296           s = "=" & s
297           equalsAdded = True
298       End If
299       Application.Calculation = xlCalculationManual
300       Application.DisplayAlerts = False
               
          
          ThisWorkbook.Sheets(1).Cells(1, 1).FormulaLocal = s
          On Error GoTo DecimalFixer

302       s = ThisWorkbook.Sheets(1).Cells(1, 1).Formula
303       If equalsAdded Then
304           If left(s, 1) = "=" Then s = Mid(s, 2)
305       End If
306       ConvertFromCurrentLocale = s
          ThisWorkbook.Sheets(1).Cells(1, 1).Clear
308       Application.Calculation = oldCalculation
309       Application.DisplayAlerts = oldDisplayAlerts
310       Exit Function
errorHandler:
        
          ThisWorkbook.Sheets(1).Cells(1, 1).Clear
312       Application.Calculation = oldCalculation
313       Application.DisplayAlerts = oldDisplayAlerts
314       ConvertFromCurrentLocale = ""

DecimalFixer: 'Ensures decimal character used is correct.
        If Application.DecimalSeparator = "." Then
            s = Replace(s, ".", ",")
            ThisWorkbook.Sheets(1).Cells(1, 1).FormulaLocal = s
        ElseIf Application.DecimalSeparator = "," Then
            s = Replace(s, ",", ".")
            ThisWorkbook.Sheets(1).Cells(1, 1).FormulaLocal = s
        End If
        Resume

End Function

Function ConvertToCurrentLocale(ByVal s As String) As String
          ' Convert a formula or a range from the current locale into US locale
          ' This will add a leading "=" if its not already there
          ' A blank string is returned if any errors occur
          ' This works by putting the expression into cell A1 on Sheet1 of the add-in; crude but seems to work
          ' We turn off calculation & hide alerts as we don't want Excel popping up dialogs asking for references to other sheets
          Dim oldCalculation As Long
315       oldCalculation = Application.Calculation
          Dim oldDisplayAlerts As Boolean
316       oldDisplayAlerts = Application.DisplayAlerts
317       On Error GoTo errorHandler
318       s = Trim(s)
          Dim equalsAdded As Boolean
319       If left(s, 1) <> "=" Then
320           s = "=" & s
321           equalsAdded = True
322       End If
323       Application.Calculation = xlCalculationManual
324       Application.DisplayAlerts = False
325       ThisWorkbook.Sheets(1).Cells(1, 1).Formula = s
326       s = ThisWorkbook.Sheets(1).Cells(1, 1).FormulaLocal
327       If equalsAdded Then
328           If left(s, 1) = "=" Then s = Mid(s, 2)
329       End If
330       ConvertToCurrentLocale = s
331       ThisWorkbook.Sheets(1).Cells(1, 1).Clear
332       Application.Calculation = oldCalculation
333       Exit Function
errorHandler:
334       ThisWorkbook.Sheets(1).Cells(1, 1).Clear
335       Application.DisplayAlerts = oldDisplayAlerts
336       Application.Calculation = oldCalculation
337       ConvertToCurrentLocale = ""
End Function

Function ValidLPFileVarName(s As String)
      ' http://lpsolve.sourceforge.net/5.5/CPLEX-format.htm
      'The letter E or e, alone or followed by other valid symbols, or followed by another E or e, should be avoided as this notation is reserved for exponential entries. Thus, variables cannot be named e9, E-24, E8cats, or other names that could be interpreted as an exponent. Even variable names such as eels or example can cause a read error, depending on their placement in an input line.
338       If left(s, 1) = "E" Then
339           ValidLPFileVarName = "_" & s
340       Else
341           ValidLPFileVarName = s
342       End If
End Function

'Function FullLPFileVarName(cell As Range, AdjCellsSheetIndex As Long)
' NO LONGER USED
' Get a valid name for the LP variable of the form A1_2 meaing cell A1 on the 2nd worksheet,
' or _E1 meaning cell E1 on the 'default' worksheet. We need to prefix E with _ to be safe; otherwise it can clash with exponential notation
' http://lpsolve.sourceforge.net/5.5/CPLEX-format.htm
'The letter E or e, alone or followed by other valid symbols, or followed by another E or e, should be avoided as this notation is reserved for exponential entries. Thus, variables cannot be named e9, E-24, E8cats, or other names that could be interpreted as an exponent. Even variable names such as eels or example can cause a read error, depending on their placement in an input line.
'    Dim sheetIndex As Long, s As String
'    sheetIndex = cell.Worksheet.Index
'    s = cell.Address(False, False)
'    If left(s, 1) = "E" Then s = "_" & s
'    If sheetIndex <> AdjCellsSheetIndex Then s = s & "_" & str(sheetIndex)
'    FullLPFileVarName = s
'End Function

'Function ConvertFullLPFileVarNameToRange(s As String, AdjCellsSheetIndex As Long) As Range
' COnvert an encoded LP variable name back into a range on the appropriate sheet
''    Dim i As Long, sheetIndex As Long
'    If left(s, 1) = "_" Then s = Mid(s, 2) ' Remove any protective initial _ for addresses starting E
'    i = InStr(1, s, "_")
'    If i = 0 Then
'        sheetIndex = AdjCellsSheetIndex
'    Else
'        sheetIndex = Val(Mid(s, i + 1))
'        s = left(s, i - 1)
'    End If
'    Set ConvertFullLPFileVarNameToRange = Worksheets(sheetIndex).Range(s)
'End Function

Function SolverRelationAsUnicodeChar(rel As Long) As String
343       Select Case rel
              Case RelationGE
344               SolverRelationAsUnicodeChar = ChrW(&H2265) ' ">" gg
345           Case RelationEQ
346               SolverRelationAsUnicodeChar = "="
347           Case RelationLE
348               SolverRelationAsUnicodeChar = ChrW(&H2264) ' "<"
349           Case Else
350               SolverRelationAsUnicodeChar = "(unknown)"
351       End Select
End Function

'Function SolverRelationAsChar(rel As Long) As String
'1740      Select Case rel
'              Case RelationGE
'1750              SolverRelationAsChar = ">" ' ChrW(&H2265) ' ">" gg
'1760          Case RelationEQ
'1770              SolverRelationAsChar = "="
'1780          Case RelationLE
'1790              SolverRelationAsChar = "<" ' ChrW(&H2264) ' "<"
'1800          Case Else
'1810              SolverRelationAsChar = "(unknown)"
'1820      End Select
'End Function

Function SolverRelationAsString(rel As Long) As String
352       Select Case rel
              Case RelationGE
353               SolverRelationAsString = ">="
354           Case RelationEQ
355               SolverRelationAsString = "="
356           Case RelationLE
357               SolverRelationAsString = "<="
358           Case Else
359               SolverRelationAsString = "(unknown)"
360       End Select
End Function

Function ReverseRelation(rel As Long) As Long
361       ReverseRelation 4 - rel
End Function

Function UserSetQuickSolveParameterRange() As Boolean
362       UserSetQuickSolveParameterRange = False
363       If Application.Workbooks.Count = 0 Then
364           MsgBox "Error: No active workbook available", , "OpenSolver" & sOpenSolverVersion & " Error"
365           Exit Function
366       End If

          Dim sheetName As String
367       On Error Resume Next
368       sheetName = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!" ' NB: We have to double any ' when we quote the name
369       If Err.Number <> 0 Then
370           MsgBox "Error: Unable to access the active sheet", , "OpenSolver" & sOpenSolverVersion & " Error"
371           Exit Function
372       End If
373       On Error GoTo 0
          
          ' Find the Parameter range
          Dim ParamRange As Range
374       On Error Resume Next
375       Set ParamRange = Range(sheetName & ParamRangeName)
376       On Error GoTo 0
          
          ' Get a range from the user
          Dim NewRange As Range
377       On Error Resume Next
378       If ParamRange Is Nothing Then
379           Set NewRange = Application.InputBox(prompt:="Please select the 'parameter' cells that you will be changing between successsive solves of the model.", Type:=8, title:="OpenSolver Quick Solve Parameters")
380       Else
381           Set NewRange = Application.InputBox(prompt:="Please select the 'parameter' cells that you will be changing between successsive solves of the model.", Type:=8, Default:=ParamRange.Address, title:="OpenSolver Quick Solve Parameters")
382       End If
383       On Error GoTo 0
          
384       If Not NewRange Is Nothing Then
385           If NewRange.Worksheet.Name <> ActiveSheet.Name Then
386               MsgBox "Error: The parameter cells need to be on the current worksheet.", , "OpenSolver" & sOpenSolverVersion & " Error"
387               Exit Function
388           End If
              'On Error Resume Next
389           If Not ParamRange Is Nothing Then
                  ' Name needs to be deleted first
390               ActiveWorkbook.Names(sheetName & ParamRangeName).Delete
391           End If
392           Names.Add Name:=sheetName & ParamRangeName, RefersTo:=NewRange 'ActiveWorkbook.
              ' Return true as we have succeeded
393           UserSetQuickSolveParameterRange = True
394       End If
End Function

Function CheckModelHasParameterRange()
395       If Application.Workbooks.Count = 0 Then
396           MsgBox "Error: No active workbook available", , "OpenSolver" & sOpenSolverVersion & " Error"
397           Exit Function
398       End If

          Dim sheetName As String
399       On Error Resume Next
400       sheetName = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!" ' NB: We have to double any ' when we quote the name
401       If Err.Number <> 0 Then
402           MsgBox "Error: Unable to access the active sheet", , "OpenSolver" & sOpenSolverVersion & " Error"
403           Exit Function
404       End If
405       On Error GoTo 0
          
406       CheckModelHasParameterRange = True
          ' Find the Parameter range
          Dim ParamRange As Range
407       On Error Resume Next
408       Set ParamRange = Range(sheetName & ParamRangeName)
409       If Err.Number <> 0 Then
410           MsgBox "Error: No parameter range could be found on the worksheet. Please use the Initialize Quick Solve Parameters menu item to define the cells that you wish to change between successive OpenSolver solves. Note that changes to these cells must lead to changes in the underlying model's right hand side values for its constraints.", title:="OpenSolver" & sOpenSolverVersion & " Error"
411           CheckModelHasParameterRange = False
412           Exit Function
413       End If
End Function

Sub GetSolveOptions(sheetName As String, SolveOptions As SolveOptionsType, errorString As String)
          ' Get the Solver Options, stored in named ranges with values such as "=0.12"
          ' Because these are NAMEs, they are always in English, not the local language, so get their value using Val
414       On Error GoTo errorHandler
415       errorString = ""
416       SetAnyMissingDefaultExcel2007SolverOptions ' This can happen if they have created the model using an old version of OpenSolver
417       With SolveOptions
418           .maxTime = Val(Mid(Names(sheetName & "solver_tim").value, 2)) ' Trim the "="; use Val to get a conversion in English, not the local language
419           .MaxIterations = Val(Mid(Names(sheetName & "solver_itr").value, 2))
420           .Precision = Val(Mid(Names(sheetName & "solver_pre").value, 2))
421           .Tolerance = Val(Mid(Names(sheetName & "solver_tol").value, 2))  ' Stored as a value between 0 and 1 by Excel's Solver (representing a percentage)
              ' .Convergence = Val(Mid(Names(SheetName & "solver_cvg").Value, 2)) NOT USED BY OPEN SOLVER, YET!
              ' Excel stores ...!solver_sho=1 if Show Iteration Results is turned on, 2 if off (NB: Not 0!)
422           .ShowIterationResults = Names(sheetName & "solver_sho").value = "=1"
423       End With
ExitSub:
424       Exit Sub
errorHandler:
425       errorString = "No Solve options (such as Tolerance) could be found - perhaps a model has not been defined on this sheet?"
End Sub

Sub SetAnyMissingDefaultExcel2007SolverOptions()
          ' We set all the default values, as per Solver in Excel 2007, but with some changes. This ensures Solver does not delete the few values we actually use
426       If ActiveWorkbook Is Nothing Then Exit Sub
427       If ActiveSheet Is Nothing Then Exit Sub
          Dim s As String
428       If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_drv", s) Then SetSolverNameOnSheet "drv", "=1"
429       If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_est", s) Then SetSolverNameOnSheet "est", "=1"
430       If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_itr", s) Then SetSolverNameOnSheet "itr", "=100" ' OpenSolver ignores this
431       If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_neg", s) Then SetSolverNameOnSheet "neg", "=1"  ' Not "=2" as we want >=0 constraints
432       If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_num", s) Then SetSolverNameOnSheet "num", "=0"
433       If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_nwt", s) Then SetSolverNameOnSheet "nwt", "=1"
434       If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_pre", s) Then SetSolverNameOnSheet "pre", "=0.000001"
435       If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_scl", s) Then SetSolverNameOnSheet "scl", "=2"
436       If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_sho", s) Then SetSolverNameOnSheet "sho", "=2"
437       If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_tim", s) Then SetSolverNameOnSheet "tim", "=9999999999"  ' not "=100" as we want longer runs; Solver will force this to be no more than 9999
438       If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_tol", s) Then SetSolverNameOnSheet "tol", "=0.05"
439       If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_typ", s) Then SetSolverNameOnSheet "typ", "=1"
440       If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_val", s) Then SetSolverNameOnSheet "val", "=0"
441       If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_cvg", s) Then SetSolverNameOnSheet "cvg", "=0.0001"  ' probably not needed, but set it to be safe
End Sub

'Code Courtesy of
'Dev Ashish
Public Function fHandleFile(stFile As String, lShowHow As Long)
      Dim lRet As Long, varTaskID As Variant
      Dim stRet As String
          Dim hwnd
          ' Dim StartDoc
          ' hwnd = apiFindWindow("OPUSAPP", "0")
          'First try ShellExecute
442       lRet = apiShellExecute(hwnd, vbNullString, _
                  stFile, vbNullString, vbNullString, lShowHow)
                  
443       If lRet > ERROR_SUCCESS Then
444           stRet = vbNullString
445           lRet = -1
446       Else
447           Select Case lRet
                  Case ERROR_NO_ASSOC:
                      'Try the OpenWith dialog
448                   varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " _
                              & stFile, WIN_NORMAL)
449                   lRet = (varTaskID <> 0)
450               Case ERROR_OUT_OF_MEM:
451                   stRet = "Error: Out of Memory/Resources. Couldn't Execute!"
452               Case ERROR_FILE_NOT_FOUND:
453                   stRet = "Error: File not found.  Couldn't Execute!"
454               Case ERROR_PATH_NOT_FOUND:
455                   stRet = "Error: Path not found. Couldn't Execute!"
456               Case ERROR_BAD_FORMAT:
457                   stRet = "Error:  Bad File Format. Couldn't Execute!"
458               Case Else:
459           End Select
460       End If
461       fHandleFile = lRet & IIf(stRet = "", vbNullString, ", " & stRet)
End Function

Function GetExistingFilePathName(Directory As String, FileName As String, ByRef pathName As String) As Boolean
462      pathName = JoinPaths(Directory, FileName)
463      GetExistingFilePathName = FileOrDirExists(pathName)
End Function

Function CheckWorksheetAvailable(Optional SuppressDialogs As Boolean = False, Optional ThrowError As Boolean = False) As Boolean
464       CheckWorksheetAvailable = False
          ' Check there is a workbook
465       If Application.Workbooks.Count = 0 Then
466           If ThrowError Then Err.Raise Number:=OpenSolver_NoWorkbook, Source:="OpenSolver", Description:="No active workbook available."
467           If Not SuppressDialogs Then MsgBox "Error: No active workbook available", , "OpenSolver" & sOpenSolverVersion & " Error"
468           Exit Function
469       End If
          ' Check we can access the worksheet
          Dim w As Worksheet
470       On Error Resume Next
471       Set w = ActiveWorkbook.ActiveSheet
472       If Err.Number <> 0 Then
473           If ThrowError Then Err.Raise Number:=OpenSolver_NoWorksheet, Source:="OpenSolver", Description:="The active sheet is not a worksheet."
474           If Not SuppressDialogs Then MsgBox "Error: The active sheet is not a worksheet.", , "OpenSolver" & sOpenSolverVersion & " Error"
475           Exit Function
476       End If
          ' OK
477       CheckWorksheetAvailable = True
End Function

Function GetOneCellInRange(r As Range, instance As Long) As Range
          ' Given an 'instance' between 1 and r.Count, return the instance'th cell in the range, where our count goes cross each row in turn (as does 'for each in range')
          Dim RowOffset As Long, ColOffset As Long
          Dim NumCols As Long
          ' Debug.Assert r.Areas.count = 1
478       NumCols = r.Columns.Count
479       RowOffset = ((instance - 1) \ NumCols)
480       ColOffset = ((instance - 1) Mod NumCols)
481       Set GetOneCellInRange = r.Cells(1 + RowOffset, 1 + ColOffset)
End Function

Function Max_Int(a As Long, b As Long) As Long
482       If a > b Then
483           Max_Int = a
484       Else
485           Max_Int = b
486       End If
End Function

Function Max_Double(a As Double, b As Double) As Double
487       If a > b Then
488           Max_Double = a
489       Else
490           Max_Double = b
491       End If
End Function

Function Create1x1Array(X As Variant) As Variant
          ' Create a 1x1 array containing the value x
          Dim v(1, 1) As Variant
492       v(1, 1) = X
493       Create1x1Array = v
End Function

Function ForceCalculate(prompt As String, Optional MinimiseUserInteraction As Boolean = False) As Boolean
           'There appears to be a bug in Excel 2010 where the .Calculate does not always complete. We handle up to 3 such failures.
           ' We have seen this problem arise on large models.
#If Mac Then
          'In Excel 2011 the Application.CalculationState is not included:
          'http://sysmod.wordpress.com/2011/10/24/more-differences-mainly-vba/
          'Try calling 'Calculate' two times just to be safe? This will probably cause problems down the line, maybe Office 2014 will fix it?
494       Application.Calculate
495       Application.Calculate
496       ForceCalculate = True
#Else
497       Application.Calculate
498       If Application.CalculationState <> xlDone Then
499           Application.Calculate
              Dim i As Long
500           For i = 1 To 10
501               DoEvents
502               Sleep (100)
503           Next i
504       End If
505       If Application.CalculationState <> xlDone Then Application.Calculate
506       If Application.CalculationState <> xlDone Then
507           DoEvents
508           Application.CalculateFullRebuild
509           DoEvents
510       End If
          
          ' Check for circular references causing problems, which can happen if iterative calculation mode is enabled.
511       If Application.CalculationState <> xlDone Then
512           If Application.Iteration Then
513               If MinimiseUserInteraction Then
514                   Application.Iteration = False
515                   Application.Calculate
516               ElseIf MsgBox("Iterative calculation mode is enabled and may be interfering with the inital calculation. Would you like to try disabling iterative calculation mode to see if this fixes the problem?", _
                            vbYesNo, _
                            "OpenSolver: Iterative Calculation Mode Detected...") = vbYes Then
517                   Application.Iteration = False
518                   Application.Calculate
519               End If
520           End If
521       End If
          
522       While Application.CalculationState <> xlDone
523           If MinimiseUserInteraction Then
524               ForceCalculate = False
525               Exit Function
526           ElseIf MsgBox(prompt, _
                          vbCritical + vbRetryCancel + vbDefaultButton1, _
                          "OpenSolver: Calculation Error Occured...") = vbCancel Then
527               ForceCalculate = False
528               Exit Function
529           Else 'Recalculate the workbook if the user wants to retry
530               Application.Calculate
531           End If
532       Wend
533       ForceCalculate = True
#End If
End Function

Function ProperUnion(R1 As Range, R2 As Range) As Range
          ' Return the union of r1 and r2, where r1 may be Nothing
          ' TODO: Handle the fact that Union will return a range with multiple copies of overlapping cells - does this matter?
534       If R1 Is Nothing Then
535           Set ProperUnion = R2
536       ElseIf R2 Is Nothing Then
537           Set ProperUnion = R1
538       ElseIf R1 Is Nothing And R2 Is Nothing Then
539           Set ProperUnion = Nothing
540       Else
541           Set ProperUnion = Union(R1, R2)
542       End If
End Function

Function GetRangeValues(r As Range) As Variant()
          ' This copies the values from a possible multi-area range into a variant
          Dim v() As Variant, i As Long
543       ReDim v(r.Areas.Count)
544       For i = 1 To r.Areas.Count
545           v(i) = r.Areas(i).Value2 ' Copy the entire area into the i'th entry of v
546       Next i
547       GetRangeValues = v
End Function

Sub SetRangeValues(r As Range, v() As Variant)
          ' This copies the values from a variant into a possibly multi-area range; see GetRangeValues
          Dim i As Long
548       For i = 1 To r.Areas.Count
549           r.Areas(i).Value2 = v(i)
550       Next i
End Sub

Function MergeRangesCellByCell(R1 As Range, R2 As Range) As Range
          ' This merges range r2 into r1 cell by cell.
          ' This shoulsd be fastest if range r2 is smaller than r1
          Dim result As Range, cell As Range
551       Set result = R1
552       For Each cell In R2
553           Set result = Union(result, cell)
554       Next cell
555       Set MergeRangesCellByCell = result
End Function

Function RemoveRangeOverlap(r As Range) As Range
          ' This creates a new range from r which does not contain any multiple repetitions of cells
          ' This works around the fact that Excel allows range like "A1:A2,A2:A3", which has a .count of 4 cells
          ' The Union function does NOT remove all overlaps; call this after the union to
556       If r.Areas.Count = 1 Then
557           Set RemoveRangeOverlap = r
558           Exit Function
559       End If
          Dim s As Range, i As Long
560       Set s = r.Areas(1)
561       For i = 2 To r.Areas.Count
562           If Intersect(s, r.Areas(i)) Is Nothing Then
                  ' Just take the standard union
563               Set s = Union(s, r.Areas(i))
564           Else
                  ' Merge these two ranges cell by cell; this seems to remove the overlap in my tests, but also see http://www.cpearson.com/excel/BetterUnion.aspx
                  ' Merge the smaller range into the larger
565               If s.Count < r.Areas(i).Count Then
566                   Set s = MergeRangesCellByCell(r.Areas(i), s)
567               Else
568                   Set s = MergeRangesCellByCell(s, r.Areas(i))
569               End If
570           End If
571       Next i
572       Set RemoveRangeOverlap = s
End Function

Function CheckRangeContainsNoAmbiguousMergedCells(r As Range, BadCell As Range) As Boolean
          ' This checks that if the range contains any merged cells, those cells are the 'home' cell (top left) in the merged cell block
          ' and thus references to these cells are indeed to a unique cell
          ' If we have a cell that is not the top left of a merged cell, then this will be read as blank, and writing to this will effect other cells.
573       CheckRangeContainsNoAmbiguousMergedCells = True
574       If Not r.MergeCells Then
575           Exit Function
576       End If
          Dim cell As Range
577       For Each cell In r
578           If cell.MergeCells Then
579               If cell.Address <> cell.MergeArea.Cells(1, 1).Address Then
580                   Set BadCell = cell
581                   CheckRangeContainsNoAmbiguousMergedCells = False
582                   Exit Function
583               End If
584           End If
585       Next cell
End Function

Function StripWorksheetNameAndDollars(s As String, currentSheet As Worksheet) As String
          ' Remove the current worksheet name from a formula, along with any $
          ' Shorten the formula (eg Test4!$M$11/4+Test4!$A$3) by removing the current sheet name and all $
586       StripWorksheetNameAndDollars = Replace(s, currentSheet.Name & "!", "")   ' Remove names like Test4!
587       StripWorksheetNameAndDollars = Replace(StripWorksheetNameAndDollars, "'" & Replace(currentSheet.Name, "'", "''") & "'!", "") ' Remove names with spaces that are quoted, like 'Test 4'!; we have to double any ' when we quote the name
588       StripWorksheetNameAndDollars = Replace(StripWorksheetNameAndDollars, "$", "")
End Function

' If a key doesn't exist we have to add it, otherwise we just set it
' Note: Numeric values should be passed as strings in English (not the local language)
Sub SetSolverNameOnSheet(Name As String, value As String)
589       Name = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!solver_" + Name ' NB: We have to double any ' when we quote the name
    On Error GoTo doesntExist:
590       Names(Name).value = value
591       Exit Sub
doesntExist:
592       Names.Add Name, value, False
End Sub

' NB: This is a different functiom to SetSolverNameOnSheet as we want to pass a range (on an arbitrary sheet)
'     If we use a variant,it fails as passing a range may simply pass its cell value
' If a key doesn't exist we have to add it, otherwise we just set it
' Solver stores names like Sheet1!$A$1; the sheet name is always given
Sub SetSolverNamedRangeOnSheet(Name As String, value As Range)
593       Name = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!solver_" + Name ' NB: We have to double any ' when we quote the name
    On Error GoTo doesntExist:
594       Names(Name).value = "=" & GetDisplayAddress(value, False) ' Cannot simply assign Names(name).Value=Value as this assigns the value in a single cell, not its address
595       Exit Sub
doesntExist:
596       Names.Add Name, "=" & GetDisplayAddress(value, False), False ' GetDisplayAddress(value), False
End Sub

Sub DeleteSolverNameOnSheet(Name As String)
597       Name = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!solver_" + Name ' NB: We have to double any ' when we quote the name
598       On Error Resume Next
599       Names(Name).Delete
doesntExist:
End Sub

' If a key doesn't exist we have to add it, otherwise we just set it
' Note: Numeric values should be passed as strings in English (not the local language)
Sub SetNameOnSheet(Name As String, value As String)
600       Name = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!" + Name ' NB: We have to double any ' when we quote the name
    On Error GoTo doesntExist:
601       Names(Name).value = value
602       Exit Sub
doesntExist:
603       Names.Add Name, value, False
End Sub

' NB: Simply using a variant in SetSolverNameOnSheet fails as passing a range can simply pass its cell value
' If a key doesn't exist we have to add it, otherwise we just set it
' Note: Numeric values should be passed as strings in English (not the local language)
Sub SetNamedRangeOnSheet(Name As String, value As Range)
604       Name = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!" + Name ' NB: We have to double any ' when we quote the name
    On Error GoTo doesntExist:
605       Names(Name).value = "=" & GetDisplayAddress(value, False) ' "=" & GetDisplayAddress(Value)
606       Exit Sub
doesntExist:
607       Names.Add Name, "=" & GetDisplayAddress(value, False), False ' "=" & GetDisplayAddress(Value, False), False
End Sub

' If a key doesn't exist we have to add it, otherwise we just set it
' Note: Numeric values should be passed as strings in English (not the local language)
Sub DeleteNameOnSheet(Name As String)
608       Name = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!" + Name '  ' NB: We have to double any ' when we quote the name
609       On Error Resume Next
610       Names(Name).Delete
doesntExist:
End Sub

Function TrimBlankLines(s As String) As String
          ' Remove any blank lines at the beginning or end of s
          Dim Done As Boolean
611       While Not Done
612           If Len(s) < Len(vbNewLine) Then
613               Done = True
614           ElseIf left(s, Len(vbNewLine)) = vbNewLine Then
615              s = Mid(s, 3)
616           Else
617               Done = True
618           End If
619       Wend
620       Done = False
621       While Not Done
622           If Len(s) < Len(vbNewLine) Then
623               Done = True
624           ElseIf right(s, Len(vbNewLine)) = vbNewLine Then
625              s = left(s, Len(s) - 2)
626           Else
627               Done = True
628           End If
629       Wend
630       TrimBlankLines = s
End Function

Function IsAmericanNumber(s As String, Optional i As Long = 1) As Boolean
          ' Check this is a number like 3.45  or +1.23e-34
          ' This does NOT test for regional variations such as 12,34
          ' This code exists because
          '   val("12+3") gives 12 with no error
          '   Assigning a string to a double uses region-specific translation, so x="1,2" works in French
          '   IsNumeric("12,45") is true even on a US English system (and even worse...)
          '   IsNumeric("($1,23,,3.4,,,5,,E67$)")=True! See http://www.eggheadcafe.com/software/aspnet/31496070/another-vba-bug.aspx)

          Dim MustBeInteger As Boolean, SeenDot As Boolean, SeenDigit As Boolean
631       MustBeInteger = i > 1   ' We call this a second time after seeing the "E", when only an int is allowed
632       IsAmericanNumber = False    ' Assume we fail
633       If Len(s) = 0 Then Exit Function ' Not a number
634       If Mid(s, i, 1) = "+" Or Mid(s, i, 1) = "-" Then i = i + 1 ' Skip leading sign
635       For i = i To Len(s)
636           Select Case Asc(Mid(s, i, 1))
              Case Asc("E"), Asc("e")
637               If MustBeInteger Or Not SeenDigit Then Exit Function ' No exponent allowed (as must be a simple integer)
638               IsAmericanNumber = IsAmericanNumber(s, i + 1)   ' Process an int after the E
639               Exit Function
640           Case Asc(".")
641               If SeenDot Then Exit Function
642               SeenDot = True
643           Case Asc("0") To Asc("9")
644               SeenDigit = True
645           Case Else
646               Exit Function   ' Not a valid char
647           End Select
648       Next i
          ' i As Long, AllowDot As Boolean
649       IsAmericanNumber = SeenDigit
End Function

Sub TestIsAmericanNumber()
650       Debug.Assert (IsAmericanNumber("12.34") = True)
651       Debug.Assert (IsAmericanNumber("12.34e3") = True)
652       Debug.Assert (IsAmericanNumber("+12.34e3") = True)
653       Debug.Assert (IsAmericanNumber("-12.34e-3") = True)
654       Debug.Assert (IsAmericanNumber("12.34e") = False)
655       Debug.Assert (IsAmericanNumber("1e") = False)
656       Debug.Assert (IsAmericanNumber("+") = False)
657       Debug.Assert (IsAmericanNumber("+1e-") = False)
658       Debug.Assert (IsAmericanNumber("E1") = False)
659       Debug.Assert (IsAmericanNumber("12.3.4") = False)
660       Debug.Assert (IsAmericanNumber("-") = False)
661       Debug.Assert (IsAmericanNumber("-+3") = False)
End Sub

Sub test()
          Dim r As Range
662       Set r = Range("A1")
663       Debug.Print OpenSolver.GetDisplayAddress(r, False)
End Sub

Function SystemIs64Bit() As Boolean
#If Mac Then
          ' Check bitness of Mac by attempting to load 64-bit kernel
          ' http://macscripter.net/viewtopic.php?pid=137569#p137569
          Dim script As String
664       script = "try" & vbNewLine & _
                 "return ((do shell script ""sysctl -n hw.optional.x86_64"") as integer) as boolean" & vbNewLine & _
             "on error" & vbNewLine & _
                 "return false" & vbNewLine & _
             "end try"
665       SystemIs64Bit = MacScript(script)
#Else
          ' Is true if the Windows system is a 64 bit one
          ' If Not Environ("ProgramFiles(x86)") = "" Then Is64Bit=True, or
          ' Is64bit = Len(Environ("ProgramW6432")) > 0; see:
          ' http://blog.johnmuellerbooks.com/2011/06/06/checking-the-vba-environment.aspx and
          ' http://www.mrexcel.com/forum/showthread.php?542727-Determining-If-OS-Is-32-Bit-Or-64-Bit-Using-VBA and
          ' http://stackoverflow.com/questions/6256140/how-to-detect-if-the-computer-is-x32-or-x64 and
          ' http://msdn.microsoft.com/en-us/library/ms684139%28v=vs.85%29.aspx
666      SystemIs64Bit = Environ("ProgramFiles(x86)") <> ""
#End If
End Function

'Public Function GetDefinedNameFromRange(theSheet As Worksheet, DefinedRange As String) As String
'    ' Given a defined name 'name' that refers to a range, get the name (if any) of this range; otherwise get its RefersTo string, or "" if no name
'    Dim Book As Workbook, NameRange As String
'    Set Book = theSheet.Parent
'    Dim RefersTo As String
'    On Error GoTo RangeNotDefined
'    GetDefinedNameFromRange = DefinedRange
'    Dim n As Name
'    NameRange = "=" & theSheet.Name & "!" & DefinedRange
'    For Each n In Book.Names
'        If n.Visible Then
'            If n.RefersTo = NameRange Then
'                GetDefinedNameFromRange = n.Name
'            End If
'        End If
'    Next
'RangeNotDefined:
'End Function

Function MakeNewSheet(namePrefix As String, sheetName As String) As String
          Dim NeedSheet As Boolean, newSheet As Worksheet, nameSheet As String, i As Long
667       On Error Resume Next
668       Application.ScreenUpdating = False
          Dim s As String, value As String
669       s = Sheets(namePrefix).Name
670       If Err.Number <> 0 Then
671           Set newSheet = Sheets.Add
672           newSheet.Name = namePrefix
673           nameSheet = namePrefix
674           ActiveWindow.DisplayGridlines = False
675       Else
676           Call GetNameValueIfExists(ActiveWorkbook, sheetName & "OpenSolver_UpdateSensitivity", value)
677           If value Then
678               Sheets(namePrefix).Cells.Delete
679               nameSheet = namePrefix
680           Else
681               i = 1
682               Set newSheet = Sheets.Add
683               NeedSheet = True
684               On Error Resume Next
685               While NeedSheet
686                   nameSheet = namePrefix & " " & i
687                   newSheet.Name = nameSheet
688                   If Err.Number = 0 Then NeedSheet = False
689                   i = i + 1
690                   Err.Number = 0
691               Wend
692               ActiveWindow.DisplayGridlines = False
693           End If
694       End If
          
695       MakeNewSheet = nameSheet
696       Application.ScreenUpdating = True
End Function

'==================================================================
'Save visible defined names in book in a cache to find them quickly
'==================================================================

'''''' ASL NEW FUNCTION 2012-01-23 - Andres Sommerhoff
Public Sub SearchRangeName_DestroyCache()
697       If Not SearchRangeNameCACHE Is Nothing Then
698           While SearchRangeNameCACHE.Count > 0
699               SearchRangeNameCACHE.Remove 1
700           Wend
701       End If
702       Set SearchRangeNameCACHE = Nothing
End Sub

'''''' ASL NEW FUNCTION 2012-01-23- Andres Sommerhoff
Private Sub SearchRangeName_LoadCache(sheet As Worksheet)
          Dim TestName As Name
          Dim rComp As Range
          Dim i As Long
          
          'Some checks in case the cache is an obsolete version
          Static LastNamesCount As Long
          Static LastSheetName As String
          Static LastFileName As String
          Dim CurrNamesCount As Long
          Dim CurrSheetName As String
          Dim CurrFileName As String

703       On Error Resume Next
704       CurrNamesCount = sheet.Parent.Names.Count
705       CurrSheetName = sheet.Name
706       CurrFileName = sheet.Parent.Name
          
707       If LastNamesCount <> CurrNamesCount _
             Or LastSheetName <> CurrSheetName _
             Or LastFileName <> CurrFileName Then
708                   SearchRangeName_DestroyCache  'Check confirm it is obsolote. Cache need to redone...
709       End If
              
710       If SearchRangeNameCACHE Is Nothing Then
711           Set SearchRangeNameCACHE = New Collection  'Start building a new Cache
712       Else
713           Exit Sub 'Cache is ok -> return back
714       End If

          'Here the Cache will be filled with visible range names only
715       For i = 1 To ActiveWorkbook.Names.Count
716           Set TestName = ActiveWorkbook.Names(i)
              
717           If TestName.Visible = True Then  'Iterate through the visible names only
                      ' Skip any references to external workbooks
718                   If left$(TestName.RefersTo, 1) = "=" And InStr(TestName.RefersTo, "[") > 1 Then GoTo tryNext
719                   On Error GoTo tryerror
                      'Build the Cache with the range address as key (='sheet1'!$A$1:$B$3)
720                   Set rComp = TestName.RefersToRange
721                   SearchRangeNameCACHE.Add TestName, (rComp.Name)
722                   GoTo tryNext
tryerror:
723                   Resume tryNext
tryNext:
724           End If
725       Next i
          
726       LastNamesCount = CurrNamesCount
727       LastSheetName = CurrSheetName
728       LastFileName = CurrFileName
          
End Sub

' ASL NEW FUNCTION 2012-01-23 - Andres Sommerhoff
Public Function SearchRangeInVisibleNames(r As Range) As Name
729       SearchRangeName_LoadCache r.Parent  'Use a collection as cache. Without cache is a little bit slow.
                                              'To refresh the cache use SearchRangeName_DestroyCache()
730       On Error Resume Next
731       Set SearchRangeInVisibleNames = SearchRangeNameCACHE.Item((r.Name))
          
End Function

Public Sub OpenURL(URL As String)

#If Mac Then
          ' Use applescript to open the webpage
          Dim s As String
732       s = "open location """ + URL + """"
733       MacScript s
#Else
          ' Use windows file handler to open webpage
734       Call fHandleFile(URL, WIN_NORMAL)
#End If

End Sub

Public Sub SetCurrentDirectory(NewPath As String)
#If Mac Then
735       ChDir NewPath
#Else
736       SetCurrentDirectoryA NewPath
#End If
End Sub

Public Function ReplaceDelimitersMac(Path As String) As String
737       ReplaceDelimitersMac = Replace(Path, ":", "/")
End Function

Public Function ConvertHfsPath(Path As String) As String
      ' Any direct file system access (using 'system' or in script files) on Mac requires
      ' that HFS-style paths are converted to normal posix paths. On windows this
      ' function does nothing, so it can safely wrap all file system calls on any platform
      ' Input (HFS path):   "Macintosh HD:Users:jack:filename.txt"
      ' Output (posix path): "/Volumes/Macintosh HD/Users/jack/filename.txt"
#If Mac Then
          ' Check we have an HFS path and not posix
738       If InStr(Path, ":") > 0 Then
              ' Prefix disk name with :Volumes:
739           ConvertHfsPath = ":Volumes:" & Path
              ' Convert path delimiters
740           ConvertHfsPath = ReplaceDelimitersMac(ConvertHfsPath)
741       Else
              ' Path is already posix
742           ConvertHfsPath = Path
743       End If
#Else
744       ConvertHfsPath = Path
#End If
End Function

Public Function GetDriveName() As String
#If Mac Then
    If CachedDriveName = "" Then
        Dim Path As String
        Path = GetTempFolder()
        CachedDriveName = left(Path, InStr(Path, ":") - 1)
    End If
    
    GetDriveName = CachedDriveName
#End If
End Function

Public Function QuotePath(Path As String) As String
          ' Quote path
745       QuotePath = """" & Path & """"
End Function

Public Function MakePathSafe(Path As String) As String
    MakePathSafe = QuotePath(ConvertHfsPath(Path))
End Function

Public Sub CreateScriptFile(ByRef ScriptFilePath As String, FileContents As String, Optional EnableEcho As Boolean)
      ' Create a script file with the specified contents.
746       On Error GoTo ErrHandler
747       Open ScriptFilePath For Output As 1
          
#If Win32 Then
          ' Add echo off for windows
748       If Not EnableEcho Then
749           Print #1, "@echo off" & vbCrLf
750       End If
#End If
751       Print #1, FileContents
752       Close #1
          
          ' Make shell script executable on Mac
#If Mac Then
753       system ("chmod +x " & MakePathSafe(ScriptFilePath))
#End If

754       Exit Sub
          
ErrHandler:
755       Close #1
756       Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Sub

Public Sub DeleteFileAndVerify(FilePath As String, errorPrefix As String, errorDesc As String)
      ' Deletes file and raises error if not successful
757       If FileOrDirExists(FilePath) Then Kill FilePath
758       If FileOrDirExists(FilePath) Then
759           Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:=errorDesc
760       End If
End Sub

Public Sub OpenFile(FilePath As String, notFoundMessage As String)
761       On Error GoTo errorHandler
762       If Not FileOrDirExists(FilePath) Then
763           MsgBox notFoundMessage, , "OpenSolver" & sOpenSolverVersion & " Error"
764       Else
              ' Check that there is no workbook open with the same name
              Dim w As Workbook
765           On Error Resume Next
766           Set w = Workbooks(right(FilePath, InStr(FilePath, Application.PathSeparator)))
767           On Error GoTo errorHandler
768           Workbooks.Open FileName:=FilePath, ReadOnly:=True ' , Format:=Tabs
769       End If
ExitSub:
770       Exit Sub
errorHandler:
771       MsgBox "OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source, , "OpenSolver Code Error"
772       Resume ExitSub
End Sub

'==============================================================================
' ConvertRelationToEnum
' Given the value of a solver_relX Name, pick the equivalent OpenSolver operator
Function ConvertRelationToEnum(ByVal strNameContents As String) As Variant
773       Select Case Mid(strNameContents, 2)
              Case "1": ConvertRelationToEnum = RelationConsts.RelationLE
774           Case "2": ConvertRelationToEnum = RelationConsts.RelationEQ
775           Case "3": ConvertRelationToEnum = RelationConsts.RelationGE
776       End Select
End Function

' WriteToFile
' Writes a string to the given file number, adds a newline, and can easily
' uncomment debug line to print to Immediate if needed. Adds number of spaces to front if specified
Sub WriteToFile(intFileNum As Long, strData As String, Optional numSpaces As Long = 0)
          Dim spaces As String, i As Long
777       spaces = ""
778       For i = 1 To numSpaces
779           spaces = spaces + " "
780       Next i
781       Print #intFileNum, spaces & strData
          'Debug.Print strData
End Sub

'==============================================================================
Function TestIntersect(ByRef R1 As Range, ByRef R2 As Range) As Boolean
          Dim r3 As Range
782       Set r3 = Intersect(R1, R2)
783       TestIntersect = Not (r3 Is Nothing)
          ' Below: a test to see if I could do it faster - I couldn't
          'Dim R1_X1 As Long, R1_Y1 As Long, R1_X2 As Long, R1_Y2 As Long
          'Dim R2_X1 As Long, R2_Y1 As Long, R2_X2 As Long, R2_Y2 As Long
          'R1_X1 = R1.Column
          'R1_X2 = R1_X1 + R1.Columns.Count - 1
          'R2_X1 = R2.Column
          'R2_X2 = R2_X1 + R2.Columns.Count - 1
          'R1_Y1 = R1.Row
          'R1_Y2 = R1_Y2 + R1.Rows.Height - 1
          'R2_Y1 = R2.Row
          'R2_Y2 = R2_Y2 + R2.Rows.Height - 1
          'TestIntersect = _
          '    R1_X1 <= R2_X2 And _  ' Cond A
          '    R1_X2 >= R2_X1 And _  ' Cond B
          '    R1_Y1 <= R2_Y2 And _  ' Cond C
          '    R1_Y2 >= R2_Y1        ' Cond D
End Function

' Replaces all spaces with NBSP char
Function MakeSpacesNonBreaking(Text As String) As String
784       MakeSpacesNonBreaking = Replace(Text, Chr(32), Chr(NBSP))
End Function

' Returns true if a number is zero (within tolerance)
Function IsZero(num As Double) As Boolean
785       If Abs(num) < OpenSolver.EPSILON Then
786           IsZero = True
787       Else
788           IsZero = False
789       End If
End Function

' Update error cache
Sub UpdateErrorCache(ErrorNumber As Long, ErrorSource As String, ErrorDescription As String)
#If Mac Then
    OpenSolver_ErrNumber = ErrorNumber
    OpenSolver_ErrSource = ErrorSource
    OpenSolver_ErrDescription = ErrorDescription
#End If
End Sub

' Clear any cached errors
Sub ResetErrorCache()
#If Mac Then
    OpenSolver_ErrNumber = 0
    OpenSolver_ErrSource = ""
    OpenSolver_ErrDescription = ""
#End If
End Sub

Sub MBox(errorMessage As String, Optional linkTarget As String, Optional linkText As String)
    'This function replaces msgbox for reporting errors, and allows us to do a number of things to improve user feedback when somethign goes wrong.
    
    'The string "Help_" is used to denote a helpful message that guides the user's actions (an "intentional" error), as opposed to an error report, which we didn't expect to happen.
    'If this string is present, line numbers and other identification info are are not added to the error message, otherwise they are.
    
    'Some line numbers need to be added at the source of the error in order to be stored. So we'll strip the help message of line numbers, which have this form: "(at line XXX)"
    
    'The programmer may also wish to provide the user with a help link.
    'If it is an unintended error, the help link is set to the opensolver help page by default, unless the user has entered another link.
    'If it is a help message, a link is only displayed when the programmer enters a link.
    'The linkTarget is stored in the tooltip of the linkLabel attribute of the form. This lets the user see the url before clicking it.
    'If no linkText is supplied, the linkTarget is used as the hyperlinked text.
        
    If InStr(errorMessage, "Help_") Then
        'This is a help message, so strip the Help_ from it
        errorMessage = Replace(errorMessage, "Help_", "")
        
        'Strip the unneeded line numbers from it too.
        Dim linNumStartPos As Integer
        Dim linNumEndPos As Integer
        
        'find line number start and end
        linNumStartPos = InStr(errorMessage, "(at line ")
        ' Sometimes error messages on mac get garbled so this check is needed
        If linNumStartPos > 0 Then
            linNumEndPos = InStr(linNumStartPos, errorMessage, ")")
            'Remove this bit from the string
            errorMessage = left(errorMessage, linNumStartPos - 1) & right(errorMessage, Len(errorMessage) - linNumEndPos)
        End If
    Else
        'this is an error message, so add the line number reporting and other info
        errorMessage = "OpenSolver" & sOpenSolverVersion & " encountered an error:" & vbCrLf & errorMessage & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & vbCrLf & "Source = " & Err.Source & ", ErrNumber=" & Err.Number
        
        'If no link is provided, add the opensolver help link
        If isMissing(linkTarget) Then
            linkTarget = "http://opensolver.org/help/"
            linkText = "OpenSolver Help Forum"
        End If
    End If
    
    'Catching errors
    If isMissing(linkText) Then
        If Not isMissing(linkTarget) Then 'print the url as the link text
            linkText = linkTarget
        Else
            linkText = ""
        End If
    End If
        
    If isMissing(linkTarget) Then
        linkTarget = ""
    End If
    
    ' We need to unlock the textbox before writing to it on Mac
    MessageBox.TextBox1.Locked = False
    MessageBox.TextBox1.Text = errorMessage
    MessageBox.TextBox1.Locked = True
    
    MessageBox.LinkLabel.Caption = linkText
    MessageBox.LinkLabel.ControlTipText = linkTarget
    MessageBox.Show
End Sub

' Case-insensitive InStr helper
Function InStrText(String1 As String, String2 As String)
    InStrText = InStr(1, String1, String2, vbTextCompare)
End Function
