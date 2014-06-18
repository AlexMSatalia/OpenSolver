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

'Solution results, as reported by Excel Solver
' FROM http://msdn.microsoft.com/en-us/library/ff197237.aspx
' 0 Solver found a solution. All constraints and optimality conditions are satisfied.
' 1 Solver has converged to the current solution. All constraints are satisfied.
' 2 Solver cannot improve the current solution. All constraints are satisfied.
' 3 Stop chosen when the maximum iteration limit was reached.
' 4 The Objective Cell values do not converge.
' 5 Solver could not find a feasible solution.
' 6 Solver stopped at user’s request.
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
   ' NotLinear = 7 We throw an error instead
   ' ErrorInTargetOrConstraint = 9  We throw an error instead
   ' ErrorInModel = 13 We throw an error instead
   ' IntegerOptimal = 14 We just return Optimal
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

Public Const ModelFileName As String = "model.lp"  ' Open Solver writes this file
Public Const XMLFileName As String = "job.xml"  ' Open Solver writes this file
Public Const PuLPFileName As String = "opensolver.py"  ' Open Solver writes this file
Public Const SolutionFileName = "modelsolution.txt"    ' CBC writes this file for us to read back in
Public Const PathDelimeter = "\"
Public Const ExternalSolverExeName As String = "cbc.exe"   ' The Executable to run (with no path)
Public Const ExternalSolverExeName64 As String = "cbc64.exe"   ' The Executable to run (with no path) in 64 bit systems (if it exists)

' TODO: These & other declarations, and type definitons, need to be updated for 64 bit systems; see:
'   http://msdn.microsoft.com/en-us/library/ee691831.aspx
'   http://technet.microsoft.com/en-us/library/ee833946.aspx
#If VBA7 Then
    Private Declare PtrSafe Function GetTempPath Lib "kernel32" _
    Alias "GetTempPathA" (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long
#Else
    Private Declare Function GetTempPath Lib "kernel32" _
    Alias "GetTempPathA" (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long
#End If

#If VBA7 Then
    Declare PtrSafe Function SetCurrentDirectory Lib "kernel32" _
    Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
#Else
    Declare Function SetCurrentDirectory Lib "kernel32" Alias _
    "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
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
        wShowWindow As Integer
        cbReserved2 As Integer
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
        wShowWindow As Integer
        cbReserved2 As Integer
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
'NON-LINEAR XLL FUNCTION - Matthew Milner
#If VBA7 Then
    #If Win64 Then
        Public Declare PtrSafe Function NomadMain Lib "OpenSolverNomadDll64.dll" (ByVal SolveRelaxation As Boolean) As Long
    #Else
        Public Declare PtrSafe Function NomadMain Lib "OpenSolverNomadDll.dll" (ByVal SolveRelaxation As Boolean) As Long
    #End If
#Else
    Public Declare Function NomadMain Lib "OpenSolverNomadDll.dll" (ByVal SolveRelaxation As Boolean) As Long
#End If

'=====================================================================


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
Function OSSolveSync(SolverPath As String, pathName As String, PrintingOptionString As String, logPath As String, Optional WindowStyle As Long, Optional WaitForCompletion As Boolean, Optional userCancelled As Boolean, Optional exeResult As Long) As Boolean
      'TODO: Optional for Boolean doesn't seem to work IsMissing is always false and value is false?
      ' Returns true if successful completion, false if escape was pressed
50        OSSolveSync = False
          userCancelled = False
          exeResult = -1
          Dim proc As PROCESS_INFORMATION
          Dim start As STARTUPINFO
          Dim ret As Long
          ' Initialize the STARTUPINFO structure:
60        With start
70            .cb = Len(start)
80        If Not isMissing(WindowStyle) Then
90            .dwFlags = STARTF_USESHOWWINDOW
100           .wShowWindow = WindowStyle
110       End If
120       End With
          ' Start the shelled application:
130       ret& = CreateProcessA(0&, SolverPath & pathName & PrintingOptionString & logPath, 0&, 0&, 1&, _
                                NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
140       If ret& = 0 Then
              pathName = SolverPath & " " & pathName
150           Err.Raise Number:=OpenSolver_CBCExecutionError, Source:="OpenSolver", _
              Description:="Unable to run the external program: " & pathName & ". " & vbCrLf & vbCrLf _
              & "Error " & Err.LastDllError & ": " & DLLErrorText(Err.LastDllError)
160       End If
170       If Not isMissing(WaitForCompletion) Then
180           If Not WaitForCompletion Then GoTo ExitSuccessfully
190       End If
          
          ' Wait for the shelled application to finish:
          ' Allow the user to cancel the run. Pressing ESC seems to be well detected with this loop structure
          ' if the new process is hidden; if it is just minimized, then Escape does not seem to be well detected.
          'TODO: Put up a modal dialog for long runs....
200       On Error GoTo errorHandler
210       Do
              ' ret& = WaitForSingleObject(proc.hProcess, INFINITE)
220           ret& = WaitForSingleObject(proc.hProcess, 50) ' Wait for up to 50 milliseconds
              ' Application.CheckAbort  ' We don't need this as the escape key already causes any error
230       Loop Until ret& <> 258

          ' Get the return code for the executable; http://msdn.microsoft.com/en-us/library/windows/desktop/ms683189%28v=vs.85%29.aspx
          Dim lExitCode As Long
231       If GetExitCodeProcess(proc.hProcess, lExitCode) = 0 Then GoTo DLLErrorHandler
232       If Not isMissing(exeResult) Then
233           exeResult = lExitCode
234       End If

ExitSuccessfully:
240       OSSolveSync = True
          
ExitSub:
250       On Error Resume Next
260       ret& = CloseHandle(proc.hProcess)
270       Exit Function
          
errorHandler:
          Dim ErrorNumber As Long, ErrorDescription As String, ErrorSource As String
280       ErrorNumber = Err.Number
290       ErrorDescription = Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
300       ErrorSource = Err.Source
          
310       If Err.Number = 18 Then
              ' Firstly show the CBC
              ' m_dwSirenProcessID = proc.dwProcessID;
              ' hWnd = GetWindowHandle(m_dwSirenProcessID); enumerates windows, using GetWindowThreadProcessId
              ' ::ShowWindowAsync(hWnd,sw_WindowState);
              ' See http://www.everythingaccess.com/tutorials.asp?ID=Bring-an-external-application-window-to-the-foreground
              '     for an example of finding a given running application's window

              Dim f As UserFormInterrupt
320           Set f = New UserFormInterrupt
330           Application.Cursor = xlDefault
340           f.Show
              'If msgbox("You have pressed the Escape key. Do you wish to cancel?", _
                         vbQuestion + vbYesNo + vbDefaultButton1, _
                         "OpenSolver: User Interrupt Occured...") = vbNo Then
              Dim StopSolving As Boolean
350           StopSolving = f.Tag = vbCancel
360           Unload f
370           Application.Cursor = xlWait
380           If Not StopSolving Then
390               Resume 'continue on from where error occured
400           Else
                  ' Kill CBC (if it is still running?)
410               TerminateProcess proc.hProcess, 0   ' Give an exit code of 0?
415               userCancelled = True
420               Resume ExitSub
430           End If
440       End If
          
450       On Error Resume Next
460       ret& = CloseHandle(proc.hProcess)
470       Err.Raise ErrorNumber, "OpenSolver OSSolveSync", ErrorDescription
          Exit Function
DLLErrorHandler:
471       On Error Resume Next
472       ret& = CloseHandle(proc.hProcess)
          Err.Raise Err.LastDllError, "OpenSolver OSSolverSync", DLLErrorText(Err.LastDllError) & IIf(Erl = 0, "", " (at line " & Erl & ")")
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

600       lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
610       If lCount Then
620           DLLErrorText = left$(sBuff, lCount - 2) ' Remove line feeds
630       End If

End Function


Function GetTempFolder() As String
          'Get Temp Folder
          ' See http://www.pcreview.co.uk/forums/thread-934893.php
          Dim fctRet As Long
640       GetTempFolder = String$(255, 0)
650       fctRet = GetTempPath(255, GetTempFolder)
660       If fctRet <> 0 Then
670           GetTempFolder = left(GetTempFolder, fctRet)
680           If right(GetTempFolder, 1) <> "\" Then GetTempFolder = GetTempFolder & "\"
690       Else
700           GetTempFolder = ""
710       End If

        '  NEW CODE 2013-01-22 - Andres Sommerhoff (ASL) - Country: Chile
        '  Use Environment Var to have the option to a different Temp path for Opensolver.
        '  To allow have independent configuration in different computers, Environment Var
        '  is used instead of saving the option in the excel.
        '  This also work as workaround to avoid problem with spaces in the temp path.
720       If Environ("OpenSolverTempPath") <> "" Then
730             GetTempFolder = Environ("OpenSolverTempPath")
740       End If
        '  ASL END NEW CODE
End Function

Function GetModelFileName(Optional SolveNEOS As Boolean = False, Optional SolvePulp As Boolean = False) As String
          If SolveNEOS Then
              GetModelFileName = XMLFileName
          ElseIf SolvePulp Then
              GetModelFileName = PuLPFileName
          Else
750           GetModelFileName = ModelFileName
          End If
End Function

Function GetSolutionFileName() As String
760       GetSolutionFileName = SolutionFileName
End Function

Function GetModelFullPath(Optional SolveNEOS As Boolean = False, Optional SolvePulp As Boolean = False) As String
770       GetModelFullPath = GetTempFolder & GetModelFileName(SolveNEOS, SolvePulp)
End Function

Function GetSolutionFullPath() As String
780       GetSolutionFullPath = GetTempFolder & GetSolutionFileName
End Function

Function GetParamRangeName() As String
790       GetParamRangeName = ParamRangeName
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
          
800       On Error Resume Next
810       Set NM = w.Names(theName)
820       If Err.Number <> 0 Then ' Name does not exist
830           value = ""
840           GetNameValueIfExists = False
850           Exit Function
860       End If
          
870       On Error Resume Next
880       Set r = NM.RefersToRange
890       If Err.Number = 0 Then
900           HasRef = True
910       Else
920           HasRef = False
930       End If
940       If HasRef = True Then
950           value = r.value
960       Else
970           s = NM.RefersTo
980           If StrComp(Mid(s, 2, 1), Chr(34), vbBinaryCompare) = 0 Then
                  ' text constant
990               value = Mid(s, 3, Len(s) - 3)
1000          Else
                  ' numeric contant (AJM: or Formula)
1010              value = Mid(s, 2)
1020          End If
1030      End If
1040      GetNameValueIfExists = True
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
1050      On Error Resume Next
1060      Set o = book.Names(Name)
1070      NameExistsInWorkbook = (Err.Number = 0)
End Function

Function GetNameRefersToIfExists(book As Workbook, Name As String, RefersTo As String) As Boolean
          ' WARNING: If the name has a sheet prefix, eg Sheet1!OpenSolverCBCParameters, then this will NOT find the range
          ' if the range has been defined globally (which happens when the user defines a name if that name exists only once)
1080      On Error Resume Next
1090      RefersTo = book.Names(Name).RefersTo
1100      GetNameRefersToIfExists = (Err.Number = 0)
End Function

Function GetNamedRangeIfExists(book As Workbook, Name As String, r As Range) As Boolean
          ' WARNING: If the name has a sheet prefix, eg Sheet1!OpenSolverCBCParameters, then this will NOT find the range
          ' if the range has been defined globally (which happens when the user defines a name if that name exists only once)
1110      On Error Resume Next
1120      Set r = book.Names(Name).RefersToRange
1130      GetNamedRangeIfExists = (Err.Number = 0)
End Function

Function GetNamedRangeIfExistsOnSheet(sheet As Worksheet, Name As String, r As Range) As Boolean
          ' This finds a named range (either local or global) if it exists, and if it refers to the specified sheet.
          ' It will not find a globally defined name
          ' GetNamedRangeIfExistsOnSheet = False
1140      On Error Resume Next
1150      Set r = sheet.Range(Name)   ' This will return either a local or globally defined named range, that must refer to the specified sheet. OTherwise there is an error
1160      GetNamedRangeIfExistsOnSheet = Err.Number = 0
          ' If r.Worksheet.Name <> Sheet.Name Then Exit Function
          ' GetNamedRangeIfExistsOnSheet = True
End Function

Function GetNamedNumericValueIfExists(book As Workbook, Name As String, value As Double) As Boolean
          ' Get a named range that must contain a double value or the form "=12.34" or "=12" etc, with no spaces
          Dim isRange As Boolean, r As Range, RefersToFormula As Boolean, RangeRefersToError As Boolean, RefersTo As String, isMissing As Boolean
1170      GetNameAsValueOrRange book, Name, isMissing, isRange, r, RefersToFormula, RangeRefersToError, RefersTo, value
1180      GetNamedNumericValueIfExists = Not isMissing And Not isRange And Not RefersToFormula And Not RangeRefersToError
End Function

Function GetNamedIntegerIfExists(book As Workbook, Name As String, IntegerValue As Integer) As Boolean
          ' Get a named range that must contain an integer value
          Dim value As Double
1190      If GetNamedNumericValueIfExists(book, Name, value) Then
1200          IntegerValue = Int(value)
1210          GetNamedIntegerIfExists = IntegerValue = value
1220      Else
1230          GetNamedIntegerIfExists = False
1240      End If
End Function

Function GetNamedStringIfExists(book As Workbook, Name As String, value As String) As Boolean
          ' Get a named range that must contain a string value (probably with quotes)
1250      If GetNameRefersToIfExists(book, Name, value) Then
1260          If left(value, 2) = "=""" Then ' Remove delimiters and equals in: ="...."
1270              value = Mid(value, 3, Len(value) - 3)
1280          ElseIf left(value, 1) = "=" Then
1290              value = Mid(value, 2)
1300          End If
1310          GetNamedStringIfExists = True
1320      Else
1330          GetNamedStringIfExists = False
1340      End If
End Function

Sub GetNameAsValueOrRange(book As Workbook, theName As String, isMissing As Boolean, isRange As Boolean, r As Range, RefersToFormula As Boolean, RangeRefersToError As Boolean, RefersTo As String, value As Double)
          ' See http://www.cpearson.com/excel/DefinedNames.aspx, but see below for internationalisation problems with this code
1350      RangeRefersToError = False
1360      RefersToFormula = False
          ' Dim r As Range
          Dim NM As Name
1370      On Error Resume Next
1380      Set NM = book.Names(theName)
1390      If Err.Number <> 0 Then
1400          isMissing = True
1410          Exit Sub
1420      End If
1430      isMissing = False
1440      On Error Resume Next
1450      Set r = NM.RefersToRange
1460      If Err.Number = 0 Then
1470          isRange = True
1480      Else
1490          isRange = False
1500      End If
1510      If Not isRange Then
              ' String will be of form: "=5", or "=Sheet1!#REF!" or "=Test4!$M$11/4+Test4!$A$3"
1520          RefersTo = Mid(NM.RefersTo, 2)
1530          If right(RefersTo, 6) = "!#REF!" Then
1540              RangeRefersToError = True
1550          Else
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
1560              If IsAmericanNumber(RefersTo) Then
1570                  value = Val(RefersTo)   ' Force a conversion to a number using Val which uses US settings (no regionalisation)
1580              Else
1590                  RefersToFormula = True
1600              End If
1610          End If
1620      End If
End Sub

Function GetDisplayAddress(r As Range, Optional showRangeName As Boolean = False) As String
             ' Get a name to display for this range which includes a sheet name if this range is not on the active sheet
              Dim s As String
              Dim R2 As Range
              Dim Rname As Name
              Dim i As Integer
          
              'Find if the range has a defined name
1630          If r.Worksheet.Name = ActiveSheet.Name Then
1640              GetDisplayAddress = r.Address
1650              If showRangeName Then
1660                  Set Rname = SearchRangeInVisibleNames(r)
1670                  If Not Rname Is Nothing Then
1680                      GetDisplayAddress = StripWorksheetNameAndDollars(Rname.Name, ActiveSheet)
1690                  End If
1700              End If
1710              Exit Function
1720          End If

              ' We first attempt converting without quoting the worksheet name
1730          On Error GoTo Try2
1740          Set R2 = r.Areas(1)
1750          s = R2.Worksheet.Name & "!" & R2.Address
1760          If showRangeName Then
1770              Set Rname = SearchRangeInVisibleNames(R2)
1780              If Not Rname Is Nothing Then
1790                  s = R2.Worksheet.Name & "!" & StripWorksheetNameAndDollars(Rname.Name, R2.Worksheet)
1800              End If
1810          End If

              Dim pre As String
              ' Conversion must also work with multiple areas, eg: A1,B5 converts to Sheet1!A1,Sheet1!B5
1820          For i = 2 To r.Areas.Count
1830              Set R2 = r.Areas(i)
1840              pre = R2.Worksheet.Name & "!" & R2.Address
1850              If showRangeName Then
1860                  Set Rname = SearchRangeInVisibleNames(R2)
1870                  If Not Rname Is Nothing Then
1880                      pre = R2.Worksheet.Name & "!" & StripWorksheetNameAndDollars(Rname.Name, R2.Worksheet)
1890                  End If
1900              End If
1910              s = s & "," & pre
1920          Next i
1930          Set R2 = Range(s) ' Check it has worked!
1940          GetDisplayAddress = s
1950          Exit Function

Try2:
              ' We now try with quotes around the worksheet name
              ' TODO: This can probably be done more efficiently!
              ' Note that we need to double any single quotes in the name to double quotes in the process (2012.10.29)
1960          On Error GoTo 0 ' Turn back on error handling; a failure now shoudl throw an error
1970          Set R2 = r.Areas(1)
1980          s = "'" & Replace(R2.Worksheet.Name, "'", "''") & "'!" & R2.Address
1990          If showRangeName Then
2000              Set Rname = SearchRangeInVisibleNames(R2)
2010              If Not Rname Is Nothing Then
2020                  s = "'" & Replace(R2.Worksheet.Name, "'", "''") & "'!" & StripWorksheetNameAndDollars(Rname.Name, R2.Worksheet)
2030              End If
2040          End If
              ' Conversion must also work with multiple areas, eg: A1,B5 converts to Sheet1!A1,Sheet1!B5
2050          For i = 2 To r.Areas.Count
2060              Set R2 = r.Areas(i)
2070              pre = "'" & Replace(R2.Worksheet.Name, "'", "''") & "'!" & R2.Address
2080              If showRangeName Then
2090                  Set Rname = SearchRangeInVisibleNames(R2)
2100                  If Not Rname Is Nothing Then
2110                      pre = "'" & Replace(R2.Worksheet.Name, "'", "''") & "'!" & StripWorksheetNameAndDollars(Rname.Name, R2.Worksheet)
2120                  End If
2130              End If
2140              s = s & "," & pre
2150          Next i
2160          Set R2 = Range(s) ' Check it has worked!
              'Show the proper sheet name without the doubled quotes
              's = Replace(s, "''", "'")
2170          GetDisplayAddress = s
2180          Exit Function
End Function

Function GetDisplayAddressInCurrentLocale(r As Range) As String
      ' Get a name to display for this range which includes a sheet name if this range is not on the active sheet
          Dim s As String, R2 As Range
2190      If r.Worksheet.Name = ActiveSheet.Name Then
2200          GetDisplayAddressInCurrentLocale = r.AddressLocal
2210          Exit Function
2220      End If
2230      On Error GoTo Try2
          Dim i As Integer
2240      Set R2 = r.Areas(1)
2250      s = R2.Worksheet.Name & "!" & R2.AddressLocal
          ' Conversion must also work with multiple areas, eg: A1,B5 converts to Sheet1!A1,Sheet1!B5
2260      For i = 2 To r.Areas.Count
2270         Set R2 = r.Areas(i)
2280         s = s & Application.International(xlListSeparator) & R2.Worksheet.Name & "!" & R2.AddressLocal
2290      Next i
2300      Set R2 = Range(ConvertFromCurrentLocale(s)) ' Check it has worked!
2310      GetDisplayAddressInCurrentLocale = s
2320      Exit Function
Try2:
2330      On Error GoTo 0 ' Turn back on error handling; a failure now should throw an error
2340      Set R2 = r.Areas(1)
2350      s = "'" & Replace(R2.Worksheet.Name, "'", "''") & "'!" & R2.Address ' NB: We jhave to double any single quotes when we quote the name
          ' Conversion must also work with multiple areas, eg: A1,B5 converts to Sheet1!A1,Sheet1!B5
2360      For i = 2 To r.Areas.Count
2370         Set R2 = r.Areas(i)
2380         s = s & Application.International(xlListSeparator) & "'" & Replace(R2.Worksheet.Name, "'", "''") & "'!" & R2.AddressLocal
2390      Next i
2400      Set R2 = Range(ConvertFromCurrentLocale(s)) ' Check it has worked!
2410      GetDisplayAddressInCurrentLocale = s
2420      Exit Function
End Function

Function RemoveActiveSheetNameFromString(s As String) As String
          ' Try the active sheet name in quotes first
          Dim sheetName As String
2430      sheetName = "'" & Replace(ActiveSheet.Name, "'", "''") & "'!" ' We double any single quotes when we quote the name
2440      If InStr(s, sheetName) Then
2450          RemoveActiveSheetNameFromString = Replace(s, sheetName, "")
2460          Exit Function
2470      End If
2480      sheetName = ActiveSheet.Name & "!"
2490      If InStr(s, sheetName) Then
2500          RemoveActiveSheetNameFromString = Replace(s, sheetName, "")
2510          Exit Function
2520      End If
2530      RemoveActiveSheetNameFromString = s
End Function

Function ConvertFromCurrentLocale(ByVal s As String) As String
          ' Convert a formula or a range from the current locale into US locale
          ' This will add a leading "=" if its not already there
          ' A blank string is returned if any errors occur
          ' This works by putting the expression into cell A1 on Sheet1 of the add-in!
          ' We turn off calculation & hide alerts as we don't want Excel popping up dialogs asking for references to other sheets
          Dim oldCalculation As Integer
2540      oldCalculation = Application.Calculation
          Dim oldDisplayAlerts As Boolean
2550      oldDisplayAlerts = Application.DisplayAlerts
2560      On Error GoTo errorHandler
2570      s = Trim(s)
          Dim equalsAdded As Boolean
2580      If left(s, 1) <> "=" Then
2590          s = "=" & s
2600          equalsAdded = True
2610      End If
2620      Application.Calculation = xlCalculationManual
2630      Application.DisplayAlerts = False
2640      ThisWorkbook.Sheets(1).Cells(1, 1).FormulaLocal = s
2650      s = ThisWorkbook.Sheets(1).Cells(1, 1).Formula
2660      If equalsAdded Then
2670          If left(s, 1) = "=" Then s = Mid(s, 2)
2680      End If
2690      ConvertFromCurrentLocale = s
2700      ThisWorkbook.Sheets(1).Cells(1, 1).Clear
2710      Application.Calculation = oldCalculation
2720      Application.DisplayAlerts = oldDisplayAlerts
2730      Exit Function
errorHandler:
2740      ThisWorkbook.Sheets(1).Cells(1, 1).Clear
2750      Application.Calculation = oldCalculation
2760      Application.DisplayAlerts = oldDisplayAlerts
2770      ConvertFromCurrentLocale = ""
End Function

Function ConvertToCurrentLocale(ByVal s As String) As String
          ' Convert a formula or a range from the current locale into US locale
          ' This will add a leading "=" if its not already there
          ' A blank string is returned if any errors occur
          ' This works by putting the expression into cell A1 on Sheet1 of the add-in; crude but seems to work
          ' We turn off calculation & hide alerts as we don't want Excel popping up dialogs asking for references to other sheets
          Dim oldCalculation As Integer
2780      oldCalculation = Application.Calculation
          Dim oldDisplayAlerts As Boolean
2790      oldDisplayAlerts = Application.DisplayAlerts
2800      On Error GoTo errorHandler
2810      s = Trim(s)
          Dim equalsAdded As Boolean
2820      If left(s, 1) <> "=" Then
2830          s = "=" & s
2840          equalsAdded = True
2850      End If
2860      Application.Calculation = xlCalculationManual
2870      Application.DisplayAlerts = False
2880      ThisWorkbook.Sheets(1).Cells(1, 1).Formula = s
2890      s = ThisWorkbook.Sheets(1).Cells(1, 1).FormulaLocal
2900      If equalsAdded Then
2910          If left(s, 1) = "=" Then s = Mid(s, 2)
2920      End If
2930      ConvertToCurrentLocale = s
2940      ThisWorkbook.Sheets(1).Cells(1, 1).Clear
2950      Application.Calculation = oldCalculation
2960      Exit Function
errorHandler:
2970      ThisWorkbook.Sheets(1).Cells(1, 1).Clear
2980      Application.DisplayAlerts = oldDisplayAlerts
2990      Application.Calculation = oldCalculation
3000      ConvertToCurrentLocale = ""
End Function

Function ValidLPFileVarName(s As String)
      ' http://lpsolve.sourceforge.net/5.5/CPLEX-format.htm
      'The letter E or e, alone or followed by other valid symbols, or followed by another E or e, should be avoided as this notation is reserved for exponential entries. Thus, variables cannot be named e9, E-24, E8cats, or other names that could be interpreted as an exponent. Even variable names such as eels or example can cause a read error, depending on their placement in an input line.
3010      If left(s, 1) = "E" Then
3020          ValidLPFileVarName = "_" & s
3030      Else
3040          ValidLPFileVarName = s
3050      End If
End Function

'Function FullLPFileVarName(cell As Range, AdjCellsSheetIndex As Integer)
' NO LONGER USED
' Get a valid name for the LP variable of the form A1_2 meaing cell A1 on the 2nd worksheet,
' or _E1 meaning cell E1 on the 'default' worksheet. We need to prefix E with _ to be safe; otherwise it can clash with exponential notation
' http://lpsolve.sourceforge.net/5.5/CPLEX-format.htm
'The letter E or e, alone or followed by other valid symbols, or followed by another E or e, should be avoided as this notation is reserved for exponential entries. Thus, variables cannot be named e9, E-24, E8cats, or other names that could be interpreted as an exponent. Even variable names such as eels or example can cause a read error, depending on their placement in an input line.
'    Dim sheetIndex As Integer, s As String
'    sheetIndex = cell.Worksheet.Index
'    s = cell.Address(False, False)
'    If left(s, 1) = "E" Then s = "_" & s
'    If sheetIndex <> AdjCellsSheetIndex Then s = s & "_" & str(sheetIndex)
'    FullLPFileVarName = s
'End Function

'Function ConvertFullLPFileVarNameToRange(s As String, AdjCellsSheetIndex As Integer) As Range
' COnvert an encoded LP variable name back into a range on the appropriate sheet
''    Dim i As Integer, sheetIndex As Integer
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

Function SolverRelationAsUnicodeChar(rel As Integer) As String
3060      Select Case rel
              Case RelationGE
3070              SolverRelationAsUnicodeChar = ChrW(&H2265) ' ">" gg
3080          Case RelationEQ
3090              SolverRelationAsUnicodeChar = "="
3100          Case RelationLE
3110              SolverRelationAsUnicodeChar = ChrW(&H2264) ' "<"
3120          Case Else
3130              SolverRelationAsUnicodeChar = "(unknown)"
3140      End Select
End Function

'Function SolverRelationAsChar(rel As Integer) As String
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

Function SolverRelationAsString(rel As Integer) As String
3150      Select Case rel
              Case RelationGE
3160              SolverRelationAsString = ">="
3170          Case RelationEQ
3180              SolverRelationAsString = "="
3190          Case RelationLE
3200              SolverRelationAsString = "<="
3210          Case Else
3220              SolverRelationAsString = "(unknown)"
3230      End Select
End Function

Function ReverseRelation(rel As Integer) As Integer
3240      ReverseRelation 4 - rel
End Function

Function UserSetQuickSolveParameterRange() As Boolean
3250      UserSetQuickSolveParameterRange = False
3260      If Application.Workbooks.Count = 0 Then
3270          MsgBox "Error: No active workbook available", , "OpenSolver" & sOpenSolverVersion & " Error"
3280          Exit Function
3290      End If

          Dim sheetName As String
3300      On Error Resume Next
3310      sheetName = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!" ' NB: We have to double any ' when we quote the name
3320      If Err.Number <> 0 Then
3330          MsgBox "Error: Unable to access the active sheet", , "OpenSolver" & sOpenSolverVersion & " Error"
3340          Exit Function
3350      End If
3360      On Error GoTo 0
          
          ' Find the Parameter range
          Dim ParamRange As Range
3370      On Error Resume Next
3380      Set ParamRange = Range(sheetName & ParamRangeName)
3390      On Error GoTo 0
          
          ' Get a range from the user
          Dim NewRange As Range
3400      On Error Resume Next
3410      If ParamRange Is Nothing Then
3420          Set NewRange = Application.InputBox(prompt:="Please select the 'parameter' cells that you will be changing between successsive solves of the model.", Type:=8, title:="OpenSolver Quick Solve Parameters")
3430      Else
3440          Set NewRange = Application.InputBox(prompt:="Please select the 'parameter' cells that you will be changing between successsive solves of the model.", Type:=8, Default:=ParamRange.Address, title:="OpenSolver Quick Solve Parameters")
3450      End If
3460      On Error GoTo 0
          
3470      If Not NewRange Is Nothing Then
3480          If NewRange.Worksheet.Name <> ActiveSheet.Name Then
3490              MsgBox "Error: The parameter cells need to be on the current worksheet.", , "OpenSolver" & sOpenSolverVersion & " Error"
3500              Exit Function
3510          End If
              'On Error Resume Next
3520          If Not ParamRange Is Nothing Then
                  ' Name needs to be deleted first
3530              ActiveWorkbook.Names(sheetName & ParamRangeName).Delete
3540          End If
3550          Names.Add Name:=sheetName & ParamRangeName, RefersTo:=NewRange 'ActiveWorkbook.
              ' Return true as we have succeeded
3560          UserSetQuickSolveParameterRange = True
3570      End If
End Function

Function CheckModelHasParameterRange()
3580      If Application.Workbooks.Count = 0 Then
3590          MsgBox "Error: No active workbook available", , "OpenSolver" & sOpenSolverVersion & " Error"
3600          Exit Function
3610      End If

          Dim sheetName As String
3620      On Error Resume Next
3630      sheetName = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!" ' NB: We have to double any ' when we quote the name
3640      If Err.Number <> 0 Then
3650          MsgBox "Error: Unable to access the active sheet", , "OpenSolver" & sOpenSolverVersion & " Error"
3660          Exit Function
3670      End If
3680      On Error GoTo 0
          
3690      CheckModelHasParameterRange = True
          ' Find the Parameter range
          Dim ParamRange As Range
3700      On Error Resume Next
3710      Set ParamRange = Range(sheetName & ParamRangeName)
3720      If Err.Number <> 0 Then
3730          MsgBox "Error: No parameter range could be found on the worksheet. Please use the Initialize Quick Solve Parameters menu item to define the cells that you wish to change between successive OpenSolver solves. Note that changes to these cells must lead to changes in the underlying model's right hand side values for its constraints.", title:="OpenSolver" & sOpenSolverVersion & " Error"
3740          CheckModelHasParameterRange = False
3750          Exit Function
3760      End If
End Function

Sub GetSolveOptions(sheetName As String, SolveOptions As SolveOptionsType, ErrorString As String)
          ' Get the Solver Options, stored in named ranges with values such as "=0.12"
          ' Because these are NAMEs, they are always in English, not the local language, so get their value using Val
3770      On Error GoTo errorHandler
3780      ErrorString = ""
3790      SetAnyMissingDefaultExcel2007SolverOptions ' This can happen if they have created the model using an old version of OpenSolver
3800      With SolveOptions
3810          .maxTime = Val(Mid(Names(sheetName & "solver_tim").value, 2)) ' Trim the "="; use Val to get a conversion in English, not the local language
3820          .MaxIterations = Val(Mid(Names(sheetName & "solver_itr").value, 2))
3830          .Precision = Val(Mid(Names(sheetName & "solver_pre").value, 2))
3840          .Tolerance = Val(Mid(Names(sheetName & "solver_tol").value, 2))  ' Stored as a value between 0 and 1 by Excel's Solver (representing a percentage)
              ' .Convergence = Val(Mid(Names(SheetName & "solver_cvg").Value, 2)) NOT USED BY OPEN SOLVER, YET!
              ' Excel stores ...!solver_sho=1 if Show Iteration Results is turned on, 2 if off (NB: Not 0!)
3850          .ShowIterationResults = Names(sheetName & "solver_sho").value = "=1"
3860      End With
ExitSub:
3870      Exit Sub
errorHandler:
3880      ErrorString = "No Solve options (such as Tolerance) could be found - perhaps a model has not been defined on this sheet?"
End Sub

Sub SetAnyMissingDefaultExcel2007SolverOptions()
          ' We set all the default values, as per Solver in Excel 2007, but with some changes. This ensures Solver does not delete the few values we actually use
3890      If ActiveWorkbook Is Nothing Then Exit Sub
3900      If ActiveSheet Is Nothing Then Exit Sub
          Dim s As String
3910      If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_drv", s) Then SetSolverNameOnSheet "drv", "=1"
3920      If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_est", s) Then SetSolverNameOnSheet "est", "=1"
3930      If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_itr", s) Then SetSolverNameOnSheet "itr", "=100" ' OpenSolver ignores this
3940      If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_lin", s) Then SetSolverNameOnSheet "lin", "=1"  ' Not "=2" as we want a linear model
3950      If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_neg", s) Then SetSolverNameOnSheet "neg", "=1"  ' Not "=2" as we want >=0 constraints
3960      If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_num", s) Then SetSolverNameOnSheet "num", "=0"
3970      If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_nwt", s) Then SetSolverNameOnSheet "nwt", "=1"
3980      If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_pre", s) Then SetSolverNameOnSheet "pre", "=0.000001"
3990      If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_scl", s) Then SetSolverNameOnSheet "scl", "=2"
4000      If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_sho", s) Then SetSolverNameOnSheet "sho", "=2"
4010      If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_tim", s) Then SetSolverNameOnSheet "tim", "=9999999999"  ' not "=100" as we want longer runs; Solver will force this to be no more than 9999
4020      If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_tol", s) Then SetSolverNameOnSheet "tol", "=0.05"
4030      If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_typ", s) Then SetSolverNameOnSheet "typ", "=1"
4040      If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_val", s) Then SetSolverNameOnSheet "val", "=0"
4050      If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_cvg", s) Then SetSolverNameOnSheet "cvg", "=0.0001"  ' probably not needed, but set it to be safe
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
4060      lRet = apiShellExecute(hwnd, vbNullString, _
                  stFile, vbNullString, vbNullString, lShowHow)
                  
4070      If lRet > ERROR_SUCCESS Then
4080          stRet = vbNullString
4090          lRet = -1
4100      Else
4110          Select Case lRet
                  Case ERROR_NO_ASSOC:
                      'Try the OpenWith dialog
4120                  varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " _
                              & stFile, WIN_NORMAL)
4130                  lRet = (varTaskID <> 0)
4140              Case ERROR_OUT_OF_MEM:
4150                  stRet = "Error: Out of Memory/Resources. Couldn't Execute!"
4160              Case ERROR_FILE_NOT_FOUND:
4170                  stRet = "Error: File not found.  Couldn't Execute!"
4180              Case ERROR_PATH_NOT_FOUND:
4190                  stRet = "Error: Path not found. Couldn't Execute!"
4200              Case ERROR_BAD_FORMAT:
4210                  stRet = "Error:  Bad File Format. Couldn't Execute!"
4220              Case Else:
4230          End Select
4240      End If
4250      fHandleFile = lRet & IIf(stRet = "", vbNullString, ", " & stRet)
End Function

Function GetExistingFilePathName(Directory As String, FileName As String, ByRef pathName As String) As Boolean
4260     If right(" " & Directory, 1) <> PathDelimeter Then Directory = Directory & PathDelimeter
4270     pathName = Directory & FileName
4280     GetExistingFilePathName = Dir(pathName) <> ""
End Function

'Function GetExternalSolverPath(ExternalSolverPath As String)
'          ' Location of the external solver; we look in various locations to find it every time we solve the model
'          ' This will throw an exception if the solver cannot be found
'          GetExternalSolverPath = False
'          Dim Try1 As String, Try2 As String, Try3 As String
'          ' Check that the external solver CBC can be found
'3750      ExternalSolverPath = ThisWorkbook.Path ' In same directory as OpenSolver.xlam
'3760      If right(" " & ExternalSolverPath, 1) <> PathDelimeter Then ExternalSolverPath = ExternalSolverPath & PathDelimeter
'3770      Try1 = ExternalSolverPath & ExternalSolverExeName
'3780      If Dir(Try1) = "" Then
'              ' Try in c:\temp\OpenSolver\
'3790          ExternalSolverPath = ExternalSolverPath1
'3800          Try2 = ExternalSolverPath & ExternalSolverExeName
'3810          If Dir(Try2) = "" Then
'                  ' Try in the active workbook's location
'3820              ExternalSolverPath = ActiveWorkbook.Path   ' May be blank if no active workbook
'3830              If right(" " & ExternalSolverPath, 1) <> PathDelimeter Then ExternalSolverPath = ExternalSolverPath & PathDelimeter
'3840              Try3 = ExternalSolverPath & ExternalSolverExeName
'3850              If Dir(Try3) = "" Then
'                      ' Give up
'3860                  Err.Raise Number:=OpenSolver_CBCMissingError, Source:="OpenSolver GetExternalSolverPath", _
'                            Description:="Unable to find the external solver `" + ExternalSolverExeName + "' at any of the following locations:" & vbCrLf & vbCrLf _
'                            & Try1 & vbCrLf & Try2 & vbCrLf & Try3 & vbCrLf & vbCrLf _
'                            & "Please ensure the file `" + ExternalSolverExeName + "' is in the folder:" + vbCrLf + ThisWorkbook.Path _
'                            + vbCrLf + "that contains the `OpenSolver.xlam' file." & vbCrLf _
'                            & "Note: After downloading the OpenSolver compressed (zipped) file, you must extract all the files before running OpenSolver; " _
'                            & "running OpenSolver from within the zipped file will not work."
'3870              End If
'3880          End If
'3890      End If
'End Function

Sub GetExternalSolverPathName(ByRef CombinedPathName As String, solverName As String)
          ' Location of the external solver (path and file name); we look in various locations to find it every time we solve the model
          ' We look for a 64 bit version if we are running a 64-bit system
          ' This will throw an exception if the solver cannot be found
          Dim Try1 As String, Try2 As String
          
4290      If SystemIs64Bit Then
4300          If GetExistingFilePathName(ThisWorkbook.Path, Replace(solverName, ".exe", "64.exe"), CombinedPathName) Then Exit Sub ' Found a 64 bit solver
4310          Try1 = CombinedPathName
4320      End If
          ' Look for the 32 bit version
4330      If GetExistingFilePathName(ThisWorkbook.Path, solverName, CombinedPathName) Then Exit Sub
4340      Try2 = CombinedPathName
          ' Fail
4350      If solverName = "cbc.exe" Then
4360            Err.Raise Number:=OpenSolver_CBCMissingError, Source:="OpenSolver GetExternalSolverPathName", _
                      Description:="Unable to find the external solver `" + ExternalSolverExeName + "' at any of the following location(s):" & vbCrLf & vbCrLf _
                      & Try1 & IIf(Try1 <> "", vbCrLf, "") & Try2 & vbCrLf & vbCrLf _
                      & "Please ensure the file `" + ExternalSolverExeName + "' is in the folder:" + vbCrLf + ThisWorkbook.Path _
                      + vbCrLf + "that contains the `OpenSolver.xlam' file." & vbCrLf & vbCrLf _
                      & "Notes: After downloading the OpenSolver compressed (zipped) file, you must extract all the files before running OpenSolver; " _
                      & "running OpenSolver from within the zipped file will not work." & vbCrLf _
                      & "On a 64 bit system, the '" & ExternalSolverExeName64 & "' file will be used if it exists."
4370      Else
                'Give an error if the user has selected a solver other then cbc and give the option of changing to cbc
4380            If MsgBox("Unable to find the external solver `" + solverName + "' at any of the following location(s):" & vbCrLf & vbCrLf _
                      & Try1 & IIf(Try1 <> "", vbCrLf, "") & vbCrLf & vbCrLf _
                      & "Please ensure the file `" + solverName + "' is in the folder" _
                      + "that contains the `OpenSolver.xlam' file." & vbCrLf & vbCrLf _
                      & "Notes: If you do not have " & solverName & " installed then change the chosen solver to the default (cbc) or one of the other solver engines in the model dialogue under 'Solver Engine...'" & vbCrLf _
                      & "The default CBC solver should be found when downloading the OpenSolver compressed (zipped) file; " & vbCrLf _
                      & "Running OpenSolver from within the zipped file will not work." & vbCrLf & vbCrLf _
                      & "Would you like to change your preferred solver to the default 'cbc' solver and continue solving?", vbYesNo, "OpenSolver" & sOpenSolverVersion & " Error") _
                      = vbYes Then
                    
                    'If they want to change it to cbc then set the solver for that sheet as cbc and continue solving
4390                solverName = "cbc.exe"
4400                MsgBox prompt:="Your preferred solver has been changed to 'cbc' for this model", title:="OpenSolver"
4410                Call SetNameOnSheet("OpenSolver_ChosenSolver", "=CBC")
4420                GetExternalSolverPathName CombinedPathName, solverName
        
4430            Else
                    'Raise an error if they choose to not use cbc
4440                Err.Raise Number:=OpenSolver_CBCMissingError, Source:="OpenSolver GetExternalSolverPathName", _
                      Description:="Unable to find the external solver '" & solverName & "'."
              
                'Err.Raise Number:=OpenSolver_CBCMissingError, Source:="OpenSolver GetExternalSolverPathName", _
                      Description:="Unable to find the external solver `" + SolverName + "' at any of the following location(s):" & vbCrLf & vbCrLf _
                      & Try1 & IIf(Try1 <> "", vbCrLf, "") & Try2 & vbCrLf & vbCrLf _
                      & "Please ensure the file `" + SolverName + "' is in the folder:" + vbCrLf + ThisWorkbook.Path _
                      + vbCrLf + "that contains the `OpenSolver.xlam' file." & vbCrLf & vbCrLf _
                      & "Notes: If you do not have " & SolverName & " installed then change the chosen solver to the default (cbc) in the model dialogue under 'Solver Engine...'" & vbCrLf _
                      & "The default CBC solver should be found when downloading the OpenSolver compressed (zipped) file; " & vbCrLf _
                      & "Running OpenSolver from within the zipped file will not work."
4450            End If
4460      End If
End Sub

Function GetCBCExtraParametersString(sheet As Worksheet, ErrorString As String) As String
          ' The user can define a set of parameters they want to pass to CBC; this gets them as a string
          ' Note: The named range MUST be on the current sheet
          Dim CBCParametersRange As Range, CBCExtraParametersString As String, i As Long
4470      ErrorString = ""
4480      If GetNamedRangeIfExistsOnSheet(sheet, "OpenSolver_CBCParameters", CBCParametersRange) Then
4490          If CBCParametersRange.Columns.Count <> 2 Then
4500              ErrorString = "The range OpenSolver_CBCParameters must be a two-column table."
4510              Exit Function
4520          End If
4530          For i = 1 To CBCParametersRange.Rows.Count
                  Dim ParamName As String, ParamValue As String
4540              ParamName = Trim(CBCParametersRange.Cells(i, 1))
4550              If ParamName <> "" Then
4560                  If left(ParamName, 1) <> "-" Then ParamName = "-" & ParamName
4570                  ParamValue = Trim(CBCParametersRange.Cells(i, 2))
4580                  CBCExtraParametersString = CBCExtraParametersString & " " & ParamName & " " & ParamValue
4590              End If
4600          Next i
4610      End If
4620      GetCBCExtraParametersString = CBCExtraParametersString
End Function

Function CheckWorksheetAvailable(Optional SuppressDialogs As Boolean = False, Optional ThrowError As Boolean = False) As Boolean
4630      CheckWorksheetAvailable = False
          ' Check there is a workbook
4640      If Application.Workbooks.Count = 0 Then
4650          If ThrowError Then Err.Raise Number:=OpenSolver_NoWorkbook, Source:="OpenSolver", Description:="No active workbook available."
4660          If Not SuppressDialogs Then MsgBox "Error: No active workbook available", , "OpenSolver" & sOpenSolverVersion & " Error"
4670          Exit Function
4680      End If
          ' Check we can access the worksheet
          Dim w As Worksheet
4690      On Error Resume Next
4700      Set w = ActiveWorkbook.ActiveSheet
4710      If Err.Number <> 0 Then
4720          If ThrowError Then Err.Raise Number:=OpenSolver_NoWorksheet, Source:="OpenSolver", Description:="The active sheet is not a worksheet."
4730          If Not SuppressDialogs Then MsgBox "Error: The active sheet is not a worksheet.", , "OpenSolver" & sOpenSolverVersion & " Error"
4740          Exit Function
4750      End If
          ' OK
4760      CheckWorksheetAvailable = True
End Function

Function GetOneCellInRange(r As Range, instance As Long) As Range
          ' Given an 'instance' between 1 and r.Count, return the instance'th cell in the range, where our count goes cross each row in turn (as does 'for each in range')
          Dim RowOffset As Long, ColOffset As Long
          Dim NumCols As Long
          ' Debug.Assert r.Areas.count = 1
4770      NumCols = r.Columns.Count
4780      RowOffset = ((instance - 1) \ NumCols)
4790      ColOffset = ((instance - 1) Mod NumCols)
4800      Set GetOneCellInRange = r.Cells(1 + RowOffset, 1 + ColOffset)
End Function

Function Max_Int(a As Integer, b As Integer) As Integer
4810      If a > b Then
4820          Max_Int = a
4830      Else
4840          Max_Int = b
4850      End If
End Function

Function Max_Double(a As Double, b As Double) As Double
4860      If a > b Then
4870          Max_Double = a
4880      Else
4890          Max_Double = b
4900      End If
End Function

Function Create1x1Array(X As Variant) As Variant
          ' Create a 1x1 array containing the value x
          Dim v(1, 1) As Variant
4910      v(1, 1) = X
4920      Create1x1Array = v
End Function

Function ForceCalculate(prompt As String) As Boolean
           'There appears to be a bug in Excel 2010 where the .Calculate does not always complete. We handle up to 3 such failures.
           ' We have seen this problem arise on large models.
4930       Application.Calculate

4940      If Application.CalculationState <> xlDone Then Application.Calculate
4950      If Application.CalculationState <> xlDone Then Application.Calculate
4960      If Application.CalculationState <> xlDone Then
4970          DoEvents
4980          Application.CalculateFullRebuild
4990          DoEvents
5000      End If
          
          ' Check for circular references causing problems, which can happen if iterative calculation mode is enabled.
          If Application.CalculationState <> xlDone Then
              If Application.Iteration Then
                  If MsgBox("Iterative calculation mode is enabled and may be interfering with the inital calculation. Would you like to try disabling iterative calculation mode to see if this fixes the problem?", _
                            vbYesNo, _
                            "OpenSolver: Iterative Calculation Mode Detected...") = vbYes Then
                      Application.Iteration = False
                      Application.Calculate
                  End If
              End If
          End If
          
5010      While Application.CalculationState <> xlDone
5020          If MsgBox(prompt, _
                          vbCritical + vbRetryCancel + vbDefaultButton1, _
                          "OpenSolver: Calculation Error Occured...") = vbCancel Then
5030              ForceCalculate = False
5040              Exit Function
5050          Else 'Recalculate the workbook if the user wants to retry
5060              Application.Calculate
5070          End If
5080      Wend
5090      ForceCalculate = True
End Function

Function ProperUnion(R1 As Range, R2 As Range) As Range
          ' Return the union of r1 and r2, where r1 may be Nothing
          ' TODO: Handle the fact that Union will return a range with multiple copies of overlapping cells - does this matter?
5100      If R1 Is Nothing Then
5110          Set ProperUnion = R2
5120      ElseIf R2 Is Nothing Then
5130          Set ProperUnion = R1
5140      ElseIf R1 Is Nothing And R2 Is Nothing Then
5150          Set ProperUnion = Nothing
5160      Else
5170          Set ProperUnion = Union(R1, R2)
5180      End If
End Function

Function GetRangeValues(r As Range) As Variant()
          ' This copies the values from a possible multi-area range into a variant
          Dim v() As Variant, i As Long
5190      ReDim v(r.Areas.Count)
5200      For i = 1 To r.Areas.Count
5210          v(i) = r.Areas(i).Value2 ' Copy the entire area into the i'th entry of v
5220      Next i
5230      GetRangeValues = v
End Function

Sub SetRangeValues(r As Range, v() As Variant)
          ' This copies the values from a variant into a possibly multi-area range; see GetRangeValues
          Dim i As Long
5240      For i = 1 To r.Areas.Count
5250          r.Areas(i).Value2 = v(i)
5260      Next i
End Sub

Function MergeRangesCellByCell(R1 As Range, R2 As Range) As Range
          ' This merges range r2 into r1 cell by cell.
          ' This shoulsd be fastest if range r2 is smaller than r1
          Dim result As Range, cell As Range
5270      Set result = R1
5280      For Each cell In R2
5290          Set result = Union(result, cell)
5300      Next cell
5310      Set MergeRangesCellByCell = result
End Function

Function RemoveRangeOverlap(r As Range) As Range
          ' This creates a new range from r which does not contain any multiple repetitions of cells
          ' This works around the fact that Excel allows range like "A1:A2,A2:A3", which has a .count of 4 cells
          ' The Union function does NOT remove all overlaps; call this after the union to
5320      If r.Areas.Count = 1 Then
5330          Set RemoveRangeOverlap = r
5340          Exit Function
5350      End If
          Dim s As Range, i As Long
5360      Set s = r.Areas(1)
5370      For i = 2 To r.Areas.Count
5380          If Intersect(s, r.Areas(i)) Is Nothing Then
                  ' Just take the standard union
5390              Set s = Union(s, r.Areas(i))
5400          Else
                  ' Merge these two ranges cell by cell; this seems to remove the overlap in my tests, but also see http://www.cpearson.com/excel/BetterUnion.aspx
                  ' Merge the smaller range into the larger
5410              If s.Count < r.Areas(i).Count Then
5420                  Set s = MergeRangesCellByCell(r.Areas(i), s)
5430              Else
5440                  Set s = MergeRangesCellByCell(s, r.Areas(i))
5450              End If
5460          End If
5470      Next i
5480      Set RemoveRangeOverlap = s
End Function

Function CheckRangeContainsNoAmbiguousMergedCells(r As Range, BadCell As Range) As Boolean
          ' This checks that if the range contains any merged cells, those cells are the 'home' cell (top left) in the merged cell block
          ' and thus references to these cells are indeed to a unique cell
          ' If we have a cell that is not the top left of a merged cell, then this will be read as blank, and writing to this will effect other cells.
5490      CheckRangeContainsNoAmbiguousMergedCells = True
5500      If Not r.MergeCells Then
5510          Exit Function
5520      End If
          Dim cell As Range
5530      For Each cell In r
5540          If cell.MergeCells Then
5550              If cell.Address <> cell.MergeArea.Cells(1, 1).Address Then
5560                  Set BadCell = cell
5570                  CheckRangeContainsNoAmbiguousMergedCells = False
5580                  Exit Function
5590              End If
5600          End If
5610      Next cell
End Function

Function StripWorksheetNameAndDollars(s As String, currentSheet As Worksheet) As String
          ' Remove the current worksheet name from a formula, along with any $
          ' Shorten the formula (eg Test4!$M$11/4+Test4!$A$3) by removing the current sheet name and all $
5620      StripWorksheetNameAndDollars = Replace(s, currentSheet.Name & "!", "")   ' Remove names like Test4!
5630      StripWorksheetNameAndDollars = Replace(StripWorksheetNameAndDollars, "'" & Replace(currentSheet.Name, "'", "''") & "'!", "") ' Remove names with spaces that are quoted, like 'Test 4'!; we have to double any ' when we quote the name
5640      StripWorksheetNameAndDollars = Replace(StripWorksheetNameAndDollars, "$", "")
End Function

' If a key doesn't exist we have to add it, otherwise we just set it
' Note: Numeric values should be passed as strings in English (not the local language)
Sub SetSolverNameOnSheet(Name As String, value As String)
5650      Name = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!solver_" + Name ' NB: We have to double any ' when we quote the name
    On Error GoTo doesntExist:
5660      Names(Name).value = value
5670      Exit Sub
doesntExist:
5680      Names.Add Name, value, False
End Sub

' NB: This is a different functiom to SetSolverNameOnSheet as we want to pass a range (on an arbitrary sheet)
'     If we use a variant,it fails as passing a range may simply pass its cell value
' If a key doesn't exist we have to add it, otherwise we just set it
' Solver stores names like Sheet1!$A$1; the sheet name is always given
Sub SetSolverNamedRangeOnSheet(Name As String, value As Range)
5690      Name = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!solver_" + Name ' NB: We have to double any ' when we quote the name
    On Error GoTo doesntExist:
5700      Names(Name).value = "=" & GetDisplayAddress(value, False) ' Cannot simply assign Names(name).Value=Value as this assigns the value in a single cell, not its address
5710      Exit Sub
doesntExist:
5720      Names.Add Name, "=" & GetDisplayAddress(value, False), False ' GetDisplayAddress(value), False
End Sub

Sub DeleteSolverNameOnSheet(Name As String)
5730      Name = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!solver_" + Name ' NB: We have to double any ' when we quote the name
5740      On Error Resume Next
5750      Names(Name).Delete
doesntExist:
End Sub

' If a key doesn't exist we have to add it, otherwise we just set it
' Note: Numeric values should be passed as strings in English (not the local language)
Sub SetNameOnSheet(Name As String, value As String)
5760      Name = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!" + Name ' NB: We have to double any ' when we quote the name
    On Error GoTo doesntExist:
5770      Names(Name).value = value
5780      Exit Sub
doesntExist:
5790      Names.Add Name, value, False
End Sub

' NB: Simply using a variant in SetSolverNameOnSheet fails as passing a range can simply pass its cell value
' If a key doesn't exist we have to add it, otherwise we just set it
' Note: Numeric values should be passed as strings in English (not the local language)
Sub SetNamedRangeOnSheet(Name As String, value As Range)
5800      Name = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!" + Name ' NB: We have to double any ' when we quote the name
    On Error GoTo doesntExist:
5810      Names(Name).value = "=" & GetDisplayAddress(value, False) ' "=" & GetDisplayAddress(Value)
5820      Exit Sub
doesntExist:
5830      Names.Add Name, "=" & GetDisplayAddress(value, False), False ' "=" & GetDisplayAddress(Value, False), False
End Sub

' If a key doesn't exist we have to add it, otherwise we just set it
' Note: Numeric values should be passed as strings in English (not the local language)
Sub DeleteNameOnSheet(Name As String)
5840      Name = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!" + Name '  ' NB: We have to double any ' when we quote the name
5850      On Error Resume Next
5860      Names(Name).Delete
doesntExist:
End Sub

Function TrimBlankLines(s As String) As String
          ' Remove any blank lines at the beginning or end of s
          Dim Done As Boolean
5870      While Not Done
5880          If Len(s) < Len(vbNewLine) Then
5890              Done = True
5900          ElseIf left(s, Len(vbNewLine)) = vbNewLine Then
5910             s = Mid(s, 3)
5920          Else
5930              Done = True
5940          End If
5950      Wend
5960      Done = False
5970      While Not Done
5980          If Len(s) < Len(vbNewLine) Then
5990              Done = True
6000          ElseIf right(s, Len(vbNewLine)) = vbNewLine Then
6010             s = left(s, Len(s) - 2)
6020          Else
6030              Done = True
6040          End If
6050      Wend
6060      TrimBlankLines = s
End Function

Function IsAmericanNumber(s As String, Optional i As Integer = 1) As Boolean
          ' Check this is a number like 3.45  or +1.23e-34
          ' This does NOT test for regional variations such as 12,34
          ' This code exists because
          '   val("12+3") gives 12 with no error
          '   Assigning a string to a double uses region-specific translation, so x="1,2" works in French
          '   IsNumeric("12,45") is true even on a US English system (and even worse...)
          '   IsNumeric(“($1,23,,3.4,,,5,,E67$)”)=True! See http://www.eggheadcafe.com/software/aspnet/31496070/another-vba-bug.aspx)

          Dim MustBeInteger As Boolean, SeenDot As Boolean, SeenDigit As Boolean
6070      MustBeInteger = i > 1   ' We call this a second time after seeing the "E", when only an int is allowed
6080      IsAmericanNumber = False    ' Assume we fail
6090      If Len(s) = 0 Then Exit Function ' Not a number
6100      If Mid(s, i, 1) = "+" Or Mid(s, i, 1) = "-" Then i = i + 1 ' Skip leading sign
6110      For i = i To Len(s)
6120          Select Case Asc(Mid(s, i, 1))
              Case Asc("E"), Asc("e")
6130              If MustBeInteger Or Not SeenDigit Then Exit Function ' No exponent allowed (as must be a simple integer)
6140              IsAmericanNumber = IsAmericanNumber(s, i + 1)   ' Process an int after the E
6150              Exit Function
6160          Case Asc(".")
6170              If SeenDot Then Exit Function
6180              SeenDot = True
6190          Case Asc("0") To Asc("9")
6200              SeenDigit = True
6210          Case Else
6220              Exit Function   ' Not a valid char
6230          End Select
6240      Next i
          ' i As Integer, AllowDot As Boolean
6250      IsAmericanNumber = SeenDigit
End Function

Sub TestIsAmericanNumber()
6260      Debug.Assert (IsAmericanNumber("12.34") = True)
6270      Debug.Assert (IsAmericanNumber("12.34e3") = True)
6280      Debug.Assert (IsAmericanNumber("+12.34e3") = True)
6290      Debug.Assert (IsAmericanNumber("-12.34e-3") = True)
6300      Debug.Assert (IsAmericanNumber("12.34e") = False)
6310      Debug.Assert (IsAmericanNumber("1e") = False)
6320      Debug.Assert (IsAmericanNumber("+") = False)
6330      Debug.Assert (IsAmericanNumber("+1e-") = False)
6340      Debug.Assert (IsAmericanNumber("E1") = False)
6350      Debug.Assert (IsAmericanNumber("12.3.4") = False)
6360      Debug.Assert (IsAmericanNumber("-") = False)
6370      Debug.Assert (IsAmericanNumber("-+3") = False)
End Sub

Sub test()
          Dim r As Range
6380      Set r = Range("A1")
6390      Debug.Print OpenSolver.GetDisplayAddress(r, False)
End Sub

Function SystemIs64Bit() As Boolean
          ' Is true if the Windows system is a 64 bit one
          ' If Not Environ("ProgramFiles(x86)") = "" Then Is64Bit=True, or
          ' Is64bit = Len(Environ("ProgramW6432")) > 0; see:
          ' http://blog.johnmuellerbooks.com/2011/06/06/checking-the-vba-environment.aspx and
          ' http://www.mrexcel.com/forum/showthread.php?542727-Determining-If-OS-Is-32-Bit-Or-64-Bit-Using-VBA and
          ' http://stackoverflow.com/questions/6256140/how-to-detect-if-the-computer-is-x32-or-x64 and
          ' http://msdn.microsoft.com/en-us/library/ms684139%28v=vs.85%29.aspx
6400     SystemIs64Bit = Environ("ProgramFiles(x86)") <> ""
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
          Dim NeedSheet As Boolean, newSheet As Worksheet, nameSheet As String, i As Integer
6410      On Error Resume Next
6420      Application.ScreenUpdating = False
          Dim s As String, value As String
6430      s = Sheets(namePrefix).Name
6440      If Err.Number <> 0 Then
6450          Set newSheet = Sheets.Add
6460          newSheet.Name = namePrefix
6470          nameSheet = namePrefix
6480          ActiveWindow.DisplayGridlines = False
6490      Else
6500          Call GetNameValueIfExists(ActiveWorkbook, sheetName & "OpenSolver_UpdateSensitivity", value)
6510          If value Then
6520              Sheets(namePrefix).Cells.Delete
6530              nameSheet = namePrefix
6540          Else
6550              i = 1
6560              Set newSheet = Sheets.Add
6570              NeedSheet = True
6580              On Error Resume Next
6590              While NeedSheet
6600                  nameSheet = namePrefix & " " & i
6610                  newSheet.Name = nameSheet
6620                  If Err.Number = 0 Then NeedSheet = False
6630                  i = i + 1
6640                  Err.Number = 0
6650              Wend
6660              ActiveWindow.DisplayGridlines = False
6670          End If
6680      End If
          
6690      MakeNewSheet = nameSheet
6700      Application.ScreenUpdating = True
End Function

'==================================================================
'Save visible defined names in book in a cache to find them quickly
'==================================================================

'''''' ASL NEW FUNCTION 2012-01-23 - Andres Sommerhoff
Public Sub SearchRangeName_DestroyCache()
6710      If Not SearchRangeNameCACHE Is Nothing Then
6720          While SearchRangeNameCACHE.Count > 0
6730              SearchRangeNameCACHE.Remove 1
6740          Wend
6750      End If
6760      Set SearchRangeNameCACHE = Nothing
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

6770      On Error Resume Next
6780      CurrNamesCount = sheet.Parent.Names.Count
6790      CurrSheetName = sheet.Name
6800      CurrFileName = sheet.Parent.Name
          
6810      If LastNamesCount <> CurrNamesCount _
             Or LastSheetName <> CurrSheetName _
             Or LastFileName <> CurrFileName Then
6820                  SearchRangeName_DestroyCache  'Check confirm it is obsolote. Cache need to redone...
6830      End If
              
6840      If SearchRangeNameCACHE Is Nothing Then
6850          Set SearchRangeNameCACHE = New Collection  'Start building a new Cache
6860      Else
6870          Exit Sub 'Cache is ok -> return back
6880      End If

          'Here the Cache will be filled with visible range names only
6890      For i = 1 To ActiveWorkbook.Names.Count
6900          Set TestName = ActiveWorkbook.Names(i)
              
6910          If TestName.Visible = True Then  'Iterate through the visible names only
6920                  On Error GoTo tryerror
                      'Build the Cache with the range address as key (='sheet1'!$A$1:$B$3)
6930                  Set rComp = TestName.RefersToRange
6940                  SearchRangeNameCACHE.Add TestName, (rComp.Name)
6950                  GoTo tryNext
tryerror:
6960                  Resume tryNext
tryNext:
6970          End If
6980      Next i
          
6990      LastNamesCount = CurrNamesCount
7000      LastSheetName = CurrSheetName
7010      LastFileName = CurrFileName
          
End Sub

' ASL NEW FUNCTION 2012-01-23 - Andres Sommerhoff
Public Function SearchRangeInVisibleNames(r As Range) As Name
          Dim ret As Name
          Dim TestName As Name
          Dim i As Long
          Dim rComp As Range
                    
7020      SearchRangeName_LoadCache r.Parent  'Use a collection as cache. Without cache is a little bit slow.
                                              'To refresh the cache use SearchRangeName_DestroyCache()
7030      On Error Resume Next
7040      Set SearchRangeInVisibleNames = SearchRangeNameCACHE.Item((r.Name))
          
End Function


