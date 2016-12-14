Attribute VB_Name = "OpenSolverErrorHandler"
Option Explicit

Const DEBUG_MODE As Boolean = False
Const USER_CANCEL As Long = 18
 
Public Const SILENT_ERROR As String = "Cancelled by user"
Private Const FILE_ERROR_LOG As String = "error.log"
 
Public ErrMsg As String
Public ErrNum As Long
Public ErrSource As String
Public ErrLinkTarget As String

' OpenSolver error numbers.
Public Const OpenSolver_Error = vbObjectError + 1000 ' A general unexpected error in our code
Public Const OpenSolver_UserError = vbObjectError + 1001 ' An error caused by the user making a mistake
Public Const OpenSolver_UserCancelledError = vbObjectError + 1002 ' The user requested the solve be cancelled

Public Const LINK_NO_SOLUTION_FILE As String = "http://opensolver.org/help/#cbcfails"
Public Const LINK_SOLVER_CRASH As String = "http://opensolver.org/help/#cbccrashes"
Public Const LINK_PARAMETER_DOCS As String = "http://opensolver.org/using-opensolver/#extra-parameters"

Sub ClearError()
          ' Clear all saved error details
1         ErrNum = 0
2         ErrSource = vbNullString
3         ErrMsg = vbNullString
4         ErrLinkTarget = vbNullString
End Sub
 
Function ReportError(ModuleName As String, ProcedureName As String, Optional IsEntryPoint = False, Optional MinimiseUserInteraction As Boolean = False, Optional UserMessage As String = "", Optional StackTraceMessage As String = "") As Boolean
          ' See if we should clear the log file
          Dim NewLogFile
1         NewLogFile = (ErrNum = 0)
          
          ' Grab the error info before it's cleared by On Error Resume Next below.
          Dim ErrLine As Long
2         ErrNum = Err.Number
3         ErrLine = Erl
          
4         If ErrNum = USER_CANCEL Then
5             If ShowEscapeCancelMessage() = vbNo Then
6                 ReportError = False  'Continue on from where error occured in original code
7                 Exit Function
8             Else
9                 ErrNum = OpenSolver_UserCancelledError
10                ErrMsg = SILENT_ERROR
11            End If
12        End If
          
13        If DEBUG_MODE Then
14            ReportError = False
15            Stop ' Break execution if we are running in debug mode
16            Exit Function
17        End If
          
          ' If this is the originating error, the static error message variable will be empty.
          ' In that case, store the originating error message in the static variable.
18        If Len(ErrMsg) = 0 Then ErrMsg = Err.Description

          ' Add the additional user error message, if it is present.
          If Len(UserMessage) > 0 Then ErrMsg = ErrMsg & vbNewLine & UserMessage
          
          ' We don't want errors in the error logging to matter.
19        On Error Resume Next
          
          ' Load the default filename if required.
          Dim FileName As String
20        FileName = ThisWorkbook.Name
          
          Dim Path As String
21        Path = GetErrorLogFilePath()
22        If NewLogFile Then DeleteFileAndVerify Path
          
          ' Construct the fully-qualified error source name.
23        ErrSource = Format(Now, "dd mmm yy hh:mm:ss") & " [" & FileName & "] " & ModuleName & "." & ProcedureName
          
          ' Create the error text to be logged.
          Dim LogText As String
24        LogText = ErrSource & ": Line " & ErrLine
          
          ' Add the additional stack trace error message, if it is present.
          If Len(StackTraceMessage) > 0 Then LogText = LogText & ": " & StackTraceMessage
          
          ' Get the solver info if we need it, avoiding clashes in file handles while writing the log file
          ' TODO fix our file IO throughout so that this extra step isn't needed
          Dim SolverInfo As String
25        If IsEntryPoint Then
26            SolverInfo = SolverSummary()
27        End If
          
          ' Open the log file, write out the error information and close the log file.
          Dim FileNum As Integer
28        FileNum = FreeFile()
29        Open Path For Append As #FileNum
30        Print #FileNum, LogText
31        If IsEntryPoint Then
32            Print #FileNum, vbNewLine & "Error " & CStr(ErrNum) & ": " & ErrMsg & vbNewLine
33            If Len(LastUsedSolver) <> 0 Then Print #FileNum, "Solver: " & LastUsedSolver & vbNewLine
34            Print #FileNum, EnvironmentDetail() & vbNewLine
35            Print #FileNum, StripNonBreakingSpaces(SolverInfo)
36        End If
37        Close #FileNum
          
38        If IsEntryPoint Then
39            If Not MinimiseUserInteraction And ErrNum <> OpenSolver_UserCancelledError Then
                  ' We are at an entry point - report the error to the user
                  Dim prompt As String, LinkTarget As String, MoreDetailsButton As Boolean, ReportIssueButton As Boolean
40                prompt = ErrMsg
41                ErrMsg = vbNullString  ' Reset error message in case there's an error while showing the form
                  
                  ' A message with an OpenSolver_UserError denotes an error caused by the user, as opposed to an error we didn't expect to happen.
                  ' For these messages, other info isn't shown with the error message.
42                If ErrNum = OpenSolver_UserError Then
                      ' Intentional error
43                    MoreDetailsButton = False
44                    ReportIssueButton = False
45                Else
                      ' Unintentional error, so add extra info
46                    prompt = "OpenSolver " & sOpenSolverVersion & " encountered an error:" & vbNewLine & _
                               prompt & vbNewLine & vbNewLine & _
                               "An error log with more details has been saved, which you can see by clicking 'More Details'. " & _
                               "If you continue to have trouble, please use the 'Report Issue' button or visit the OpenSolver website for assistance:"
                      ' Add the OpenSolver help link
47                    If Len(ErrLinkTarget) = 0 Then ErrLinkTarget = "http://opensolver.org/help/"
                      
48                    MoreDetailsButton = True
49                    ReportIssueButton = True
50                End If
                  
51                MsgBoxEx prompt, vbOKOnly, "OpenSolver - Error", LinkTarget:=ErrLinkTarget, MoreDetailsButton:=MoreDetailsButton, ReportIssueButton:=ReportIssueButton
52            End If
53        End If
          
54        ReportError = True
End Function

Public Function GetErrorLogFilePath() As String
1         GetTempFilePath FILE_ERROR_LOG, GetErrorLogFilePath
End Function

Public Sub RaiseGeneralError(ErrorMessage As String, Optional HelpLink As String)
1         ErrLinkTarget = HelpLink
2         Err.Raise OpenSolver_Error, Description:=ErrorMessage
End Sub
Public Sub RaiseUserError(ErrorMessage As String, Optional HelpLink As String)
1         ErrLinkTarget = HelpLink
2         Err.Raise OpenSolver_UserError, Description:=ErrorMessage
End Sub
Public Sub RaiseUserCancelledError()
1         Err.Raise OpenSolver_UserCancelledError, Description:=SILENT_ERROR
End Sub

Public Sub RethrowError(Optional CurrentError As ErrObject)
1         If CurrentError Is Nothing Then
2             Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
3         Else
4             Err.Raise CurrentError.Number, Description:=CurrentError.Description
5         End If
End Sub
