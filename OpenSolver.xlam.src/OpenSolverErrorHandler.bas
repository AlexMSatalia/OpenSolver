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
    ErrNum = 0
    ErrSource = vbNullString
    ErrMsg = vbNullString
    ErrLinkTarget = vbNullString
End Sub
 
Function ReportError(ModuleName As String, ProcedureName As String, Optional IsEntryPoint = False, Optional MinimiseUserInteraction As Boolean = False) As Boolean
    ' See if we should clear the log file
    Dim NewLogFile
    NewLogFile = (ErrNum = 0)
    
    ' Grab the error info before it's cleared by On Error Resume Next below.
    Dim ErrLine As Long
    ErrNum = Err.Number
    ErrLine = Erl
    
    If ErrNum = USER_CANCEL Then
        If ShowEscapeCancelMessage() = vbNo Then
            ReportError = False  'Continue on from where error occured in original code
            Exit Function
        Else
            ErrNum = OpenSolver_UserCancelledError
            ErrMsg = SILENT_ERROR
        End If
    End If
    
    If DEBUG_MODE Then
        ReportError = False
        Stop ' Break execution if we are running in debug mode
        Exit Function
    End If
    
    ' If this is the originating error, the static error message variable will be empty.
    ' In that case, store the originating error message in the static variable.
    If Len(ErrMsg) = 0 Then ErrMsg = Err.Description
    
    ' We don't want errors in the error logging to matter.
    On Error Resume Next
    
    ' Load the default filename if required.
    Dim FileName As String
    FileName = ThisWorkbook.Name
    
    Dim Path As String
    Path = GetErrorLogFilePath()
    If NewLogFile Then DeleteFileAndVerify Path
    
    ' Construct the fully-qualified error source name.
    ErrSource = Format(Now, "dd mmm yy hh:mm:ss") & " [" & FileName & "] " & ModuleName & "." & ProcedureName
    
    ' Create the error text to be logged.
    Dim LogText As String
    LogText = ErrSource & ": Line " & ErrLine
    
    ' Get the solver info if we need it, avoiding clashes in file handles while writing the log file
    ' TODO fix our file IO throughout so that this extra step isn't needed
    Dim SolverInfo As String
    If IsEntryPoint Then
        SolverInfo = SolverSummary()
    End If
    
    ' Open the log file, write out the error information and close the log file.
    Dim FileNum As Integer
    FileNum = FreeFile()
    Open Path For Append As #FileNum
    Print #FileNum, LogText
    If IsEntryPoint Then
        Print #FileNum, vbNewLine & "Error " & CStr(ErrNum) & ": " & ErrMsg & vbNewLine
        If Len(LastUsedSolver) <> 0 Then Print #FileNum, "Solver: " & LastUsedSolver & vbNewLine
        Print #FileNum, EnvironmentDetail() & vbNewLine
        Print #FileNum, StripNonBreakingSpaces(SolverInfo)
    End If
    Close #FileNum
    
    If IsEntryPoint Then
        If Not MinimiseUserInteraction And ErrNum <> OpenSolver_UserCancelledError Then
            ' We are at an entry point - report the error to the user
            Dim prompt As String, LinkTarget As String, MoreDetailsButton As Boolean, ReportIssueButton As Boolean
            prompt = ErrMsg
            ErrMsg = vbNullString  ' Reset error message in case there's an error while showing the form
            
            ' A message with an OpenSolver_UserError denotes an error caused by the user, as opposed to an error we didn't expect to happen.
            ' For these messages, other info isn't shown with the error message.
            If ErrNum = OpenSolver_UserError Then
                ' Intentional error
                MoreDetailsButton = False
                ReportIssueButton = False
            Else
                ' Unintentional error, so add extra info
                prompt = "OpenSolver " & sOpenSolverVersion & " encountered an error:" & vbNewLine & _
                         prompt & vbNewLine & vbNewLine & _
                         "An error log with more details has been saved, which you can see by clicking 'More Details'. " & _
                         "If you continue to have trouble, please use the 'Report Issue' button or visit the OpenSolver website for assistance:"
                ' Add the OpenSolver help link
                If Len(ErrLinkTarget) = 0 Then ErrLinkTarget = "http://opensolver.org/help/"
                
                MoreDetailsButton = True
                ReportIssueButton = True
            End If
            
            MsgBoxEx prompt, vbOKOnly, "OpenSolver - Error", LinkTarget:=ErrLinkTarget, MoreDetailsButton:=MoreDetailsButton, ReportIssueButton:=ReportIssueButton
        End If
    End If
    
    ReportError = True
End Function

Public Function GetErrorLogFilePath() As String
    GetTempFilePath FILE_ERROR_LOG, GetErrorLogFilePath
End Function

Public Sub RaiseGeneralError(ErrorMessage As String, Optional HelpLink As String)
    ErrLinkTarget = HelpLink
    Err.Raise OpenSolver_Error, Description:=ErrorMessage
End Sub
Public Sub RaiseUserError(ErrorMessage As String, Optional HelpLink As String)
    ErrLinkTarget = HelpLink
    Err.Raise OpenSolver_UserError, Description:=ErrorMessage
End Sub
Public Sub RaiseUserCancelledError()
    Err.Raise OpenSolver_UserCancelledError, Description:=SILENT_ERROR
End Sub

Public Sub RethrowError(Optional CurrentError As ErrObject)
    If CurrentError Is Nothing Then
        Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Else
        Err.Raise CurrentError.Number, Description:=CurrentError.Description
    End If
End Sub
