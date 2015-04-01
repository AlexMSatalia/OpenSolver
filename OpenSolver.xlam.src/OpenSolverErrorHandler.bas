Attribute VB_Name = "OpenSolverErrorHandler"
Public Const DEBUG_MODE As Boolean = False
Public Const USER_CANCEL As Long = 18
 
Private Const SILENT_ERROR As String = "UserCancel"
Private Const FILE_ERROR_LOG As String = "error.log"
 
Public ErrMsg As String
Public ErrNum As Long
Public ErrSource As String

' OpenSolver error numbers.
Public Const OpenSolver_ModelError = vbObjectError + 1001 ' An error occured while building the model
Public Const OpenSolver_BuildError = vbObjectError + 1002 ' An error occured while building the model
Public Const OpenSolver_SolveError = vbObjectError + 1003 ' An error occured while solving the model
Public Const OpenSolver_UserCancelledError = vbObjectError + 1004 ' The user cancelled the model build or the model solve

Public Const OpenSolver_ExecutableError = vbObjectError + 1011 ' Something went wrong trying to run an external program
Public Const OpenSolver_CBCError = vbObjectError + 1012 ' Something went wrong trying to run CBC
Public Const OpenSolver_GurobiError = vbObjectError + 1013 ' Something went wrong trying to run Gurobi
Public Const OpenSolver_NeosError = vbObjectError + 1014 ' Something went wrong trying to run NEOS
Public Const OpenSolver_NomadError = vbObjectError + 1015 ' Something went wrong trying to run NOMAD

Public Const OpenSolver_NoFile = vbObjectError + 1021
Public Const OpenSolver_NoWorksheet = vbObjectError + 1022 ' There is no active workbook
Public Const OpenSolver_NoWorkbook = vbObjectError + 1023 ' There is no active worksheet

Public Const OpenSolver_VisualizerError = vbObjectError + 1031 ' An error occured while running the visualizer

 
Function ReportError(ModuleName As String, ProcedureName As String, Optional IsEntryPoint = False, Optional MinimiseUserInteraction As Boolean = False) As Boolean
    ' See if we should clear the log file
    Dim NewLogFile
    NewLogFile = (ErrNum = 0)
    
    ' Grab the error info before it’s cleared by On Error Resume Next below.
    Dim ErrLine As Long
    ErrNum = Err.Number
    ErrLine = Erl
    
    If ErrNum = USER_CANCEL Then
        If MsgBox("You have pressed the Escape key. Do you wish to cancel?", _
                  vbCritical + vbYesNo + vbDefaultButton1, _
                  "OpenSolver: User Interrupt Occured...") = vbNo Then
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
        Print #FileNum, EnvironmentSummary() & vbNewLine
        Print #FileNum, StripNonBreakingSpaces(SolverInfo)
    End If
    Close #FileNum
    
    If IsEntryPoint Then
        If Not MinimiseUserInteraction And ErrNum <> OpenSolver_UserCancelledError Then
            ' We are at an entry point - report the error to the user
            Dim prompt As String, LinkTarget As String, MoreDetailsButton As Boolean, ReportIssueButton As Boolean
            prompt = ErrMsg
            ErrMsg = ""  ' Reset error message in case there's an error while showing the form
            LinkTarget = ""
            
            ' A message with an OpenSolver_*Error denotes an "intentional" error being reported, as opposed to an error we didn't expect to happen.
            ' For these messages, other info isn't shown with the error message.
            If False Then
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
                LinkTarget = "http://opensolver.org/help/"
                
                MoreDetailsButton = True
                ReportIssueButton = True
            End If
            
            MsgBoxEx prompt, vbOKOnly, "OpenSolver - Error", LinkTarget:=LinkTarget, MoreDetailsButton:=MoreDetailsButton, ReportIssueButton:=ReportIssueButton
        End If
        
        ' Clear all saved error details
        ErrNum = 0
        ErrSource = ""
        ErrMsg = ""
    End If
    
    ReportError = True
End Function

Public Function GetErrorLogFilePath() As String
    GetTempFilePath FILE_ERROR_LOG, GetErrorLogFilePath
End Function
