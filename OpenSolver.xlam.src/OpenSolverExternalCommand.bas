Attribute VB_Name = "OpenSolverExternalCommand"
Option Explicit

#If Mac Then
    ' Declare libc functions
    Private Declare Function system Lib "libc.dylib" (ByVal command As String) As Long
    Private Declare Function popen Lib "libc.dylib" (ByVal command As String, ByVal mode As String) As Long
    Private Declare Function pclose Lib "libc.dylib" (ByVal file As Long) As Long
    Private Declare Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal Size As Long, ByVal Items As Long, ByVal stream As Long) As Long
    Private Declare Function feof Lib "libc.dylib" (ByVal file As Long) As Long
#Else
    #If VBA7 Then
        Type SECURITY_ATTRIBUTES
            nLength As Long
            lpSecurityDescriptor As LongPtr
            bInheritHandle As Long
        End Type

        Type PROCESS_INFORMATION
            hProcess As LongPtr
            hThread As LongPtr
            dwProcessId As Long
            dwThreadId As Long
        End Type

        Type STARTUPINFO
            cb As Long
            lpReserved As Long
            lpDesktop As Long
            lpTitle As Long
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
            lpReserved2 As Byte
            hStdInput As LongPtr
            hStdOutput As LongPtr
            hStdError As LongPtr
        End Type
    #Else
        Private Type SECURITY_ATTRIBUTES
            nLength As Long
            lpSecurityDescriptor As Long
            bInheritHandle As Long
        End Type

        Private Type PROCESS_INFORMATION
            hProcess As Long
            hThread As Long
            dwProcessId As Long
            dwThreadId As Long
        End Type

        Private Type STARTUPINFO
            cb As Long
            lpReserved As Long
            lpDesktop As Long
            lpTitle As Long
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
            lpReserved2 As Byte
            hStdInput As Long
            hStdOutput As Long
            hStdError As Long
        End Type
    #End If

    Private Const NORMAL_PRIORITY_CLASS As Long = &H20&
    Private Const STARTF_USESHOWWINDOW  As Long = &H1
    Private Const STARTF_USESTDHANDLES  As Long = &H100

    Private Const SW_HIDE               As Long = 0&
    Private Const SW_SHOWNORMAL         As Long = 1&
    Private Const SW_SHOWMINIMIZED      As Long = 2&

    #If VBA7 Then
        Private Declare PtrSafe Function CreatePipe Lib "kernel32" (phReadPipe As LongPtr, phWritePipe As LongPtr, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
        Private Declare PtrSafe Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
        Private Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
        Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
        Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
        Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As LongPtr, lpExitCode As Long) As Long
        Private Declare PtrSafe Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)
        Private Declare PtrSafe Function GetFileSize Lib "kernel32" (ByVal hFile As LongPtr, lpFileSizeHigh As Long) As Long
        Private Declare PtrSafe Function TerminateProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long
        Private Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As LongPtr) As Long
    #Else
        Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
        Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
        Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
        Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
        Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
        Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
        Private Declare Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)
        Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
        Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
        Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
    #End If
#End If

Public Enum WindowStyleType
    Hide
    Normal
    Minimize
End Enum

Public Function RunExternalCommand(CommandString As String, Optional LogPath As String, Optional WindowStyle As WindowStyleType = Hide, Optional WaitForCompletion As Boolean = True, Optional ExitCode As Long) As Boolean
' Runs an external command, returning false if the command doesn't run
'     CommandString:      the command to run
'     LogPath:            if specified, echos stdout to this file
'     WindowStyle:          the visibility of the process
'     WaitForCompletion:    waits for the process to complete before returning
'     ExitCode:             the return code of the process

    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    ' Combine logging and command strings
    Dim FullCommand As String
    FullCommand = AddLoggingToCommand(CommandString, LogPath, WindowStyle <> Hide)
    
    #If Mac Then
        If WaitForCompletion Then
            ' We are waiting for completion
            If WindowStyle = Hide Then
                Dim ret As Long
                ret = system(FullCommand)
                If ret = 0 Then RunExternalCommand = True
            Else
                ' Applescript escapes double quotes with a backslash
                FullCommand = Replace(FullCommand, """", "\""")
            
                ' Applescript for opening a terminal to run our command
                ' 1. Create window if terminal not already open, then activate window
                ' 2. Run our shell command in the terminal, saving a reference to the open window
                ' 3. Loop until the window is no longer busy
                Dim script As String
                script = _
                    "tell application ""Terminal""" & vbNewLine & _
                    "    activate" & vbNewLine & _
                    "    set w to do script """ & FullCommand & """" & vbNewLine & _
                    "    repeat" & vbNewLine & _
                    "        delay 1" & vbNewLine & _
                    "        if not busy of w then exit repeat" & vbNewLine & _
                    "    end repeat" & vbNewLine & _
                    "    do script ""exit"" in w" & vbNewLine & _
                    "end tell" & vbNewLine & _
                    "tell application ""Microsoft Excel""" & vbNewLine & _
                    "    activate" & vbNewLine & _
                    "end tell"
                MacScript (script)
                RunExternalCommand = True
            End If
        Else
            ' We need to start the command asynchronously.
            ' To do this, we dump the command to script file then start using "open -a Terminal <file>"
            Dim TempScript As String
            TempScript = GetTempFilePath("tempscript.sh")
            If FileOrDirExists(TempScript) Then Kill TempScript
            
            CreateScriptFile TempScript, FullCommand
            
            ' TODO: Figure out if we can do this silently
            ret = system("open -a Terminal " & MakePathSafe(TempScript))
            If ret = 0 Then RunExternalCommand = True
        End If

    #Else
        RunExternalCommand = RunCommand(FullCommand, WindowStyle, ExitCode, False, 0&, WaitForCompletion)
    #End If

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverExternalCommand", "RunExternalCommand") Then Resume
    RaiseError = True
    GoTo ExitFunction
    
End Function

Public Function ReadExternalCommandOutput(CommandString As String, Optional LogPath As String, Optional ExitCode As Long) As String
' Runs an external command and pipes output back into VBA. From https://stackoverflow.com/questions/2784367/capture-output-value-from-a-shell-command-in-vba
'     CommandString:      the command to run
'     LogPath:            if specified, writes the whole output to this file after pipe finishes
'     ExitCode:           the exit code of the command

    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    #If Mac Then
        Application.EnableCancelKey = xlErrorHandler
        
        Dim file As Long
        file = popen(CommandString, "r")
    
        If file = 0 Then
            Exit Function
        End If
    
        While feof(file) = 0
            Dim chunk As String
            Dim read As Long
            chunk = Space(50)
            read = fread(chunk, 1, Len(chunk) - 1, file)
            If read > 0 Then
                chunk = left$(chunk, read)
                ReadExternalCommandOutput = ReadExternalCommandOutput & chunk
            End If
        Wend
    #Else
        Dim tSA_CreatePipe As SECURITY_ATTRIBUTES
        tSA_CreatePipe.nLength = Len(tSA_CreatePipe)
        tSA_CreatePipe.lpSecurityDescriptor = 0&
        tSA_CreatePipe.bInheritHandle = True
        
        ' Create the pipe
        Dim hRead As Long, hWrite As Long
        If (CreatePipe(hRead, hWrite, tSA_CreatePipe, 0&) <> 0&) Then
            ' Run the command inside our pipe
            RunCommand CommandString, Hide, ExitCode, True, hWrite, True
            
            ' Read the results from the pipe
            Dim lngSizeOf As Long
            lngSizeOf = GetFileSize(hRead, 0&)
            If (lngSizeOf > 0) Then
                Dim abytBuff() As Byte, bRead As Long
                ReDim abytBuff(lngSizeOf - 1)
                If ReadFile(hRead, abytBuff(0), UBound(abytBuff) + 1, bRead, ByVal 0&) Then
                    ReadExternalCommandOutput = StrConv(abytBuff, vbUnicode)
                End If
            End If
        End If
    #End If

    ' Dump the output to the LogPath if present
    If Len(LogPath) <> 0 Then
        Open LogPath For Output As #1
        Print #1, ReadExternalCommandOutput
        Close #1
    End If
    
    #If Mac Then
ExitMac:
        ExitCode = pclose(file)
    #Else
ExitWindows:
        CloseHandle hWrite
        CloseHandle hRead
    #End If
    
ExitFunction:
    Close #1
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverExternalCommand", "ReadExternalCommandOutput") Then Resume
    RaiseError = True
    #If Mac Then
        GoTo ExitMac
    #Else
        GoTo ExitWindows
    #End If
End Function

Private Function RunCommand(CommandString As String, Optional WindowStyle As WindowStyleType = Hide, Optional ExitCode As Long, Optional IsBeingPiped As Boolean, Optional hWrite As Long, Optional WaitForCompletion As Boolean = True) As Boolean
' Helper function to spawn a new process for our command. Returns true if process starts/runs successfully
'     CommandString:        the command to run
'     WindowStyle:          the visibility of the process
'     ExitCode:             the return code of the process
'     WaitForCompletion:    waits for the process to complete before returning

    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler

    Dim tStartupInfo As STARTUPINFO
    tStartupInfo.cb = Len(tStartupInfo)
    
    With tStartupInfo
        .cb = Len(tStartupInfo)
        If IsBeingPiped Then
            GetStartupInfo tStartupInfo
            ' Set the process to run in our pipe
            .hStdOutput = hWrite
            .hStdError = hWrite
            .dwFlags = STARTF_USESTDHANDLES
        End If
        .dwFlags = STARTF_USESHOWWINDOW Or .dwFlags
        .wShowWindow = WindowStyleToSW(WindowStyle)
    End With
    
    ' Start the process
    Dim tSA_CreateProcessPrcInfo As PROCESS_INFORMATION, lngResult As Long
    lngResult = CreateProcess(0&, CommandString, 0&, 0&, True, _
                              NORMAL_PRIORITY_CLASS, 0&, 0&, tStartupInfo, tSA_CreateProcessPrcInfo)
    
    ' Check process has started correctly
    If lngResult = 0& Then
        Err.Raise OpenSolver_ExecutableError, Description:="Unable to run the external program: " & CommandString & ". " & vbNewLine & vbNewLine & _
                                                           "Error " & Err.LastDllError & ": " & DLLErrorText(Err.LastDllError)
    End If
    
    If Not WaitForCompletion Then
        RunCommand = True
        GoTo ExitWindows
    End If
    
    ' Loop until completion with sleep time for escape handling
    Do While True
        ' Split time between Excel and checking the solver process in 20:1 ratio, so hopefully a 20:1 chance of catching an escape press
        lngResult = WaitForSingleObject(tSA_CreateProcessPrcInfo.hProcess, 10) ' Wait for up to 1 millisecond
        ' Break if we are done to avoid sleeping unnecessarily below
        If lngResult <> 258 Then Exit Do
        Sleep 50 ' Sleep in Excel to keep escape detection responsive
        DoEvents
    Loop
    
    ' Get exit code
    Call GetExitCodeProcess(tSA_CreateProcessPrcInfo.hProcess, ExitCode)
    RunCommand = True
    
ExitWindows:
    CloseHandle tSA_CreateProcessPrcInfo.hThread
    CloseHandle tSA_CreateProcessPrcInfo.hProcess

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverExternalCommand", "ReadExternalCommandOutput") Then Resume
    RaiseError = True
    #If Mac Then
        GoTo ExitMac
    #Else
        On Error Resume Next
        If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then
            TerminateProcess tSA_CreateProcessPrcInfo.hProcess, 0
            GoTo ExitFunction
        Else
            GoTo ExitWindows
        End If
    #End If
End Function

Private Function WindowStyleToSW(WindowStyle As WindowStyleType) As Long
    Select Case WindowStyle
    Case WindowStyleType.Hide
        WindowStyleToSW = SW_HIDE
    Case WindowStyleType.Normal
        WindowStyleToSW = SW_SHOWNORMAL
    Case WindowStyleType.Minimize
        WindowStyleToSW = SW_SHOWMINIMIZED
    End Select
End Function

Private Function AddLoggingToCommand(CommandString As String, LogPath As String, NeedsStdOut As Boolean) As String
    Dim LoggingOperator As String
    If Len(LogPath) <> 0 Then
        If NeedsStdOut Then
            LoggingOperator = " | " & MakePathSafe(TeePath) & " "
        Else
            LoggingOperator = " > "
        End If
    End If
    AddLoggingToCommand = CommandString & LoggingOperator & MakePathSafe(LogPath)
End Function

Private Function TeePath() As String
    #If Mac Then
        TeePath = "tee"
    #Else
        GetExistingFilePathName JoinPaths(ThisWorkbook.Path, "Utils", "mtee"), "mtee.exe", TeePath
    #End If
End Function

Private Function DLLErrorText(ByVal lLastDLLError As Long) As String
' From http://stackoverflow.com/questions/1439200/vba-shell-and-wait-with-exit-code
    Dim sBuff As String * 256
    Dim lCount As Long
    Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
    Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200

    lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
    If lCount Then
        DLLErrorText = left$(sBuff, lCount - 2) ' Remove line feeds
    End If
End Function
