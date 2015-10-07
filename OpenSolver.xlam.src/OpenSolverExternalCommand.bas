Attribute VB_Name = "OpenSolverExternalCommand"
Option Explicit

#If Mac Then
    ' Declare libc functions
    Private Declare Function system Lib "libc.dylib" (ByVal Command As String) As Long
    Private Declare Function popen Lib "libc.dylib" (ByVal Command As String, ByVal mode As String) As Long
    Private Declare Function pclose Lib "libc.dylib" (ByVal file As Long) As Long
    Private Declare Function read Lib "libc.dylib" (ByVal fd As Long, ByVal buffer As String, ByVal Size As Long) As Long
    Private Declare Function feof Lib "libc.dylib" (ByVal file As Long) As Long
    Private Declare Function fileno Lib "libc.dylib" (ByVal file As Long) As Long
    Private Declare Function fcntl Lib "libc.dylib" (ByVal fd As Long, ByVal cmd As Long, funcargs As Any) As Long

    Private Const O_NONBLOCK As Integer = 4
    Private Const F_GETFL As Long = 3&
    Private Const F_SETFL As Long = 4&
#Else
    #If VBA7 Then
        Private Type PROCESS_INFORMATION
            hProcess As LongPtr
            hThread As LongPtr
            dwProcessId As Long
            dwThreadId As Long
        End Type
        
        Private Type STARTUPINFO
            cb As Long
            lpReserved As LongPtr
            lpDesktop As LongPtr
            lpTitle As LongPtr
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
        
        Private Type SECURITY_ATTRIBUTES
            nLength As Long
            lpSecurityDescriptor As LongPtr
            bInheritHandle As Long
        End Type
    #Else
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
        
        Private Type SECURITY_ATTRIBUTES
            nLength As Long
            lpSecurityDescriptor As Long
            bInheritHandle As Long
        End Type
    #End If

    Private Const STARTF_USESHOWWINDOW = &H1
    Private Const STARTF_USESTDHANDLES  As Long = &H100
    
    
    Private Enum enPriority_Class
        NORMAL_PRIORITY_CLASS = &H20
        IDLE_PRIORITY_CLASS = &H40
        HIGH_PRIORITY_CLASS = &H80
    End Enum

    #If VBA7 Then
        Private Declare PtrSafe Function CreatePipe Lib "kernel32" (phReadPipe As LongPtr, phWritePipe As LongPtr, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
        Private Declare PtrSafe Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As LongPtr
        Private Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
        Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
        Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
        Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As LongPtr, lpExitCode As Long) As Long
        Private Declare PtrSafe Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)
        Private Declare PtrSafe Function GetFileSize Lib "kernel32" (ByVal hFile As LongPtr, lpFileSizeHigh As Long) As Long
        Private Declare PtrSafe Function TerminateProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long
        Private Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As LongPtr) As Long
        Private Declare PtrSafe Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As LongPtr, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
    #Else
        Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
        Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
        Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
        Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
        Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
        Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
        Private Declare Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)
        Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
        Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
        Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
        Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
    #End If
#End If

Private Enum enSW
    SW_HIDE = 0
    SW_NORMAL = 1
    SW_MAXIMIZE = 3
    SW_MINIMIZE = 6
End Enum

Public Enum WindowStyleType
    Hide = enSW.SW_HIDE
    Normal = enSW.SW_NORMAL
    Maximize = enSW.SW_MAXIMIZE
    Minimize = enSW.SW_MINIMIZE
End Enum

' Our custom type
#If Mac Then
    Private Type ExecInformation
        file As Long
        fd As Long
        BinaryName As String
    End Type
#ElseIf VBA7 Then
    Private Type ExecInformation
        ProcInfo As PROCESS_INFORMATION
        hWrite As LongPtr
        hRead As LongPtr
    End Type
#Else
    Private Type ExecInformation
        ProcInfo As PROCESS_INFORMATION
        hWrite As Long
        hRead As Long
    End Type
#End If

Public Function RunExternalCommand(CommandString As String, Optional StartDir As String, Optional WindowStyle As WindowStyleType = Hide, Optional WaitForCompletion As Boolean = True, Optional ExitCode As Long) As Boolean
' Runs an external command, returning false if the command doesn't run
'     CommandString:      the command to run
'     WindowStyle:          the visibility of the process
'     WaitForCompletion:    waits for the process to complete before returning
'     ExitCode:             the return code of the process

    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
    
    #If Mac Then
        RunExternalCommand = RunExternalCommand_Mac(CommandString, StartDir, WindowStyle, WaitForCompletion, ExitCode)
    #Else
        RunExternalCommand = RunExternalCommand_Win(CommandString, StartDir, WindowStyle, WaitForCompletion, ExitCode)
    #End If  ' Mac

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverExternalCommand", "RunExternalCommand") Then Resume
    RaiseError = True
    GoTo ExitFunction
    
End Function

#If Mac Then
Private Function RunExternalCommand_Mac(CommandString As String, Optional StartDir As String, Optional WindowStyle As WindowStyleType = Hide, Optional WaitForCompletion As Boolean = True, Optional ExitCode As Long) As Boolean
Dim FullCommand As String
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    If Len(StartDir) <> 0 Then FullCommand = "cd " & MakePathSafe(StartDir) & ScriptSeparator
    FullCommand = FullCommand & CommandString

    If WaitForCompletion Then
        ' We are waiting for completion
        If WindowStyle = Hide Then
            Dim ret As Long
            ret = system(FullCommand)
            If ret = 0 Then RunExternalCommand_Mac = True
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
            RunExternalCommand_Mac = True
        End If
    Else
        ' We need to start the command asynchronously.
        ' To do this, we dump the command to script file then start using "open -a Terminal <file>"
        Dim TempScript As String
        If GetTempFilePath("tempscript.sh", TempScript) Then DeleteFileAndVerify TempScript
        
        CreateScriptFile TempScript, FullCommand
        
        If WindowStyle = Hide Then
            ' Applescript escapes double quotes with a backslash
            FullCommand = Replace(MakePathSafe(TempScript), """", "\""")
            MacScript "do shell script """ & FullCommand & " > /dev/null 2>&1 &"""
            RunExternalCommand_Mac = True
        Else
            ret = system("open -a Terminal " & MakePathSafe(TempScript))
            If ret = 0 Then RunExternalCommand_Mac = True
        End If
    End If

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverExternalCommand", "RunExternalCommand_Mac") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function
#Else
Private Function RunExternalCommand_Win(CommandString As String, Optional StartDir As String, Optional WindowStyle As WindowStyleType = Hide, Optional WaitForCompletion As Boolean = True, Optional ExitCode As Long) As Boolean
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler

    Dim tStartupInfo As STARTUPINFO
    With tStartupInfo
        .cb = Len(tStartupInfo)
        .dwFlags = STARTF_USESHOWWINDOW Or .dwFlags
        .wShowWindow = WindowStyle
    End With
    
    'Not used, but needed
    Dim sec1 As SECURITY_ATTRIBUTES
    Dim sec2 As SECURITY_ATTRIBUTES
    'Set the structure size
    sec1.nLength = Len(sec1)
    sec2.nLength = Len(sec2)
    
    ' Start the process
    Dim tSA_CreateProcessPrcInfo As PROCESS_INFORMATION
    #If VBA7 Then
        Dim lngResult As LongPtr
    #Else
        Dim lngResult As Long
    #End If
    lngResult = CreateProcess(vbNullString, CommandString, sec1, sec2, True, _
                              NORMAL_PRIORITY_CLASS, ByVal 0&, StartDir, tStartupInfo, tSA_CreateProcessPrcInfo)
    
    ' Check process has started correctly
    If lngResult = 0& Then
        Err.Raise OpenSolver_ExecutableError, Description:="Unable to run the external program: " & CommandString & ". " & vbNewLine & vbNewLine & _
                                                           "Error " & Err.LastDllError & ": " & DLLErrorText(Err.LastDllError)
    End If
    
    If Not WaitForCompletion Then
        RunExternalCommand_Win = True
        GoTo ExitFunction
    End If
    
    ' Loop until completion with sleep time for escape handling
    Do While True
        ' Split time between Excel and checking the solver process in 20:1 ratio, so hopefully a 20:1 chance of catching an escape press
        lngResult = WaitForSingleObject(tSA_CreateProcessPrcInfo.hProcess, 10) ' Wait for up to 1 millisecond
        ' Break if we are done to avoid sleeping unnecessarily below
        If lngResult <> 258 Then Exit Do
        mSleep 50 ' Sleep in Excel to keep escape detection responsive
        DoEvents
    Loop
    
    ' Get exit code
    Call GetExitCodeProcess(tSA_CreateProcessPrcInfo.hProcess, ExitCode)
    RunExternalCommand_Win = True
    
ExitClose:
    CloseHandle tSA_CreateProcessPrcInfo.hThread
    CloseHandle tSA_CreateProcessPrcInfo.hProcess
ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverExternalCommand", "RunExternalCommand_Win") Then Resume
    RaiseError = True
    On Error Resume Next
    If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then
        TerminateProcess tSA_CreateProcessPrcInfo.hProcess, 0
    End If
    GoTo ExitClose
End Function
#End If  ' Mac

Public Function ReadExternalCommandOutput(CommandString As String, Optional LogPath As String, Optional StartDir As String, Optional DisplayOutput As Boolean, Optional ExitCode As Long) As String
' Runs an external command and pipes output back into VBA. From https://stackoverflow.com/questions/2784367/capture-output-value-from-a-shell-command-in-vba
'     CommandString:      the command to run
'     LogPath:            if specified, writes the whole output to this file after pipe finishes
'     ExitCode:           the exit code of the command

    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler

    #If Mac Then
        ReadExternalCommandOutput = ReadExternalCommandOutput_Mac(CommandString, LogPath, StartDir, DisplayOutput, ExitCode)
    #Else
        ReadExternalCommandOutput = ReadExternalCommandOutput_Win(CommandString, LogPath, StartDir, DisplayOutput, ExitCode)
    #End If  ' Mac
    
ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverExternalCommand", "ReadExternalCommandOutput") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

#If Mac Then
Private Function ReadExternalCommandOutput_Mac(CommandString As String, Optional LogPath As String, Optional StartDir As String, Optional DisplayOutput As Boolean, Optional ExitCode As Long) As String
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler
    
    Dim Command As String
    Command = "cd " & MakePathSafe(StartDir) & ScriptSeparator & CommandString

    Dim file As Long, fd As Long
    file = popen(Command, "r")
    
    ' Set non-blocking read flag on the pipe
    fd = fileno(file)
    fcntl fd, F_SETFL, ByVal (O_NONBLOCK Or fcntl(fd, F_GETFL, ByVal 0&))
    
    If file = 0 Then
        Exit Function
    End If
    
    Dim DoingLogging As Boolean, FileNum As Long
    DoingLogging = Len(LogPath) <> 0
    If DoingLogging Then
        FileNum = FreeFile()
        Open LogPath For Output As #FileNum
    End If

    Do While True
        Dim Chunk As String
        Dim num_read As Long
        Chunk = Space(4096)
        num_read = read(fd, Chunk, Len(Chunk) - 1)
        If num_read = 0 Then
            Exit Do
        ElseIf num_read > 0 Then
            Chunk = Left$(Chunk, num_read)
            Debug.Print Chunk
            If DoingLogging Then Print #FileNum, Chunk;
            ReadExternalCommandOutput_Mac = ReadExternalCommandOutput_Mac & Chunk
            mSleep 50
            DoEvents
        End If
    Loop
    
ExitFunction:
    ExitCode = pclose(file)
    Close #FileNum
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverExternalCommand", "ReadExternalCommandOutput_Mac") Then Resume
    RaiseError = True
    If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then
        RunExternalCommand "pkill " & GetExecutableName(CommandString)
    End If
    GoTo ExitFunction
End Function
#Else
Private Function ReadExternalCommandOutput_Win(CommandString As String, Optional LogPath As String, Optional StartDir As String, Optional DisplayOutput As Boolean, Optional ExitCode As Long) As String
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler
        
    Dim tSA_CreatePipe As SECURITY_ATTRIBUTES
    With tSA_CreatePipe
        .nLength = Len(tSA_CreatePipe)
        .lpSecurityDescriptor = 0&
        .bInheritHandle = True
    End With
    
    ' Create the pipe
    #If VBA7 Then
        Dim hRead As LongPtr, hWrite As LongPtr
    #Else
        Dim hRead As Long, hWrite As Long
    #End If
    If CreatePipe(hRead, hWrite, tSA_CreatePipe, 0&) = 0& Then
        Err.Raise OpenSolver_ExecutableError, Description:="Couldn't create pipe"
    End If
    
    Dim tStartupInfo As STARTUPINFO
    With tStartupInfo
        .cb = Len(tStartupInfo)
        GetStartupInfo tStartupInfo
        ' Set the process to run in our pipe
        .hStdOutput = hWrite
        .hStdError = hWrite
        .hStdInput = hRead
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = Hide
    End With
    
    'Not used, but needed
    Dim sec1 As SECURITY_ATTRIBUTES
    Dim sec2 As SECURITY_ATTRIBUTES
    sec1.nLength = Len(sec1)
    sec2.nLength = Len(sec2)
    
    ' Start the process
    Dim tSA_CreateProcessPrcInfo As PROCESS_INFORMATION
    #If VBA7 Then
        Dim lngResult As LongPtr
    #Else
        Dim lngResult As Long
    #End If
    lngResult = CreateProcess(vbNullString, CommandString, sec1, sec2, True, _
                              NORMAL_PRIORITY_CLASS, ByVal 0&, StartDir, tStartupInfo, tSA_CreateProcessPrcInfo)
    
    ' Check process has started correctly
    If lngResult = 0& Then
        Err.Raise OpenSolver_ExecutableError, Description:="Unable to run the external program: " & CommandString & ". " & vbNewLine & vbNewLine & _
                                                           "Error " & Err.LastDllError & ": " & DLLErrorText(Err.LastDllError)
    End If
    
    CloseHandle hWrite
    CloseHandle tSA_CreateProcessPrcInfo.hThread
    hWrite = 0&
    
    Dim sOutput As String
    
    Dim DoingLogging As Boolean, FileNum As Long
    DoingLogging = Len(LogPath) <> 0
    If DoingLogging Then
        FileNum = FreeFile()
        Open LogPath For Output As #FileNum
    End If
    
    Do
        ' Check whether the proc has finished yet
        lngResult = WaitForSingleObject(tSA_CreateProcessPrcInfo.hProcess, 10&)
        
        ' Read the size of results from the pipe
        Dim lngSizeOf As Long
        lngSizeOf = 0&
        'lngSizeOf = GetFileSize(hRead, 0&)
        PeekNamedPipe hRead, ByVal 0&, 0&, ByVal 0&, lngSizeOf, ByVal 0&
        
        If lngSizeOf = 0 Then
            ' Exit if no more data in pipe and process has finished
            ' It's possible that the process can still be running and we have no new data since last read.
            If lngResult <> 258 Then Exit Do
        Else
            Dim abytBuff() As Byte, bRead As Long
            ReDim abytBuff(lngSizeOf - 1)
            If ReadFile(hRead, abytBuff(0), UBound(abytBuff) + 1, bRead, ByVal 0&) Then
                sOutput = Left$(StrConv(abytBuff(), vbUnicode), lngSizeOf)
                Debug.Print "new read", vbNewLine, sOutput, vbNewLine
                If DoingLogging Then Print #FileNum, sOutput;  ' Stream to log file
                ReadExternalCommandOutput_Win = ReadExternalCommandOutput_Win & sOutput
            End If
        End If
        mSleep 50 ' Sleep in Excel to keep escape detection responsive
        DoEvents
    Loop
    
    ' Get exit code
    GetExitCodeProcess tSA_CreateProcessPrcInfo.hProcess, ExitCode
    CloseHandle tSA_CreateProcessPrcInfo.hProcess
    
ExitFunction:
    CloseHandle hWrite
    CloseHandle hRead
    Close #FileNum
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverExternalCommand", "ReadExternalCommandOutput_Win") Then Resume
    RaiseError = True
    On Error Resume Next
    If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then
        TerminateProcess tSA_CreateProcessPrcInfo.hProcess, 0
    End If
    GoTo ExitFunction
End Function
#End If  ' Mac

#If Win32 Then
Private Function DLLErrorText(ByVal lLastDLLError As Long) As String
' From http://stackoverflow.com/questions/1439200/vba-shell-and-wait-with-exit-code
    Dim sBuff As String * 256
    Dim lCount As Long
    Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
    Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200

    lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
    If lCount Then
        DLLErrorText = Left$(sBuff, lCount - 2) ' Remove line feeds
    End If
End Function
#End If  ' Win32

Private Function GetExecutableName(CommandString As String) As String
    Dim start As Long, finish As Long
    ' Get position of the first space, or the end of the first arg if quoted
    If Left(CommandString, 1) = """" Then
        finish = InStr(2, CommandString, """")
    Else
        finish = InStr(CommandString, " ")
    End If
    If finish = 0 Then finish = Len(CommandString) + 1
    
    ' Read backwards to first path delimiter
    Dim Delimiter As String
    #If Mac Then
        Delimiter = "/"
    #Else
        Delimiter = "\"
    #End If
    start = InStrRev(CommandString, Delimiter, finish)
    
    GetExecutableName = Mid(CommandString, start + 1, finish - start - 1)
End Function

Private Function StartProcess(Command As String, StartDir As String, PipeCommand As Boolean) As ExecInformation
    #If Mac Then
        ' Save the name of the binary in case we need to kill it later
        StartProcess.BinaryName = GetExecutableName(Command)
        
        ' cd to the starting directory if supplied
        Dim FullCommand As String
        If Len(StartDir) <> 0 Then
            FullCommand = "cd " & MakePathSafe(StartDir) & ScriptSeparator
        End If
        FullCommand = FullCommand & Command

        ' Start command in pipe
        StartProcess.file = popen(Command, "r")
        If StartProcess.file = 0 Then
            Err.Raise OpenSolver_ExecutableError, _
                Description:="Unable to run the command: " & vbNewLine & Command
        End If
    
        ' Set non-blocking read flag on the pipe
        StartProcess.fd = fileno(StartProcess.file)
        fcntl StartProcess.fd, F_SETFL, ByVal (O_NONBLOCK Or fcntl(StartProcess.fd, F_GETFL, ByVal 0&))
    #Else
        If PipeCommand Then
            ' Create the pipe
            Dim tSA_CreatePipe As SECURITY_ATTRIBUTES
            With tSA_CreatePipe
                .nLength = Len(tSA_CreatePipe)
                .lpSecurityDescriptor = 0&
                .bInheritHandle = True
            End With
        
            If CreatePipe(StartProcess.hRead, StartProcess.hWrite, tSA_CreatePipe, 0&) = 0& Then
                Err.Raise OpenSolver_ExecutableError, Description:="Couldn't create pipe"
            End If
        End If
        
        Dim tStartupInfo As STARTUPINFO
        With tStartupInfo
            .cb = Len(tStartupInfo)
            GetStartupInfo tStartupInfo
            
            If PipeCommand Then
                ' Set the process to run in our pipe
                .hStdOutput = StartProcess.hWrite
                .hStdError = StartProcess.hWrite
                .hStdInput = StartProcess.hRead
                .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
                .wShowWindow = Hide
            Else
                .dwFlags = STARTF_USESHOWWINDOW Or .dwFlags
                .wShowWindow = Hide  ' WindowStyle
            End If
        End With
        
        'Not used, but needed for CreateProcess
        Dim sec1 As SECURITY_ATTRIBUTES
        Dim sec2 As SECURITY_ATTRIBUTES
        sec1.nLength = Len(sec1)
        sec2.nLength = Len(sec2)
        
        ' Start the process
        #If VBA7 Then
            Dim result As LongPtr
        #Else
            Dim result As Long
        #End If
        result = CreateProcess(vbNullString, Command, sec1, sec2, True, NORMAL_PRIORITY_CLASS, _
                               ByVal 0&, StartDir, tStartupInfo, StartProcess.ProcInfo)
        
        ' Check process has started correctly
        If result = 0 Then
            Err.Raise OpenSolver_ExecutableError, _
                Description:="Unable to run the external program: " & Command & vbNewLine & vbNewLine & _
                             "Error " & Err.LastDllError & ": " & DLLErrorText(Err.LastDllError)
        End If
        
        ' Close unneeded handles
        CloseHandle StartProcess.hWrite
        CloseHandle StartProcess.ProcInfo.hThread
        StartProcess.hWrite = 0&
    #End If
End Function

Private Function GetExitCode(ExecInfo As ExecInformation) As Long
    ' Tidy up everything to do with the execution
    #If Mac Then
        GetExitCode = pclose(ExecInfo.file)
    #Else
        GetExitCodeProcess ExecInfo.ProcInfo.hProcess, GetExitCode
    #End If
End Function

Private Sub CloseProcess(ExecInfo As ExecInformation)
    #If Mac Then
    #Else
        CloseHandle ExecInfo.ProcInfo.hProcess
        CloseHandle ExecInfo.ProcInfo.hThread
        CloseHandle ExecInfo.hWrite
        CloseHandle ExecInfo.hRead
    #End If
End Sub

Private Sub KillProcess(ExecInfo As ExecInformation)
    #If Mac Then
        system "pkill " & ExecInfo.BinaryName
    #Else
        TerminateProcess ExecInfo.ProcInfo.hProcess, 0
    #End If
End Sub

Private Function ReadChunk(ExecInfo As ExecInformation, ByRef NewData As String) As Boolean
    Const CHUNK_SIZE As Long = 4096
    Dim NumCharsRead As Long
    
    #If Mac Then
        Dim Chunk As String
        Chunk = Space(CHUNK_SIZE)
        NumCharsRead = read(StartProcess.fd, Chunk, Len(Chunk) - 1)
        
        ' NumCharsRead = -1 if nothing new written but process alive
        '              =  0 if nothing new written and process ended
        '              >  0 if new data has been read - need to read again to get state of process
        If NumCharsRead > 0 Then
            NewData = Left$(Chunk, NumCharsRead)
        End If
        ReadChunk = (NumCharsRead <> 0)  ' True if process is alive
    #Else
        ' Read the size of results from the pipe, without retrieving data
        Dim lngSizeOf As Long
        lngSizeOf = 0&  ' Reset the size to zero - this isn't done by PeekNamedPipe
        PeekNamedPipe ExecInfo.hRead, ByVal 0&, 0&, ByVal 0&, lngSizeOf, ByVal 0&
        
        If lngSizeOf = 0 Then
            ' It's possible that the process can still be running and we have no new data since last read.
            If WaitForSingleObject(ExecInfo.ProcInfo.hProcess, 10&) <> 258 Then
                ReadChunk = False  ' No more data coming into pipe
            Else
                ReadChunk = True  ' Process is still alive
            End If
        Else
            Dim abytBuff() As Byte
            ReDim abytBuff(lngSizeOf - 1)
            If ReadFile(ExecInfo.hRead, abytBuff(0), UBound(abytBuff) + 1, NumCharsRead, ByVal 0&) Then
                NewData = Left$(StrConv(abytBuff(), vbUnicode), NumCharsRead)
            End If
            ReadChunk = True
        End If
    #End If
End Function

