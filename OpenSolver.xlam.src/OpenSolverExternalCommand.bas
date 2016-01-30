Attribute VB_Name = "OpenSolverExternalCommand"
Option Explicit

#If Mac Then
    ' Declare libc functions
    Private Declare Function system Lib "libc.dylib" (ByVal Command As String) As Long
    Private Declare Function popen Lib "libc.dylib" (ByVal Command As String, ByVal Mode As String) As Long
    Private Declare Function pclose Lib "libc.dylib" (ByVal file As Long) As Long
    Private Declare Function read Lib "libc.dylib" (ByVal fd As Long, ByVal buffer As String, ByVal Size As Long) As Long
    Private Declare Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal Size As Long, ByVal Items As Long, ByVal stream As Long) As Long
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

Public StartTime As Single

' Our custom type
#If Mac Then
    Private Type ExecInformation
        file As Long
        fd As Long
        BinaryName As String
        pid As Long
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

#If Mac Then
Private Sub GetPid(ExecInfo As ExecInformation)
    ' Gets the pid of the last spawned process with the same binary name
    ' We use popen directly rather than ExecCapture to avoid creating an infinite loop
    
    ' Get the pid of the newest process (-n) with the same name
    Dim file As Long
    file = popen("pgrep -n " & ExecInfo.BinaryName, "r")
    If file = 0 Then Exit Sub
    
    Dim result As String
    Do While feof(file) = 0
        Dim chunk As String, read As Long
        chunk = Space(4096)
        read = fread(chunk, 1, Len(chunk) - 1, file)
        If read > 0 Then
            chunk = Left$(chunk, read)
            result = result & chunk
        End If
    Loop
    pclose file
    ExecInfo.pid = Int(Val(result))
End Sub
#End If  ' Mac

Private Function StartProcess(Command As String, StartDir As String, Async As Boolean, Display As Boolean) As ExecInformation
    #If Mac Then
        ' Save the name of the binary in case we need to kill it later
        StartProcess.BinaryName = GetExecutableName(Command)
        
        ' cd to the starting directory if supplied
        Dim FullCommand As String
        If Len(StartDir) <> 0 Then
            FullCommand = "cd " & MakePathSafe(StartDir) & ScriptSeparator
        End If
        FullCommand = FullCommand & Command
        
        If Async And Display Then
            ' We need to start the command asynchronously.
            ' To do this, we dump the command to script file then start using "open -a Terminal <file>"
            Dim TempScript As String
            If GetTempFilePath("tempscript.sh", TempScript) Then DeleteFileAndVerify TempScript
            CreateScriptFile TempScript, FullCommand
            FullCommand = "open -a Terminal " & MakePathSafe(TempScript)
        End If

        ' Start command in pipe
        StartProcess.file = popen(FullCommand, "r")
        If StartProcess.file = 0 Then
            Err.Raise OpenSolver_ExecutableError, _
                Description:="Unable to run the command: " & vbNewLine & Command
        End If
        
        ' Log the pid of the spawned process
        GetPid StartProcess
        
        ' Set non-blocking read flag on the pipe
        StartProcess.fd = fileno(StartProcess.file)
        fcntl StartProcess.fd, F_SETFL, ByVal (O_NONBLOCK Or fcntl(StartProcess.fd, F_GETFL, ByVal 0&))
    #Else
        If Not Async Then
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
            
            If Async Then
                .dwFlags = STARTF_USESHOWWINDOW Or .dwFlags
                .wShowWindow = IIf(Display, enSW.SW_NORMAL, enSW.SW_HIDE)
            Else
                ' Set the process to run in our pipe
                .hStdOutput = StartProcess.hWrite
                .hStdError = StartProcess.hWrite
                .hStdInput = StartProcess.hRead
                .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
                .wShowWindow = enSW.SW_HIDE
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
        StartProcess.ProcInfo.hThread = 0&
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
        ' Try to kill using the saved pid
        If ExecInfo.pid <> 0 Then
            If system("kill " & ExecInfo.pid) = 0 Then
                Exit Sub
            End If
        End If

        ' Fallback to kill according to the binary name
        system "pkill " & ExecInfo.BinaryName
    #Else
        TerminateProcess ExecInfo.ProcInfo.hProcess, 0
    #End If
End Sub

Private Function ReadChunk(ExecInfo As ExecInformation, ByRef NewData As String) As Boolean
    Dim NumCharsRead As Long
    NewData = vbNullString
    
    #If Mac Then
        Const CHUNK_SIZE As Long = 4096
        Dim chunk As String
        chunk = Space(CHUNK_SIZE)
        NumCharsRead = read(ExecInfo.fd, chunk, Len(chunk) - 1)
        
        ' NumCharsRead = -1 if nothing new written but process alive
        '              =  0 if nothing new written and process ended
        '              >  0 if new data has been read - need to read again to get state of process
        If NumCharsRead > 0 Then
            NewData = Left$(chunk, NumCharsRead)
        End If
        ReadChunk = (NumCharsRead <> 0)  ' True if process is alive
    #Else
        #If VBA7 Then
            Dim lngResult As LongPtr
        #Else
            Dim lngResult As Long
        #End If
        lngResult = WaitForSingleObject(ExecInfo.ProcInfo.hProcess, 10&)
    
        ' Read the size of results from the pipe, without retrieving data
        Dim lngSizeOf As Long
        lngSizeOf = 0&  ' Reset the size to zero - this isn't done by PeekNamedPipe
        PeekNamedPipe ExecInfo.hRead, ByVal 0&, 0&, ByVal 0&, lngSizeOf, ByVal 0&
        
        If lngSizeOf = 0 Then
            ' It's possible that the process can still be running and we have no new data since last read.
            If lngResult <> 258 Then
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

Public Function ExecCapture(Command As String, Optional LogPath As String, Optional StartDir As String, Optional DisplayOutput As Boolean, Optional ExitCode As Long) As String
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
    
    Dim InteractiveStatus As Boolean, CursorStatus As XlMousePointer
    InteractiveStatus = Application.Interactive
    CursorStatus = Application.Cursor

    If DisplayOutput Then
        Dim frmConsole As FConsole
        Set frmConsole = New FConsole
        frmConsole.SetInput Command, LogPath, StartDir
        
        Application.Interactive = True
        Application.Cursor = xlDefault
        frmConsole.Show
        Application.Cursor = CursorStatus
        Application.Interactive = InteractiveStatus
        
        ' Get all data from form and destroy it
        frmConsole.GetOutput ExitCode, ExecCapture
        Dim status As String
        status = frmConsole.Tag
        Unload frmConsole
        
        If Len(status) > 0 Then
            If status = "Aborted" Then
                Err.Raise OpenSolver_UserCancelledError, Description:="Execution was aborted"
            Else
                Err.Raise OpenSolver_ExecutableError, Description:=status
            End If
        End If
    Else
        ExecCapture = RunCommand(Command, LogPath, StartDir, False, DisplayOutput, ExitCode)
    End If

ExitFunction:
    Application.Interactive = InteractiveStatus
    Application.Cursor = CursorStatus
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverExternalCommand", "ExecCapture") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

Public Function Exec(Command As String, Optional StartDir As String, Optional DisplayOutput As Boolean, Optional ExitCode As Long) As Boolean
    RunCommand Command, vbNullString, StartDir, False, DisplayOutput, ExitCode
    Exec = True
End Function

Public Function ExecAsync(Command As String, Optional StartDir As String, Optional DisplayOutput As Boolean) As Boolean
    RunCommand Command, vbNullString, StartDir, True, DisplayOutput, 0&
    ExecAsync = True
End Function

Public Function ExecConsole(frmConsole As FConsole, Command As String, Optional LogPath As String, Optional StartDir As String, Optional ExitCode As Long)
    ExecConsole = RunCommand(Command, LogPath, StartDir, False, True, ExitCode, frmConsole)
End Function

Private Function RunCommand(Command As String, LogPath As String, StartDir As String, Async As Boolean, Display As Boolean, ExitCode As Long, Optional frmConsole As FConsole) As String
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler
    
    StartTime = Timer()
    
    Dim DoingConsole As Boolean
    DoingConsole = Not frmConsole Is Nothing
    If DoingConsole Then
        frmConsole.AppendText "Starting process:" & vbNewLine & Command & vbNewLine & vbNewLine
        If Len(StartDir) <> 0 Then
           frmConsole.AppendText "Startup directory:" & vbNewLine & MakePathSafe(StartDir) & vbNewLine & vbNewLine
        End If
    End If
    
    Dim ExecInfo As ExecInformation
    ExecInfo = StartProcess(Command, StartDir, Async, Display)
    
    If Async Then GoTo ExitFunction
    
    ' Take care of things before we start looping
    Dim DoingLogging As Boolean, FileNum As Long
    DoingLogging = Len(LogPath) <> 0
    If DoingLogging Then
        FileNum = FreeFile()
        Open LogPath For Output As #FileNum
    End If
    If DoingConsole Then
        Dim PidMessage As String
        #If Mac Then
            PidMessage = ": PID " & ExecInfo.pid
        #End If
        frmConsole.AppendText "Process started" & PidMessage & "." & vbNewLine & vbNewLine
    End If
    
    Dim OriginalStatus As String
    If Application.StatusBar = False Then
        OriginalStatus = "OpenSolver: Solving Model..."
    Else
        OriginalStatus = Application.StatusBar
    End If

    Dim NewData As String
    Do While ReadChunk(ExecInfo, NewData)
        If Len(NewData) > 0 Then
            If DoingLogging Then
                Print #FileNum, NewData;
            End If
            RunCommand = RunCommand & NewData
        End If
        If DoingConsole Then
            If frmConsole.Tag = "Cancelled" Then
                Err.Raise OpenSolver_UserCancelledError, Description:=SILENT_ERROR
            End If
        
            ' Update console even if no new data to update the elapsed time
            frmConsole.AppendText NewData
        End If
        If DoingLogging Then UpdateStatusBar OriginalStatus & " Time elapsed: " & Int(Timer() - StartTime) & " seconds"
        mSleep 50
        DoEvents
    Loop
    
ExitFunction:
    Close #FileNum
    If Not Async Then ExitCode = GetExitCode(ExecInfo)
    CloseProcess ExecInfo
    If DoingConsole Then frmConsole.MarkCompleted
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverExternalCommand", "RunCommand") Then Resume
    RaiseError = True
    If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then
        KillProcess ExecInfo
    End If
    GoTo ExitFunction
End Function


