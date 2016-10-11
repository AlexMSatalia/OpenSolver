Attribute VB_Name = "OpenSolverExternalCommand"
Option Explicit

#If Mac Then
    ' Declare libc functions
    #If VBA7 Then
         ' TODO not sure here on longptr vs long
        Private Declare PtrSafe Function system Lib "libc.dylib" (ByVal Command As String) As Long
        Private Declare PtrSafe Function popen Lib "libc.dylib" (ByVal Command As String, ByVal Mode As String) As LongPtr
        Private Declare PtrSafe Function pclose Lib "libc.dylib" (ByVal file As LongPtr) As Long
        Private Declare PtrSafe Function read Lib "libc.dylib" (ByVal fd As Long, ByVal buffer As String, ByVal Size As Long) As Long
        Private Declare PtrSafe Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal Size As Long, ByVal Items As Long, ByVal stream As LongPtr) As Long
        Private Declare PtrSafe Function feof Lib "libc.dylib" (ByVal file As LongPtr) As Long
        Private Declare PtrSafe Function fileno Lib "libc.dylib" (ByVal file As LongPtr) As Long
        Private Declare PtrSafe Function fcntl Lib "libc.dylib" (ByVal fd As Long, ByVal cmd As Long, funcargs As Any) As Long
    #Else
        Private Declare Function system Lib "libc.dylib" (ByVal Command As String) As Long
        Private Declare Function popen Lib "libc.dylib" (ByVal Command As String, ByVal Mode As String) As Long
        Private Declare Function pclose Lib "libc.dylib" (ByVal file As Long) As Long
        Private Declare Function read Lib "libc.dylib" (ByVal fd As Long, ByVal buffer As String, ByVal Size As Long) As Long
        Private Declare Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal Size As Long, ByVal Items As Long, ByVal stream As Long) As Long
        Private Declare Function feof Lib "libc.dylib" (ByVal file As Long) As Long
        Private Declare Function fileno Lib "libc.dylib" (ByVal file As Long) As Long
        Private Declare Function fcntl Lib "libc.dylib" (ByVal fd As Long, ByVal cmd As Long, funcargs As Any) As Long
    #End If

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
    #If VBA7 Then
        Private Type ExecInformation
            file As LongPtr
            fd As Long
            BinaryName As String
            pid As Long
        End Type
    #Else
        Private Type ExecInformation
            file As Long
            fd As Long
            BinaryName As String
            pid As Long
        End Type
    #End If
#Else
    #If VBA7 Then
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
#End If

#If Win32 Then
Private Function DLLErrorText(ByVal lLastDLLError As Long) As String
      ' From http://stackoverflow.com/questions/1439200/vba-shell-and-wait-with-exit-code
          Dim sBuff As String * 256
          Dim lCount As Long
          Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
          Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200

1         lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
2         If lCount Then
3             DLLErrorText = Left$(sBuff, lCount - 2) ' Remove line feeds
4         End If
End Function
#End If  ' Win32

Private Function GetExecutableName(CommandString As String) As String
          Dim start As Long, finish As Long
          ' Get position of the first space, or the end of the first arg if quoted
1         If Left(CommandString, 1) = """" Then
2             finish = InStr(2, CommandString, """")
3         Else
4             finish = InStr(CommandString, " ")
5         End If
6         If finish = 0 Then finish = Len(CommandString) + 1
          
          ' Read backwards to first path delimiter
          Dim Delimiter As String
    #If Mac Then
7             Delimiter = "/"
    #Else
8             Delimiter = "\"
    #End If
9         start = InStrRev(CommandString, Delimiter, finish)
          
10        GetExecutableName = Mid(CommandString, start + 1, finish - start - 1)
End Function

#If Mac Then
Private Sub GetPid(ExecInfo As ExecInformation)
          ' Gets the pid of the last spawned process with the same binary name
          ' We use popen directly rather than ExecCapture to avoid creating an infinite loop
          
          ' MAC2016: `fread` on `pgrep` seems to hang on Mac 2016? We just skip this for now
1         If Val(Application.Version) >= 15 Then Exit Sub
          
          ' Get the pid of the newest process (-n) with the same name
    #If VBA7 Then
              Dim file As LongPtr
    #Else
              Dim file As Long
    #End If
2         file = popen("pgrep -n " & ExecInfo.BinaryName, "r")
3         If file = 0 Then Exit Sub
          
          ' TODO merge into read chunk?
          Dim result As String
4         Do While feof(file) = 0
              Dim chunk As String, NumCharsRead As Long
5             chunk = String(4096, Chr$(0))
6             NumCharsRead = fread(chunk, 1, Len(chunk) - 1, file)
              
              ' On Mac 2016, `fread` returns 0 always? So we initialize with nulls
              ' and trim the result to first null char instead
7             If Val(Application.Version) >= 15 Then
8                 NumCharsRead = InStr(chunk, Chr$(0)) - 1
9                 If NumCharsRead < 0 Then
10                    NumCharsRead = Len(chunk)
11                End If
12            End If
              
13            If NumCharsRead > 0 Then
14                chunk = Left$(chunk, NumCharsRead)
15                result = result & chunk
16            End If
17        Loop
18        pclose file
19        ExecInfo.pid = Int(Val(result))
End Sub
#End If  ' Mac

Private Function StartProcess(Command As String, StartDir As String, Async As Boolean, Display As Boolean) As ExecInformation
    #If Mac Then
              ' Save the name of the binary in case we need to kill it later
1             StartProcess.BinaryName = GetExecutableName(Command)
              
              ' cd to the starting directory if supplied
              Dim FullCommand As String
2             If Len(StartDir) <> 0 Then
3                 FullCommand = "cd " & MakePathSafe(StartDir) & ScriptSeparator
4             End If
5             FullCommand = FullCommand & Command
              
6             If Async And Display Then
                  ' We need to start the command asynchronously.
                  ' To do this, we dump the command to script file then start using "open -a Terminal <file>"
                  Dim TempScript As String
7                 If GetTempFilePath("tempscript.sh", TempScript) Then DeleteFileAndVerify TempScript
8                 CreateScriptFile TempScript, FullCommand
9                 FullCommand = "open -a Terminal " & MakePathSafe(TempScript)
10            End If

              ' Start command in pipe
11            StartProcess.file = popen(FullCommand, "r")
12            If StartProcess.file = 0 Then
13                RaiseGeneralError "Unable to run the command: " & vbNewLine & Command
14            End If
              
              ' Log the pid of the spawned process
15            GetPid StartProcess
              
              ' Set non-blocking read flag on the pipe
16            StartProcess.fd = fileno(StartProcess.file)
17            fcntl StartProcess.fd, F_SETFL, ByVal (O_NONBLOCK Or fcntl(StartProcess.fd, F_GETFL, ByVal 0&))
    #Else
18            If Not Async Then
                  ' Create the pipe
                  Dim tSA_CreatePipe As SECURITY_ATTRIBUTES
19                With tSA_CreatePipe
20                    .nLength = Len(tSA_CreatePipe)
21                    .lpSecurityDescriptor = 0&
22                    .bInheritHandle = True
23                End With
              
24                If CreatePipe(StartProcess.hRead, StartProcess.hWrite, tSA_CreatePipe, 0&) = 0& Then
25                    RaiseGeneralError "Couldn't create pipe"
26                End If
27            End If
              
              'Not used, but needed for CreateProcess
              ' This needs to be above creation of `tStartupInfo` or it will crash x64 Excel on Windows 7
              Dim sec1 As SECURITY_ATTRIBUTES
              Dim sec2 As SECURITY_ATTRIBUTES
28            sec1.nLength = Len(sec1)
29            sec2.nLength = Len(sec2)
              
              Dim tStartupInfo As STARTUPINFO
30            With tStartupInfo
31                .cb = Len(tStartupInfo)
32                GetStartupInfo tStartupInfo
                  
33                If Async Then
34                    .dwFlags = STARTF_USESHOWWINDOW Or .dwFlags
35                    .wShowWindow = IIf(Display, enSW.SW_NORMAL, enSW.SW_HIDE)
36                Else
                      ' Set the process to run in our pipe
37                    .hStdOutput = StartProcess.hWrite
38                    .hStdError = StartProcess.hWrite
39                    .hStdInput = StartProcess.hRead
40                    .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
41                    .wShowWindow = enSW.SW_HIDE
42                End If
43            End With
              
              ' Start the process
        #If VBA7 Then
                  Dim result As LongPtr
        #Else
                  Dim result As Long
        #End If
44            result = CreateProcess(vbNullString, Command, sec1, sec2, True, NORMAL_PRIORITY_CLASS, _
                                     ByVal 0&, StartDir, tStartupInfo, StartProcess.ProcInfo)
              
              ' Check process has started correctly
45            If result = 0 Then
46                RaiseGeneralError "Unable to run the external program: " & Command & vbNewLine & vbNewLine & _
                                    "Error " & Err.LastDllError & ": " & DLLErrorText(Err.LastDllError)
47            End If
              
              ' Close unneeded handles
48            CloseHandle StartProcess.hWrite
49            CloseHandle StartProcess.ProcInfo.hThread
50            StartProcess.hWrite = 0&
51            StartProcess.ProcInfo.hThread = 0&
    #End If
End Function

Private Function GetExitCode(ExecInfo As ExecInformation) As Long
          ' Tidy up everything to do with the execution
    #If Mac Then
1             GetExitCode = pclose(ExecInfo.file)
    #Else
2             GetExitCodeProcess ExecInfo.ProcInfo.hProcess, GetExitCode
    #End If
End Function

Private Sub CloseProcess(ExecInfo As ExecInformation)
    #If Mac Then
    #Else
1             CloseHandle ExecInfo.ProcInfo.hProcess
2             CloseHandle ExecInfo.ProcInfo.hThread
3             CloseHandle ExecInfo.hWrite
4             CloseHandle ExecInfo.hRead
    #End If
End Sub

Private Sub KillProcess(ExecInfo As ExecInformation)
    #If Mac Then
              ' Try to kill using the saved pid
1             If ExecInfo.pid <> 0 Then
2                 If system("kill " & ExecInfo.pid) = 0 Then
3                     Exit Sub
4                 End If
5             End If

              ' Fallback to kill according to the binary name
6             system "pkill " & ExecInfo.BinaryName
    #Else
7             TerminateProcess ExecInfo.ProcInfo.hProcess, 0
    #End If
End Sub

Private Function ReadChunk(ExecInfo As ExecInformation, ByRef NewData As String) As Boolean
          Dim NumCharsRead As Long
1         NewData = vbNullString
          
    #If Mac Then
              Const CHUNK_SIZE As Long = 4096
              Dim chunk As String
              
2             If Val(Application.Version) >= 15 Then
                  ' `read` seems to hang on Mac 2016, use `fread` instead.
                  ' This means the read will block until the stream closes.
                  ' MAC2016: find a fix for this so we can get non-blocking reads?
3                 ReadChunk = feof(ExecInfo.file) = 0
4                 If ReadChunk Then
5                     chunk = String(CHUNK_SIZE, Chr$(0))
6                     NumCharsRead = fread(chunk, 1, Len(chunk) - 1, ExecInfo.file)
                      
                      ' NumCharsRead always seems to be zero on Mac 2016
                      ' We initialize the string with nulls and trim to first null instead
7                     NumCharsRead = InStr(chunk, Chr$(0)) - 1
8                     If NumCharsRead < 0 Then
9                         NumCharsRead = Len(chunk)
10                    End If
                      
11                    If NumCharsRead > 0 Then
12                        NewData = Left$(chunk, NumCharsRead)
13                    End If
14                End If
15            Else
16                chunk = Space(CHUNK_SIZE)
17                NumCharsRead = read(ExecInfo.fd, chunk, Len(chunk) - 1)
              
                  ' NumCharsRead = -1 if nothing new written but process alive
                  '              =  0 if nothing new written and process ended
                  '              >  0 if new data has been read - need to read again to get state of process
18                If NumCharsRead > 0 Then
19                    NewData = Left$(chunk, NumCharsRead)
20                End If
21                ReadChunk = (NumCharsRead <> 0)  ' True if process is alive
22            End If
    #Else
        #If VBA7 Then
                  Dim lngResult As LongPtr
        #Else
                  Dim lngResult As Long
        #End If
23            lngResult = WaitForSingleObject(ExecInfo.ProcInfo.hProcess, 10&)
          
              ' Read the size of results from the pipe, without retrieving data
              Dim lngSizeOf As Long
24            lngSizeOf = 0&  ' Reset the size to zero - this isn't done by PeekNamedPipe
25            PeekNamedPipe ExecInfo.hRead, ByVal 0&, 0&, ByVal 0&, lngSizeOf, ByVal 0&
              
26            If lngSizeOf = 0 Then
                  ' It's possible that the process can still be running and we have no new data since last read.
27                If lngResult <> 258 Then
28                    ReadChunk = False  ' No more data coming into pipe
29                Else
30                    ReadChunk = True  ' Process is still alive
31                End If
32            Else
                  Dim abytBuff() As Byte
33                ReDim abytBuff(lngSizeOf - 1)
34                If ReadFile(ExecInfo.hRead, abytBuff(0), UBound(abytBuff) + 1, NumCharsRead, ByVal 0&) Then
35                    NewData = Left$(StrConv(abytBuff(), vbUnicode), NumCharsRead)
36                End If
37                ReadChunk = True
38            End If
    #End If
End Function

Public Function ExecCapture(Command As String, Optional LogPath As String, Optional StartDir As String, Optional DisplayOutput As Boolean, Optional ExitCode As Long) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim InteractiveStatus As Boolean, CursorStatus As XlMousePointer
3         InteractiveStatus = Application.Interactive
4         CursorStatus = Application.Cursor

5         If DisplayOutput Then
              Dim frmConsole As FConsole
6             Set frmConsole = New FConsole
7             frmConsole.SetInput Command, LogPath, StartDir
              
8             Application.Interactive = True
9             Application.Cursor = xlDefault
10            frmConsole.Show
11            Application.Cursor = CursorStatus
12            Application.Interactive = InteractiveStatus
              
              ' Get all data from form and destroy it
13            frmConsole.GetOutput ExitCode, ExecCapture
              Dim status As String
14            status = frmConsole.Tag
15            Unload frmConsole
              
16            If Len(status) > 0 Then
17                If status = "Aborted" Then
18                    RaiseUserCancelledError
19                Else
20                    RaiseGeneralError status
21                End If
22            End If
23        Else
24            ExecCapture = RunCommand(Command, LogPath, StartDir, False, DisplayOutput, ExitCode)
25        End If

ExitFunction:
26        Application.Interactive = InteractiveStatus
27        Application.Cursor = CursorStatus
28        If RaiseError Then RethrowError
29        Exit Function

ErrorHandler:
30        If Not ReportError("OpenSolverExternalCommand", "ExecCapture") Then Resume
31        RaiseError = True
32        GoTo ExitFunction
End Function

Public Function Exec(Command As String, Optional StartDir As String, Optional DisplayOutput As Boolean, Optional ExitCode As Long) As Boolean
1         RunCommand Command, vbNullString, StartDir, False, DisplayOutput, ExitCode
2         Exec = True
End Function

Public Function ExecAsync(Command As String, Optional StartDir As String, Optional DisplayOutput As Boolean) As Boolean
1         RunCommand Command, vbNullString, StartDir, True, DisplayOutput, 0&
2         ExecAsync = True
End Function

Public Function ExecConsole(frmConsole As FConsole, Command As String, Optional LogPath As String, Optional StartDir As String, Optional ExitCode As Long)
1         ExecConsole = RunCommand(Command, LogPath, StartDir, False, True, ExitCode, frmConsole)
End Function

Private Function RunCommand(Command As String, LogPath As String, StartDir As String, Async As Boolean, Display As Boolean, ExitCode As Long, Optional frmConsole As FConsole) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
3         Application.EnableCancelKey = xlErrorHandler
          
4         StartTime = Timer()
          
          Dim DoingConsole As Boolean
5         DoingConsole = Not frmConsole Is Nothing
6         If DoingConsole Then
7             frmConsole.AppendText "Starting process:" & vbNewLine & Command & vbNewLine & vbNewLine
8             If Len(StartDir) <> 0 Then
9                frmConsole.AppendText "Startup directory:" & vbNewLine & MakePathSafe(StartDir) & vbNewLine & vbNewLine
10            End If
11        End If
          
          Dim ExecInfo As ExecInformation
12        ExecInfo = StartProcess(Command, StartDir, Async, Display)
          
13        If Async Then GoTo ExitFunction
          
          ' Take care of things before we start looping
          Dim DoingLogging As Boolean, FileNum As Long
14        DoingLogging = Len(LogPath) <> 0
15        If DoingLogging Then
16            FileNum = FreeFile()
17            Open LogPath For Output As #FileNum
18        End If
19        If DoingConsole Then
              Dim PidMessage As String
        #If Mac Then
20                PidMessage = ": PID " & ExecInfo.pid
        #End If
21            frmConsole.AppendText "Process started" & PidMessage & "." & vbNewLine & vbNewLine
22        End If
          
          Dim OriginalStatus As String
23        If Application.StatusBar = False Then
24            OriginalStatus = "OpenSolver: Solving Model..."
25        Else
26            OriginalStatus = Application.StatusBar
27        End If

          Dim NewData As String
28        Do While ReadChunk(ExecInfo, NewData)
29            If Len(NewData) > 0 Then
30                If DoingLogging Then
31                    Print #FileNum, NewData;
32                End If
33                RunCommand = RunCommand & NewData
34            End If
35            If DoingConsole Then
36                If frmConsole.Tag = "Cancelled" Then
37                    RaiseUserCancelledError
38                End If
              
                  ' Update console even if no new data to update the elapsed time
39                frmConsole.AppendText NewData
40            End If
41            If DoingLogging Then UpdateStatusBar OriginalStatus & " Time elapsed: " & Int(Timer() - StartTime) & " seconds"
42            mSleep 50
43            DoEvents
44        Loop
          
ExitFunction:
45        Close #FileNum
46        If Not Async Then ExitCode = GetExitCode(ExecInfo)
47        CloseProcess ExecInfo
48        If DoingConsole Then frmConsole.MarkCompleted
49        If RaiseError Then RethrowError
50        Exit Function

ErrorHandler:
51        If Not ReportError("OpenSolverExternalCommand", "RunCommand") Then Resume
52        RaiseError = True
53        If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then
54            KillProcess ExecInfo
55        End If
56        GoTo ExitFunction
End Function
