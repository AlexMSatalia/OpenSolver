Attribute VB_Name = "OpenSolverIO"
Option Explicit

' All interactions with file system and sheets in Excel

#If Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Function mktemp Lib "libc.dylib" (ByVal template As String) As String
    #Else
        Private Declare Function mktemp Lib "libc.dylib" (ByVal template As String) As String
    #End If
#Else
    #If VBA7 Then
        Private Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
        Private Declare PtrSafe Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long
    #Else
        Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
        Private Declare Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long
    #End If
#End If

Sub CheckLocationValid()
1               If StringHasUnicode(ThisWorkbook.Path) Then
2                   MsgBoxEx "The path that OpenSolver is being loaded from contains unicode characters. " & _
                             "This means the solvers are very unlikely to work. " & _
                             "Please move the OpenSolver folder so that there are no unicode characters in the complete path to the folder " & vbNewLine & vbNewLine & _
                             "The OpenSolver folder is currently located at: " & vbNewLine & _
                             ThisWorkbook.Path
3               End If
End Sub

Function SolverDirIsPresent() As Boolean
1               SolverDirIsPresent = FileOrDirExists(SolverDir)
End Function

Function ActiveSheetWithValidation() As Worksheet
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' Check there is a workbook
3         If Application.Workbooks.Count = 0 Then RaiseGeneralError "No active workbook available."

          ' Check we can access the worksheet
4         On Error Resume Next
5         Set ActiveSheetWithValidation = ActiveSheet
6         If Err.Number <> 0 Then
7             On Error GoTo ErrorHandler
8             RaiseGeneralError "The active sheet is not a worksheet."
9         End If

ExitFunction:
10        If RaiseError Then RethrowError
11        Exit Function

ErrorHandler:
12        If Not ReportError("OpenSolverIO", "CheckActiveWorksheetAvailable") Then Resume
13        RaiseError = True
14        GoTo ExitFunction
End Function

Sub GetActiveSheetIfMissing(sheet As Worksheet)
1         If sheet Is Nothing Then Set sheet = ActiveSheetWithValidation
End Sub

Function MakeNewSheet(namePrefix As String, OverwriteExisting As Boolean) As Worksheet
          Dim ScreenStatus As Boolean
1         ScreenStatus = Application.ScreenUpdating
2         Application.ScreenUpdating = False

          Dim newSheet As Worksheet
3         On Error Resume Next
4         Set newSheet = Sheets(namePrefix)
5         If Err.Number <> 0 Then
6             Set newSheet = Sheets.Add
7             newSheet.Name = namePrefix
8         Else
9             If OverwriteExisting Then
10                newSheet.Cells.Delete
11            Else
12                Set newSheet = Sheets.Add
                  Dim i As Long, NeedSheet As Boolean
13                i = 1
14                NeedSheet = True
15                On Error Resume Next
16                While NeedSheet
17                    newSheet.Name = namePrefix & " " & i
18                    If Err.Number = 0 Then NeedSheet = False
19                    i = i + 1
20                    Err.Number = 0
21                Wend
22            End If
23        End If
24        Set MakeNewSheet = newSheet
25        Application.ScreenUpdating = ScreenStatus
End Function

Function GetExistingFilePathName(Directory As String, FileName As String, ByRef pathName As String) As Boolean
1         pathName = JoinPaths(Directory, FileName)
2         GetExistingFilePathName = FileOrDirExists(pathName)
End Function

Function GetRootDriveName() As String
                Dim RaiseError As Boolean
1               RaiseError = False
2               On Error GoTo ErrorHandler

          #If Mac Then
                    Static DriveName As String
                    ' We assume that the temp folder is on the root drive
                    ' Seems reasonable, the user might be able to mess this up if they really try.
3                   If DriveName = "" Then
                        Dim Path As String
4                       Path = GetTempFolder(False)
5                       DriveName = Left(Path, InStr(Path, Application.PathSeparator) - 1)
6                   End If
7                   GetRootDriveName = DriveName
          #End If

ExitFunction:
8               If RaiseError Then RethrowError
9               Exit Function

ErrorHandler:
10              If Not ReportError("OpenSolverIO", "GetRootDriveName") Then Resume
11              RaiseError = True
12              GoTo ExitFunction
End Function

Function ConvertHfsPathToPosix(Path As String) As String
      ' Any direct file system access (using 'system' or in script files) on Mac requires
      ' that HFS-style paths are converted to normal POSIX paths. On Windows this
      ' function does nothing, so it can safely wrap all file system calls on any platform
      ' Input (HFS path):   "Macintosh HD:Users:jack:filename.txt"
      ' Output (POSIX path): "/Volumes/Macintosh HD/Users/jack/filename.txt"

          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          #If Mac Then
              ' Check we have an HFS path and not POSIX
3             If InStr(Path, "/") = 0 Then
                  Dim RootDriveName As String
4                 RootDriveName = GetRootDriveName()
5                 If Left(Path, Len(RootDriveName)) = RootDriveName Then
6                     ConvertHfsPathToPosix = Mid(Path, Len(RootDriveName) + 1)
7                 Else
                      ' Prefix disk name with :Volumes:
8                     ConvertHfsPathToPosix = ":Volumes:" & Path
9                 End If
                  ' Convert path delimiters
10                ConvertHfsPathToPosix = Replace(ConvertHfsPathToPosix, ":", "/")
11            Else
                  ' Path is already POSIX
12                ConvertHfsPathToPosix = Path
13            End If
          #Else
14            ConvertHfsPathToPosix = Path
          #End If

ExitFunction:
15        If RaiseError Then RethrowError
16        Exit Function

ErrorHandler:
17        If Not ReportError("OpenSolverIO", "ConvertHfsPathToPosix") Then Resume
18        RaiseError = True
19        GoTo ExitFunction
End Function

Function ConvertPosixPathToHfs(Path As String) As String
      ' Converts a POSIX path back to HFS
      ' Input (POSIX path): "/Volumes/Macintosh HD/Users/jack/filename.txt"
      ' Output (HFS path): "Macintosh HD:Users:jack:filename.txt"
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

    #If Mac Then
              ' Make sure we have a POSIX path
3             If InStr(Path, ":") = 0 Then
                  Const VolumePrefix = "/Volumes/"
4                 If Left(Path, Len(VolumePrefix)) = VolumePrefix Then
                      ' If the POSIX path starts with /Volumes/, then we keep the drive name after /Volumes/
5                     ConvertPosixPathToHfs = Mid(Path, Len(VolumePrefix) + 1)
6                 Else
                      ' If the POSIX path starts at the root, we add the name of the root drive
7                     ConvertPosixPathToHfs = GetRootDriveName() & Path
8                 End If
                  ' Convert Path delimiters
9                 ConvertPosixPathToHfs = Replace(ConvertPosixPathToHfs, "/", ":")
10            Else
11                ConvertPosixPathToHfs = Path
12            End If
    #Else
13            ConvertPosixPathToHfs = Path
    #End If

ExitFunction:
14        If RaiseError Then RethrowError
15        Exit Function

ErrorHandler:
16        If Not ReportError("OpenSolverIO", "ConvertPosixPathToHfs") Then Resume
17        RaiseError = True
18        GoTo ExitFunction
End Function

Function MakePathSafe(Path As String) As String
      ' Prepares a path for command-line invocation
1         MakePathSafe = IIf(Len(Path) = 0, vbNullString, Quote(ConvertHfsPathToPosix(Path)))
End Function

Function JoinPaths(ParamArray Paths() As Variant) As String
                Dim i As Long
1               For i = LBound(Paths) To UBound(Paths) - 1
2                   JoinPaths = JoinPaths & Paths(i) & IIf(Right(Paths(i), 1) <> Application.PathSeparator, Application.PathSeparator, vbNullString)
3               Next i
4               JoinPaths = JoinPaths & Paths(UBound(Paths))
End Function

Function FileOrDirExists(pathName As String) As Boolean
      ' from http://www.vbaexpress.com/kb/getarticle.php?kb_id=559
          
1         If IsMac And Val(Application.Version) >= 15 Then
              ' On Mac 2016, any file access via VBA seems to set the extended attribute
              ' `com.apple.quarantine` on the file.
              ' This attribute blocks execution of the file later, like downloaded files on Windows
              
              ' Instead, we use the `test` shell function via libc to check existence
              Dim result As String
2             result = ExecCapture("test -e " & MakePathSafe(pathName) & " && echo exists")
3             FileOrDirExists = Len(result) > 0
4         Else
              Dim iTemp As Long
5             On Error Resume Next
6             iTemp = GetAttr(pathName)
              'Check if error exists and set response appropriately
7             FileOrDirExists = (Err.Number = 0)
8         End If
           
End Function

Sub DeleteFileAndVerify(FilePath As String)
      ' Deletes file and raises error if not successful
1         On Error GoTo DeleteError
2         If FileOrDirExists(FilePath) Then Kill FilePath
3         If FileOrDirExists(FilePath) Then
4             GoTo DeleteError
5         End If
6         Exit Sub
          
DeleteError:
7         RaiseUserError "Unable to delete the file: " & FilePath & vbNewLine & vbNewLine & _
                         Err.Description & vbNewLine & vbNewLine & _
                         "To fix this, try restarting Excel and check Task Manager to make sure no solver is running. " & _
                         "If this error still appears after that, try restarting the computer."
End Sub

Sub OpenFile(FilePath As String, NotFoundMessage As String)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If Not FileOrDirExists(FilePath) Then
4             RaiseGeneralError NotFoundMessage
5         Else
              ' Check that there is no workbook open with the same name
              Dim w As Workbook
6             On Error Resume Next
7             Set w = Workbooks(Right(FilePath, InStr(FilePath, Application.PathSeparator)))
8             On Error GoTo ErrorHandler
9             Workbooks.Open FileName:=FilePath, ReadOnly:=True ' , Format:=Tabs
10        End If
11        Exit Sub

ErrorHandler:
12        MsgBox NotFoundMessage
End Sub

Sub OpenFolder(FolderPath As String, NotFoundMessage As String)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If Not FileOrDirExists(FolderPath) Then
4             RaiseGeneralError NotFoundMessage
5         Else
        #If Mac Then
6                 ExecAsync "open " & MakePathSafe(FolderPath)
        #Else
7                 ExecAsync "explorer.exe " & FolderPath, DisplayOutput:=True
        #End If
8         End If
9         Exit Sub

ErrorHandler:
10        MsgBox NotFoundMessage
End Sub

Sub DeleteFolderAndContents(FolderPath As String)
1         On Error Resume Next

    #If Mac Then
2             ExecAsync "rm -rf " & MakePathSafe(FolderPath)
    #Else
3             Kill JoinPaths(FolderPath, "*.*")
4             RmDir FolderPath
    #End If
5         On Error GoTo 0
End Sub

Sub DeleteTempFolder()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim TempFolderPath As String
3         TempFolderPath = GetTempFolder()
4         If Len(TempFolderPath) <> 0 Then
5             DeleteFolderAndContents TempFolderPath
6         End If

ExitSub:
7         If RaiseError Then RethrowError
8         Exit Sub

ErrorHandler:
9         If Not ReportError("OpenSolverIO", "DeleteTempFolder") Then Resume
10        RaiseError = True
11        GoTo ExitSub
End Sub

Function CreateTempName(Prefix As String) As String
1         On Error GoTo Failed
    #If Mac Then
2             CreateTempName = mktemp(Prefix & "-XXXX")
    #Else
              Dim fso As Object
3             Set fso = CreateObject("Scripting.FileSystemObject")
              
              Dim RandomName As String
4             RandomName = fso.GetTempName()
5             RandomName = Mid(RandomName, 4, Len(RandomName) - 8)
6             CreateTempName = Prefix & "-" & RandomName
    #End If
7         Exit Function
          
Failed:
8         CreateTempName = Prefix
End Function

Function GetTempFolder(Optional AllowEnvironOverride As Boolean = True) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Static TempFolderPath As String
3         If Len(TempFolderPath) = 0 Then
              #If Mac Then
4                 If Val(Application.Version) >= 15 Then
5                     TempFolderPath = MacScript("return (POSIX path of (path to temporary items)) as string")
6                 Else
7                     TempFolderPath = MacScript("return (path to temporary items) as string")
8                 End If
              #Else
                  ' Get Temp Folder
                  ' See http://www.pcreview.co.uk/forums/thread-934893.php
                  Dim ret As Long
9                 TempFolderPath = String$(255, 0)
10                ret = GetTempPath(255, TempFolderPath)
11                If ret <> 0 Then
12                    TempFolderPath = Left(TempFolderPath, ret)
13                Else
14                    TempFolderPath = vbNullString
15                End If
              #End If
              
              ' Andres Sommerhoff (ASL) - Country: Chile
              ' Allow user to specify a temp path using an environment variable
              ' This can also be a workaround to avoid problem with spaces in the temp path.
16            If AllowEnvironOverride Then
17                If Len(Environ("OpenSolverTempPath")) > 0 Then
18                    TempFolderPath = Environ("OpenSolverTempPath")
19                End If
20            End If

              ' Append OpenSolver to dir
21            If Len(TempFolderPath) <> 0 Then
22                TempFolderPath = JoinPaths(TempFolderPath, CreateTempName("OpenSolver"))
23                If Right(TempFolderPath, 1) <> Application.PathSeparator Then TempFolderPath = TempFolderPath & Application.PathSeparator
24            End If
25        End If
          
26        If Not FileOrDirExists(TempFolderPath) Then MkDir TempFolderPath
27        GetTempFolder = TempFolderPath

ExitFunction:
28        If RaiseError Then RethrowError
29        Exit Function

ErrorHandler:
30        If Not ReportError("OpenSolverIO", "GetTempFolder") Then Resume
31        RaiseError = True
32        GoTo ExitFunction
End Function

Function GetTempFilePath(FileName As String, ByRef Path As String) As Boolean
1         GetTempFilePath = GetExistingFilePathName(GetTempFolder, FileName, Path)
End Function

Sub CreateScriptFile(ByRef ScriptFilePath As String, FileContents As String, Optional EnableEcho As Boolean)
      ' Create a script file with the specified contents.
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         Open ScriptFilePath For Output As #1
          
          #If Win32 Then
              ' Add echo off for windows
4             If Not EnableEcho Then
5                 Print #1, "@echo off" & vbCrLf
6             End If
          #End If
7         Print #1, FileContents
8         Close #1
          
          #If Mac Then
9             Exec "chmod +x " & MakePathSafe(ScriptFilePath)
          #End If

ExitSub:
10        Close #1
11        If RaiseError Then RethrowError
12        Exit Sub

ErrorHandler:
13        If Not ReportError("OpenSolverIO", "CreateScriptFile") Then Resume
14        RaiseError = True
15        GoTo ExitSub
End Sub

Sub SetCurrentDirectory(NewPath As String)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          #If Mac Then
3             ChDir NewPath
          #Else
4             SetCurrentDirectoryA NewPath
          #End If

ExitSub:
5         If RaiseError Then RethrowError
6         Exit Sub

ErrorHandler:
7         If Not ReportError("OpenSolverIO", "SetCurrentDirectory") Then Resume
8         RaiseError = True
9         GoTo ExitSub
End Sub

Function GetAddInIfExists(AddInObj As Excel.AddIn, Title As String) As Boolean
          ' See http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.addins.aspx
1         Set AddInObj = Nothing
2         On Error Resume Next
3         Set AddInObj = Application.AddIns.Item(Title)
4         GetAddInIfExists = (Err = 0)
End Function

Function GetOpenSolverAddInIfExists(OpenSolverAddIn As Excel.AddIn) As Boolean
          Dim Title As String
1         Title = "OpenSolver"
2         If IsMac And Val(Application.Version) < 15 Then
              ' On Mac 2011, the Application.AddIns collection is indexed by filename.ext rather than just filename as on Windows
3             Title = Title & ".xlam"
4         End If
5         GetOpenSolverAddInIfExists = GetAddInIfExists(OpenSolverAddIn, Title)
End Function

Function ChangeOpenSolverAutoload(loadAtStartup As Boolean) As Boolean
          ' NOTE: If Mac and no workbooks are open, this will crash
1         ChangeOpenSolverAutoload = False

2         If loadAtStartup Then  ' User is changing from True to False
3             If MsgBox("This will configure Excel to automatically load the OpenSolver add-in (from its current location) when Excel starts.  Continue?", vbOKCancel) <> vbOK Then GoTo ExitSub
4         Else ' User is turning off auto load
5             If MsgBox("This will re-configure Excel's Add-In settings so that OpenSolver does not load automatically at startup. You will need to re-load OpenSolver when you wish to use it next, or re-enable it using Excel's Add-In settings." & vbCrLf & vbCrLf _
                        & "WARNING: OpenSolver will also be shut down right now by Excel, and so will disappear immediately. No data will be lost." & vbCrLf & vbCrLf _
                        & "Continue?", vbOKCancel) <> vbOK Then GoTo ExitSub
6         End If
          
          ' On older versions of Excel, Add-ins can only be added if we have at least one workbook open; see http://vbadud.blogspot.com/2007/06/excel-vba-install-excel-add-in-xla-or.html
          Dim TempBook As Workbook
7         If Workbooks.Count = 0 Then Set TempBook = Workbooks.Add
          
          Dim OpenSolverAddIn As Excel.AddIn
8         If Not GetOpenSolverAddInIfExists(OpenSolverAddIn) Then
9             Set OpenSolverAddIn = Application.AddIns.Add(ThisWorkbook.FullName, False)
10        End If
          
          ' Closing the temp book can throw an error on Mac, we just ignore
11        On Error Resume Next
12        If Not TempBook Is Nothing Then TempBook.Close
13        On Error GoTo 0
          
14        If OpenSolverAddIn Is Nothing Then
15            MsgBox "Unable to load or access addin " & ThisWorkbook.FullName
16        Else
17            OpenSolverAddIn.Installed = loadAtStartup ' OpenSolver will quit immediately when this is set to false, unless a reference is set to OpenSolver
18            ChangeOpenSolverAutoload = True
19        End If
ExitSub:
End Function

Public Sub AppendFile(ByVal Path As String, ByVal txt As String)
    Dim FileNum As Integer
    
    FileNum = FreeFile()
    
    Open Path For Append As #FileNum
        Print #FileNum, txt
    Close #FileNum
End Sub
