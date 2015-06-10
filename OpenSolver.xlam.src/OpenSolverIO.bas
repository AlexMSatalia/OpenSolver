Attribute VB_Name = "OpenSolverIO"
Option Explicit

' All interactions with file system and sheets in Excel

#If Win32 Then
    #If VBA7 Then
        Private Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
        Private Declare PtrSafe Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long
    #Else
        Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
        Private Declare Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long
    #End If
#End If

Sub CheckLocationValid()
          If StringHasUnicode(ThisWorkbook.Path) Then
              MsgBoxEx "The path that OpenSolver is being loaded from contains unicode characters. " & _
                       "This means the solvers are very unlikely to work. " & _
                       "Please move the OpenSolver folder so that there are no unicode characters in the complete path to the folder " & vbNewLine & vbNewLine & _
                       "The OpenSolver folder is currently located at: " & vbNewLine & _
                       ThisWorkbook.Path
          End If
End Sub

Function CheckWorksheetAvailable(Optional SuppressDialogs As Boolean = False, Optional ThrowError As Boolean = False) As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

464       CheckWorksheetAvailable = False
          ' Check there is a workbook
465       If Application.Workbooks.Count = 0 Then
466           If ThrowError Then Err.Raise Number:=OpenSolver_NoWorkbook, Description:="No active workbook available."
467           If Not SuppressDialogs Then MsgBox "Error: No active workbook available", , "OpenSolver" & sOpenSolverVersion & " Error"
468           GoTo ExitFunction
469       End If
          ' Check we can access the worksheet
          Dim w As Worksheet
470       On Error Resume Next
471       Set w = ActiveWorkbook.ActiveSheet
472       If Err.Number <> 0 Then
              On Error GoTo ErrorHandler
473           If ThrowError Then Err.Raise Number:=OpenSolver_NoWorksheet, Description:="The active sheet is not a worksheet."
474           If Not SuppressDialogs Then MsgBox "Error: The active sheet is not a worksheet.", , "OpenSolver" & sOpenSolverVersion & " Error"
475           GoTo ExitFunction
476       End If

477       CheckWorksheetAvailable = True

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverIO", "CheckWorksheetAvailable") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Sub GetActiveBookAndSheetIfMissing(book As Workbook, Optional sheet As Worksheet)
    If book Is Nothing Then Set book = ActiveWorkbook
    If sheet Is Nothing Then Set sheet = book.ActiveSheet
End Sub

Function MakeNewSheet(namePrefix As String, OverwriteExisting As Boolean) As String
          Dim NeedSheet As Boolean, newSheet As Worksheet, nameSheet As String, i As Long
          Dim ScreenStatus As Boolean
          ScreenStatus = Application.ScreenUpdating
668       Application.ScreenUpdating = False
          Dim s As String
          On Error Resume Next
669       s = Sheets(namePrefix).Name
670       If Err.Number <> 0 Then
671           Set newSheet = Sheets.Add
672           newSheet.Name = namePrefix
673           nameSheet = namePrefix
675       Else
677           If OverwriteExisting Then
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
693           End If
694       End If
695       MakeNewSheet = nameSheet
696       Application.ScreenUpdating = ScreenStatus
End Function

Function GetExistingFilePathName(Directory As String, FileName As String, ByRef pathName As String) As Boolean
462       pathName = JoinPaths(Directory, FileName)
463       GetExistingFilePathName = FileOrDirExists(pathName)
End Function

Function GetRootDriveName() As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          #If Mac Then
              Static DriveName As String
              ' We assume that the temp folder is on the root drive
              ' Seems reasonable, the user might be able to mess this up if they really try.
              If DriveName = "" Then
                  Dim Path As String
                  Path = GetTempFolder(False)
                  DriveName = left(Path, InStr(Path, ":") - 1)
              End If
              GetRootDriveName = DriveName
          #End If

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverIO", "GetRootDriveName") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function ConvertHfsPathToPosix(Path As String) As String
' Any direct file system access (using 'system' or in script files) on Mac requires
' that HFS-style paths are converted to normal POSIX paths. On Windows this
' function does nothing, so it can safely wrap all file system calls on any platform
' Input (HFS path):   "Macintosh HD:Users:jack:filename.txt"
' Output (POSIX path): "/Volumes/Macintosh HD/Users/jack/filename.txt"

          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          #If Mac Then
              ' Check we have an HFS path and not POSIX
738           If InStr(Path, "/") = 0 Then
                  Dim RootDriveName As String
                  RootDriveName = GetRootDriveName()
                  If left(Path, Len(RootDriveName)) = RootDriveName Then
                      ConvertHfsPathToPosix = Mid(Path, Len(RootDriveName) + 1)
                  Else
                      ' Prefix disk name with :Volumes:
739                   ConvertHfsPathToPosix = ":Volumes:" & Path
                  End If
                  ' Convert path delimiters
740               ConvertHfsPathToPosix = Replace(ConvertHfsPathToPosix, ":", "/")
741           Else
                  ' Path is already POSIX
742               ConvertHfsPathToPosix = Path
743           End If
          #Else
744           ConvertHfsPathToPosix = Path
          #End If

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverIO", "ConvertHfsPathToPosix") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function ConvertPosixPathToHfs(Path As String) As String
' Converts a POSIX path back to HFS
' Input (POSIX path): "/Volumes/Macintosh HD/Users/jack/filename.txt"
' Output (HFS path): "Macintosh HD:Users:jack:filename.txt"
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    #If Mac Then
        ' Make sure we have an HFS path
        If InStr(Path, ":") = 0 Then
            Const VolumePrefix = "/Volumes/"
            If left(Path, Len(VolumePrefix)) = VolumePrefix Then
                ' If the POSIX path starts with /Volumes/, then we keep the drive name after /Volumes/
                ConvertPosixPathToHfs = Mid(Path, Len(VolumePrefix) + 1)
            Else
                ' If the POSIX path starts at the root, we add the name of the root drive
                ConvertPosixPathToHfs = GetRootDriveName() & Path
            End If
            ' Convert Path delimiters
            ConvertPosixPathToHfs = Replace(ConvertPosixPathToHfs, "/", ":")
        Else
            ConvertPosixPathToHfs = Path
        End If
    #Else
        ConvertPosixPathToHfs = Path
    #End If

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverIO", "ConvertPosixPathToHfs") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

Function MakePathSafe(Path As String) As String
' Prepares a path for command-line invocation
    MakePathSafe = IIf(Len(Path) = 0, "", QuotePath(ConvertHfsPathToPosix(Path)))
End Function

Function JoinPaths(ParamArray Paths() As Variant) As String
          Dim i As Long
          For i = LBound(Paths) To UBound(Paths) - 1
              JoinPaths = JoinPaths & Paths(i) & IIf(right(Paths(i), 1) <> Application.PathSeparator, Application.PathSeparator, "")
          Next i
          JoinPaths = JoinPaths & Paths(UBound(Paths))
End Function

Function FileOrDirExists(pathName As String) As Boolean
' from http://www.vbaexpress.com/kb/getarticle.php?kb_id=559
    
          Dim iTemp As Long
105       On Error Resume Next
106       iTemp = GetAttr(pathName)
           
           'Check if error exists and set response appropriately
107       FileOrDirExists = (Err.Number = 0)
End Function

Sub DeleteFileAndVerify(FilePath As String)
      ' Deletes file and raises error if not successful
          On Error GoTo DeleteError
757       If FileOrDirExists(FilePath) Then Kill FilePath
758       If FileOrDirExists(FilePath) Then
759           GoTo DeleteError
760       End If
          Exit Sub
          
DeleteError:
          Err.Raise Number:=Err.Number, Description:="Unable to delete the file: " & FilePath & vbNewLine & vbNewLine & Err.Description
End Sub

Sub OpenFile(FilePath As String, notFoundMessage As String)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

762       If Not FileOrDirExists(FilePath) Then
763           Err.Raise OpenSolver_NoFile, Description:=notFoundMessage
764       Else
              ' Check that there is no workbook open with the same name
              Dim w As Workbook
765           On Error Resume Next
766           Set w = Workbooks(right(FilePath, InStr(FilePath, Application.PathSeparator)))
767           On Error GoTo ErrorHandler
768           Workbooks.Open FileName:=FilePath, ReadOnly:=True ' , Format:=Tabs
769       End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("OpenSolverIO", "OpenFile") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Function GetTempFolder(Optional AllowEnvironOverride As Boolean = True) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Static TempFolderPath As String
88        If Len(TempFolderPath) = 0 Then
              #If Mac Then
89                TempFolderPath = MacScript("return (path to temporary items) as string")

              #Else
                  ' Get Temp Folder
                  ' See http://www.pcreview.co.uk/forums/thread-934893.php
                  Dim ret As Long
90                TempFolderPath = String$(255, 0)
91                ret = GetTempPath(255, TempFolderPath)
92                If ret <> 0 Then
93                    TempFolderPath = left(TempFolderPath, ret)
94                    If right(TempFolderPath, 1) <> "\" Then TempFolderPath = TempFolderPath & "\"
95                Else
96                    TempFolderPath = ""
97                End If
              #End If
              ' Andres Sommerhoff (ASL) - Country: Chile
              ' Allow user to specify a temp path using an environment variable
              ' This can also be a workaround to avoid problem with spaces in the temp path.
98            If AllowEnvironOverride Then
                  If Environ("OpenSolverTempPath") <> "" Then
99                    TempFolderPath = Environ("OpenSolverTempPath")
                  End If
100           End If
101       End If
102       GetTempFolder = TempFolderPath

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverIO", "GetTempFolder") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function GetTempFilePath(FileName As String, ByRef Path As String) As Boolean
104       GetTempFilePath = GetExistingFilePathName(GetTempFolder, FileName, Path)
End Function

Sub CreateScriptFile(ByRef ScriptFilePath As String, FileContents As String, Optional EnableEcho As Boolean)
' Create a script file with the specified contents.
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

747       Open ScriptFilePath For Output As #1
          
          #If Win32 Then
              ' Add echo off for windows
748           If Not EnableEcho Then
749               Print #1, "@echo off" & vbCrLf
750           End If
          #End If
751       Print #1, FileContents
752       Close #1
          
          #If Mac Then
753           RunExternalCommand "chmod +x " & MakePathSafe(ScriptFilePath)
          #End If

ExitSub:
          Close #1
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("OpenSolverIO", "CreateScriptFile") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Sub SetCurrentDirectory(NewPath As String)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          #If Mac Then
735           ChDir NewPath
          #Else
736           SetCurrentDirectoryA NewPath
          #End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("OpenSolverIO", "SetCurrentDirectory") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Function GetAddInIfExists(AddInObj As Excel.AddIn, Title As String) As Boolean
          ' See http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.addins.aspx
3474      Set AddInObj = Nothing
3475      On Error Resume Next
3476      Set AddInObj = Application.AddIns.Item(Title)
3477      GetAddInIfExists = (Err = 0)
End Function

Function GetOpenSolverAddInIfExists(OpenSolverAddIn As Excel.AddIn) As Boolean
          Dim Title As String
3481      Title = "OpenSolver"
          #If Mac Then
              ' On Mac, the Application.AddIns collection is indexed by filename.ext rather than just filename as on Windows
3482          Title = Title & ".xlam"
          #End If
          GetOpenSolverAddInIfExists = GetAddInIfExists(OpenSolverAddIn, Title)
End Function

Function ChangeOpenSolverAutoload(loadAtStartup As Boolean) As Boolean
          ' NOTE: If Mac and no workbooks are open, this will crash
          ChangeOpenSolverAutoload = False

3495      If loadAtStartup Then  ' User is changing from True to False
3496          If MsgBox("This will configure Excel to automatically load the OpenSolver add-in (from its current location) when Excel starts.  Continue?", vbOKCancel) <> vbOK Then GoTo ExitSub
3497      Else ' User is turning off auto load
3498          If MsgBox("This will re-configure Excel's Add-In settings so that OpenSolver does not load automatically at startup. You will need to re-load OpenSolver when you wish to use it next, or re-enable it using Excel's Add-In settings." & vbCrLf & vbCrLf _
                        & "WARNING: OpenSolver will also be shut down right now by Excel, and so will disappear immediately. No data will be lost." & vbCrLf & vbCrLf _
                        & "Continue?", vbOKCancel) <> vbOK Then GoTo ExitSub
3499      End If
          
          ' On older versions of Excel, Add-ins can only be added if we have at least one workbook open; see http://vbadud.blogspot.com/2007/06/excel-vba-install-excel-add-in-xla-or.html
          Dim TempBook As Workbook
3501      If Workbooks.Count = 0 Then Set TempBook = Workbooks.Add
          
          Dim OpenSolverAddIn As Excel.AddIn
3502      If Not GetOpenSolverAddInIfExists(OpenSolverAddIn) Then
3503          Set OpenSolverAddIn = Application.AddIns.Add(ThisWorkbook.FullName, False)
3504      End If
          
          ' Closing the temp book can throw an error on Mac, we just ignore
          On Error Resume Next
3505      If Not TempBook Is Nothing Then TempBook.Close
          On Error GoTo 0
          
3506      If OpenSolverAddIn Is Nothing Then
3507          MsgBox "Unable to load or access addin " & ThisWorkbook.FullName
3508      Else
3509          OpenSolverAddIn.Installed = loadAtStartup ' OpenSolver will quit immediately when this is set to false, unless a reference is set to OpenSolver
              ChangeOpenSolverAutoload = True
          End If
ExitSub:
End Function

