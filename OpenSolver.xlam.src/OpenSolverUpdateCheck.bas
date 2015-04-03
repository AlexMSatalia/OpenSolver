Attribute VB_Name = "OpenSolverUpdateCheck"
Const FilesPageUrl = "http://opensolver.org/latest-version/"
' The link below is a useful tool for testing async-ness, timeouts and slow connections
' It delays the response of the server (2s by default). Add "?sleep=5" to change timeout to 5s etc
' Const FilesPageUrl = "https://fake-response.appspot.com/"
Private HasCheckedForUpdate As Boolean

Private DoSilentFail As Boolean

Const OpenSolverRegName = "OpenSolver"
Const PreferencesRegName = "Preferences"
Const CheckForUpdatesRegName = "CheckForUpdates"
Const LastUpdateCheckRegName = "LastUpdateCheck"

Const MinTimeBetweenChecks As Double = 1 ' 1 day between checks

#If Mac Then
    Private Const UpdateLogName = "update.log"
    Dim LogFilePath As String
    Dim NumChecks As Long
    Const MaxTime As Long = 10
#End If

Sub InitialiseUpdateCheck(Optional ByVal SilentFail As Boolean = False)
    HasCheckedForUpdate = True
    SetLastCheckTime Now
    
    DoSilentFail = SilentFail
    
    ' We initiate a request for the version info.
    ' This check should be asynchronous, and fire "CompleteUpdateCheck" when the response is returned
    #If Mac Then
        InitialiseUpdateCheck_Mac
    #Else
        InitialiseUpdateCheck_Windows
    #End If
End Sub

Sub InitialiseUpdateCheck_Windows()
'http://dailydoseofexcel.com/archives/2006/10/09/async-xmlhttp-calls/
    Dim xmlHttpRequest As Object ' MSXML2.XMLHTTP
    Set xmlHttpRequest = CreateObject("MSXML2.ServerXMLHTTP")
    ' Set timeout limits - 5 secs for each part of the request
    xmlHttpRequest.setTimeouts 5000, 5000, 5000, 5000
    
    On Error GoTo FailedState

    ' Create an instance of the wrapper class.
    Dim MyXmlAsyncHandler As XmlAsyncHandler
    Set MyXmlAsyncHandler = New XmlAsyncHandler
    MyXmlAsyncHandler.Initialize xmlHttpRequest, "CompleteUpdateCheck"

    ' Assign the wrapper class object to onreadystatechange.
    xmlHttpRequest.OnReadyStateChange = MyXmlAsyncHandler

    ' Get the page asynchronously.
    xmlHttpRequest.Open "GET", FilesPageUrl, True
    xmlHttpRequest.send ""
    Exit Sub

FailedState:
    MsgBox Err.Number & ": " & Err.Description
End Sub

' On Mac, we use `cURL` via command line which is included by default
#If Mac Then
Private Function InitialiseUpdateCheck_Mac() As String
    Dim Cmd As String
    Dim result As String
    Dim ExitCode As Long
    
    If GetTempFilePath(UpdateLogName, LogFilePath) Then DeleteFileAndVerify (LogFilePath)

    ' -L follows redirects, -m sets Max Time
    Cmd = "curl -L -m " & MaxTime & " " & FilesPageUrl
    RunExternalCommand Cmd, LogFilePath, Hide, False
    
    NumChecks = 0
    
    Application.OnTime Now + TimeSerial(0, 0, 1), "CheckForCompletion_Mac"
End Function
#End If

#If Mac Then
Public Sub CheckForCompletion_Mac()
    Dim CheckAgain As Boolean
    CheckAgain = True
    
    Dim Response As String
    
    If FileOrDirExists(LogFilePath) Then
        Open LogFilePath For Input As #1
            Response = Input$(LOF(1), 1)
        Close #1
        
        If Len(Response) > 0 Then CheckAgain = False
    End If
    
    If CheckAgain And NumChecks < MaxTime Then
        NumChecks = NumChecks + 1
        Application.OnTime Now + TimeSerial(0, 0, 1), "CheckForCompletion_Mac"
    Else
        CompleteUpdateCheck Response
    End If
End Sub
#End If

' Gets version number of current release from our response text.
Private Function GetLatestOpenSolverVersion(Response As String) As String
    GetLatestOpenSolverVersion = ""
    
    ' We are looking for the following message:
    '   "Latest version: (x.x.x)"
    Dim startString As String
    startString = "Latest version: "
    
    Dim start As Long, openingParen As Long, closingParen As Long
    start = InStrText(Response, startString)
    If start > 0 Then
        openingParen = InStr(start, Response, "(") + 1
        closingParen = InStr(openingParen, Response, ")")
        GetLatestOpenSolverVersion = Mid(Response, openingParen, closingParen - openingParen)
    End If
End Function

' Function to run once our request has completed
Sub CompleteUpdateCheck(Response As String)
    Dim LatestVersion As String
    LatestVersion = GetLatestOpenSolverVersion(Response)
    If Len(LatestVersion) = 0 Then GoTo ConnectionError

    Dim LatestNumbers() As String, CurrentNumbers() As String
    LatestNumbers() = Split(LatestVersion, ".")
    CurrentNumbers() = Split(sOpenSolverVersion, ".")
    
    Dim UpdateAvailable As Boolean
    UpdateAvailable = False
    For i = 0 To 2
        If CInt(LatestNumbers(i)) > CInt(CurrentNumbers(i)) Then
            UpdateAvailable = True
            Exit For
        End If
    Next
    
    If UpdateAvailable Then
        frmUpdate.ShowUpdate LatestVersion
    ElseIf Not DoSilentFail Then
        MsgBox "No updates for OpenSolver are available at this time.", vbOKOnly, "OpenSolver - Update Check"
    End If
    
ExitSub:
    Exit Sub
    
ConnectionError:
    If Not SilentFail Then
        MsgBox "The update checker was unable to determine the latest version of OpenSolver. Please try again later."
    End If
    GoTo ExitSub
End Sub

Sub AutoUpdateCheck()
    ' Don't check the saved setting if we have already run the checker
    If Not HasCheckedForUpdate Then
        If GetUpdateSetting() Then
            ' Check time since last check
            If Now - GetLastCheckTime() > MinTimeBetweenChecks Then
                InitialiseUpdateCheck True
            End If
        End If
    End If
End Sub

Public Function GetUpdateSetting() As Boolean
    ' From rondebruin.nl: The GetSetting default argument can't be an empty string on Mac
    Dim result As Variant
    result = GetSetting(OpenSolverRegName, PreferencesRegName, CheckForUpdatesRegName, "?")
    
    ' If registry key is missing, then check with user whether to autocheck
    If result = "?" Then
        result = MsgBox("Would you like OpenSolver to automatically check for updates? " & vbNewLine & vbNewLine & _
                        "You can change this option at any time by going to ""About OpenSolver"". " & _
                        "You can also run update checks manually from there.", vbYesNoCancel, "OpenSolver - Check for Updates?")
        If result = vbCancel Then
            ' Set result to false (without saving it) so that the check doesn't run this time
            result = False
        Else
            ' Save result
            result = (result = vbYes)
            SaveUpdateSetting (CBool(result))
        End If
    End If
    
    GetUpdateSetting = CBool(result)
End Function

Public Sub SaveUpdateSetting(UpdateSetting As Boolean)
    SaveSetting OpenSolverRegName, PreferencesRegName, CheckForUpdatesRegName, UpdateSetting
End Sub

' Useful for testing update check
Private Sub DeleteUpdateSetting()
    DeleteSetting OpenSolverRegName, PreferencesRegName, CheckForUpdatesRegName
End Sub

Private Function GetLastCheckTime() As Double
    Dim result As Variant
    result = GetSetting(OpenSolverRegName, PreferencesRegName, LastUpdateCheckRegName, 0)
    
    GetLastCheckTime = CDbl(result)
End Function

Private Sub SetLastCheckTime(CheckTime As Double)
    SaveSetting OpenSolverRegName, PreferencesRegName, LastUpdateCheckRegName, CStr(CheckTime)
End Sub

Private Sub DeleteLastCheckTime()
    DeleteSetting OpenSolverRegName, PreferencesRegName, LastUpdateCheckRegName
End Sub
