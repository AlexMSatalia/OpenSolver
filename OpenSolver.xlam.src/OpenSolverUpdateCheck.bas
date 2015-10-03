Attribute VB_Name = "OpenSolverUpdateCheck"
Private HasCheckedForUpdate As Boolean

Private DoSilentFail As Boolean
Private DoWaitForResponse As Boolean
Private AvoidPromptForBeta As Boolean

Const OpenSolverRegName = "OpenSolver"
Const PreferencesRegName = "Preferences"
Const CheckForUpdatesRegName = "CheckForUpdates"
Const CheckForBetaUpdatesRegName = "CheckForBetaUpdates"
Const LastUpdateCheckRegName = "LastUpdateCheck"
Const GuidRegName = "Guid"

Const MinTimeBetweenChecks As Double = 1 ' 1 day between checks

#If Mac Then
    Private Const UpdateLogName = "update.log"
    Dim LogFilePath As String
    Dim NumChecks As Long
    Const MaxTime As Long = 10
#End If
 
Function GetUserAgent() As String
    GetUserAgent = OSFamily() & "/" & OSVersion() & "x" & OSBitness() & " " & _
                   "Excel/" & Application.Version & "x" & ExcelBitness() & " " & _
                   "OpenSolver/" & sOpenSolverVersion & "x" & OpenSolverDistribution() '& " " & _
                   '"GUID/" & GetGuid()
End Function

Private Function GetPageUrl() As String
    If DEBUG_MODE Then
        ' The link below is a useful tool for async testing
        ' It delays the response of the server (2s by default). Add "?sleep=5" to change timeout to 5s etc
        GetPageUrl = "https://fake-response.appspot.com/"
    ElseIf GetBetaUpdateSetting(AvoidPromptForBeta) Then
        GetPageUrl = "http://opensolver.org/download/731/"
    Else
        GetPageUrl = "http://opensolver.org/download/726/"
    End If
End Function

Sub InitialiseUpdateCheck(Optional ByVal SilentFail As Boolean = False, Optional WaitForResponse As Boolean = False)
    HasCheckedForUpdate = True
    SetLastCheckTime Now
    
    DoSilentFail = SilentFail
    DoWaitForResponse = WaitForResponse
    If DoWaitForResponse Then Application.Cursor = xlWait
    
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
    xmlHttpRequest.Open "GET", GetPageUrl(), True
    xmlHttpRequest.setRequestHeader "User-Agent", GetUserAgent()
    xmlHttpRequest.send ""
    Exit Sub

FailedState:
    MsgBox Err.Number & ": " & Err.Description
End Sub

' On Mac, we use `cURL` via command line which is included by default
#If Mac Then
Private Function InitialiseUpdateCheck_Mac() As String
    Dim Command As String
    
    If GetTempFilePath(UpdateLogName, LogFilePath) Then DeleteFileAndVerify (LogFilePath)

    ' -L follows redirects, -m sets Max Time
    Command = "curl -L" & _
              " -m " & MaxTime & _
              " -o " & MakePathSafe(LogFilePath) & _
              " -A " & Quote(GetUserAgent()) & _
              " " & GetPageUrl()
    RunExternalCommand Command, "", Hide, False
    
    NumChecks = 0
    
    If DoWaitForResponse Then
        CheckForCompletion_Mac
    Else
        Application.OnTime Now + TimeSerial(0, 0, 1), "CheckForCompletion_Mac"
    End If
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
        
        ' If the log file is not empty then we may be finished
        If Len(Response) > 0 Then CheckAgain = False
    End If
    
    If CheckAgain And NumChecks < MaxTime Then
        NumChecks = NumChecks + 1
        If DoWaitForResponse Then
            mSleep 1000  ' 1 second
            CheckForCompletion_Mac
        Else
            Application.OnTime Now + TimeSerial(0, 0, 1), "CheckForCompletion_Mac"
        End If
    Else
        CompleteUpdateCheck Response
    End If
End Sub
#End If

' Function to run once our request has completed
Sub CompleteUpdateCheck(Response As String)
    On Error GoTo ConnectionError
    If Len(Response) < 5 Or _
       Mid(Response, 2, 1) <> "." Or _
       Mid(Response, 4, 1) <> "." Then
        GoTo ConnectionError
    End If

    Dim LatestNumbers() As String, CurrentNumbers() As String
    LatestNumbers() = Split(Response, ".")
    CurrentNumbers() = Split(sOpenSolverVersion, ".")
    
    Dim UpdateAvailable As Boolean, LatestNumber As Long, CurrentNumber As Long
    UpdateAvailable = False
    For i = 0 To 2
        LatestNumber = CInt(LatestNumbers(i))
        CurrentNumber = CInt(CurrentNumbers(i))
        If LatestNumber > CurrentNumber Then
            UpdateAvailable = True
            Exit For
        ElseIf LatestNumber < CurrentNumber Then
            ' No need to check any more
            Exit For
        End If
    Next
    
    Application.Cursor = xlDefault
    
    If UpdateAvailable Then
        Dim frmUpdateNotification As FUpdateNotification
        Set frmUpdateNotification = New FUpdateNotification
        frmUpdateNotification.ShowUpdate Response
        Unload frmUpdateNotification
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
        Dim SettingWasMissing As Boolean, DoCheck As Boolean
        ' Get the entry, and show the update settings form if missing
        DoCheck = GetUpdateSetting(False, SettingWasMissing)
        
        ' If the setting was missing, then we have shown the update settings form already
        ' We don't want to show it again when we get the beta entry
        AvoidPromptForBeta = SettingWasMissing
        
        If DoCheck Then
            ' Check time since last check
            If Now - GetLastCheckTime() > MinTimeBetweenChecks Then
                InitialiseUpdateCheck True
            End If
        End If
    End If
End Sub

Public Function GetUpdateSetting(Optional SilentFail As Boolean = True, Optional Missing As Boolean) As Boolean
    Dim result As Variant
    ' From rondebruin.nl: The GetSetting default argument can't be an empty string on Mac
    result = GetSetting(OpenSolverRegName, PreferencesRegName, CheckForUpdatesRegName, "?")
    
    If result = "?" Then
        Missing = True
        ' Handle a missing entry
        If SilentFail Then
            ' In silent mode, return false without saving anything
            result = False
        Else
            ' Otherwise, show the dialog and get the setting
            Dim frmUpdateSettings As FUpdateSettings
            Set frmUpdateSettings = New FUpdateSettings
            frmUpdateSettings.Show
            Unload frmUpdateSettings
            result = GetUpdateSetting(True)
        End If
    Else
        Missing = False
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

Public Function GetBetaUpdateSetting(Optional SilentFail As Boolean = True, Optional Missing As Boolean) As Boolean
    Dim result As Variant
    ' From rondebruin.nl: The GetSetting default argument can't be an empty string on Mac
    result = GetSetting(OpenSolverRegName, PreferencesRegName, CheckForBetaUpdatesRegName, "?")
    
    If result = "?" Then
        Missing = True
        ' Handle a missing entry
        If SilentFail Then
            ' In silent mode, return false without saving anything
            result = False
        Else
            ' Otherwise, show the dialog and get the setting
            Dim frmUpdateSettings As FUpdateSettings
            Set frmUpdateSettings = New FUpdateSettings
            frmUpdateSettings.Show
            Unload frmUpdateSettings
            result = GetUpdateSetting(True)
        End If
    Else
        Missing = False
    End If
    
    GetBetaUpdateSetting = CBool(result)
End Function

Public Sub SaveBetaUpdateSetting(BetaUpdateSetting As Boolean)
    SaveSetting OpenSolverRegName, PreferencesRegName, CheckForBetaUpdatesRegName, BetaUpdateSetting
End Sub

Private Sub DeleteBetaUpdateSetting()
    DeleteSetting OpenSolverRegName, PreferencesRegName, CheckForBetaUpdatesRegName
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

Private Sub ResetHasChecked()
    HasCheckedForUpdate = False
End Sub

Private Function GetGuid() As String
    ' From rondebruin.nl: The GetSetting default argument can't be an empty string on Mac
    Dim result As Variant
    result = GetSetting(OpenSolverRegName, PreferencesRegName, GuidRegName, "?")
    
    If result = "?" Then
        result = MakeGuid()
        SetGuid CStr(result)
    End If
    
    GetGuid = CStr(result)
End Function

Private Sub SetGuid(Guid As String)
    SaveSetting OpenSolverRegName, PreferencesRegName, GuidRegName, Guid
End Sub

Private Sub DeleteGuid()
    DeleteSetting OpenSolverRegName, PreferencesRegName, GuidRegName
End Sub

Private Function MakeGuid() As String
    #If Mac Then
        MakeGuid = Application.Clean(ReadExternalCommandOutput("uuidgen"))
    #Else
        MakeGuid = Mid$(CreateObject("Scriptlet.TypeLib").Guid, 2, 36)
    #End If
End Function

Private Sub ResetAllUpdateSettings()
    DeleteBetaUpdateSetting
    DeleteUpdateSetting
    DeleteLastCheckTime
    HasCheckedForUpdate = False
End Sub
