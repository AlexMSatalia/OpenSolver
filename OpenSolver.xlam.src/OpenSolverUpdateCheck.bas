Attribute VB_Name = "OpenSolverUpdateCheck"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private HasCheckedForUpdate As Boolean

Private DoSilentFail As Boolean
Private DoWaitForResponse As Boolean
Private AvoidPromptForBeta As Boolean

Private Const OpenSolverRegName = "OpenSolver"
Private Const PreferencesRegName = "Preferences"
Private Const CheckForUpdatesRegName = "CheckForUpdates"
Private Const CheckForBetaUpdatesRegName = "CheckForBetaUpdates"
Private Const LastUpdateCheckRegName = "LastUpdateCheck"
Private Const FirstUpdateCheckRegName = "FirstUpdateCheck"
Private Const GuidRegName = "Guid"

' From rondebruin.nl:
' The GetSetting default argument can't be an emptystring on Mac
Private Const VALUE_IF_MISSING As String = "?"

Private Const MinTimeBetweenChecks As Double = 1 ' 1 day between checks
Private Const MinTimeBeforeCheck As Double = 1 ' 1 day before checks begin

#If Mac Then
    Private Const UpdateLogName = "update.log"
    Private LogFilePath As String
    Private NumChecks As Long
    Private Const MaxTime As Long = 10
#End If
 
Function GetUserAgent() As String
    GetUserAgent = EnvironmentString() ' & " " & "GUID/" & GetGuid()
End Function

Private Function GetPageUrl() As String
    If False Then
        ' The link below is a useful tool for async testing
        ' It delays the response of the server (2s by default).
        ' Add "?sleep=5" to change timeout to 5s etc
        GetPageUrl = "https://fake-response.appspot.com/"
    ElseIf GetBetaUpdateSetting(AvoidPromptForBeta) Then
        GetPageUrl = "http://opensolver.org/download/731/"
    Else
        GetPageUrl = "http://opensolver.org/download/726/"
    End If
End Function

Sub InitialiseUpdateCheck(Optional ByVal SilentFail As Boolean = False, _
        Optional WaitForResponse As Boolean = False)
    HasCheckedForUpdate = True
    SetLastCheckTime Now
    
    DoSilentFail = SilentFail
    DoWaitForResponse = WaitForResponse
    If DoWaitForResponse Then Application.Cursor = xlWait
    
    ' We initiate a request for the version info.
    ' This check should be asynchronous, and fire "CompleteUpdateCheck" when
    ' the response is returned
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
    
    If GetTempFilePath(UpdateLogName, LogFilePath) Then
        DeleteFileAndVerify (LogFilePath)
    End If

    ' -L follows redirects, -m sets Max Time
    Command = "curl -L" & _
              " -m " & MaxTime & _
              " -o " & MakePathSafe(LogFilePath) & _
              " -A " & Quote(GetUserAgent()) & _
              " " & GetPageUrl()
    ExecAsync Command
    
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
            Application.OnTime Now + TimeSerial(0, 0, 1), _
                               "CheckForCompletion_Mac"
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
    
    Dim UpdateAvailable As Boolean, i As Long
    UpdateAvailable = False
    For i = 0 To 2
        Dim LatestNumber As Long, CurrentNumber As Long
        LatestNumber = CLng(LatestNumbers(i))
        CurrentNumber = CLng(CurrentNumbers(i))
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
        MsgBox "No updates for OpenSolver are available at this time.", _
               vbOKOnly, "OpenSolver - Update Check"
    End If
    
ExitSub:
    Exit Sub
    
ConnectionError:
    If Not DoSilentFail Then
        MsgBox "The update checker was unable to determine the latest " _
             & "version of OpenSolver. Please try again later."
    End If
    GoTo ExitSub
End Sub

Sub AutoUpdateCheck()
    ' Don't check the saved setting if we have already run the checker
    If Not HasCheckedForUpdate Then
        ' Make sure we don't keep asking the user for their preference if
        ' we have already shown them the settings form
        HasCheckedForUpdate = True
        
        Dim FirstUpdateCheckTime As Double
        FirstUpdateCheckTime = GetFirstCheckTime()
        ' Record if this is the first update check
        If FirstUpdateCheckTime = 0 Then
            SetFirstCheckTime Now
            Exit Sub
        End If
        
        ' Make sure 24 hours has passed since the first update check (ref #245)
        If Now - FirstUpdateCheckTime < MinTimeBeforeCheck Then
            Exit Sub
        End If
        
        Dim SettingWasMissing As Boolean, DoCheck As Boolean
        ' Get the entry, and show the update settings form if missing
        DoCheck = GetUpdateSetting(False, SettingWasMissing)
        
        ' If the setting was missing, then we have shown the update settings
        ' form, and we don't want to show it again when we get the beta entry
        AvoidPromptForBeta = SettingWasMissing
        
        If DoCheck Then
            ' Check time since last check
            If Now - GetLastCheckTime() > MinTimeBetweenChecks Then
                InitialiseUpdateCheck True
            End If
        End If
    End If
End Sub

Private Function GetUpdateRegName(Beta As Boolean) As String
    GetUpdateRegName = IIf(Beta, CheckForBetaUpdatesRegName, _
                                 CheckForUpdatesRegName)
End Function

Public Function GetUpdateSetting(Optional SilentFail As Boolean = True, _
        Optional Missing As Boolean, _
        Optional Beta As Boolean = False) As Boolean
        
    Dim result As Variant
    result = GetSetting(OpenSolverRegName, PreferencesRegName, _
                        GetUpdateRegName(Beta), VALUE_IF_MISSING)
    
    If result = VALUE_IF_MISSING Then
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
    
    ' Check for stable updates only by default
    Dim DefaultSetting As Boolean
    DefaultSetting = IIf(Beta, False, True)
    GetUpdateSetting = SafeCBool(result, DefaultSetting)
End Function
Public Sub SaveUpdateSetting(UpdateSetting As Boolean, _
        Optional Beta As Boolean = False)
    SaveSetting OpenSolverRegName, PreferencesRegName, _
                GetUpdateRegName(Beta), BoolToInt(UpdateSetting)
End Sub
Private Sub DeleteUpdateSetting(Optional Beta As Boolean = False)
    On Error Resume Next
    DeleteSetting OpenSolverRegName, PreferencesRegName, GetUpdateRegName(Beta)
End Sub

Public Function GetBetaUpdateSetting(Optional SilentFail As Boolean = True, _
        Optional Missing As Boolean) As Boolean
    GetBetaUpdateSetting = GetUpdateSetting(SilentFail, Missing, Beta:=True)
End Function
Public Sub SaveBetaUpdateSetting(UpdateSetting As Boolean)
    SaveUpdateSetting UpdateSetting, Beta:=True
End Sub
Private Sub DeleteBetaUpdateSetting()
    DeleteUpdateSetting Beta:=True
End Sub

Private Function GetLastCheckTime() As Double
    Dim result As Variant
    result = GetSetting(OpenSolverRegName, PreferencesRegName, _
                        LastUpdateCheckRegName, 0)
    
    GetLastCheckTime = Val(result)
End Function
Private Sub SetLastCheckTime(CheckTime As Double)
    SaveSetting OpenSolverRegName, PreferencesRegName, _
                LastUpdateCheckRegName, StrExNoPlus(CheckTime)
End Sub
Private Sub DeleteLastCheckTime()
    On Error Resume Next
    DeleteSetting OpenSolverRegName, PreferencesRegName, LastUpdateCheckRegName
End Sub

Private Function GetFirstCheckTime() As Double
    Dim result As Variant
    result = GetSetting(OpenSolverRegName, PreferencesRegName, _
                        FirstUpdateCheckRegName, 0)
    
    GetFirstCheckTime = Val(result)
End Function
Private Sub SetFirstCheckTime(CheckTime As Double)
    SaveSetting OpenSolverRegName, PreferencesRegName, _
                FirstUpdateCheckRegName, StrExNoPlus(CheckTime)
End Sub
Private Sub DeleteFirstCheckTime()
    On Error Resume Next
    DeleteSetting OpenSolverRegName, PreferencesRegName, FirstUpdateCheckRegName
End Sub

Private Sub ResetHasChecked()
    HasCheckedForUpdate = False
End Sub

Private Function GetGuid() As String
    ' From rondebruin.nl:
    ' The GetSetting default argument can't be an empty string on Mac
    Dim result As Variant
    result = GetSetting(OpenSolverRegName, PreferencesRegName, GuidRegName, _
                        VALUE_IF_MISSING)
    
    If result = VALUE_IF_MISSING Then
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
        MakeGuid = Application.Clean(ExecCapture("uuidgen"))
    #Else
        MakeGuid = Mid$(CreateObject("Scriptlet.TypeLib").Guid, 2, 36)
    #End If
End Function

Private Sub ResetAllUpdateSettings()
    DeleteBetaUpdateSetting
    DeleteUpdateSetting
    DeleteLastCheckTime
    DeleteFirstCheckTime
    HasCheckedForUpdate = False
End Sub
