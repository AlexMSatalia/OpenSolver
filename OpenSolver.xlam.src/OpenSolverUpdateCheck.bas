Attribute VB_Name = "OpenSolverUpdateCheck"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private HasCheckedForUpdate As Boolean

Private DoSilentFail As Boolean
Private DoWaitForResponse As Boolean
Private AvoidPromptForBeta As Boolean

Private Const CheckForUpdatesRegName = "CheckForUpdates"
Private Const CheckForBetaUpdatesRegName = "CheckForBetaUpdates"
Private Const LastUpdateCheckRegName = "LastUpdateCheck"
Private Const FirstUpdateCheckRegName = "FirstUpdateCheck"
Private Const GuidRegName = "Guid"

Private Const MinTimeBetweenChecks As Double = 1 ' 1 day between checks
Private Const MinTimeBeforeCheck As Double = 1 ' 1 day before checks begin

#If Mac Then
    Private Const UpdateLogName = "update.log"
    Private LogFilePath As String
    Private NumChecks As Long
    Private Const MaxTime As Long = 10
#End If
 
Function GetUserAgent() As String
1         GetUserAgent = EnvironmentString() ' & " " & "GUID/" & GetGuid()
End Function

Private Function GetPageUrl() As String
1         If False Then
              ' The link below is a useful tool for async testing
              ' It delays the response of the server (2s by default).
              ' Add "?sleep=5" to change timeout to 5s etc
2             GetPageUrl = "https://fake-response.appspot.com/"
3         ElseIf GetBetaUpdateSetting(AvoidPromptForBeta) Then
4             GetPageUrl = "http://opensolver.org/download/731/"
5         Else
6             GetPageUrl = "http://opensolver.org/download/726/"
7         End If
End Function

Sub InitialiseUpdateCheck(Optional ByVal SilentFail As Boolean = False, _
        Optional WaitForResponse As Boolean = False)
1         HasCheckedForUpdate = True
2         SetLastCheckTime Now
          
3         DoSilentFail = SilentFail
4         DoWaitForResponse = WaitForResponse
5         If DoWaitForResponse Then Application.Cursor = xlWait
          
          ' We initiate a request for the version info.
          ' This check should be asynchronous, and fire "CompleteUpdateCheck" when
          ' the response is returned
    #If Mac Then
6             InitialiseUpdateCheck_Mac
    #Else
7             InitialiseUpdateCheck_Windows
    #End If
End Sub

Sub InitialiseUpdateCheck_Windows()
      'http://dailydoseofexcel.com/archives/2006/10/09/async-xmlhttp-calls/
          Dim xmlHttpRequest As Object ' MSXML2.XMLHTTP
1         Set xmlHttpRequest = CreateObject("MSXML2.ServerXMLHTTP")
          ' Set timeout limits - 5 secs for each part of the request
2         xmlHttpRequest.setTimeouts 5000, 5000, 5000, 5000
          
3         On Error GoTo FailedState

          ' Create an instance of the wrapper class.
          Dim MyXmlAsyncHandler As XmlAsyncHandler
4         Set MyXmlAsyncHandler = New XmlAsyncHandler
5         MyXmlAsyncHandler.Initialize xmlHttpRequest, "CompleteUpdateCheck"

          ' Assign the wrapper class object to onreadystatechange.
6         xmlHttpRequest.OnReadyStateChange = MyXmlAsyncHandler

          ' Get the page asynchronously.
7         xmlHttpRequest.Open "GET", GetPageUrl(), True
8         xmlHttpRequest.setRequestHeader "User-Agent", GetUserAgent()
9         xmlHttpRequest.send vbNullString
10        Exit Sub

FailedState:
11        MsgBox Err.Number & ": " & Err.Description
End Sub

' On Mac, we use `cURL` via command line which is included by default
#If Mac Then
Private Function InitialiseUpdateCheck_Mac() As String
          Dim Command As String
          
1         If GetTempFilePath(UpdateLogName, LogFilePath) Then
2             DeleteFileAndVerify (LogFilePath)
3         End If

          ' -L follows redirects, -m sets Max Time
4         Command = "curl -L" & _
                    " -m " & MaxTime & _
                    " -o " & MakePathSafe(LogFilePath) & _
                    " -A " & Quote(GetUserAgent()) & _
                    " " & GetPageUrl()
5         ExecAsync Command
          
6         NumChecks = 0
          
7         If DoWaitForResponse Then
8             CheckForCompletion_Mac
9         Else
10            Application.OnTime Now + TimeSerial(0, 0, 1), "CheckForCompletion_Mac"
11        End If
End Function
#End If

#If Mac Then
Public Sub CheckForCompletion_Mac()
          Dim CheckAgain As Boolean
1         CheckAgain = True
          
          Dim Response As String
          
2         If FileOrDirExists(LogFilePath) Then
3             Open LogFilePath For Input As #1
4                 Response = Input$(LOF(1), 1)
5             Close #1
              
              ' If the log file is not empty then we may be finished
6             If Len(Response) > 0 Then CheckAgain = False
7         End If
          
8         If CheckAgain And NumChecks < MaxTime Then
9             NumChecks = NumChecks + 1
10            If DoWaitForResponse Then
11                mSleep 1000  ' 1 second
12                CheckForCompletion_Mac
13            Else
14                Application.OnTime Now + TimeSerial(0, 0, 1), _
                                     "CheckForCompletion_Mac"
15            End If
16        Else
17            CompleteUpdateCheck Response
18        End If
End Sub
#End If

' Function to run once our request has completed
Sub CompleteUpdateCheck(Response As String)
1         On Error GoTo ConnectionError
2         If Len(Response) < 5 Or _
             Mid(Response, 2, 1) <> "." Or _
             Mid(Response, 4, 1) <> "." Then
3             GoTo ConnectionError
4         End If

          Dim LatestNumbers() As String, CurrentNumbers() As String
5         LatestNumbers() = Split(Response, ".")
6         CurrentNumbers() = Split(sOpenSolverVersion, ".")
          
          Dim UpdateAvailable As Boolean, i As Long
7         UpdateAvailable = False
8         For i = 0 To 2
              Dim LatestNumber As Long, CurrentNumber As Long
9             LatestNumber = CLng(LatestNumbers(i))
10            CurrentNumber = CLng(CurrentNumbers(i))
11            If LatestNumber > CurrentNumber Then
12                UpdateAvailable = True
13                Exit For
14            ElseIf LatestNumber < CurrentNumber Then
                  ' No need to check any more
15                Exit For
16            End If
17        Next
          
18        Application.Cursor = xlDefault
          
19        If UpdateAvailable Then
              Dim frmUpdateNotification As FUpdateNotification
20            Set frmUpdateNotification = New FUpdateNotification
21            frmUpdateNotification.ShowUpdate Response
22            Unload frmUpdateNotification
23        ElseIf Not DoSilentFail Then
24            MsgBox "No updates for OpenSolver are available at this time.", _
                     vbOKOnly, "OpenSolver - Update Check"
25        End If
          
ExitSub:
26        Exit Sub
          
ConnectionError:
27        If Not DoSilentFail Then
28            MsgBox "The update checker was unable to determine the latest " _
                   & "version of OpenSolver. Please try again later."
29        End If
30        GoTo ExitSub
End Sub

Sub AutoUpdateCheck()
          ' Don't check the saved setting if we have already run the checker
1         If Not HasCheckedForUpdate Then
              ' Make sure we don't keep asking the user for their preference if
              ' we have already shown them the settings form
2             HasCheckedForUpdate = True
              
              Dim FirstUpdateCheckTime As Double
3             FirstUpdateCheckTime = GetFirstCheckTime()
              ' Record if this is the first update check
4             If FirstUpdateCheckTime = 0 Then
5                 SetFirstCheckTime Now
6                 Exit Sub
7             End If
              
              ' Make sure 24 hours has passed since the first update check (ref #245)
8             If Now - FirstUpdateCheckTime < MinTimeBeforeCheck Then
9                 Exit Sub
10            End If
              
              Dim SettingWasMissing As Boolean, DoCheck As Boolean
              ' Get the entry, and show the update settings form if missing
11            DoCheck = GetUpdateSetting(False, SettingWasMissing)
              
              ' If the setting was missing, then we have shown the update settings
              ' form, and we don't want to show it again when we get the beta entry
12            AvoidPromptForBeta = SettingWasMissing
              
13            If DoCheck Then
                  ' Check time since last check
14                If Now - GetLastCheckTime() > MinTimeBetweenChecks Then
15                    InitialiseUpdateCheck True
16                End If
17            End If
18        End If
End Sub

Private Function GetUpdateRegName(Beta As Boolean) As String
1         GetUpdateRegName = IIf(Beta, CheckForBetaUpdatesRegName, _
                                       CheckForUpdatesRegName)
End Function

Public Function GetUpdateSetting(Optional SilentFail As Boolean = True, _
        Optional Missing As Boolean, _
        Optional Beta As Boolean = False) As Boolean
              
          Dim result As Variant
1         result = GetSetting(OpenSolverRegName, PreferencesRegName, _
                              GetUpdateRegName(Beta), VALUE_IF_MISSING)
          
2         If result = VALUE_IF_MISSING Then
3             Missing = True
              ' Handle a missing entry
4             If SilentFail Then
                  ' In silent mode, return false without saving anything
5                 result = False
6             Else
                  ' Otherwise, show the dialog and get the setting
                  Dim frmUpdateSettings As FUpdateSettings
7                 Set frmUpdateSettings = New FUpdateSettings
8                 frmUpdateSettings.Show
9                 Unload frmUpdateSettings
10                result = GetUpdateSetting(True)
11            End If
12        Else
13            Missing = False
14        End If
          
          ' Check for stable updates only by default
          Dim DefaultSetting As Boolean
15        DefaultSetting = IIf(Beta, False, True)
16        GetUpdateSetting = SafeCBool(result, DefaultSetting)
End Function
Public Sub SaveUpdateSetting(UpdateSetting As Boolean, _
        Optional Beta As Boolean = False)
1         SaveSetting OpenSolverRegName, PreferencesRegName, _
                      GetUpdateRegName(Beta), BoolToInt(UpdateSetting)
End Sub
Private Sub DeleteUpdateSetting(Optional Beta As Boolean = False)
1         On Error Resume Next
2         DeleteSetting OpenSolverRegName, PreferencesRegName, GetUpdateRegName(Beta)
End Sub

Public Function GetBetaUpdateSetting(Optional SilentFail As Boolean = True, _
        Optional Missing As Boolean) As Boolean
1         GetBetaUpdateSetting = GetUpdateSetting(SilentFail, Missing, Beta:=True)
End Function
Public Sub SaveBetaUpdateSetting(UpdateSetting As Boolean)
1         SaveUpdateSetting UpdateSetting, Beta:=True
End Sub
Private Sub DeleteBetaUpdateSetting()
1         DeleteUpdateSetting Beta:=True
End Sub

Private Function GetLastCheckTime() As Double
          Dim result As Variant
1         result = GetSetting(OpenSolverRegName, PreferencesRegName, _
                              LastUpdateCheckRegName, 0)
          
2         GetLastCheckTime = Val(result)
End Function
Private Sub SetLastCheckTime(CheckTime As Double)
1         SaveSetting OpenSolverRegName, PreferencesRegName, _
                      LastUpdateCheckRegName, StrExNoPlus(CheckTime)
End Sub
Private Sub DeleteLastCheckTime()
1         On Error Resume Next
2         DeleteSetting OpenSolverRegName, PreferencesRegName, LastUpdateCheckRegName
End Sub

Private Function GetFirstCheckTime() As Double
          Dim result As Variant
1         result = GetSetting(OpenSolverRegName, PreferencesRegName, _
                              FirstUpdateCheckRegName, 0)
          
2         GetFirstCheckTime = Val(result)
End Function
Private Sub SetFirstCheckTime(CheckTime As Double)
1         SaveSetting OpenSolverRegName, PreferencesRegName, _
                      FirstUpdateCheckRegName, StrExNoPlus(CheckTime)
End Sub
Private Sub DeleteFirstCheckTime()
1         On Error Resume Next
2         DeleteSetting OpenSolverRegName, PreferencesRegName, FirstUpdateCheckRegName
End Sub

Private Sub ResetHasChecked()
1         HasCheckedForUpdate = False
End Sub

Private Function GetGuid() As String
          ' From rondebruin.nl:
          ' The GetSetting default argument can't be an empty string on Mac
          Dim result As Variant
1         result = GetSetting(OpenSolverRegName, PreferencesRegName, GuidRegName, _
                              VALUE_IF_MISSING)
          
2         If result = VALUE_IF_MISSING Then
3             result = MakeGuid()
4             SetGuid CStr(result)
5         End If
          
6         GetGuid = CStr(result)
End Function

Private Sub SetGuid(Guid As String)
1         SaveSetting OpenSolverRegName, PreferencesRegName, GuidRegName, Guid
End Sub

Private Sub DeleteGuid()
1         DeleteSetting OpenSolverRegName, PreferencesRegName, GuidRegName
End Sub

Private Function MakeGuid() As String
    #If Mac Then
1             MakeGuid = Application.Clean(ExecCapture("uuidgen"))
    #Else
2             MakeGuid = Mid$(CreateObject("Scriptlet.TypeLib").Guid, 2, 36)
    #End If
End Function

Private Sub ResetAllUpdateSettings()
1         DeleteBetaUpdateSetting
2         DeleteUpdateSetting
3         DeleteLastCheckTime
4         DeleteFirstCheckTime
5         HasCheckedForUpdate = False
End Sub
