Attribute VB_Name = "OpenSolverUpdateCheck"
Const FilesPageUrl = "http://sourceforge.net/projects/opensolver/files/"
Private HasCheckedForUpdate As Boolean

Const OpenSolverRegName = "OpenSolver"
Const PreferencesRegName = "Preferences"
Const CheckForUpdatesRegName = "CheckForUpdates"

Private Function GetFilesPageText() As String
    #If Mac Then
        GetFilesPageText = GetFilesPageText_Mac()
    #Else
        GetFilesPageText = GetFilesPageText_Windows()
    #End If
End Function

' On Windows we use an MSXML Http request to get the page.
' Late binding is required so we don't have the (failing on Mac) reference to MSXML
Private Function GetFilesPageText_Windows() As String
    GetFilesPageText_Windows = ""
    
    Dim WinHttpReq As Object  ' MSXML2.XMLHTTP
    Set WinHttpReq = CreateObject("MSXML2.XMLHTTP")
    WinHttpReq.Open "GET", FilesPageUrl, False
    WinHttpReq.send
    
    If WinHttpReq.status = 200 Then
        GetFilesPageText_Windows = WinHttpReq.responseText
    End If
End Function

' On Mac, we use `cURL` via command line which is included by default
#If Mac Then
Private Function GetFilesPageText_Mac() As String
    GetFilesPageText_Mac = ""
    
    Dim Cmd As String
    Dim result As String
    Dim ExitCode As Long

    ' -L follows redirects
    Cmd = "curl -L " & FilesPageUrl
    result = execShell(Cmd, ExitCode)
    
    If ExitCode = 0 Then
        GetFilesPageText_Mac = result
    End If
End Function
#End If

' Gets version number of current release from Sourceforge.
' The release script updates readme.txt on Sourceforge with the current version.
' The readme gets displayed at the bottom of the files page, so we can scrape it for the version.
' An alternative would be to download the readme directly, but Sourceforge's download redirection
' makes this difficult to do using MSXML (cURL works fine).
Private Function GetLatestOpenSolverVersion() As String
    GetLatestOpenSolverVersion = ""
    
    Dim Response As String
    Response = GetFilesPageText()
    
    ' We are looking for the following message:
    '   "Please download the latest version listed here (x.x.x)."
    Dim startString As String
    startString = "the latest version listed here"
    
    Dim start As Long, openingParen As Long, closingParen As Long
    start = InStrText(Response, startString)
    If start > 0 Then
        openingParen = InStr(start, Response, "(") + 1
        closingParen = InStr(openingParen, Response, ")")
        GetLatestOpenSolverVersion = Mid(Response, openingParen, closingParen - openingParen)
    End If
End Function

Sub CheckForUpdate(Optional ByVal SilentFail As Boolean = False)
    Application.Cursor = xlWait
    Application.StatusBar = "Checking for updates to OpenSolver..."
    
    HasCheckedForUpdate = True
    
    Dim LatestVersion As String
    LatestVersion = GetLatestOpenSolverVersion()

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
    Application.Cursor = xlDefault
    Application.StatusBar = False
    
    If UpdateAvailable Then
        frmUpdate.ShowUpdate LatestVersion
    ElseIf Not SilentFail Then
        MsgBox "No updates for OpenSolver are available at this time.", vbOKOnly, "OpenSolver - Update Check"
    End If
End Sub

Sub AutoUpdateCheck()
    ' Don't check the saved setting if we have already run the checker
    If Not HasCheckedForUpdate Then
        If GetUpdateSetting() Then
            CheckForUpdate True
        End If
    End If
End Sub

Public Function GetUpdateSetting() As Boolean
    Dim result As Variant
    result = GetSetting(OpenSolverRegName, PreferencesRegName, CheckForUpdatesRegName, "")
    
    ' If registry key is missing, then check with user whether to autocheck
    If result = "" Then
        result = MsgBox("Would you like OpenSolver to automatically check for updates? " & vbNewLine & vbNewLine & _
                        "You can change this option at any time by going to ""About OpenSolver"". " & _
                        "You can also run update checks manually from there.", vbYesNo, "OpenSolver - Check for Updates?") = vbYes
        SaveUpdateSetting (CBool(result))
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
