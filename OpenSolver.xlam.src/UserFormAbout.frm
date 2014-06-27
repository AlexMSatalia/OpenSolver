VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormAbout 
   Caption         =   "About OpenSolver"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8880.001
   OleObjectBlob   =   "UserFormAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const OpenSolverStudioAddInName = "OpenSolverStudio"

' See http://support.microsoft.com/kb/145679
' http://www.tek-tips.com/faqs.cfm?fid=3719

Const REG_SZ As Long = 1
Const REG_DWORD As Long = 4

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003

Const ERROR_NONE = 0
Const ERROR_BADDB = 1
Const ERROR_BADKEY = 2
Const ERROR_CANTOPEN = 3
Const ERROR_CANTREAD = 4
Const ERROR_CANTWRITE = 5
Const ERROR_OUTOFMEMORY = 6
Const ERROR_ARENA_TRASHED = 7
Const ERROR_ACCESS_DENIED = 8
Const ERROR_INVALID_PARAMETERS = 87
Const ERROR_NO_MORE_ITEMS = 259

Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_ALL_ACCESS = &H3F

Const REG_OPTION_NON_VOLATILE = 0

' NB: The following two definitions differ slightly in RegSetValueExLong,RegSetValueExString,
' RegQueryValueExNULL,RegQueryValueExString as to whether or not the lpData (aka lpValue) is ByVal or not
' Googling suggests both are ok.
' See also http://www.tudosobrexcel.com/vba/vba_32bits/Win32API_PtrSafe.TXT
#If VBA7 Then    ' VBA7; see http://code.google.com/p/excel-connector/issues/attachmentText?id=8&aid=6686646755651591166&name=Office+connector+Patch.txt&token=7e6e4d37a97cacb36874c4b6d23db9c7
Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As Long
Private Declare PtrSafe Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
     (ByVal hKey As LongPtr, _
     ByVal lpSubKey As String, _
     ByVal Reserved As Long, _
     ByVal lpClass As String, _
     ByVal dwOptions As Long, _
     ByVal samDesired As Long, _
     ByVal lpSecurityAttributes As Long, _
     phkResult As LongPtr, _
     lpdwDisposition As Long) As Long
Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" _
     Alias "RegOpenKeyExA" ( _
     ByVal hKey As LongPtr, _
     ByVal lpSubKey As String, _
     ByVal ulOptions As Long, _
     ByVal samDesired As Long, _
     phkResult As LongPtr) As Long
Private Declare PtrSafe Function RegQueryValueExString Lib "advapi32.dll" _
     Alias "RegQueryValueExA" ( _
     ByVal hKey As LongPtr, _
     ByVal lpValueName As String, _
     ByVal lpReserved As Long, _
     lpType As Long, _
     lpData As String, _
     lpcbData As Long) As Long
Private Declare PtrSafe Function RegQueryValueExLong Lib "advapi32.dll" _
     Alias "RegQueryValueExA" ( _
     ByVal hKey As LongPtr, _
     ByVal lpValueName As String, _
     ByVal lpReserved As Long, _
     lpType As Long, _
     lpData As Long, _
     lpcbData As Long) As Long
Private Declare PtrSafe Function RegQueryValueExNULL Lib "advapi32.dll" _
     Alias "RegQueryValueExA" ( _
     ByVal hKey As LongPtr, _
     ByVal lpValueName As String, _
     ByVal lpReserved As Long, _
     lpType As Long, _
     lpData As Long, _
     lpcbData As Long) As Long
Private Declare PtrSafe Function RegSetValueExString Lib "advapi32.dll" _
     Alias "RegSetValueExA" ( _
     ByVal hKey As LongPtr, _
     ByVal lpValueName As String, _
     ByVal Reserved As Long, _
     ByVal dwType As Long, _
     lpData As String, _
     ByVal cbData As Long) As Long
Private Declare PtrSafe Function RegSetValueExLong Lib "advapi32.dll" _
     Alias "RegSetValueExA" ( _
     ByVal hKey As LongPtr, _
     ByVal lpValueName As String, _
     ByVal Reserved As Long, _
     ByVal dwType As Long, _
     lpData As Long, _
     ByVal cbData As Long) As Long
Private Declare PtrSafe Function RegDeleteValue Lib "advapi32.dll" _
     Alias "RegDeleteValueA" (ByVal hKey As LongPtr, ByVal lpValueName As String) As Long
     
#Else '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
     "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
     ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
     As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
     As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
     "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
     ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
     Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
     "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
     String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
     As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
     "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
     String, ByVal lpReserved As Long, lpType As Long, lpData As _
     Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
     "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
     String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
     As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
     "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
     ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
     String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
     "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
     ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
     ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
    (ByVal hKey As Long, ByVal lpValueName As String) As Long

#End If

Private EventsEnabled As Boolean

         
' See http://support.microsoft.com/kb/145679
Public Function SetValueEx(ByVal hKey As Long, sValueName As String, _
           lType As Long, vValue As Variant) As Long
           Dim lValue As Long
           Dim sValue As String
34870      Select Case lType
               Case REG_SZ
34880              sValue = vValue & Chr$(0)
34890              SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
34900          Case REG_DWORD
34910              lValue = vValue
34920              SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
34930          End Select
 End Function

' See http://support.microsoft.com/kb/145679
 Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As _
 String, vValue As Variant) As Long
           Dim cch As Long
           Dim lrc As Long
           Dim lType As Long
           Dim lValue As Long
           Dim sValue As String

34940      On Error GoTo QueryValueExError

           ' Determine the size and type of data to be read
34950      lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
34960      If lrc <> ERROR_NONE Then Error 5

34970      Select Case lType
               ' For strings
               Case REG_SZ:
34980              sValue = String(cch, 0)

34990  lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, _
       sValue, cch)
35000              If lrc = ERROR_NONE Then
35010                  vValue = left$(sValue, cch - 1)
35020              Else
35030                  vValue = Empty
35040              End If
               ' For DWORDS
35050          Case REG_DWORD:
35060  lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, _
       lValue, cch)
35070              If lrc = ERROR_NONE Then vValue = lValue
35080          Case Else
                   'all other data types not supported
35090              lrc = -1
35100      End Select

QueryValueExExit:
35110      QueryValueEx = lrc
35120      Exit Function

QueryValueExError:
35130      Resume QueryValueExExit
 End Function

' See http://support.microsoft.com/kb/145679
Private Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
             Dim hNewKey As Long         'handle to the new key
             Dim lRetVal As Long         'result of the RegCreateKeyEx function

35140        lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
                       vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                       0&, hNewKey, lRetVal)
35150        RegCloseKey (hNewKey)
End Sub
    
' See http://support.microsoft.com/kb/145679
Private Sub SetKeyValue(sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
          Dim lRetVal As Long         'result of the SetValueEx function
          Dim hKey As Long         'handle of open key

          'open the specified key
35160     lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, _
                                    KEY_SET_VALUE, hKey)
35170     lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
35180     RegCloseKey (hKey)
End Sub
    
' See http://support.microsoft.com/kb/145679
Private Sub QueryValue(sKeyName As String, sValueName As String)
             Dim lRetVal As Long         'result of the API functions
             Dim hKey As Long         'handle of opened key
             Dim vValue As Variant      'setting of queried value

35190        lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, KEY_QUERY_VALUE, hKey)
35200        lRetVal = QueryValueEx(hKey, sValueName, vValue)
35210        MsgBox vValue
35220        RegCloseKey (hKey)
   End Sub

Private Function GetValueIfExists(sKeyName As String, sValueName As String, vValue As Variant) As Boolean
          Dim lRetVal As Long         'result of the API functions
          Dim hKey As Long         'handle of opened key
35230     lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, KEY_QUERY_VALUE, hKey)
35240     If lRetVal = ERROR_NONE Then
35250         lRetVal = QueryValueEx(hKey, sValueName, vValue)
35260         RegCloseKey hKey
35270     End If
35280     GetValueIfExists = lRetVal = ERROR_NONE
   End Function

Private Function KeyExists(sKeyName As String, sValueName As String, vValue As Variant) As Boolean
          Dim lRetVal As Long         'result of the API functions
          Dim hKey As Long         'handle of opened key
35290     lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, KEY_QUERY_VALUE, hKey)
35300     If lRetVal = ERROR_NONE Then RegCloseKey hKey
35310     KeyExists = lRetVal = ERROR_NONE
End Function

Function DeleteValue(sKeyName As String, sValueName As String) As Boolean
          Dim lRetVal As Long
          Dim hKey As Long
35320     lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, KEY_ALL_ACCESS, hKey)
35330     If lRetVal = ERROR_NONE Then
35340         lRetVal = RegDeleteValue(hKey, sValueName)
35350         RegCloseKey hKey
35360     End If
35370     DeleteValue = lRetVal = ERROR_NONE
End Function

' See http://support.microsoft.com/kb/145679
Sub CreateKeyAndSetValue(sNewKeyName As String, lPredefinedKey As Long, sValueName As String, vValueSetting As Variant, lValueType As Long)
          Dim hKey As Long         'handle to the new key
          Dim lRetVal As Long         'result of the RegCreateKeyEx function

35380     lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
                    vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                    0&, hKey, lRetVal)
35390     lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
35400     RegCloseKey (hKey)
End Sub

Private Function GetCOMAddInIfExists(AddIn As Variant, progID As String) As Boolean
          ' See http://msdn.microsoft.com/en-us/library/aa432088(v=office.12).aspx for COMAddIn
35410     Set AddIn = Nothing
35420     On Error Resume Next
35430     Set AddIn = Application.COMAddIns.Item(progID)
35440     GetCOMAddInIfExists = Err = 0
End Function

Private Function GetAddInIfExists(AddIn As Variant, title As String) As Boolean
          ' See http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.addins.aspx
          ' http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.addin.aspx
35450     Set AddIn = Nothing
35460     On Error Resume Next
35470     Set AddIn = Application.AddIns.Item(title)
35480     GetAddInIfExists = Err = 0
End Function

Private Sub btnInstallOpenSolverStudio_Click()
          Dim AddInName As String, key As String, FullPath As String
35490     key = "Software\Microsoft\Office\Excel\Addins\" & OpenSolverStudioAddInName
35500     FullPath = ThisWorkbook.Path
35510     If right(" " & FullPath, 1) <> PathDelimeter Then FullPath = FullPath & PathDelimeter
35520     FullPath = FullPath & OpenSolverStudioAddInName & ".vsto"
35530     If Dir(FullPath) = "" Then
35540         MsgBox "Unable to find the OpenSolverStudio file: " & FullPath, , "OpenSolver"
35550         Exit Sub
35560     End If
          ' FullPath = "file:///" & FullPath & "|vstolocal" The Visual Studio  ' The values created by Visual Studio are of this form
          
          ' For a description of these registry entries, see
          ' http://msdn.microsoft.com/en-us/library/bb386106.aspx
35570     CreateNewKey key, HKEY_CURRENT_USER
35580     SetKeyValue key, "Description", "Open Source Optimisation for Excel", REG_SZ
35590     SetKeyValue key, "FriendlyName", "OpenSolver Studio", REG_SZ
35600     SetKeyValue key, "LoadBehavior", 3, REG_DWORD
35610     SetKeyValue key, "Manifest", FullPath & "|vstolocal", REG_SZ
          
          'Dim AddIn As Variant, progID As String
          'For Each AddIn In Application.COMAddIns
          '    Debug.Print AddIn.Description, AddIn.GUID, AddIn.progID, AddIn.Connect
          'Next AddIn
          'progID = "vv"
          
          ' Now refresh the addins to force a load
35620     Application.COMAddIns.Update

          Dim AddIn As Variant, AddInStatus As String
35630     If GetCOMAddInIfExists(AddIn, OpenSolverStudioAddInName) Then
35640         AddIn.Connect = False   ' Force a refresh
35650         AddIn.Connect = True   ' Force a refresh?; see http://stackoverflow.com/questions/213375/excel-vba-load-addins
35660         If AddIn.Connect Then
35670             AddInStatus = "OpenSolverStudio is installed and active."
35680         Else
35690             AddInStatus = "OpenSolverStudio is installed but could not be activated."
35700         End If
35710     Else
35720         lblStudioStatus.Caption = "OpenSolverStudio installation failed."
35730     End If

35740     ShowOpenSolverStudioStatus
35750     MsgBox AddInStatus & vbCrLf & vbCrLf & "Actions Taken:" & vbCrLf & "Installed registry keys under:" & vbCrLf & key & vbCrLf & "for the VSTO manifest:" & vbCrLf & FullPath, vbOKOnly, "OpenSolver: OpenSolver Studio Installation"
          
          ' Application.COMAddIns property - COM addins including VSTO's
          ' Application.AddIns property - lists all XLA addins
End Sub

Private Sub buttonOK_Click()
35760     Me.Hide
End Sub

Sub ReflectOpenSolverStatus()
          ' Update buttons to reflect install status of OpenSolver
35770     On Error GoTo errorHandler
          Dim InstalledAndActive As Boolean
35780     InstalledAndActive = False
          Dim AddIn As Variant
35790     Set AddIn = Nothing
35800     If GetAddInIfExists(AddIn, "OpenSolver") Then
35810         Set AddIn = Application.AddIns("OpenSolver")
35820         InstalledAndActive = AddIn.Installed
35830     End If
errorHandler:
35840     EventsEnabled = False
35850     chkAutoLoad.value = InstalledAndActive
35860     chkAutoLoad.Enabled = Not InstalledAndActive
35870     buttonUninstall.Enabled = InstalledAndActive
35880     EventsEnabled = True
End Sub

Private Sub buttonUninstall_Click()
35890     ChangeAutoloadStatus False
End Sub

Private Sub chkAutoLoad_Change()
35900     If Not EventsEnabled Then Exit Sub
35910     ChangeAutoloadStatus chkAutoLoad.value
End Sub

Private Sub ChangeAutoloadStatus(loadAtStartup As Boolean)
          ' See http://www.jkp-ads.com/articles/AddinsAndSetupFactory.asp
          ' HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Excel\Add-in Manager
          ' HKEY_CURRENT_USER\Software\Microsoft\Office\10.0\Excel\Add-in Manager
          ' The name of the Entry is the path
35920     If loadAtStartup Then  ' User is changing from True to False
35930         If MsgBox("This will configure Excel to automatically load the OpenSolver add-in (from its current location) when Excel starts.  Continue?", vbOKCancel) <> vbOK Then GoTo ExitSub
35940     Else ' User is turning off auto load
35950         If MsgBox("This will re-configure Excel's Add-In settings so that OpenSolver does not load automatically at startup. You will need to re-load OpenSolver when you wish to use it next, or re-enable it using Excel's Add-In settings." & vbCrLf & vbCrLf _
                        & "WARNING: If you continue, OpenSolver will also be shut down right now by Excel, and so will disappear immediately. No data will be lost." & vbCrLf & vbCrLf _
                        & "Continue?", vbOKCancel) <> vbOK Then GoTo ExitSub
35960     End If
          Dim AddIn As Variant
35970     Set AddIn = Nothing
          
          ' Add-ins can only be added if we have at least one workbook open; see http://vbadud.blogspot.com/2007/06/excel-vba-install-excel-add-in-xla-or.html
          Dim TempBook As Workbook
35980     If Workbooks.Count = 0 Then Set TempBook = Workbooks.Add

35990     If Not GetAddInIfExists(AddIn, "OpenSolver") Then
36000         Set AddIn = Application.AddIns.Add(ThisWorkbook.FullName, False)
36010     End If
          
36020     If Not TempBook Is Nothing Then TempBook.Close
          
36030     If AddIn Is Nothing Then
36040         MsgBox "Unable to load or access addin " & ThisWorkbook.FullName
36050     Else
36060         AddIn.Installed = loadAtStartup ' OpenSolver will quit immediately when this is set to false
36070     End If
ExitSub:
36080     ReflectOpenSolverStatus
          
          'Dim Keys(4) As String
          'Keys(1) = "Software\Microsoft\Office\10.0\Excel\Add-in Manager"
          'Keys(2) = "Software\Microsoft\Office\11.0\Excel\Add-in Manager"
          'Keys(3) = "Software\Microsoft\Office\12.0\Excel\Add-in Manager"
          'Keys(4) = "Software\Microsoft\Office\13.0\Excel\Add-in Manager"
          'key = "Software\Microsoft\Office\10.0\Excel\Add-in Manager"
          'For Each key In Keys
          '    If KeyExists(key) Then
          '        If chkAutoLoad.Value = 0 Then
          '            DeleteValue key, ThisWorkbook.FullName
          '        Else
          '            SetKeyValue key, ThisWorkbook.FullName, "", REG_SZ
          '        End If
          '    End If
          'Next key
End Sub


Private Sub labelOpenSolverOrg_Click()
36090     Call fHandleFile("http://www.opensolver.org", WIN_NORMAL)
          ' or ThisWorkbook.FollowHyperlink Address :="http://www.opensolver.org", NewWindow := true
End Sub

Private Sub ShowOpenSolverStudioStatus()
          Dim AddIn As Variant
36100     If GetCOMAddInIfExists(AddIn, OpenSolverStudioAddInName) Then
36110         If AddIn.Connect Then
36120             lblStudioStatus.Caption = "OpenSolverStudio: Installed and Active"
36130         Else
36140             lblStudioStatus.Caption = "OpenSolverStudio: Installed but inactive"
36150         End If
36160     Else
36170         lblStudioStatus.Caption = "OpenSolverStudio: Not Installed"
36180     End If
End Sub

Private Sub UserForm_Activate()
          Dim VBAversion As String
36190     VBAversion = "VBA"
#If VBA7 Then
36200     VBAversion = "VBA7"
#ElseIf VBA6 Then
36210     VBAversion = "VBA6"
#End If

          Dim ExcelBitness As String
#If Win64 Then
          ExcelBitness = "64"
#Else
          ExcelBitness = "32"
#End If

36220     labelVersion.Caption = "Version " & sOpenSolverVersion & " (" & sOpenSolverDate & ") running on " & IIf(SystemIs64Bit, "64", "32") & " bit Windows in " & VBAversion & " in " & ExcelBitness & " bit Excel " & Application.Version
36230     labelFilePath = "OpenSolverFile: " & ThisWorkbook.FullName
          ' ShowOpenSolverStudioStatus
36240     ReflectOpenSolverStatus
36250     EventsEnabled = True
          txtAbout.SetFocus
          txtAbout.SelStart = 0
End Sub


