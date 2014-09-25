VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormAbout 
   Caption         =   "About OpenSolver"
   ClientHeight    =   7140
   ClientLeft      =   42
   ClientTop       =   343
   ClientWidth     =   8883
   OleObjectBlob   =   "UserFormAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EventsEnabled As Boolean

Private Function GetAddInIfExists(AddIn As Variant, title As String) As Boolean
          ' See http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.addins.aspx
          ' http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.addin.aspx
35450     Set AddIn = Nothing
35460     On Error Resume Next
35470     Set AddIn = Application.AddIns.Item(title)
35480     GetAddInIfExists = Err = 0
End Function

Private Sub buttonOK_Click()
35760     Me.Hide
End Sub

Public Sub ReflectOpenSolverStatus(f As UserForm)
          ' Update buttons to reflect install status of OpenSolver
35770     On Error GoTo errorHandler
          Dim InstalledAndActive As Boolean
35780     InstalledAndActive = False

          Dim title As String
          title = "OpenSolver"
#If Mac Then
          ' On Mac, the Application.AddIns collection is indexed by filename.ext rather than just filename as on Windows
          title = title & ".xlam"
#End If
          Dim AddIn As Variant
35790     Set AddIn = Nothing
35800     If GetAddInIfExists(AddIn, title) Then
35810         Set AddIn = Application.AddIns(title)
35820         InstalledAndActive = AddIn.Installed
35830     End If
errorHandler:
35840     EventsEnabled = False
35850     f.chkAutoLoad.value = InstalledAndActive
35860     f.chkAutoLoad.Enabled = Not InstalledAndActive
35880     EventsEnabled = True
End Sub

Private Sub buttonUninstall_Click()
35890     ChangeAutoloadStatus False, Me
End Sub

Private Sub chkAutoLoad_Change()
35900     If Not EventsEnabled Then Exit Sub
35910     ChangeAutoloadStatus chkAutoLoad.value, Me
End Sub

Public Sub ChangeAutoloadStatus(loadAtStartup As Boolean, f As UserForm)
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
36080     ReflectOpenSolverStatus f
End Sub


Private Sub labelOpenSolverOrg_Click()
36090     Call OpenURL("http://www.opensolver.org")
End Sub

Private Sub UserForm_Activate()
    ActivateAboutForm Me
End Sub

Public Sub ActivateAboutForm(f As UserForm)
          Application.StatusBar = "OpenSolver: Fetching solver information..."
          Application.Cursor = xlWait

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
          Dim OS As String
#If Mac Then
          OS = "Mac"
#Else
          OS = "Windows"
#End If

36220     f.labelVersion.Caption = "Version " & sOpenSolverVersion & " (" & sOpenSolverDate & ") running on " & IIf(SystemIs64Bit, "64", "32") & " bit " & OS & " in " & VBAversion & " in " & ExcelBitness & " bit Excel " & Application.Version
          
          f.txtAbout.Locked = False
          f.txtAbout.Text = About_OpenSolver
          f.txtAbout.Text = f.txtAbout.Text & About_CBC & vbNewLine & vbNewLine
          f.txtAbout.Text = f.txtAbout.Text & About_Gurobi & vbNewLine & vbNewLine
          f.txtAbout.Text = f.txtAbout.Text & About_NOMAD & vbNewLine & vbNewLine
          f.txtAbout.Text = f.txtAbout.Text & About_Bonmin & vbNewLine & vbNewLine
          f.txtAbout.Text = f.txtAbout.Text & About_Couenne & vbNewLine & vbNewLine
          
36230     f.labelFilePath = "OpenSolverFile: " & MakeSpacesNonBreaking(ThisWorkbook.FullName)
36240     ReflectOpenSolverStatus f
36250     EventsEnabled = True

          f.txtAbout.Locked = True
          f.txtAbout.SetFocus
          f.txtAbout.SelStart = 0
          
          Application.StatusBar = False
          Application.Cursor = xlDefault
End Sub


Public Function About_OpenSolver() As String
About_OpenSolver = _
"(c) Andrew J Mason 2011 , 2012" & vbNewLine & _
"Developed by Andrew Mason, Iain Dunning and Jack Dunn, with coding assistance by Kat Gilbert, Matthew Milner, Kris Atkins. Assistance with Mac version given by Zhanibek Datbayev." & vbNewLine & _
"Department of Engineering Science" & vbNewLine & _
"University of Auckland, New Zealand" & vbNewLine & _
vbNewLine & _
"Excel 2003 Menu Code" & vbNewLine & _
"Provided by Paul Becker of Eclipse Engineering (http://www.eclipseeng.com)" & vbNewLine & _
vbNewLine & _
"OpenSolver allows the Coin-OR CBC optimization engine to be used to solve linear integer programming problems in Excel as well as the NOMAD optimization engine to solve non-linear programming problems. OpenSolver also offers the choice of solving linear problems with the Gurobi optimizer if this is installed." & vbNewLine & _
vbNewLine & _
"OpenSolver is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.  The COIN-OR solvers (CBC, Couenne and Bonmin) are licensed under the Eclipse Public License while the NOMAD software is subject to the terms of the GNU Lesser General Public License." & vbNewLine & _
vbNewLine & _
"OpenSolver is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.  You should have received a copy of the GNU General Public License along with OpenSolver.  If not, see http://www.gnu.org/licenses/" & vbNewLine & _
vbNewLine & _
"Excel Solver is a product developed by Frontline Systems (www.solver.com) for Microsoft. OpenSolver has no affiliation with, nor is recommend by, Microsoft or Frontline Systems. All trademark terms are the property of their respective owners." & vbNewLine & _
vbNewLine

End Function
