VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormAbout 
   Caption         =   "About OpenSolver"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8880
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
3474      Set AddIn = Nothing
3475      On Error Resume Next
3476      Set AddIn = Application.AddIns.Item(title)
3477      GetAddInIfExists = Err = 0
End Function

Private Sub buttonOK_Click()
3478      Me.Hide
End Sub

Public Sub ReflectOpenSolverStatus(f As UserForm)
          ' Update buttons to reflect install status of OpenSolver
3479      On Error GoTo errorHandler
          Dim InstalledAndActive As Boolean
3480      InstalledAndActive = False

          Dim title As String
3481      title = "OpenSolver"
#If Mac Then
          ' On Mac, the Application.AddIns collection is indexed by filename.ext rather than just filename as on Windows
3482      title = title & ".xlam"
#End If
          Dim AddIn As Variant
3483      Set AddIn = Nothing
3484      If GetAddInIfExists(AddIn, title) Then
3485          Set AddIn = Application.AddIns(title)
3486          InstalledAndActive = AddIn.Installed
3487      End If
errorHandler:
3488      EventsEnabled = False
3489      f.chkAutoLoad.value = InstalledAndActive
3490      f.chkAutoLoad.Enabled = Not InstalledAndActive
3491      EventsEnabled = True
End Sub

Private Sub buttonUninstall_Click()
3492      ChangeAutoloadStatus False, Me
End Sub

Private Sub chkAutoLoad_Change()
3493      If Not EventsEnabled Then Exit Sub
3494      ChangeAutoloadStatus chkAutoLoad.value, Me
End Sub

Public Sub ChangeAutoloadStatus(loadAtStartup As Boolean, f As UserForm)
          ' See http://www.jkp-ads.com/articles/AddinsAndSetupFactory.asp
          ' HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Excel\Add-in Manager
          ' HKEY_CURRENT_USER\Software\Microsoft\Office\10.0\Excel\Add-in Manager
          ' The name of the Entry is the path
3495      If loadAtStartup Then  ' User is changing from True to False
3496          If MsgBox("This will configure Excel to automatically load the OpenSolver add-in (from its current location) when Excel starts.  Continue?", vbOKCancel) <> vbOK Then GoTo ExitSub
3497      Else ' User is turning off auto load
3498          If MsgBox("This will re-configure Excel's Add-In settings so that OpenSolver does not load automatically at startup. You will need to re-load OpenSolver when you wish to use it next, or re-enable it using Excel's Add-In settings." & vbCrLf & vbCrLf _
                        & "WARNING: If you continue, OpenSolver will also be shut down right now by Excel, and so will disappear immediately. No data will be lost." & vbCrLf & vbCrLf _
                        & "Continue?", vbOKCancel) <> vbOK Then GoTo ExitSub
3499      End If
          Dim AddIn As Variant
3500      Set AddIn = Nothing
          
          ' Add-ins can only be added if we have at least one workbook open; see http://vbadud.blogspot.com/2007/06/excel-vba-install-excel-add-in-xla-or.html
          Dim TempBook As Workbook
3501      If Workbooks.Count = 0 Then Set TempBook = Workbooks.Add

3502      If Not GetAddInIfExists(AddIn, "OpenSolver") Then
3503          Set AddIn = Application.AddIns.Add(ThisWorkbook.FullName, False)
3504      End If
          
3505      If Not TempBook Is Nothing Then TempBook.Close
          
3506      If AddIn Is Nothing Then
3507          MsgBox "Unable to load or access addin " & ThisWorkbook.FullName
3508      Else
3509          AddIn.Installed = loadAtStartup ' OpenSolver will quit immediately when this is set to false
3510      End If
ExitSub:
3511      ReflectOpenSolverStatus f
End Sub


Private Sub labelOpenSolverOrg_Click()
3512      Call OpenURL("http://www.opensolver.org")
End Sub

Private Sub UserForm_Activate()
3513      ActivateAboutForm Me
End Sub

Public Sub ActivateAboutForm(f As UserForm)
3514      Application.StatusBar = "OpenSolver: Fetching solver information..."
3515      Application.Cursor = xlWait

          Dim VBAversion As String
3516      VBAversion = "VBA"
#If VBA7 Then
3517      VBAversion = "VBA7"
#ElseIf VBA6 Then
3518      VBAversion = "VBA6"
#End If

          Dim ExcelBitness As String
#If Win64 Then
3519      ExcelBitness = "64"
#Else
3520      ExcelBitness = "32"
#End If
          Dim OS As String
#If Mac Then
3521      OS = "Mac"
#Else
3522      OS = "Windows"
#End If

3523      f.labelVersion.Caption = "Version " & sOpenSolverVersion & " (" & sOpenSolverDate & ") running on " & IIf(SystemIs64Bit, "64", "32") & " bit " & OS & " in " & VBAversion & " in " & ExcelBitness & " bit Excel " & Application.Version
          
3524      f.txtAbout.Locked = False
3525      f.txtAbout.Text = About_OpenSolver
3526      f.txtAbout.Text = f.txtAbout.Text & About_CBC & vbNewLine & vbNewLine
3527      f.txtAbout.Text = f.txtAbout.Text & About_Gurobi & vbNewLine & vbNewLine
3528      f.txtAbout.Text = f.txtAbout.Text & About_NOMAD & vbNewLine & vbNewLine
3529      f.txtAbout.Text = f.txtAbout.Text & About_Bonmin & vbNewLine & vbNewLine
3530      f.txtAbout.Text = f.txtAbout.Text & About_Couenne & vbNewLine & vbNewLine
          
3531      f.labelFilePath = "OpenSolverFile: " & MakeSpacesNonBreaking(ThisWorkbook.FullName)
3532      ReflectOpenSolverStatus f
3533      EventsEnabled = True

3534      f.txtAbout.Locked = True
3535      f.txtAbout.SetFocus
3536      f.txtAbout.SelStart = 0
          
3537      Application.StatusBar = False
3538      Application.Cursor = xlDefault
End Sub


Public Function About_OpenSolver() As String
3539  About_OpenSolver = _
      "(c) Andrew J Mason 2011 , 2012" & vbNewLine & _
      "Developed by Andrew Mason, Iain Dunning and Jack Dunn, with coding assistance by Kat Gilbert, Matthew Milner, Kris Atkins. Various contributions have been made by Andres Sommerhoff, and assistance with Mac version was given by Zhanibek Datbayev." & vbNewLine & _
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
