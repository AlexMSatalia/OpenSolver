VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "About OpenSolver"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8880
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Mac Then
    Const FormWidthAbout = 730
    Const txtAboutHeight = 300
#Else
    Const FormWidthAbout = 500
    Const txtAboutHeight = 250
#End If

Public EventsEnabled As Boolean

Private Function GetAddInIfExists(AddIn As Variant, Title As String) As Boolean
          ' See http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.addins.aspx
          ' http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.addin.aspx
3474      Set AddIn = Nothing
3475      On Error Resume Next
3476      Set AddIn = Application.AddIns.Item(Title)
3477      GetAddInIfExists = Err = 0
End Function

Private Sub chkUpdate_Change()
    SaveUpdateSetting chkUpdate.value
    chkUpdateExperimental.Enabled = chkUpdate.value
End Sub

Private Sub chkUpdateExperimental_Change()
    SaveBetaUpdateSetting chkUpdateExperimental.value
End Sub

Private Sub cmdOk_Click()
3478      Me.Hide
End Sub

Public Sub ReflectOpenSolverStatus()
          ' Update buttons to reflect install status of OpenSolver
3479      On Error GoTo ErrorHandler
          Dim InstalledAndActive As Boolean
3480      InstalledAndActive = False

          Dim Title As String
3481      Title = "OpenSolver"
#If Mac Then
          ' On Mac, the Application.AddIns collection is indexed by filename.ext rather than just filename as on Windows
3482      Title = Title & ".xlam"
#End If
          Dim AddIn As Variant
3483      Set AddIn = Nothing
3484      If GetAddInIfExists(AddIn, Title) Then
3485          Set AddIn = Application.AddIns(Title)
3486          InstalledAndActive = AddIn.Installed
3487      End If
ErrorHandler:
3488      EventsEnabled = False
3489      chkAutoLoad.value = InstalledAndActive
3490      chkAutoLoad.Enabled = Not InstalledAndActive
3491      EventsEnabled = True
End Sub

Private Sub cmdCancelLoad_Click()
3492      ChangeAutoloadStatus False
End Sub

Private Sub chkAutoLoad_Change()
3493      If Not EventsEnabled Then Exit Sub
3494      ChangeAutoloadStatus chkAutoLoad.value
End Sub

Public Sub ChangeAutoloadStatus(loadAtStartup As Boolean)
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
3511      ReflectOpenSolverStatus
End Sub


Private Sub cmdUpdate_Click()
    InitialiseUpdateCheck False, True
    chkUpdate.value = GetUpdateSetting()
End Sub

Private Sub lblUrl_Click()
3512      OpenURL "http://www.opensolver.org"
End Sub

Private Sub UserForm_Activate()
3514      UpdateStatusBar "OpenSolver: Fetching solver information...", True
3515      Application.Cursor = xlWait

          With txtFilePath
              .Locked = False
              .value = "OpenSolver file: " & MakeSpacesNonBreaking(MakePathSafe(ThisWorkbook.FullName))
              .Locked = True
          End With
          
          chkUpdate.value = GetUpdateSetting()
          chkUpdateExperimental.value = GetBetaUpdateSetting()

          With txtVersion
              .Locked = False
3523          .value = EnvironmentSummary()
              .Locked = True
              AutoHeight txtVersion, Me.width, True
              #If Mac Then
                  ' On Mac the autosizing isn't quite wide enough
                  .width = .width + 10
              #End If
          End With
          
          ReflectOpenSolverStatus
          EventsEnabled = True
          
          With txtAbout
3524          .Locked = False
3525          .Text = About_OpenSolver & vbNewLine & vbNewLine & SolverSummary()
3534          .Locked = True
3535          .SetFocus
3536          .SelStart = 0
          End With
          
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
      "Excel Solver is a product developed by Frontline Systems (www.solver.com) for Microsoft. OpenSolver has no affiliation with, nor is recommend by, Microsoft or Frontline Systems. All trademark terms are the property of their respective owners."

End Function

Private Sub AutoLayout()
    AutoFormat Me.Controls
    
    Me.width = FormWidthAbout
    
    With lblHeading
        .Font.Size = FormHeadingSize
        .width = Me.width - 2 * FormMargin
        .Caption = "OpenSolver"
        .left = FormMargin
        .top = FormMargin
        .height = FormHeadingHeight
    End With
    
    With txtVersion
        .Locked = False
        .value = "OpenSolver version information"
        .Locked = True
        .width = lblHeading.width
        .left = lblHeading.left
        .top = Below(lblHeading, False)
        .BackStyle = fmBackStyleTransparent
    End With
    
    With lblUrl
        .Caption = "http://www.OpenSolver.org"
        .ForeColor = FormLinkColor
        .left = lblHeading.left
        .top = Below(txtVersion, False)
        AutoHeight lblUrl, Me.width, True
    End With
        
    cmdUpdate.top = lblHeading.top
    
    With chkUpdate
        .Caption = "Check for updates automatically"
        .top = Below(cmdUpdate, False)
        AutoHeight chkUpdate, Me.width, True
    End With
    
    With chkUpdateExperimental
        .Caption = "Check for experimental updates"
        .top = Below(chkUpdate, False)
        AutoHeight chkUpdateExperimental, Me.width, True
    End With
    
    With cmdUpdate
        .Caption = "Check for updates"
        .width = Max(chkUpdate.width, chkUpdateExperimental.width)
        .left = LeftOfForm(Me.width, .width)
    End With
    
    chkUpdate.left = cmdUpdate.left
    chkUpdateExperimental.left = cmdUpdate.left
    
    With txtAbout
        .Locked = False
        .Text = "Loading OpenSolver info..."
        .Locked = True
        .left = lblHeading.left
        .top = Max(Below(lblUrl), Below(chkUpdateExperimental))
        .BackStyle = fmBackStyleTransparent
        .SpecialEffect = fmSpecialEffectEtched
        .height = txtAboutHeight
        .width = lblHeading.width
    End With
    
    With txtFilePath
        .Locked = False
        .Text = "OpenSolver file:"
        .Locked = True
        .left = lblHeading.left
        .top = Below(txtAbout)
        .height = 2 * FormTextHeight + 2 ' Stop the text becoming smaller
        .width = lblHeading.width
        .BackStyle = fmBackStyleTransparent
        .MultiLine = True
    End With
    
    With chkAutoLoad
        .Caption = "Load OpenSolver when Excel starts"
        AutoHeight chkAutoLoad, Me.width, True
        .left = lblHeading.left
        .top = Below(txtFilePath, False)
    End With
    
    With cmdCancelLoad
        .Caption = "Cancel loading at startup..."
        .width = FormButtonWidth * 2
        .left = RightOf(chkAutoLoad)
        .top = chkAutoLoad.top
    End With
    
    With cmdOk
        .Caption = "OK"
        .width = FormButtonWidth
        .left = LeftOfForm(Me.width, .width)
        .top = chkAutoLoad.top
    End With
    
    Me.height = FormHeight(cmdOk)
    Me.width = Me.width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - About"
End Sub

Private Sub UserForm_Initialize()
    AutoLayout
End Sub
