VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FAbout 
   Caption         =   "About OpenSolver"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8880
   OleObjectBlob   =   "FAbout.frx":0000
End
Attribute VB_Name = "FAbout"
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

Private Sub cmdOk_Click()
3478      Me.Hide
End Sub

' Make the [x] hide the form rather than unload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
          ' If CloseMode = vbFormControlMenu then we know the user
          ' clicked the [x] close button or Alt+F4 to close the form.
          If CloseMode = vbFormControlMenu Then
              cmdOk_Click
              Cancel = True
          End If
End Sub

Public Sub ReflectOpenSolverStatus()
          ' Update buttons to reflect install status of OpenSolver
          
          Dim InstalledAndActive As Boolean
3480      InstalledAndActive = False
          
          Dim OpenSolverAddIn As Excel.AddIn
3484      If GetOpenSolverAddInIfExists(OpenSolverAddIn) Then InstalledAndActive = OpenSolverAddIn.Installed

3488      EventsEnabled = False
3489      chkAutoLoad.value = InstalledAndActive
3491      EventsEnabled = True
End Sub

Private Sub chkAutoLoad_Change()
3493      If Not EventsEnabled Then Exit Sub
3494      ChangeAutoloadStatus chkAutoLoad.value
End Sub

Private Sub ChangeAutoloadStatus(loadAtStartup As Boolean)
    Dim Changed As Boolean
    Changed = ChangeOpenSolverAutoload(loadAtStartup)
    
    ' Mac doesn't close the userform when it unloads OpenSolver
    #If Mac Then
        If Not loadAtStartup And Changed Then
            Me.Hide
            Exit Sub
        End If
    #End If
    ReflectOpenSolverStatus
End Sub

Private Sub cmdUpdate_Click()
    InitialiseUpdateCheck False, True
End Sub

Private Sub cmdUpdateSettings_Click()
    With New FUpdateSettings
        .Show
    End With
End Sub

Private Sub lblUrl_Click()
3512      OpenURL "http://www.opensolver.org"
End Sub

Private Sub UserForm_Activate()
          CenterForm
    
3514      UpdateStatusBar "OpenSolver: Fetching solver information...", True
3515      Application.Cursor = xlWait

          With txtFilePath
              .Locked = False
              .value = "OpenSolver file: " & MakeSpacesNonBreaking(MakePathSafe(ThisWorkbook.FullName))
              AutoHeight txtFilePath, .width, False
              .Locked = True
          End With
          LayoutBottom
          
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
    
    With cmdUpdate
        .Caption = "Check for updates"
        .width = FormButtonWidth * 2
        .left = LeftOfForm(Me.width, .width)
        .top = lblHeading.top
    End With
    
    With cmdUpdateSettings
        .Caption = "Update Check settings..."
        .width = cmdUpdate.width
        .left = cmdUpdate.left
        .top = Below(cmdUpdate)
    End With
    
    With txtAbout
        .Locked = False
        .Text = "Loading OpenSolver info..."
        .Locked = True
        .left = lblHeading.left
        .top = Max(Below(lblUrl), Below(cmdUpdateSettings))
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
        .height = FormTextHeight + 2 ' Stop the text becoming smaller
        .width = lblHeading.width
        .BackStyle = fmBackStyleTransparent
        .MultiLine = True
    End With
    
    LayoutBottom
    
    Me.width = Me.width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - About"
End Sub

Private Sub LayoutBottom()
    With chkAutoLoad
        .Caption = "Load OpenSolver when Excel starts"
        AutoHeight chkAutoLoad, FormWidthAbout, True
        .left = lblHeading.left
        .top = Below(txtFilePath, False)
    End With
    
    With cmdOk
        .Caption = "OK"
        .width = FormButtonWidth
        .left = LeftOfForm(FormWidthAbout, .width)
        .top = chkAutoLoad.top
    End With
    
    Me.height = FormHeight(cmdOk)
End Sub

Private Sub UserForm_Initialize()
    AutoLayout
    CenterForm
End Sub

Private Sub CenterForm()
    Me.top = CenterFormTop(Me.height)
    Me.left = CenterFormLeft(Me.width)
End Sub

