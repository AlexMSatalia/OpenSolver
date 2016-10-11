VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FAbout 
   Caption         =   "About OpenSolver"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   -120
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
1         Me.Hide
End Sub

' Make the [x] hide the form rather than unload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
                ' If CloseMode = vbFormControlMenu then we know the user
                ' clicked the [x] close button or Alt+F4 to close the form.
1               If CloseMode = vbFormControlMenu Then
2                   cmdOk_Click
3                   Cancel = True
4               End If
End Sub

Public Sub ReflectOpenSolverStatus()
          ' Update buttons to reflect install status of OpenSolver
1         If SolverDirIsPresent Then
2             Me.lblInstalled.Caption = "OpenSolver is correctly installed " & ChrW(&H2714)
3         Else
4             Me.lblInstalled.Caption = "OpenSolver is not correctly installed. Make sure you have unzipped the downloaded file."
5         End If
          
          Dim InstalledAndActive As Boolean
6         InstalledAndActive = False
          
          Dim OpenSolverAddIn As Excel.AddIn
7         If GetOpenSolverAddInIfExists(OpenSolverAddIn) Then InstalledAndActive = OpenSolverAddIn.Installed

8         EventsEnabled = False
9         chkAutoLoad.value = InstalledAndActive
10        EventsEnabled = True
End Sub

Private Sub chkAutoLoad_Change()
1         If Not EventsEnabled Then Exit Sub
2         ChangeAutoloadStatus chkAutoLoad.value
End Sub

Private Sub ChangeAutoloadStatus(loadAtStartup As Boolean)
          Dim Changed As Boolean
1         Changed = ChangeOpenSolverAutoload(loadAtStartup)
          
          ' Mac doesn't close the userform when it unloads OpenSolver
          ' MAC2016 broken on Mac 2016
    #If Mac Then
2             If Not loadAtStartup And Changed Then
3                 Me.Hide
4                 Exit Sub
5             End If
    #End If
6         ReflectOpenSolverStatus
End Sub

Private Sub cmdUpdate_Click()
1         InitialiseUpdateCheck False, True
End Sub

Private Sub cmdUpdateSettings_Click()
1         With New FUpdateSettings
2             .Show
3         End With
End Sub

Private Sub lblUrl_Click()
1         OpenURL "http://www.opensolver.org"
End Sub

Private Sub UserForm_Activate()
1         CenterForm
    
2         UpdateStatusBar "OpenSolver: Fetching solver information...", True
3         Application.Cursor = xlWait

4         With txtFilePath
5             .Locked = False
6             .value = "OpenSolver file: " & MakeSpacesNonBreaking(MakePathSafe(ThisWorkbook.FullName))
7             AutoHeight txtFilePath, .Width, False
8             .Locked = True
9         End With
10        LayoutBottom
          
11        With txtVersion
12            .Locked = False
13            .value = EnvironmentSummary()
14            .Locked = True
15            AutoHeight txtVersion, Me.Width, True
              #If Mac Then
                  ' On Mac the autosizing isn't quite wide enough
16                .Width = .Width + 10
              #End If
17        End With
          
18        ReflectOpenSolverStatus
19        EventsEnabled = True
          
20        With txtAbout
21            .Locked = False
22            .Text = About_OpenSolver & vbNewLine & vbNewLine & _
                      "========== SYSTEM INFORMATION ==========" & _
                      vbNewLine & vbNewLine & _
                      EnvironmentDetail() & vbNewLine & vbNewLine & _
                      SolverSummary()
23            .Locked = True
24            .SetFocus
25            .SelStart = 0
26        End With
          
27        Application.StatusBar = False
28        Application.Cursor = xlDefault
End Sub

Public Function About_OpenSolver() As String
1     About_OpenSolver = _
      "Copyright (c) 2011, 2012, 2014-2016: Andrew J. Mason" & vbNewLine & _
      "Developed by Andrew Mason, Iain Dunning and Jack Dunn, with coding assistance by Kat Gilbert, Matthew Milner, Kris Atkins. Various contributions have been made by Andres Sommerhoff, and assistance with Mac version was given by Zhanibek Datbayev." & vbNewLine & _
      "Department of Engineering Science" & vbNewLine & _
      "University of Auckland, New Zealand" & vbNewLine & _
      vbNewLine & _
      "Excel 2003 Menu Code" & vbNewLine & _
      "Provided by Paul Becker of Eclipse Engineering (http://www.eclipseeng.com)" & vbNewLine & _
      vbNewLine & _
      "OpenSolver allows the Coin-OR CBC optimization engine to be used to solve linear integer programming problems in Excel as well as the NOMAD optimization engine to solve non-linear programming problems. OpenSolver also offers the choice of solving linear problems with the Gurobi optimizer if this is installed." & vbNewLine & _
      vbNewLine & _
      "OpenSolver is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.  License copyright years may be listed using range notation (e.g. 2011-2016) indicating that every year in the range, inclusive, is a copyrightable year that would otherwise be listed individually." & vbNewLine & _
      vbNewLine & _
      "The COIN-OR solvers (CBC, Couenne and Bonmin) are licensed under the Eclipse Public License while the NOMAD software is subject to the terms of the GNU Lesser General Public License." & vbNewLine & _
      vbNewLine & _
      "OpenSolver is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.  You should have received a copy of the GNU General Public License along with OpenSolver.  If not, see http://www.gnu.org/licenses/" & vbNewLine & _
      vbNewLine & _
      "Excel Solver is a product developed by Frontline Systems (www.solver.com) for Microsoft. OpenSolver has no affiliation with, nor is recommend by, Microsoft or Frontline Systems. All trademark terms are the property of their respective owners."

End Function

Private Sub AutoLayout()
1         AutoFormat Me.Controls
          
2         Me.Width = FormWidthAbout
          
3         With lblHeading
4             .Font.Size = FormHeadingSize
5             .Width = Me.Width - 2 * FormMargin
6             .Caption = "OpenSolver"
7             .Left = FormMargin
8             .Top = FormMargin
9             .Height = FormHeadingHeight
10        End With
          
11        With txtVersion
12            .Locked = False
13            .value = "OpenSolver version information"
14            .Locked = True
15            .Width = lblHeading.Width
16            .Left = lblHeading.Left
17            .Top = Below(lblHeading, False)
18            .BackStyle = fmBackStyleTransparent
19        End With
          
20        With lblUrl
21            .Caption = "http://www.OpenSolver.org"
22            .ForeColor = FormLinkColor
23            .Left = lblHeading.Left
24            .Top = Below(txtVersion, False)
25            AutoHeight lblUrl, Me.Width, True
26        End With
          
27        With cmdUpdate
28            .Caption = "Check for updates"
29            .Width = FormButtonWidth * 2
30            .Left = LeftOfForm(Me.Width, .Width)
31            .Top = lblHeading.Top
32        End With
          
33        With cmdUpdateSettings
34            .Caption = "Update Check settings..."
35            .Width = cmdUpdate.Width
36            .Left = cmdUpdate.Left
37            .Top = Below(cmdUpdate)
38        End With
          
39        With txtAbout
40            .Locked = False
41            .Text = "Loading OpenSolver info..."
42            .Locked = True
43            .Left = lblHeading.Left
44            .Top = Max(Below(lblUrl), Below(cmdUpdateSettings))
45            .BackStyle = fmBackStyleTransparent
46            .SpecialEffect = fmSpecialEffectEtched
47            .Height = txtAboutHeight
48            .Width = lblHeading.Width
49        End With
          
50        With lblInstalled
51            .Font.Bold = True
52            .Left = lblHeading.Left
53            .Top = Below(txtAbout)
54            .Width = lblHeading.Width
55            .BackStyle = fmBackStyleTransparent
56            AutoHeight lblInstalled, .Width
57        End With
          
58        With txtFilePath
59            .Locked = False
60            .Text = "OpenSolver file:"
61            .Locked = True
62            .Left = lblHeading.Left
63            .Top = Below(lblInstalled)
64            .Height = FormTextHeight + 2 ' Stop the text becoming smaller
65            .Width = lblHeading.Width
66            .BackStyle = fmBackStyleTransparent
67            .MultiLine = True
68        End With
          
69        LayoutBottom
          
70        Me.Width = Me.Width + FormWindowMargin
          
71        Me.BackColor = FormBackColor
72        Me.Caption = "OpenSolver - About"
End Sub

Private Sub LayoutBottom()
1         With chkAutoLoad
2             .Caption = "Load OpenSolver when Excel starts"
3             AutoHeight chkAutoLoad, FormWidthAbout, True
4             .Left = lblHeading.Left
5             .Top = Below(txtFilePath, False)
6         End With
          
7         With cmdOk
8             .Caption = "OK"
9             .Width = FormButtonWidth
10            .Left = LeftOfForm(FormWidthAbout, .Width)
11            .Top = chkAutoLoad.Top
12            .Cancel = True
13        End With
          
14        Me.Height = FormHeight(cmdOk)
End Sub

Private Sub UserForm_Initialize()
1         AutoLayout
2         CenterForm
End Sub

Private Sub CenterForm()
1         Me.Top = CenterFormTop(Me.Height)
2         Me.Left = CenterFormLeft(Me.Width)
End Sub

