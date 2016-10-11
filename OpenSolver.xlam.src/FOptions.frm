VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FOptions 
   Caption         =   "OpenSolver - Solve Options"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4140
   OleObjectBlob   =   "FOptions.frx":0000
End
Attribute VB_Name = "FOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Mac Then
    Const FormWidthOptions = 318
#Else
    Const FormWidthOptions = 212
#End If

Private SolverString As String
Private sheet As Worksheet

Private Sub cmdCancel_Click()
1         Me.Hide
End Sub

Private Sub lblExtraParametersHelp_Click()
1         OpenURL "http://opensolver.org/using-opensolver/#extra-parameters"
End Sub

' Make the [x] hide the form rather than unload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
          ' If CloseMode = vbFormControlMenu then we know the user
          ' clicked the [x] close button or Alt+F4 to close the form.
1         If CloseMode = vbFormControlMenu Then
2             cmdCancel_Click
3             Cancel = True
4         End If
End Sub

Private Sub cmdOk_Click()
1         On Error GoTo ErrorHandler
          
          ' All validation
          
          Dim SolverParametersRefersTo As String
2         SolverParametersRefersTo = RefEditToRefersTo(refExtraParameters.Text)
3         ValidateSolverParametersRefersTo SolverParametersRefersTo
          
          ' Save confirmed!

4         SetNonNegativity chkNonNeg.value, sheet
5         SetShowSolverProgress chkShowSolverProgress.value, sheet
6         SetMaxTime FormatNumberForSaving(txtMaxTime.Text), sheet
7         SetMaxIterations FormatNumberForSaving(txtMaxIter.Text), sheet
8         SetPrecision CDbl(txtPre.Text), sheet
9         SetToleranceAsPercentage CDbl(Replace(txtTol.Text, "%", vbNullString)), sheet
10        SetLinearityCheck chkPerformLinearityCheck.value, sheet
11        SetSolverParametersRefersTo SolverString, SolverParametersRefersTo, sheet
                                                                      
12        Me.Hide
13        Exit Sub

ErrorHandler:
14        MsgBox Err.Description
End Sub

Private Function FormatNumberForDisplay(Number As Double) As String
1               If Number = MAX_LONG Then
2                   FormatNumberForDisplay = vbNullString
3               Else
4                   FormatNumberForDisplay = CStr(Number)
5               End If
End Function

Private Function FormatNumberForSaving(Number As String) As Double
1               If Number = vbNullString Then
2                   FormatNumberForSaving = MAX_LONG
3               Else
4                   FormatNumberForSaving = CDbl(Number)
5               End If
End Function

Private Sub UserForm_Activate()
1         CenterForm
          
2         GetActiveSheetIfMissing sheet
          
3         SetAnyMissingDefaultSolverOptions sheet

4         chkNonNeg.value = GetNonNegativity(sheet)
5         chkShowSolverProgress.value = GetShowSolverProgress(sheet)
6         txtMaxTime.Text = FormatNumberForDisplay(GetMaxTime(sheet))
7         txtMaxIter.Text = FormatNumberForDisplay(GetMaxIterations(sheet))
8         txtTol.Text = CStr(GetToleranceAsPercentage(sheet))
9         txtPre.Text = CStr(GetPrecision(sheet))
10        chkPerformLinearityCheck.value = GetLinearityCheck(sheet)

          Dim Solver As ISolver
11        SolverString = GetChosenSolver(sheet)
12        Set Solver = CreateSolver(SolverString)

13        chkPerformLinearityCheck.Enabled = (SolverLinearity(Solver) = Linear) And _
                                             Solver.ModelType = Diff
14        txtMaxIter.Enabled = IterationLimitAvailable(Solver)
15        txtPre.Enabled = PrecisionAvailable(Solver)
16        txtMaxTime.Enabled = TimeLimitAvailable(Solver)
17        txtTol.Enabled = ToleranceAvailable(Solver)
          
18        refExtraParameters.Text = GetDisplayAddress(GetSolverParametersRefersTo(SolverString, sheet), sheet, False)
End Sub

Private Sub UserForm_Initialize()
1         AutoLayout
2         CenterForm
End Sub

Private Sub AutoLayout()
1         AutoFormat Me.Controls

2         Me.Width = FormWidthOptions
             
3         With chkNonNeg
4             .Caption = "Make unconstrained variable cells non-negative"
5             .Left = FormMargin
6             .Top = FormMargin
7             .Width = Me.Width - 2 * FormMargin
8         End With
             
9         With chkPerformLinearityCheck
10            .Caption = "Perform a quick linearity check on the model"
11            .Left = chkNonNeg.Left
12            .Top = Below(chkNonNeg, False)
13            .Width = chkNonNeg.Width
14        End With
              
15        With chkShowSolverProgress
16            .Caption = "Show optimisation progress while solving"
17            .Left = chkNonNeg.Left
18            .Top = Below(chkPerformLinearityCheck, False)
19            .Width = chkNonNeg.Width
20        End With
          
21        With txtMaxTime
22            .Width = FormButtonWidth
23            .Left = LeftOfForm(Me.Width, .Width)
24            .Top = Below(chkShowSolverProgress)
25        End With
          
26        With lblMaxTime
27            .Caption = "Maximum Solution Time (seconds):"
28            .Left = chkNonNeg.Left
29            .Width = LeftOf(txtMaxTime, .Left)
30            .Top = txtMaxTime.Top
31        End With
          
32        With txtTol
33            .Width = txtMaxTime.Width
34            .Left = txtMaxTime.Left
35            .Top = Below(txtMaxTime)
36        End With
          
37        With lblTol
38            .Caption = "Branch and Bound Tolerance (%):"
39            .Left = lblMaxTime.Left
40            .Width = lblMaxTime.Width
41            .Top = txtTol.Top
42        End With
          
43        With txtMaxIter
44            .Width = txtMaxTime.Width
45            .Left = txtMaxTime.Left
46            .Top = Below(txtTol)
47        End With
          
48        With lblMaxIter
49            .Caption = "Maximum Number of Iterations:"
50            .Left = lblMaxTime.Left
51            .Width = lblMaxTime.Width
52            .Top = txtMaxIter.Top
53        End With
          
54        With txtPre
55            .Width = txtMaxTime.Width
56            .Left = txtMaxTime.Left
57            .Top = Below(txtMaxIter)
58        End With
          
59        With lblPre
60            .Caption = "Precision:"
61            .Left = lblMaxTime.Left
62            .Width = lblMaxTime.Width
63            .Top = txtPre.Top
64        End With
          
65        With lblExtraParameters
66            .Caption = "Extra Solver Parameters Range:"
67            .Left = chkNonNeg.Left
68            .Width = chkNonNeg.Width
69            .Top = Below(txtPre)
70        End With
          
71        With lblExtraParametersHelp
72            .Caption = "What's this?"
73            .Width = Me.Width
74            AutoHeight lblExtraParametersHelp, Me.Width, True
75            .Left = LeftOfForm(Me.Width, .Width)
76            .Top = lblExtraParameters.Top
77            .Font.Underline = True
78            .ForeColor = FormLinkColor
79        End With
          
80        With refExtraParameters
81            .Width = chkNonNeg.Width
82            .Left = chkNonNeg.Left
83            .Top = Below(lblExtraParameters, False) - FormSpacing / 2
84        End With
          
85        With lblFootnote
86            .Caption = "Note: Only options that are used by the currently selected solver can be changed"
87            .Top = Below(refExtraParameters)
88            .Left = chkNonNeg.Left
89            AutoHeight lblFootnote, chkNonNeg.Width
90        End With
          
91        With cmdCancel
92            .Caption = "Cancel"
93            .Left = txtMaxTime.Left
94            .Width = txtMaxTime.Width
95            .Top = Below(lblFootnote)
96            .Cancel = True
97        End With
          
98        With cmdOk
99            .Caption = "OK"
100           .Width = txtMaxTime.Width
101           .Left = LeftOf(cmdCancel, .Width)
102           .Top = cmdCancel.Top
103       End With
          
104       Me.Height = FormHeight(cmdCancel)
105       Me.Width = Me.Width + FormWindowMargin
          
106       Me.BackColor = FormBackColor
107       Me.Caption = "OpenSolver - Solve Options"
End Sub

Private Sub CenterForm()
1         Me.Top = CenterFormTop(Me.Height)
2         Me.Left = CenterFormLeft(Me.Width)
End Sub
