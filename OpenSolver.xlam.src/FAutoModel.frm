VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FAutoModel 
   Caption         =   "OpenSolver - AutoModel"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   -75
   ClientWidth     =   8820
   OleObjectBlob   =   "FAutoModel.frx":0000
End
Attribute VB_Name = "FAutoModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Mac Then
    Const FormWidthAutoModel = 500
#Else
    Const FormWidthAutoModel = 340
#End If

Public sheet As Worksheet
Public ObjectiveFunctionCellRefersTo As String
Public ObjectiveSense As ObjectiveSenseType

Private Sub cmdCancel_Click()
1         DoEvents
2         Me.Tag = "Cancelled"
3         Me.Hide
End Sub

' Make the [x] hide the form rather than unload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
                ' If CloseMode = vbFormControlMenu then we know the user
                ' clicked the [x] close button or Alt+F4 to close the form.
1               If CloseMode = vbFormControlMenu Then
2                   cmdCancel_Click
3                   Cancel = True
4               End If
End Sub

Private Sub UserForm_Activate()
1         CenterForm
          ' Make sure sheet is up to date
2         On Error Resume Next
3         Application.Calculate
4         On Error GoTo 0

          ' Remove the 'marching ants' showing if a range is copied.
          ' Otherwise, the ants stay visible, and visually conflict with
          ' our cell selection. The ants are also left behind on the
          ' screen. This works around an apparent bug (?) in Excel 2007.
5         Application.CutCopyMode = False

          ' Show results of finding objective
6         DoEvents
7         If ObjectiveSense = UnknownObjectiveSense Then
              ' Didn't find anything
8             lblStatus.Caption = "AutoModel was unable to guess anything." & vbNewLine & _
                                  "Please enter the objective sense and the objective function cell."
9         Else
10            If ObjectiveSense = MaximiseObjective Then optMax.value = True
11            If ObjectiveSense = MinimiseObjective Then optMin.value = True
12            lblStatus.Caption = "AutoModel found the objective sense, but couldn't find the objective cell." & vbNewLine & _
                                  "Please check the objective sense and enter the objective function cell."
13        End If
14        Me.Repaint
15        DoEvents

16        Me.Tag = vbNullString
End Sub

Private Sub ResetEverything()
1         ObjectiveFunctionCellRefersTo = vbNullString
2         refObj.Text = vbNullString
3         ObjectiveSense = UnknownObjectiveSense
4         optMax.value = False
5         optMin.value = False
End Sub

Private Sub cmdFinish_Click()
          ' Get the objective sense
1         If optMax.value = True Then ObjectiveSense = MaximiseObjective
2         If optMin.value = True Then ObjectiveSense = MinimiseObjective
          
3         If ObjectiveSense = UnknownObjectiveSense Then
4             If Len(refObj.Text) = 0 Then
                  ' We allow a blank objective if no sense is set.
                  ' Set a valid sense
5                 ObjectiveSense = MinimiseObjective
6             Else
7                 GoTo BadObjSense
8             End If
9         Else
10            On Error GoTo BadObjRef
11            ObjectiveFunctionCellRefersTo = RefEditToRefersTo(refObj.Text)
12            ValidateObjectiveFunctionCellRefersTo ObjectiveFunctionCellRefersTo
13            On Error GoTo 0
14        End If
          
ExitSub:
15        Me.Hide
16        DoEvents
17        Exit Sub
          
BadObjRef:
18        MsgBox "Error: the cell address for the objective is invalid. " & _
                 "This must be a single cell. " & _
                 "Please correct this and click 'Finish AutoModel' again.", vbExclamation + vbOKOnly, "AutoModel"
19        refObj.SetFocus ' Set the focus back to the RefEdit
20        DoEvents ' Try to stop RefEdit bugs
21        Exit Sub

BadObjSense:
22        MsgBox "Error: Please select an objective sense (minimise or maximise)!", vbExclamation + vbOKOnly, "AutoModel"
23        Exit Sub
End Sub


Private Sub UserForm_Initialize()
1         AutoLayout
2         ResetEverything
3         CenterForm
End Sub

Private Sub AutoLayout()
1         AutoFormat Me.Controls
          
2         Me.Width = FormWidthAutoModel
          
3         With lblStep1
4             .Caption = "Determining the objective"
5             .Left = FormMargin
6             .Top = FormMargin
              ' Shrink width
7             AutoHeight lblStep1, Me.Width, True
8         End With
          
9         With lblStep1Explanation
10            .Left = RightOf(lblStep1)
11            .Top = lblStep1.Top
12            .Width = LeftOfForm(Me.Width, .Left)
13            .Caption = "(the objective is what you want to optimise)"
14        End With
          
15        With lblStep1How
16            .Caption = "AutoModel has tried to guess the ""sense"" you want to optimise by looking for " & _
                         """min"", ""max"", ""minimise"", etc. on the active spreadsheet. If it found it, " & _
                         "it looked in the area for something that might be the objective function cell, " & _
                         "e.g. a cell with a SUMPRODUCT() formula in it. If it cannot find anything, or " & _
                         "gets it wrong, you must enter the objective function cell so AutoModel can proceed. " & _
                         "You can also leave both the objective sense and the objective cell blank. " & _
                         "In this case, OpenSolver will just find a feasible solution to the problem."
17            .Left = lblStep1.Left
18            .Top = Below(lblStep1)
19            AutoHeight lblStep1How, Me.Width - FormMargin * 2
20        End With
          
21        With lblDiv1
22            .Left = lblStep1.Left
23            .Top = Below(lblStep1How)
24            .Height = FormDivHeight
25            .Width = lblStep1How.Width
26            .BackColor = FormDivBackColor
27        End With
          
28        With lblStatus
29            .Caption = "AutoModel was unable to guess anything." & vbNewLine & _
                         "Please enter the objective sense and objective function cell manually."
30            .Left = lblStep1.Left
31            .Top = Below(lblDiv1)
32            AutoHeight lblStatus, lblStep1How.Width
33            .Height = .Height + FormSpacing
34        End With
          
35        With lblOpt1
36            .Caption = "The objective is to:"
37            .Left = lblStep1.Left
38            .Top = Below(lblStatus) + FormSpacing * 1.5 + optMax.Height - .Height / 2
39            AutoHeight lblOpt1, Me.Width, True
40        End With
          
41        With optMax
42            .Caption = "maximise"
43            .Left = RightOf(lblOpt1)
44            .Top = Below(lblStatus)
45            AutoHeight optMax, Me.Width, True
46        End With
          
47        With optMin
48            .Caption = "minimise"
49            .Left = optMax.Left
50            .Top = Below(optMax, False)
51            AutoHeight optMin, Me.Width, True
52        End With
          
53        With lblOpt2
54            .Caption = "the value of the cell:"
55            .Left = RightOf(optMax)
56            .Top = lblOpt1.Top
57            AutoHeight lblOpt2, Me.Width, True
58        End With
          
59        With refObj
60            .Top = lblOpt2.Top - (.Height - lblOpt2.Height) / 2
61            .Left = RightOf(lblOpt2)
62            .Width = LeftOfForm(Me.Width, .Left)
63        End With
          
64        With lblStep2How
65            .Top = Below(optMin)
66            .Left = lblStep1.Left
67            AutoHeight lblStep2How, lblStep1How.Width, True
68        End With
          
69        With lblDiv2
70            .Left = lblStep1.Left
71            .Top = Below(lblStep2How)
72            .Height = FormDivHeight
73            .Width = lblStep1How.Width
74            .BackColor = FormDivBackColor
75        End With
          
76        With cmdCancel
77            .Width = FormButtonWidth * 1.2
78            .Left = LeftOfForm(Me.Width, .Width)
79            .Top = Below(lblDiv2)
80            .Caption = "Cancel"
81            .Cancel = True
82        End With
          
83        With cmdFinish
84            .Width = cmdCancel.Width
85            .Left = LeftOf(cmdCancel, .Width)
86            .Top = cmdCancel.Top
87            .Caption = "Finish AutoModel"
88        End With
          
89        With chkShow
90            .Left = lblStep1.Left
91            .Top = cmdCancel.Top
92            .Width = LeftOf(cmdFinish, .Left, False)
93            .Caption = "Show model on sheet when finished"
94            .value = True
95        End With
          
96        Me.Height = FormHeight(cmdCancel)
97        Me.Width = Me.Width + FormWindowMargin
          
98        Me.BackColor = FormBackColor
99        Me.Caption = "OpenSolver - AutoModel"
End Sub

Private Sub CenterForm()
1         Me.Top = CenterFormTop(Me.Height)
2         Me.Left = CenterFormLeft(Me.Width)
End Sub
