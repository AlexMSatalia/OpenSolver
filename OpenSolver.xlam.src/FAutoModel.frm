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

Public ObjectiveCell As Range
Public ObjectiveSense As ObjectiveSenseType

Private Sub cmdCancel_Click()
4045      DoEvents
          Me.Tag = "Cancelled"
4047      Me.Hide
End Sub

' Make the [x] hide the form rather than unload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
          ' If CloseMode = vbFormControlMenu then we know the user
          ' clicked the [x] close button or Alt+F4 to close the form.
          If CloseMode = vbFormControlMenu Then
              cmdCancel_Click
              Cancel = True
          End If
End Sub

Private Sub UserForm_Activate()
          CenterForm
          ' Make sure sheet is up to date
4050      On Error Resume Next
4051      Application.Calculate
4052      On Error GoTo 0
          ' Remove the 'marching ants' showing if a range is copied.
          ' Otherwise, the ants stay visible, and visually conflict with
          ' our cell selection. The ants are also left behind on the
          ' screen. This works around an apparent bug (?) in Excel 2007.
4053      Application.CutCopyMode = False
          ' Get ready to process
4054      DoEvents
          ' Show results of finding objective
4055      SetValues
          Me.Tag = ""
End Sub

Private Sub ResetEverything()
4056      Set ObjectiveCell = Nothing
          refObj.Text = ""
4057      ObjectiveSense = UnknownObjectiveSense
          optMax.value = False
          optMin.value = False
End Sub

Private Sub SetValues()
4059      If ObjectiveSense = UnknownObjectiveSense Then
              ' Didn't find anything
4060          lblStatus.Caption = "AutoModel was unable to guess anything." & vbNewLine & _
                                  "Please enter the objective sense and the objective function cell."
4061      Else
4062          If ObjectiveSense = MaximiseObjective Then optMax.value = True
4063          If ObjectiveSense = MinimiseObjective Then optMin.value = True
              lblStatus.Caption = "AutoModel found the objective sense, but couldn't find the objective cell." & vbNewLine & _
                                  "Please check the objective sense and enter the objective function cell."
4065      End If
          
4066      Me.Repaint
4067      DoEvents
End Sub

Private Sub cmdFinish_Click()
          ' Get the objective sense
4071      If optMax.value = True Then ObjectiveSense = MaximiseObjective
4072      If optMin.value = True Then ObjectiveSense = MinimiseObjective
          
          ' Check if user changed objective cell
4069      On Error GoTo BadObjRef
4070      Set ObjectiveCell = ActiveSheet.Range(refObj.Text)
          
          
4073      If ObjectiveSense = UnknownObjectiveSense Then
4074          MsgBox "Error: Please select an objective sense (minimise or maximise)!", vbExclamation + vbOKOnly, "AutoModel"
4075          Exit Sub
4076      End If
          
ExitSub:
4081      Me.Hide
4083      DoEvents
4084      Exit Sub
          
BadObjRef:
          ' Couldn't turn the objective cell address into a range
          ' If no objective sense is specified, show an error. We allow a blank objective if no sense is set.
4085      If ObjectiveSense <> UnknownObjectiveSense Then
              MsgBox "Error: the cell address for the objective is invalid. Please correct " & _
                     "and click 'Finish AutoModel' again.", vbExclamation + vbOKOnly, "AutoModel"
4086          refObj.SetFocus ' Set the focus back to the RefEdit
4087          DoEvents ' Try to stop RefEdit bugs
          Else
              ' Set a valid sense
              ObjectiveSense = MinimiseObjective
              Resume ExitSub
          End If
4088      Exit Sub
End Sub


Private Sub UserForm_Initialize()
    AutoLayout
    ResetEverything
    CenterForm
End Sub

Private Sub AutoLayout()
    AutoFormat Me.Controls
    
    Me.Width = FormWidthAutoModel
    
    With lblStep1
        .Caption = "Determining the objective"
        .Left = FormMargin
        .Top = FormMargin
        ' Shrink width
        AutoHeight lblStep1, Me.Width, True
    End With
    
    With lblStep1Explanation
        .Left = RightOf(lblStep1)
        .Top = lblStep1.Top
        .Width = LeftOfForm(Me.Width, .Left)
        .Caption = "(the objective is what you want to optimise)"
    End With
    
    With lblStep1How
        .Caption = "AutoModel has tried to guess the ""sense"" you want to optimise by looking for " & _
                   """min"", ""max"", ""minimise"", etc. on the active spreadsheet. If it found it, " & _
                   "it looked in the area for something that might be the objective function cell, " & _
                   "e.g. a cell with a SUMPRODUCT() formula in it. If it cannot find anything, or " & _
                   "gets it wrong, you must enter the objective function cell so AutoModel can proceed. " & _
                   "You can also leave both the objective sense and the objective cell blank. " & _
                   "In this case, OpenSolver will just find a feasible solution to the problem."
        .Left = lblStep1.Left
        .Top = Below(lblStep1)
        AutoHeight lblStep1How, Me.Width - FormMargin * 2
    End With
    
    With lblDiv1
        .Left = lblStep1.Left
        .Top = Below(lblStep1How)
        .Height = FormDivHeight
        .Width = lblStep1How.Width
        .BackColor = FormDivBackColor
    End With
    
    With lblStatus
        .Caption = "AutoModel was unable to guess anything." & vbNewLine & _
                   "Please enter the objective sense and objective function cell manually."
        .Left = lblStep1.Left
        .Top = Below(lblDiv1)
        AutoHeight lblStatus, lblStep1How.Width
        .Height = .Height + FormSpacing
    End With
    
    With lblOpt1
        .Caption = "The objective is to:"
        .Left = lblStep1.Left
        .Top = Below(lblStatus) + FormSpacing * 1.5 + optMax.Height - .Height / 2
        AutoHeight lblOpt1, Me.Width, True
    End With
    
    With optMax
        .Caption = "maximise"
        .Left = RightOf(lblOpt1)
        .Top = Below(lblStatus)
        AutoHeight optMax, Me.Width, True
    End With
    
    With optMin
        .Caption = "minimise"
        .Left = optMax.Left
        .Top = Below(optMax, False)
        AutoHeight optMin, Me.Width, True
    End With
    
    With lblOpt2
        .Caption = "the value of the cell:"
        .Left = RightOf(optMax)
        .Top = lblOpt1.Top
        AutoHeight lblOpt2, Me.Width, True
    End With
    
    With refObj
        .Top = lblOpt2.Top - (.Height - lblOpt2.Height) / 2
        .Left = RightOf(lblOpt2)
        .Width = LeftOfForm(Me.Width, .Left)
    End With
    
    With lblStep2How
        .Top = Below(optMin)
        .Left = lblStep1.Left
        AutoHeight lblStep2How, lblStep1How.Width, True
    End With
    
    With lblDiv2
        .Left = lblStep1.Left
        .Top = Below(lblStep2How)
        .Height = FormDivHeight
        .Width = lblStep1How.Width
        .BackColor = FormDivBackColor
    End With
    
    With cmdCancel
        .Width = FormButtonWidth * 1.2
        .Left = LeftOfForm(Me.Width, .Width)
        .Top = Below(lblDiv2)
        .Caption = "Cancel"
    End With
    
    With cmdFinish
        .Width = cmdCancel.Width
        .Left = LeftOf(cmdCancel, .Width)
        .Top = cmdCancel.Top
        .Caption = "Finish AutoModel"
    End With
    
    With chkShow
        .Left = lblStep1.Left
        .Top = cmdCancel.Top
        .Width = LeftOf(cmdFinish, .Left, False)
        .Caption = "Show model on sheet when finished"
        .value = True
    End With
    
    Me.Height = FormHeight(cmdCancel)
    Me.Width = Me.Width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - AutoModel"
End Sub

Private Sub CenterForm()
    Me.Top = CenterFormTop(Me.Height)
    Me.Left = CenterFormLeft(Me.Width)
End Sub
