VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAutoModel 
   Caption         =   "OpenSolver - AutoModel"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8805
   OleObjectBlob   =   "frmAutoModel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAutoModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Mac Then
    Const FormWidthAutoModel = 450
#Else
    Const FormWidthAutoModel = 340
#End If

Public model As CModel
Public GuessObjStatus As String


Private Sub cmdCancel_Click()
4045      DoEvents
4047      Unload Me
End Sub

Private Sub UserForm_Activate()
          ' Reset enabled flags
4049      ResetEverything
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
4055      GuessObj
End Sub

Private Sub ResetEverything()
4056      optMax.value = False
4057      optMin.value = False
4058      refObj.Text = ""
End Sub

Private Sub GuessObj()
4059      Select Case GuessObjStatus
              Case "NoSense"
                  ' Didn't find anything
4060              lblStatus.Caption = "AutoModel was unable to guess anything." & vbNewLine & _
                                      "Please enter the objective sense and the objective function cell."
4061          Case "SenseNoCell"
4062              If model.ObjectiveSense = MaximiseObjective Then optMax.value = True
4063              If model.ObjectiveSense = MinimiseObjective Then optMin.value = True
                  lblStatus.Caption = "AutoModel found the objective sense, but couldn't find the objective cell." & vbNewLine & _
                                      "Please check the objective sense and enter the objective function cell."
4065      End Select
          
4066      Me.Repaint
4067      DoEvents
End Sub

Private Sub cmdFinish_Click()
          ' Check if user changed objective cell
4069      On Error GoTo BadObjRef
4070      Set model.ObjectiveFunctionCell = ActiveSheet.Range(refObj.Text)
          
          ' Get the objective sense
4071      If optMax.value = True Then model.ObjectiveSense = MaximiseObjective
4072      If optMin.value = True Then model.ObjectiveSense = MinimiseObjective
4073      If model.ObjectiveSense = UnknownObjectiveSense Then
4074          MsgBox "Error: Please select an objective sense (minimise or maximise)!", vbExclamation + vbOKOnly, "AutoModel"
4075          Exit Sub
4076      End If
          
          ' Find the vars, cons
          Dim result As Boolean
4077      result = model.FindVarsAndCons(IsFirstTime:=True)

4078      If result = False Then
              ' Didn't work/error
4079          MsgBox "An unknown error occurred while trying to find the model.", vbOKOnly
4080      End If
          
4081      Unload Me
4083      DoEvents
4084      Exit Sub
          
BadObjRef:
          ' Couldn't turn the objective cell address into a range
4085      MsgBox "Error: the cell address for the objective is invalid. Please correct " & _
                  "and click 'Finish AutoModel' again.", vbExclamation + vbOKOnly, "AutoModel"
4086      refObj.SetFocus ' Set the focus back to the RefEdit
4087      DoEvents ' Try to stop RefEdit bugs
4088      Exit Sub
End Sub


Private Sub UserForm_Initialize()
    AutoLayout
End Sub

Private Sub AutoLayout()
    AutoFormat Me.Controls
    
    Me.width = FormWidthAutoModel
    
    With lblStep1
        .Caption = "Determining the objective"
        .left = FormMargin
        .top = FormMargin
        ' Shrink width
        .width = Me.width
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
    End With
    
    With lblStep1Explanation
        .left = lblStep1.left + lblStep1.width + FormSpacing
        .top = lblStep1.top
        .width = Me.width - FormMargin - .left
        .Caption = "(the objective is what you want to optimise)"
    End With
    
    With lblStep1How
        .Caption = "AutoModel has tried to guess the ""sense"" you want to optimise by looking for " & _
                   """min"", ""max"", ""minimise"", etc. on the active spreadsheet. If it found it, " & _
                   "it looked in the area for something that might be the objective function cell, " & _
                   "e.g. a cell with a SUMPRODUCT() formula in it. If it cannot find anything, or " & _
                   "gets it wrong, you must enter the objective function cell so AutoModel can proceed."
        .left = lblStep1.left
        .top = lblStep1.top + lblStep1.height + FormSpacing
        .width = Me.width - FormMargin * 2
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
        .width = Me.width - FormMargin * 2
    End With
    
    With lblDiv1
        .left = lblStep1.left
        .top = lblStep1How.top + lblStep1How.height + FormSpacing
        .height = FormDivHeight
        .width = lblStep1How.width
        .BackColor = FormDivBackColor
    End With
    
    With lblStatus
        .Caption = "AutoModel was unable to guess anything." & vbNewLine & _
                   "Please enter the objective sense and objective function cell manually."
        .left = lblStep1.left
        .top = lblDiv1.top + lblDiv1.height + FormSpacing
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
        .height = .height + FormSpacing
        .width = lblStep1How.width
    End With
    
    With lblOpt1
        .Caption = "The objective is to:"
        .left = lblStep1.left
        .top = lblStatus.top + lblStatus.height + FormSpacing * 1.5 + optMax.height - .height / 2
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
    End With
    
    With optMax
        .Caption = "maximise"
        .left = lblOpt1.left + lblOpt1.width + FormSpacing
        .top = lblStatus.top + lblStatus.height + FormSpacing
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
    End With
    
    With optMin
        .Caption = "minimise"
        .left = optMax.left
        .top = optMax.top + optMax.height
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
    End With
    
    With lblOpt2
        .Caption = "the value of the cell:"
        .left = optMax.left + optMax.width + FormSpacing
        .top = lblOpt1.top
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
    End With
    
    With refObj
        .top = lblOpt2.top
        .left = lblOpt2.left + lblOpt2.width + FormSpacing
        .width = Me.width - FormMargin - .left
        .height = FormTextHeight
    End With
    
    With lblStep2How
        .top = optMin.top + optMin.height + FormSpacing
        .left = lblStep1.left
        .width = lblStep1How.width
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
    End With
    
    With lblDiv2
        .left = lblStep1.left
        .top = lblStep2How.top + lblStep2How.height + FormSpacing
        .height = FormDivHeight
        .width = lblStep1How.width
        .BackColor = FormDivBackColor
    End With
    
    With cmdCancel
        .width = FormButtonWidth * 1.2
        .left = Me.width - FormMargin - .width
        .top = lblDiv2.top + lblDiv2.height + FormSpacing
        .Caption = "Cancel"
    End With
    
    With cmdFinish
        .width = cmdCancel.width
        .left = cmdCancel.left - FormSpacing - .width
        .top = cmdCancel.top
        .Caption = "Finish AutoModel"
    End With
    
    Me.height = cmdCancel.top + cmdCancel.height + FormMargin + FormTitleHeight
    Me.width = Me.width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - AutoModel"
End Sub
