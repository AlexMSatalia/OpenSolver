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

Private Sub cmdCancel_Click()
4089      Me.Hide
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

Private Sub cmdOk_Click()
          On Error GoTo ErrorHandler
          
          Dim ParametersRange As Range
          If Len(refExtraParameters.Text) <> 0 Then
              Set ParametersRange = Range(refExtraParameters.Text)
              ValidateParametersRange ParametersRange
          End If

4092      SetNonNegativity chkNonNeg.value
4098      SetShowSolverProgress chkShowSolverProgress.value
4102      SetMaxTime CDbl(txtMaxTime.Text)
4103      SetMaxIterations CDbl(txtMaxIter.Text)
4104      SetPrecision CDbl(txtPre.Text)
4106      SetToleranceAsPercentage CDbl(Replace(txtTol.Text, "%", ""))
4107      SetLinearityCheck chkPerformLinearityCheck.value
          SetSolverParameters SolverString, ParametersRange
                                                                      
4112      Me.Hide
          Exit Sub

ErrorHandler:
          MsgBox Err.Description
End Sub

Private Sub UserForm_Activate()
          CenterForm
          
4114      SetAnyMissingDefaultSolverOptions

4129      chkNonNeg.value = GetNonNegativity()
4130      chkShowSolverProgress.value = GetShowSolverProgress()
4131      txtMaxTime.Text = CStr(GetMaxTime())
4132      txtTol.Text = CStr(GetToleranceAsPercentage())
4133      txtMaxIter.Text = CStr(GetMaxIterations())
4134      txtPre = CStr(GetPrecision())
4135      chkPerformLinearityCheck.value = GetLinearityCheck()

          Dim Solver As ISolver
4136      SolverString = GetChosenSolver()
          Set Solver = CreateSolver(SolverString)

          chkPerformLinearityCheck.Enabled = (SolverLinearity(Solver) = Linear) And _
                                             Solver.ModelType = Diff
          txtMaxIter.Enabled = IterationLimitAvailable(Solver)
          txtPre.Enabled = PrecisionAvailable(Solver)
          txtMaxTime.Enabled = TimeLimitAvailable(Solver)
          txtTol.Enabled = ToleranceAvailable(Solver)
          
          refExtraParameters.Text = GetDisplayAddress(GetSolverParameters(SolverString), False)
End Sub

Private Sub UserForm_Initialize()
    AutoLayout
    CenterForm
End Sub

Private Sub AutoLayout()
    AutoFormat Me.Controls

    Me.Width = FormWidthOptions
       
    With chkNonNeg
        .Caption = "Make unconstrained variable cells non-negative"
        .Left = FormMargin
        .Top = FormMargin
        .Width = Me.Width - 2 * FormMargin
    End With
       
    With chkPerformLinearityCheck
        .Caption = "Perform a quick linearity check on the solution"
        .Left = chkNonNeg.Left
        .Top = Below(chkNonNeg, False)
        .Width = chkNonNeg.Width
    End With
        
    With chkShowSolverProgress
        .Caption = "Show optimisation progress while solving"
        .Left = chkNonNeg.Left
        .Top = Below(chkPerformLinearityCheck, False)
        .Width = chkNonNeg.Width
    End With
    
    With txtMaxTime
        .Width = FormButtonWidth
        .Left = LeftOfForm(Me.Width, .Width)
        .Top = Below(chkShowSolverProgress)
    End With
    
    With lblMaxTime
        .Caption = "Maximum Solution Time (seconds):"
        .Left = chkNonNeg.Left
        .Width = LeftOf(txtMaxTime, .Left)
        .Top = txtMaxTime.Top
    End With
    
    With txtTol
        .Width = txtMaxTime.Width
        .Left = txtMaxTime.Left
        .Top = Below(txtMaxTime)
    End With
    
    With lblTol
        .Caption = "Branch and Bound Tolerance (%):"
        .Left = lblMaxTime.Left
        .Width = lblMaxTime.Width
        .Top = txtTol.Top
    End With
    
    With txtMaxIter
        .Width = txtMaxTime.Width
        .Left = txtMaxTime.Left
        .Top = Below(txtTol)
    End With
    
    With lblMaxIter
        .Caption = "Maximum Number of Iterations:"
        .Left = lblMaxTime.Left
        .Width = lblMaxTime.Width
        .Top = txtMaxIter.Top
    End With
    
    With txtPre
        .Width = txtMaxTime.Width
        .Left = txtMaxTime.Left
        .Top = Below(txtMaxIter)
    End With
    
    With lblPre
        .Caption = "Precision:"
        .Left = lblMaxTime.Left
        .Width = lblMaxTime.Width
        .Top = txtPre.Top
    End With
    
    With lblExtraParameters
        .Caption = "Extra Solver Parameters:"
        .Left = chkNonNeg.Left
        .Width = chkNonNeg.Width
        .Top = Below(txtPre, False)
    End With
    
    With refExtraParameters
        .Width = chkNonNeg.Width
        .Left = chkNonNeg.Left
        .Top = Below(lblExtraParameters, False) - FormSpacing
    End With
    
    With lblFootnote
        .Caption = "Note: Only options that are used by the currently selected solver can be changed"
        .Top = Below(refExtraParameters)
        .Left = chkNonNeg.Left
        AutoHeight lblFootnote, chkNonNeg.Width
    End With
    
    With cmdCancel
        .Caption = "Cancel"
        .Left = txtMaxTime.Left
        .Width = txtMaxTime.Width
        .Top = Below(lblFootnote)
    End With
    
    With cmdOk
        .Caption = "OK"
        .Width = txtMaxTime.Width
        .Left = LeftOf(cmdCancel, .Width)
        .Top = cmdCancel.Top
    End With
    
    Me.Height = FormHeight(cmdCancel)
    Me.Width = Me.Width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - Solve Options"
End Sub

Private Sub CenterForm()
    Me.Top = CenterFormTop(Me.Height)
    Me.Left = CenterFormLeft(Me.Width)
End Sub
