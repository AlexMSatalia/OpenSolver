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

    Me.width = FormWidthOptions
       
    With chkNonNeg
        .Caption = "Make unconstrained variable cells non-negative"
        .left = FormMargin
        .top = FormMargin
        .width = Me.width - 2 * FormMargin
    End With
       
    With chkPerformLinearityCheck
        .Caption = "Perform a quick linearity check on the solution"
        .left = chkNonNeg.left
        .top = Below(chkNonNeg, False)
        .width = chkNonNeg.width
    End With
        
    With chkShowSolverProgress
        .Caption = "Show optimisation progress while solving"
        .left = chkNonNeg.left
        .top = Below(chkPerformLinearityCheck, False)
        .width = chkNonNeg.width
    End With
    
    With txtMaxTime
        .width = FormButtonWidth
        .left = LeftOfForm(Me.width, .width)
        .top = Below(chkShowSolverProgress)
    End With
    
    With lblMaxTime
        .Caption = "Maximum Solution Time (seconds):"
        .left = chkNonNeg.left
        .width = LeftOf(txtMaxTime, .left)
        .top = txtMaxTime.top
    End With
    
    With txtTol
        .width = txtMaxTime.width
        .left = txtMaxTime.left
        .top = Below(txtMaxTime)
    End With
    
    With lblTol
        .Caption = "Branch and Bound Tolerance (%):"
        .left = lblMaxTime.left
        .width = lblMaxTime.width
        .top = txtTol.top
    End With
    
    With txtMaxIter
        .width = txtMaxTime.width
        .left = txtMaxTime.left
        .top = Below(txtTol)
    End With
    
    With lblMaxIter
        .Caption = "Maximum Number of Iterations:"
        .left = lblMaxTime.left
        .width = lblMaxTime.width
        .top = txtMaxIter.top
    End With
    
    With txtPre
        .width = txtMaxTime.width
        .left = txtMaxTime.left
        .top = Below(txtMaxIter)
    End With
    
    With lblPre
        .Caption = "Precision:"
        .left = lblMaxTime.left
        .width = lblMaxTime.width
        .top = txtPre.top
    End With
    
    With lblExtraParameters
        .Caption = "Extra Solver Parameters:"
        .left = chkNonNeg.left
        .width = chkNonNeg.width
        .top = Below(txtPre, False)
    End With
    
    With refExtraParameters
        .width = chkNonNeg.width
        .left = chkNonNeg.left
        .top = Below(lblExtraParameters, False) - FormSpacing
    End With
    
    With lblFootnote
        .Caption = "Note: Only options that are used by the currently selected solver can be changed"
        .top = Below(refExtraParameters)
        .left = chkNonNeg.left
        AutoHeight lblFootnote, chkNonNeg.width
    End With
    
    With cmdCancel
        .Caption = "Cancel"
        .left = txtMaxTime.left
        .width = txtMaxTime.width
        .top = Below(lblFootnote)
    End With
    
    With cmdOk
        .Caption = "OK"
        .width = txtMaxTime.width
        .left = LeftOf(cmdCancel, .width)
        .top = cmdCancel.top
    End With
    
    Me.height = FormHeight(cmdCancel)
    Me.width = Me.width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - Solve Options"
End Sub

Private Sub CenterForm()
    Me.top = CenterFormTop(Me.height)
    Me.left = CenterFormLeft(Me.width)
End Sub
