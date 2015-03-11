VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptions 
   Caption         =   "OpenSolver - Solve Options"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4140
   OleObjectBlob   =   "frmOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOptions"
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

Private Sub cmdCancel_Click()
4089      Unload Me
End Sub

Private Sub cmdOk_Click()
4092      SetNonNegativity chkNonNeg.value
4098      SetShowSolverProgress chkShowSolverProgress.value
4102      SetMaxTime CDbl(txtMaxTime.Text)
4103      SetMaxIterations CDbl(txtMaxIter.Text)
4104      SetPrecision CDbl(txtPre.Text)
4106      SetToleranceAsPercentage CDbl(Replace(txtTol.Text, "%", ""))
4107      SetLinearityCheck chkPerformLinearityCheck.value
                                                                      
4112      Unload Me
End Sub

Private Sub UserForm_Activate()
4114      SetAnyMissingDefaultExcel2007SolverOptions

4129      chkNonNeg.value = GetNonNegativity()
4130      chkShowSolverProgress.value = GetShowSolverProgress()
4131      txtMaxTime.Text = CStr(GetMaxTime())
4132      txtTol.Text = CStr(GetToleranceAsPercentage())
4133      txtMaxIter.Text = CStr(GetMaxIterations())
4134      txtPre = CStr(GetPrecision())
4135      chkPerformLinearityCheck.value = GetLinearityCheck()

          Dim Solver As String
4136      Solver = GetChosenSolver()

          chkPerformLinearityCheck.Enabled = (SolverType(Solver) = OpenSolver_SolverType.Linear) And _
                                              Not UsesParsedModel(Solver)
          txtPre.Enabled = UsesPrecision(Solver)
          txtMaxIter.Enabled = UsesIterationLimit(Solver)
          txtTol.Enabled = UsesTolerance(Solver)
          txtMaxTime.Enabled = UsesTimeLimit(Solver)
End Sub

Private Sub UserForm_Initialize()
    AutoLayout
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
    
    With lblFootnote
        .Caption = "Note: Only options that are used by the currently selected solver can be changed"
        .top = Below(txtPre)
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
