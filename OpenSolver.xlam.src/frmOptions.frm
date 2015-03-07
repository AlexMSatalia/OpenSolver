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
4092      If chkNonNeg.value = True Then
4093          SetSolverNameOnSheet "neg", "=1"
4094      Else
4095          SetSolverNameOnSheet "neg", "=2"     ' 2 means false
4096      End If
          
4097      If chkShowSolverProgress.value = True Then
4098          SetSolverNameOnSheet "sho", "=1"
4099      Else
4100          SetSolverNameOnSheet "sho", "=2"     ' 2 means false
4101      End If
          
4102      SetSolverNameOnSheet "tim", "=" & Trim(str(CDbl(txtMaxTime.Text)))  ' Trim the leading space which str puts in for +'ve values
4103      SetSolverNameOnSheet "itr", "=" & Trim(str(CDbl(txtMaxIter.Text)))  ' Trim the leading space which str puts in for +'ve values
4104      SetSolverNameOnSheet "pre", "=" & Trim(str(CDbl(txtPre.Text)))  ' Trim the leading space which str puts in for +'ve values
4105      txtTol.Text = Replace(txtTol.Text, "%", "")
4106      SetSolverNameOnSheet "tol", "=" & Trim(str(CDbl(txtTol.Text) / 100))    ' Str() uses . for decimal
                                                                      ' CDbl respects the locale. We trim the leading space which str puts in for +'ve values
                                                                      
4107      If chkPerformLinearityCheck.value = True Then
              ' Default is "do check", so we just delete the option
4108          DeleteNameOnSheet "OpenSolver_LinearityCheck"
4109      Else
              ' Set the name, with a value of 2=off
4110          SetNameOnSheet "OpenSolver_LinearityCheck", "=2"
4111      End If
                                                                      
4112      Unload Me
End Sub

Private Sub UserForm_Activate()
4114      SetAnyMissingDefaultExcel2007SolverOptions

          Dim sheetName As String
          sheetName = EscapeSheetName(ActiveSheet)

          Dim nonNeg As Boolean, s As String
4115      If GetNameValueIfExists(ActiveWorkbook, sheetName & "solver_neg", s) Then
4116          nonNeg = s = "1"
4117      End If
          
          Dim ShowSolverProgress As Boolean
4118      If GetNameValueIfExists(ActiveWorkbook, sheetName & "solver_sho", s) Then
4119          ShowSolverProgress = s = "1"
4120      End If
          
          Dim MaxTime As Double
4121      GetNamedNumericValueIfExists ActiveWorkbook, sheetName & "solver_tim", MaxTime
          
          Dim maxIter As Double
4122      GetNamedNumericValueIfExists ActiveWorkbook, sheetName & "solver_itr", maxIter

          Dim conPre As Double
4123      GetNamedNumericValueIfExists ActiveWorkbook, sheetName & "solver_pre", conPre

          Dim tol As Double
4124      GetNamedNumericValueIfExists ActiveWorkbook, sheetName & "solver_tol", tol
          
          ' We perform a linearity check by default unless the defined name exists with value 2=off
          Dim performLinearityCheck As Boolean
4125      performLinearityCheck = True
4126      If GetNameValueIfExists(ActiveWorkbook, sheetName & "OpenSolver_LinearityCheck", s) Then
4127          performLinearityCheck = s = "1"
4128      End If

4129      chkNonNeg.value = nonNeg
4130      chkShowSolverProgress.value = ShowSolverProgress
4131      txtMaxTime.Text = CStr(MaxTime)
4132      txtTol.Text = tol * 100
4133      txtMaxIter.Text = CStr(maxIter)
4134      txtPre = CStr(conPre)
4135      chkPerformLinearityCheck.value = performLinearityCheck

          Dim Solver As String
4136      If Not GetNameValueIfExists(ActiveWorkbook, sheetName & "OpenSolver_ChosenSolver", Solver) Then
4137          Solver = "CBC"
4138          Call SetNameOnSheet("OpenSolver_ChosenSolver", "=" & Solver)
4139      End If

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
