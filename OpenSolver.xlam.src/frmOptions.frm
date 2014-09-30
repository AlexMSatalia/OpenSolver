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

Private Sub cmdCancel_Click()
4089      Unload Me
End Sub

Private Sub cmdOK_Click()
4090      frmOptions.OptionsOK Me
4091      Unload Me
End Sub

Public Sub OptionsOK(f As UserForm)
4092      If f.chkNonNeg.value = True Then
4093          SetSolverNameOnSheet "neg", "=1"
4094      Else
4095          SetSolverNameOnSheet "neg", "=2"     ' 2 means false
4096      End If
          
4097      If f.chkShowSolverProgress.value = True Then
4098          SetSolverNameOnSheet "sho", "=1"
4099      Else
4100          SetSolverNameOnSheet "sho", "=2"     ' 2 means false
4101      End If
          
4102      SetSolverNameOnSheet "tim", "=" & Trim(str(CDbl(f.txtMaxTime.Text)))  ' Trim the leading space which str puts in for +'ve values
4103      SetSolverNameOnSheet "itr", "=" & Trim(str(CDbl(f.txtMaxIter.Text)))  ' Trim the leading space which str puts in for +'ve values
4104      SetSolverNameOnSheet "pre", "=" & Trim(str(CDbl(f.txtPre.Text)))  ' Trim the leading space which str puts in for +'ve values
4105      f.txtTol.Text = Replace(f.txtTol.Text, "%", "")
4106      SetSolverNameOnSheet "tol", "=" & Trim(str(CDbl(f.txtTol.Text) / 100))    ' Str() uses . for decimal
                                                                      ' CDbl respects the locale. We trim the leading space which str puts in for +'ve values
                                                                      
4107      If f.chkPerformLinearityCheck.value = True Then
              ' Default is "do check", so we just delete the option
4108          DeleteNameOnSheet "OpenSolver_LinearityCheck"
4109      Else
              ' Set the name, with a value of 2=off
4110          SetNameOnSheet "OpenSolver_LinearityCheck", "=2"
4111      End If
                                                                      
4112      Unload Me
End Sub

Private Sub UserForm_Activate()
4113            frmOptions.OptionsActivate Me
End Sub

Public Sub OptionsActivate(f As UserForm)
4114      SetAnyMissingDefaultExcel2007SolverOptions

          Dim nonNeg As Boolean, s As String
          ' nonNeg = True   ' a sensible default
4115      If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_neg", s) Then
4116          nonNeg = s = "1"
4117      End If
          
          Dim ShowSolverProgress As Boolean
          ' ShowSolverProgress = False
4118      If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_sho", s) Then
4119          ShowSolverProgress = s = "1"
4120      End If
          
          Dim maxTime As Double
          ' maxTime = 9999 ' A default value if none is yet defined
4121      GetNamedNumericValueIfExists ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_tim", maxTime
          
          Dim maxIter As Double
4122      GetNamedNumericValueIfExists ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_itr", maxIter

          Dim conPre As Double
4123      GetNamedNumericValueIfExists ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_pre", conPre

          Dim tol As Double
          ' tol = 0.05  ' A default value if none is yet defined
4124      GetNamedNumericValueIfExists ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_tol", tol
          
          ' We perform a linearity check by default unless the defined name exists with value 2=off
          Dim performLinearityCheck As Boolean
4125      performLinearityCheck = True
4126      If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!OpenSolver_LinearityCheck", s) Then
4127          performLinearityCheck = s = "1"
4128      End If

4129      f.chkNonNeg.value = nonNeg
4130      f.chkShowSolverProgress.value = ShowSolverProgress
4131      f.txtMaxTime.Text = CStr(maxTime)
4132      f.txtTol.Text = tol * 100
4133      f.txtMaxIter.Text = CStr(maxIter)
4134      f.txtPre = CStr(conPre)
4135      f.chkPerformLinearityCheck.value = performLinearityCheck

          Dim Solver As String
4136      If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!OpenSolver_ChosenSolver", Solver) Then
4137          Solver = "CBC"
4138          Call SetNameOnSheet("OpenSolver_ChosenSolver", "=" & Solver)
4139      End If

4140      If SolverType(Solver) = OpenSolver_SolverType.NonLinear Then
              ' Disable linearity options
4141          f.chkPerformLinearityCheck.Enabled = False
4142          f.txtTol.Enabled = False
4143      End If
          
4144      If Solver <> "NOMAD" Then
              ' Disable NOMAD only options
4145          f.txtMaxIter.Enabled = False
4146          f.txtPre.Enabled = False
4147      End If
End Sub

