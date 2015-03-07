VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSolverChange 
   Caption         =   "Choose Solver"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5010
   OleObjectBlob   =   "frmSolverChange.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSolverChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Mac Then
    Const FormWidthSolverChange = 365
#Else
    Const FormWidthSolverChange = 255
#End If

Public ChosenSolver As String
Dim Solvers As Collection

Private Sub cboSolver_Change()
4724            ChosenSolver = ReverseSolverTitle(cboSolver.Text)
4725            lblDesc.Caption = SolverDesc(ChosenSolver)

                lblHyperlink.Caption = SolverLink(ChosenSolver)
                
                Dim errorString As String
4727            cmdOk.Enabled = SolverAvailable(ChosenSolver, errorString:=errorString)
4732            lblError.Caption = errorString ' empty if no errors found

                AutoLayout
End Sub

Private Sub lblHyperlink_Click()
4733      OpenURL lblHyperlink.Caption
End Sub

Private Sub UserForm_Activate()
4735      Set Solvers = New Collection
4736      Solvers.Add "CBC"
4737      Solvers.Add "Gurobi"
          'Solvers.Add "Cplex"
4738      Solvers.Add "NeosCBC"
4739      Solvers.Add "Bonmin"
4740      Solvers.Add "Couenne"
4741      Solvers.Add "NOMAD"
4742      Solvers.Add "NeosBon"
4743      Solvers.Add "NeosCou"
          'Solvers.Add "PuLP"
          
4744      cboSolver.Clear
4745      cboSolver.MatchRequired = True
4746      cboSolver.Style = fmStyleDropDownList
          
          Dim Solver As Variant
4747      For Each Solver In Solvers
4748          cboSolver.AddItem SolverTitle(CStr(Solver))
4749      Next Solver

          Dim value As String
4750      If GetNameValueIfExists(ActiveWorkbook, EscapeSheetName(ActiveWorkbook.ActiveSheet) & "OpenSolver_ChosenSolver", value) Then
4751          On Error GoTo setDefault
4752          cboSolver.Text = SolverTitle(value)
4753      Else
setDefault:
4754          cboSolver.Text = SolverTitle("CBC")
4755      End If
End Sub

Private Sub cmdOk_Click()
         'Add the chosen solver as a hidden name in the workbook
4758      Call SetNameOnSheet("OpenSolver_ChosenSolver", "=" & ChosenSolver)
4761      frmModel.FormatCurrentSolver ChosenSolver
4762      frmModel.Disabler True
4763      Unload Me
End Sub

Private Sub cmdCancel_Click()
4764      Unload Me
End Sub

Private Sub UserForm_Initialize()
    AutoLayout
End Sub

Private Sub AutoLayout()
    AutoFormat Me.Controls
    
    Me.width = FormWidthSolverChange
    
    With lblChoose
        .left = FormMargin
        .top = FormMargin
        .width = Me.width - FormMargin * 2
        .Caption = "Choose a solver from the list below:"
    End With
    
    With cboSolver
        .left = lblChoose.left
        .top = Below(lblChoose, False)
        .width = lblChoose.width
    End With
    
    With lblDesc
        .left = lblChoose.left
        .top = Below(cboSolver)
        AutoHeight lblDesc, lblChoose.width
    End With
    
    With lblHyperlink
        .left = lblChoose.left
        .top = Below(lblDesc)
        AutoHeight lblHyperlink, lblChoose.width, True
    End With
    
    With lblError
        .Visible = Len(.Caption) <> 0
        .left = lblChoose.left
        .top = Below(lblHyperlink)
        AutoHeight lblError, lblChoose.width
    End With
    
    With cmdCancel
        .Caption = "Cancel"
        .width = FormButtonWidth
        .top = Below(IIf(lblError.Visible, lblError, lblHyperlink))
        .left = LeftOfForm(Me.width, .width)
    End With
    
    With cmdOk
        .Caption = "OK"
        .width = FormButtonWidth
        .top = cmdCancel.top
        .left = LeftOf(cmdCancel, .width)
    End With
    
    
    Me.height = FormHeight(cmdCancel)
    Me.width = Me.width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - Choose Solver"
End Sub
