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
    Const FormWidthSolverChange = 330
#Else
    Const FormWidthSolverChange = 255
#End If

Public ChosenSolver As String
Dim Solvers As Collection

Private Sub cboSolver_Change()
4724            ChosenSolver = ReverseSolverTitle(cboSolver.Text)
4725            lblDesc.Caption = SolverDesc(ChosenSolver)

                With lblHyperlink
                    .Caption = SolverLink(ChosenSolver)
                    .width = Me.width
                    ' Reduce width to minimise size of link target
                    .AutoSize = False
                    .AutoSize = True
                    .AutoSize = False
                End With
                
                Dim errorString As String
4727            If SolverAvailable(ChosenSolver, errorString:=errorString) Then
4728                cmdOk.Enabled = True
4729            Else
4730                cmdOk.Enabled = False
4731            End If
4732            lblError.Caption = errorString ' empty if no errors found
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
4750      If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!OpenSolver_ChosenSolver", value) Then
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
4761      frmModel.lblSolver.Caption = "Current Solver Engine: " & UCase(left(ChosenSolver, 1)) & Mid(ChosenSolver, 2)
4762      frmModel.Disabler True, frmModel
4763      Unload Me
End Sub

Private Sub cmdCancel_Click()
4764      Unload Me
End Sub

Private Sub UserForm_Initialize()
    Me.AutoLayout
End Sub

Sub AutoLayout()
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
        .top = lblChoose.top + lblChoose.height
        .width = lblChoose.width
        .height = FormTextHeight
    End With
    
    With lblDesc
        .left = lblChoose.left
        .top = cboSolver.top + cboSolver.height + FormSpacing
        .width = lblChoose.width
        #If Mac Then
            .height = 150
        #Else
            .height = 100
        #End If
    End With
    
    With lblHyperlink
        .left = lblChoose.left
        .top = lblDesc.top + lblDesc.height + FormSpacing
        .width = lblChoose.width
        ' Reduce width to minimise size of link target
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
    End With
    
    With lblError
        .left = lblChoose.left
        .top = lblHyperlink.top + lblHyperlink.height + FormSpacing
        .width = lblChoose.width
        .height = 1.5 * FormTextHeight
    End With
    
    With cmdCancel
        .Caption = "Cancel"
        .width = FormButtonWidth
        .top = lblError.top + lblError.height + FormSpacing
        .left = Me.width - FormMargin - .width
    End With
    
    With cmdOk
        .Caption = "OK"
        .width = FormButtonWidth
        .top = cmdCancel.top
        .left = cmdCancel.left - FormSpacing - .width
    End With
    
    Me.height = cmdCancel.top + cmdCancel.height + FormMargin + FormTitleHeight
    Me.width = Me.width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - Choose Solver"
End Sub
