VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MacSolverChange 
   Caption         =   "Choose Solver"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   -6167
   ClientWidth     =   7112
   OleObjectBlob   =   "MacSolverChange.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MacSolverChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboSolver_Change()
9181      frmSolverChange.ChangeSolver Me
End Sub

Private Sub lblHyperlink_Click()
9182      OpenURL lblHyperLink.Caption
End Sub

Private Sub UserForm_Activate()
9183      frmSolverChange.ActivateSolverChange Me
End Sub

Private Sub cmdOK_Click()
9184      frmSolverChange.SolverChangeConfirm Me
9185      Unload Me
End Sub

Private Sub cmdCancel_Click()
9186      Unload frmSolverChange
9187      Unload Me
End Sub

