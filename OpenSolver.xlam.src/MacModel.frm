VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MacModel 
   Caption         =   "OpenSolver - Model"
   ClientHeight    =   10003
   ClientLeft      =   0
   ClientTop       =   -9240
   ClientWidth     =   14205
   OleObjectBlob   =   "MacModel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MacModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkGetDuals_Click()
          frmModel.UpdateGetDuals Me
End Sub

Private Sub chkGetDuals2_Click()
          frmModel.UpdateGetDuals2 Me
End Sub

Private Sub chkNameRange_Click()
          frmModel.UpdateNameRange Me
End Sub

Private Sub cmdCancelCon_Click()
          frmModel.ModelCancel Me
End Sub

Private Sub cmdChange_Click()
          frmModel.ModelSolverClick Me
End Sub

Private Sub cmdOptions_Click()
          frmModel.ModelOptionsClick Me
End Sub

Private Sub cmdReset_Click()
          frmModel.ModelReset Me
End Sub

Private Sub optMax_Click()
          frmModel.ModelMaxClick Me
End Sub

Private Sub optMin_Click()
          frmModel.ModelMinClick Me
End Sub

Private Sub optNew_Click()
          frmModel.ModelNewClick Me
End Sub

Private Sub optTarget_Click()
          frmModel.ModelTargetClick Me
End Sub

Private Sub optUpdate_Click()
          frmModel.ModelUpdateClick Me
End Sub

Private Sub refConLHS_Change()
          frmModel.ModelChangeLHS Me
End Sub

Private Sub refConRHS_Change()
          frmModel.ModelChangeRHS Me
End Sub

Private Sub UserForm_Activate()
          frmModel.ModelActivate Me
End Sub

Private Sub cmdCancel_Click()
          frmModel.ModelCancelClick Me
          ' Remove any focus taken by a RefEdit
          DoEvents
          Me.Hide
End Sub

Private Sub cmdRunAutoModel_Click()
          'Me.Hide
          DoEvents
          frmModel.ModelRunAutoModel Me
          DoEvents
          'Me.Show
End Sub

Private Sub cmdBuild_Click()
         frmModel.ModelBuild Me
         Me.Hide
End Sub

Private Sub cboConRel_Change()
          frmModel.ModelChangeConRel Me
End Sub

Private Sub cmdAddCon_Click()
          frmModel.ModelAddConstraint Me
End Sub

Private Sub cmdDelSelCon_Click()
         frmModel.ModelDeleteConstraint Me
End Sub

Private Sub lstConstraints_Change()
    frmModel.ModelLstConstraintsChange Me
End Sub

Private Sub lstConstraints_Exit(ByVal Cancel As MSForms.ReturnBoolean)
          ' When the focus leaves this list, we want to remove any highlighting shown by selected cells
47540     ActiveCell.Select
End Sub


