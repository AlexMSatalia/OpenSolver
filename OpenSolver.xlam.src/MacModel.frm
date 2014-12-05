VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MacModel 
   Caption         =   "OpenSolver - Model"
   ClientHeight    =   10005
   ClientLeft      =   0
   ClientTop       =   -9240
   ClientWidth     =   14203
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
9188            frmModel.UpdateGetDuals Me
End Sub

Private Sub chkGetDuals2_Click()
9189            frmModel.UpdateGetDuals2 Me
End Sub

Private Sub chkNameRange_Click()
9190            frmModel.UpdateNameRange Me
End Sub

Private Sub cmdCancelCon_Click()
9191            frmModel.ModelCancel Me
End Sub

Private Sub cmdChange_Click()
9192            frmModel.ModelSolverClick Me
End Sub

Private Sub cmdOptions_Click()
9193            frmModel.ModelOptionsClick Me
End Sub

Private Sub cmdReset_Click()
9194            frmModel.ModelReset Me
End Sub

Private Sub optMax_Click()
9195            frmModel.ModelMaxClick Me
End Sub

Private Sub optMin_Click()
9196            frmModel.ModelMinClick Me
End Sub

Private Sub optNew_Click()
9197            frmModel.ModelNewClick Me
End Sub

Private Sub optTarget_Click()
9198            frmModel.ModelTargetClick Me
End Sub

Private Sub optUpdate_Click()
9199            frmModel.ModelUpdateClick Me
End Sub

Private Sub refConLHS_Change()
9200            frmModel.ModelChangeLHS Me
End Sub

Private Sub refConRHS_Change()
9201            frmModel.ModelChangeRHS Me
End Sub

Private Sub UserForm_Activate()
9202            frmModel.ModelActivate Me
End Sub

Private Sub cmdCancel_Click()
9203            frmModel.ModelCancelClick Me
                ' Remove any focus taken by a RefEdit
9204            DoEvents
9205            Me.Hide
End Sub

Private Sub cmdRunAutoModel_Click()
                'Me.Hide
9206            DoEvents
9207            frmModel.ModelRunAutoModel Me
9208            DoEvents
                'Me.Show
End Sub

Private Sub cmdBuild_Click()
9209           frmModel.ModelBuild Me
9210           Me.Hide
End Sub

Private Sub cboConRel_Change()
9211            frmModel.ModelChangeConRel Me
End Sub

Private Sub cmdAddCon_Click()
9212            frmModel.ModelAddConstraint Me
End Sub

Private Sub cmdDelSelCon_Click()
9213           frmModel.ModelDeleteConstraint Me
End Sub

Private Sub lstConstraints_Change()
9214      frmModel.ModelLstConstraintsChange Me
End Sub

Private Sub lstConstraints_Exit(ByVal Cancel As MSForms.ReturnBoolean)
          ' When the focus leaves this list, we want to remove any highlighting shown by selected cells
9215      ActiveCell.Select
End Sub


