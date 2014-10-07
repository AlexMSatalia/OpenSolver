VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormInterrupt 
   Caption         =   "OpenSolver User Interrupt"
   ClientHeight    =   1743
   ClientLeft      =   42
   ClientTop       =   343
   ClientWidth     =   4711
   OleObjectBlob   =   "UserFormInterrupt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormInterrupt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonAbort_Click()
3540      Me.Tag = vbCancel
3541      Me.Hide
End Sub

Private Sub CommandButtonContinue_Click()
3542      Me.Tag = vbOK
3543      Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
3544      If CloseMode = vbFormControlMenu Then Cancel = True
End Sub
