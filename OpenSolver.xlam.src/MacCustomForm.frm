VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MacCustomForm 
   Caption         =   "OpenSolver Warning"
   ClientHeight    =   3563
   ClientLeft      =   0
   ClientTop       =   -882
   ClientWidth     =   4998
   OleObjectBlob   =   "MacCustomForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MacCustomForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_Click()

          'Unload the userform
36690     Me.Tag = "Cancelled"
36700     Me.Hide

End Sub

Private Sub CheckBox1_Click()
36710     Me.Tag = "Change"
End Sub
Private Sub Continue_Click()
36720     Me.Hide
End Sub
