VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CallingNeos 
   Caption         =   "OpenSolver Optimisation Running"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   OleObjectBlob   =   "CallingNeos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CallingNeos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()

          'Unload the userform
3583      Me.Tag = "Cancelled"
3584      Me.Hide

End Sub


