VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MacOptions 
   Caption         =   "OpenSolver - Solve Options"
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   -2640
   ClientWidth     =   6240
   OleObjectBlob   =   "MacOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MacOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
9216      Unload frmOptions
9217      Unload Me
End Sub

Private Sub cmdOK_Click()
9218            frmOptions.OptionsOK Me
9219            Unload Me
End Sub

Private Sub UserForm_Activate()
9220            frmOptions.OptionsActivate Me
End Sub


