VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MacOptions 
   Caption         =   "OpenSolver - Solve Options"
   ClientHeight    =   5805
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
          Unload frmOptions
41690     Unload Me
End Sub

Private Sub cmdOK_Click()
          frmOptions.OptionsOK Me
          Unload Me
End Sub

Private Sub UserForm_Activate()
          frmOptions.OptionsActivate Me
End Sub


