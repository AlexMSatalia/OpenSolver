VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MacAutoModel 
   Caption         =   "OpenSolver - AutoModel"
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   -1755
   ClientWidth     =   10005
   OleObjectBlob   =   "MacAutoModel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MacAutoModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
          DoEvents
          Unload frmAutoModel
          Unload Me
          DoEvents
End Sub

Private Sub UserForm_Activate()
    frmAutoModel.AutoModelActivate Me
End Sub


Private Sub cmdFinish_Click()
         DoEvents
         frmAutoModel.AutoModelFinish Me
         DoEvents
End Sub

