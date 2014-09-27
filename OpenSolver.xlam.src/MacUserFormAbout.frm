VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MacUserFormAbout 
   Caption         =   "About OpenSolver"
   ClientHeight    =   7995
   ClientLeft      =   0
   ClientTop       =   -1755
   ClientWidth     =   10005
   OleObjectBlob   =   "MacUserFormAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MacUserFormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub buttonOK_Click()
    Me.Hide
End Sub

Private Sub buttonUninstall_Click()
    UserFormAbout.ChangeAutoloadStatus False, Me
End Sub

Private Sub chkAutoLoad_Change()
    If Not UserFormAbout.EventsEnabled Then Exit Sub
    UserFormAbout.ChangeAutoloadStatus chkAutoLoad.value, Me
End Sub


Private Sub labelOpenSolverOrg_Click()
    Call OpenURL("http://www.opensolver.org")
End Sub

Private Sub UserForm_Activate()
    UserFormAbout.ActivateAboutForm Me
End Sub

