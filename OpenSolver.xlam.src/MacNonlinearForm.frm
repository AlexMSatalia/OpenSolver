VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MacNonlinearForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6405
   ClientLeft      =   0
   ClientTop       =   -885
   ClientWidth     =   11310
   OleObjectBlob   =   "MacNonlinearForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MacNonlinearForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ContinueButton_Click()
9234      Me.Hide
End Sub

Public Sub SetLinearityResult(resultString As String, IsQuickCheck As Boolean)
    NonlinearForm.CommonLinearityResult Me, resultString, IsQuickCheck
    Me.height = FullCheck.top + FullCheck.height + 30
    Caption = "OpenSolver: Linearity check "
End Sub
