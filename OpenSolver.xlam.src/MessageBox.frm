VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MessageBox 
   Caption         =   "Message Box:"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6120
   OleObjectBlob   =   "MessageBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LinkLabel_Click()
    Call OpenURL(MessageBox.LinkLabel.ControlTipText) 'this powers the hyperlink.
End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub


Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_Initialize()
    TextBox1.Text = ""
    LinkLabel.Caption = ""
End Sub

