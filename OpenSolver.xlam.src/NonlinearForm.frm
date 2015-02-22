VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NonlinearForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11310
   OleObjectBlob   =   "NonlinearForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NonlinearForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdContinue_Click()
3585      Me.Hide
End Sub

Public Sub SetLinearityResult(resultString As String, IsQuickCheck As Boolean)
    NonlinearForm.CommonLinearityResult Me, resultString, IsQuickCheck
    Me.height = chkFullCheck.top + chkFullCheck.height + 30
    Caption = "OpenSolver: Linearity check "
End Sub

Public Sub CommonLinearityResult(f As UserForm, resultString As String, IsQuickCheck As Boolean)
    f.txtNonLinearInfo.Caption = resultString
    
    'formatting of the user form f.TextBox2.AutoSize = True
    f.txtNonLinearInfo.AutoSize = False
    f.txtNonLinearInfo.height = 20
    f.txtNonLinearInfo.AutoSize = True
    f.txtNonLinearInfo.AutoSize = False
    
    Dim MaxHeight As Integer
    #If Mac Then
       MaxHeight = 350
    #Else
       MaxHeight = 250
    #End If
    If f.txtNonLinearInfo.height > MaxHeight Then f.txtNonLinearInfo.height = MaxHeight
    
    f.chkFullCheck.Caption = "Run a full linearity check. (This will destroy the current solution) "
    f.chkHighlight.Caption = "Highlight the nonlinearities"
    
    f.chkHighlight.top = f.txtNonLinearInfo.height + f.txtNonLinearInfo.top + 5
    f.chkFullCheck.top = f.chkHighlight.top + f.chkHighlight.height
    
#If Mac Then
    f.cmdContinue.top = f.chkHighlight.top + 11
#Else
    f.cmdContinue.top = f.chkHighlight.top + 6 ' Enough space around check box anyway
#End If

    f.chkFullCheck.Visible = IsQuickCheck
End Sub

