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
Private Sub ContinueButton_Click()
3585      Me.Hide
End Sub

Public Sub SetLinearityResult(resultString As String, IsQuickCheck As Boolean)
    NonlinearForm.CommonLinearityResult Me, resultString, IsQuickCheck
    Me.height = FullCheck.top + FullCheck.height + 30
    Caption = "OpenSolver: Linearity check "
End Sub

Public Sub CommonLinearityResult(f As UserForm, resultString As String, IsQuickCheck As Boolean)
    f.TextBox2.Caption = resultString
    
    'formatting of the user form f.TextBox2.AutoSize = True
    f.TextBox2.AutoSize = False
    f.TextBox2.height = 20
    f.TextBox2.AutoSize = True
    f.TextBox2.AutoSize = False
    
    Dim MaxHeight As Integer
    #If Mac Then
       MaxHeight = 300
    #Else
       MaxHeight = 250
    #End If
    If f.TextBox2.height > MaxHeight Then f.TextBox2.height = MaxHeight
    
    f.FullCheck.Caption = "Run a full linearity check. (This will destroy the current solution) "
    f.HighlightBox.Caption = "Highlight the nonlinearities"
    
    f.HighlightBox.top = f.TextBox2.height + f.TextBox2.top + 5
    f.FullCheck.top = f.HighlightBox.top + f.HighlightBox.height
    
#If Mac Then
    f.ContinueButton.top = f.HighlightBox.top + 11
#Else
    f.ContinueButton.top = f.HighlightBox.top + 6 ' Enough space around check box anyway
#End If

    f.FullCheck.Visible = IsQuickCheck
End Sub

