Attribute VB_Name = "FormConstants"
Option Explicit

#If Mac Then
    Public Const FormBackColor = &HE3E3E3
    Public Const FormFontName = "Lucida Grande"
    Public Const FormFontSize = 11
    Public Const FormButtonHeight = 22
    Public Const FormButtonWidth = 100
    Public Const FormTitleHeight = 20
    Public Const FormWindowMargin = 0
    Public Const FormMargin = 12
#Else
    Public Const FormBackColor = &H8000000F
    Public Const FormFontName = "Tahoma"
    Public Const FormFontSize = 8
    Public Const FormButtonHeight = 18
    Public Const FormButtonWidth = 66
    Public Const FormTitleHeight = 20
    Public Const FormWindowMargin = 4
    Public Const FormMargin = 6
#End If

Public Const FormSpacing = 6
Public Const FormTextBoxColor = &H80000005

Public Sub AutoFormat(ByRef Controls As Controls)
' Sets default appearances for Form controls
' Hopefully we don't run into late binding issues this way
    Dim Cont As Control, ContType As String
    For Each Cont In Controls
        ContType = TypeName(Cont)
        If ContType = "TextBox" Or ContType = "CheckBox" Or ContType = "Label" Or ContType = "CommandButton" Then
            With Cont
                .Font.Name = FormFontName
                .Font.Size = FormFontSize
                If ContType = "TextBox" Then
                    .BackColor = FormTextBoxColor
                Else
                    .BackColor = FormBackColor
                End If
                .height = FormButtonHeight
            End With
        End If
    Next
End Sub
