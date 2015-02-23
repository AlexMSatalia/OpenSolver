Attribute VB_Name = "FormConstants"
Option Explicit

#If Mac Then
    Public Const FormBackColor = &HE3E3E3
    Public Const FormFontName = "Lucida Grande"
    Public Const FormFontSize = 11
    Public Const FormHeadingSize = 18
    Public Const FormButtonHeight = 22
    Public Const FormButtonWidth = 100
    Public Const FormCheckBoxHeight = 22
    Public Const FormTitleHeight = 20
    Public Const FormWindowMargin = 0
    Public Const FormMargin = 12
    Public Const FormTextHeight = 18
#Else
    Public Const FormBackColor = &H8000000F
    Public Const FormFontName = "Tahoma"
    Public Const FormFontSize = 8
    Public Const FormHeadingSize = 16
    Public Const FormButtonHeight = 20
    Public Const FormButtonWidth = 66
    Public Const FormCheckBoxHeight = 16
    Public Const FormTitleHeight = 20
    Public Const FormWindowMargin = 4
    Public Const FormMargin = 6
    Public Const FormTextHeight = 16
#End If

Public Const FormSpacing = 6
Public Const FormTextBoxColor = &H80000005
Public Const FormLinkColor = &HFF0000
Public Const FormHeadingHeight = 24
Public Const FormDivHeight = 2
Public Const FormDivBackColor = &HC1C1C1

Public Sub AutoFormat(ByRef Controls As Controls)
' Sets default appearances for Form controls
' Hopefully we don't run into late binding issues this way
    Dim Cont As Control, ContType As String
    For Each Cont In Controls
        ContType = TypeName(Cont)
        If ContType = "TextBox" Or ContType = "CheckBox" Or ContType = "Label" Or ContType = "CommandButton" Or ContType = "OptionButton" Then
            With Cont
                .Font.Name = FormFontName
                .Font.Size = FormFontSize
                If ContType = "TextBox" Then
                    .BackColor = FormTextBoxColor
                Else
                    .BackColor = FormBackColor
                End If
                
                If ContType = "CommandButton" Then
                    .height = FormButtonHeight
                ElseIf ContType = "CheckBox" Or ContType = "OptionButton" Then
                    .height = FormCheckBoxHeight
                Else
                    .height = FormTextHeight
                End If
            End With
        End If
    Next
End Sub
