Attribute VB_Name = "FormConstants"
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
