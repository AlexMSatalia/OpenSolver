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
    Public Const FormTextBoxHeight = 22
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
    Public Const FormTextBoxHeight = 18
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
        If ContType = "TextBox" Or ContType = "CheckBox" Or ContType = "Label" Or _
          ContType = "CommandButton" Or ContType = "OptionButton" Or _
          ContType = "RefEdit" Or ContType = "ListBox" Or ContType = "ComboBox" Then
            With Cont
                .Font.Name = FormFontName
                .Font.Size = FormFontSize
                If ContType = "TextBox" Or ContType = "RefEdit" Or ContType = "ListBox" Or ContType = "ComboBox" Then
                    .BackColor = FormTextBoxColor
                Else
                    .BackColor = FormBackColor
                End If
                
                If ContType = "CommandButton" Then
                    .Height = FormButtonHeight
                ElseIf ContType = "CheckBox" Or ContType = "OptionButton" Then
                    .Height = FormCheckBoxHeight
                ElseIf ContType = "TextBox" Or ContType = "RefEdit" Or ContType = "ComboBox" Then
                    .Height = FormTextBoxHeight
                Else
                    .Height = FormTextHeight
                End If
            End With
        End If
    Next
End Sub

Public Function RightOf(OldControl As Control, Optional Spacing As Boolean = True) As Long
    RightOf = OldControl.Left + OldControl.Width + IIf(Spacing, FormSpacing, 0)
End Function

Public Function LeftOf(OldControl As Control, NewControlWidth As Long, Optional Spacing As Boolean = True) As Long
    LeftOf = OldControl.Left - NewControlWidth - IIf(Spacing, FormSpacing, 0)
End Function

Public Function LeftOfForm(FormWidth As Long, NewControlWidth As Long)
    LeftOfForm = FormWidth - FormMargin - NewControlWidth
End Function

Public Function Below(OldControl As Control, Optional Spacing As Boolean = True) As Long
    Below = OldControl.Top + OldControl.Height + IIf(Spacing, FormSpacing, 0)
End Function

Public Function FormHeight(BottomControl As Control) As Long
    FormHeight = Below(BottomControl, False) + FormMargin + FormTitleHeight
End Function

Public Sub AutoHeight(NewControl As Control, Width As Long, Optional ShrinkWidth As Boolean = False)
    With NewControl
        .Width = Width
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
        If Not ShrinkWidth Then .Width = Width
    End With
End Sub

Public Function CenterFormTop(FormHeight As Long)
    Dim BaseTop As Long, BaseHeight As Long
    
    On Error GoTo NoWindow
    BaseTop = Application.ActiveWindow.Top
    BaseHeight = Application.ActiveWindow.Height
    
    ' Excel 2010 needs Application.top instead?
    #If Win32 Then
        If Val(Application.Version) < 15 Then
            BaseTop = Application.Top - BaseTop
        End If
    #End If
    
Calculate:
    CenterFormTop = BaseTop + Max(BaseHeight / 2 - FormHeight / 2, 0)
    Exit Function
    
NoWindow:
    BaseTop = Application.Top
    #If Mac Then
        BaseHeight = Application.UsableHeight
    #Else
        BaseHeight = Application.Height
    #End If
    Resume Calculate
End Function

Public Function CenterFormLeft(FormWidth As Long)
    Dim BaseLeft As Long, BaseWidth As Long
    
    On Error GoTo NoWindow
    BaseLeft = Application.ActiveWindow.Left
    BaseWidth = Application.ActiveWindow.Width
    
    ' Excel 2010 needs Application.left instead?
    #If Win32 Then
        If Val(Application.Version) < 15 Then
            BaseLeft = Application.Left
        End If
    #End If
    
Calculate:
    CenterFormLeft = BaseLeft + Max(BaseWidth / 2 - FormWidth / 2, 0)
    Exit Function
    
NoWindow:
    BaseLeft = Application.Left
    #If Mac Then
        BaseWidth = Application.UsableWidth
    #Else
        BaseWidth = Application.Width
    #End If
    Resume Calculate
End Function

