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
    Public Const FormMargin = 6
    Public Const FormTextHeight = 16
#End If

Public Const FormSpacing = 6
Public Const FormTextBoxColor = &H80000005
Public Const FormLinkColor = &HFF0000
Public Const FormHeadingHeight = 24
Public Const FormDivHeight = 2
Public Const FormDivBackColor = &HC1C1C1

#If Win32 Then
    ' Hacks for Excel 2016 on Windows Form margins being different to earlier versions
    ' We provide functions with the same names as the old constants so that the values can
    ' change based on the version of Excel.
    ' We can then use these functions as we would the original constants
    
    Public Function FormWindowMargin() As Long
1             If Val(Application.Version) >= 16 Then
2                 FormWindowMargin = 12
3             Else
4                 FormWindowMargin = 4
5             End If
    End Function
    
    Public Function FormTitleHeight() As Long
1             If Val(Application.Version) >= 16 Then
2                 FormTitleHeight = 29
3             Else
4                 FormTitleHeight = 20
5             End If
    End Function
#End If

Public Sub AutoFormat(ByRef Controls As Controls)
      ' Sets default appearances for Form controls
      ' Hopefully we don't run into late binding issues this way
          Dim Cont As Control, ContType As String
1         For Each Cont In Controls
2             ContType = TypeName(Cont)
3             If ContType = "TextBox" Or ContType = "CheckBox" Or ContType = "Label" Or _
                ContType = "CommandButton" Or ContType = "OptionButton" Or _
                ContType = "RefEdit" Or ContType = "ListBox" Or ContType = "ComboBox" Then
4                 With Cont
5                     .Font.Name = FormFontName
6                     .Font.Size = FormFontSize
7                     If ContType = "TextBox" Or ContType = "RefEdit" Or ContType = "ListBox" Or ContType = "ComboBox" Then
8                         .BackColor = FormTextBoxColor
9                     Else
10                        .BackColor = FormBackColor
11                    End If
                      
12                    If ContType = "CommandButton" Then
13                        .Height = FormButtonHeight
14                        .Cancel = False
15                    ElseIf ContType = "CheckBox" Or ContType = "OptionButton" Then
16                        .Height = FormCheckBoxHeight
17                    ElseIf ContType = "TextBox" Or ContType = "RefEdit" Or ContType = "ComboBox" Then
18                        .Height = FormTextBoxHeight
19                    Else
20                        .Height = FormTextHeight
21                    End If
                      
22                    If ContType = "RefEdit" Then
                          ' Prevent refedit focus bugs, but lose the ability to tab into refedits, so we don't do it
                          ' See: http://peltiertech.com/using-refedit-controls-in-excel-dialogs/#comment-276990
                          '.TabStop = False
23                    End If
24                End With
25            End If
26        Next
End Sub

Public Function RightOf(OldControl As Control, Optional Spacing As Boolean = True) As Long
1         RightOf = OldControl.Left + OldControl.Width + IIf(Spacing, FormSpacing, 0)
End Function

Public Function LeftOf(OldControl As Control, NewControlWidth As Long, Optional Spacing As Boolean = True) As Long
1         LeftOf = OldControl.Left - NewControlWidth - IIf(Spacing, FormSpacing, 0)
End Function

Public Function LeftOfForm(FormWidth As Long, NewControlWidth As Long)
1         LeftOfForm = FormWidth - FormMargin - NewControlWidth
End Function

Public Function Below(OldControl As Control, Optional Spacing As Boolean = True) As Long
1         Below = OldControl.Top + OldControl.Height + IIf(Spacing, FormSpacing, 0)
End Function

Public Function FormHeight(BottomControl As Control) As Long
1         FormHeight = Below(BottomControl, False) + FormMargin + FormTitleHeight
End Function

Public Sub AutoHeight(NewControl As Control, Width As Long, Optional ShrinkWidth As Boolean = False)
1         With NewControl
2             .Width = Width
3             .AutoSize = False
4             .AutoSize = True
5             .AutoSize = False
6             If Not ShrinkWidth Then .Width = Width
7         End With
End Sub

Public Function CenterFormTop(FormHeight As Long) As Single
          Dim BaseTop As Long, BaseHeight As Long
          
1         On Error GoTo NoWindow
2         BaseTop = Application.ActiveWindow.Top
3         BaseHeight = Application.ActiveWindow.Height
          
          ' Excel 2010 needs Application.top instead?
    #If Win32 Then
4             If Val(Application.Version) < 15 Then
5                 BaseTop = Application.Top - BaseTop
6             End If
    #End If
          
Calculate:
7         CenterFormTop = BaseTop + Max(BaseHeight / 2 - FormHeight / 2, 0)
8         Exit Function
          
NoWindow:
9         BaseTop = Application.Top
    #If Mac Then
10            BaseHeight = Application.UsableHeight
    #Else
11            BaseHeight = Application.Height
    #End If
12        Resume Calculate
End Function

Public Function CenterFormLeft(FormWidth As Long) As Single
          Dim BaseLeft As Long, BaseWidth As Long
          
1         On Error GoTo NoWindow
2         BaseLeft = Application.ActiveWindow.Left
3         BaseWidth = Application.ActiveWindow.Width
          
          ' Excel 2010 needs Application.left instead?
    #If Win32 Then
4             If Val(Application.Version) < 15 Then
5                 BaseLeft = Application.Left
6             End If
    #End If
          
Calculate:
7         CenterFormLeft = BaseLeft + Max(BaseWidth / 2 - FormWidth / 2, 0)
8         Exit Function
          
NoWindow:
9         BaseLeft = Application.Left
    #If Mac Then
10            BaseWidth = Application.UsableWidth
    #Else
11            BaseWidth = Application.Width
    #End If
12        Resume Calculate
End Function
