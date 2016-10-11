VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FMsgBoxEx 
   Caption         =   "OpenSolver - Message"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9720
   OleObjectBlob   =   "FMsgBoxEx.frx":0000
End
Attribute VB_Name = "FMsgBoxEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Mac Then
    Const FormWidthMessageBox = 400
#Else
    Const FormWidthMessageBox = 250
#End If

Private Sub cmdButton1_Click()
1         Me.Tag = cmdButton1.Tag
2         Me.Hide
End Sub

Private Sub cmdButton2_Click()
1         Me.Tag = cmdButton2.Tag
2         Me.Hide
End Sub

Private Sub cmdButton3_Click()
1         Me.Tag = cmdButton3.Tag
2         Me.Hide
End Sub

' Make the [x] hide the form rather than unload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
          ' If CloseMode = vbFormControlMenu then we know the user
          ' clicked the [x] close button or Alt+F4 to close the form.
1         If CloseMode = vbFormControlMenu Then
2             Me.Hide
3             Cancel = True
4         End If
End Sub

Private Sub cmdMoreDetails_Click()
1         Me.Hide  ' Close the dialog so that we don't block interaction with the log file
2         OpenFile GetErrorLogFilePath, "No error log was found."
End Sub

Private Sub cmdReportIssue_Click()
          Dim ErrorLogFilePath As String
1         ErrorLogFilePath = GetErrorLogFilePath()
2         If Not FileOrDirExists(ErrorLogFilePath) Then
3             MsgBoxEx ("No error log was found")
4             Exit Sub
5         End If
          
          ' Read in contents of error log
          Dim ErrorLogContents As String
6         Open ErrorLogFilePath For Input As #1
7             ErrorLogContents = Input$(LOF(1), 1)
8         Close #1
          
          ' Form mailto:// link
          Dim EmailBody As String, EmailSubject As String
9         EmailSubject = "OpenSolver Error Report"
10        EmailBody = "Please insert any other information you think might be relevant here." & vbNewLine & vbNewLine & _
                      "You may want to paste in the content of the solver log file, which you can open " & _
                      "by going to the OpenSolver menu." & vbNewLine & vbNewLine & _
                      "---------- Error log follows ----------" & vbNewLine & vbNewLine & _
                      ErrorLogContents
          
          Dim MailToLink As String
11        MailToLink = "mailto:andrew@opensolver.org?cc=jack@opensolver.org" & _
                       "&subject=" & URLEncode(EmailSubject) & _
                       "&body=" & URLEncode(EmailBody)
          
12        On Error GoTo FailedOpen
13        OpenURL MailToLink
14        On Error GoTo 0
          
15        Exit Sub
          
FailedOpen:
16        MsgBoxEx "Couldn't create an email to report the issue. Please try again, or send us the error.log " & _
                   "file manually."
End Sub

Private Sub lblLink_Click()
1         Call OpenURL(lblLink.ControlTipText)
End Sub

Private Sub UserForm_Activate()
1         CenterForm
End Sub

Public Sub AutoLayout()
1         AutoFormat Me.Controls
          
2         cmdButton1.Visible = Len(cmdButton1.Caption) > 0
3         cmdButton2.Visible = Len(cmdButton2.Caption) > 0
4         cmdButton3.Visible = Len(cmdButton3.Caption) > 0
          
          ' Calculate the needed width based on how many buttons are visible
          ' We sum the values of .Visible, where True is -1 and False 0
          Dim NumButtonWidths As Double
5         NumButtonWidths = -(cmdMoreDetails.Visible + cmdReportIssue.Visible)
6         If NumButtonWidths > 0 Then NumButtonWidths = NumButtonWidths + 0.5  ' Account for spacing between main buttons and extra buttons
7         NumButtonWidths = NumButtonWidths - (cmdButton1.Visible + cmdButton2.Visible + cmdButton3.Visible)
          
          ' Calculate the width, and set to minimum constant width if button width is less than this.
8         Me.Width = Max(FormWidthMessageBox, NumButtonWidths * FormButtonWidth + (NumButtonWidths - 1) * FormSpacing + 2 * FormMargin)
          
9         With txtMessage
10            .BackColor = FormBackColor
11            .Left = FormMargin
12            .Top = FormMargin
13            AutoHeight txtMessage, Me.Width - 2 * FormMargin
14        End With
          
15        With lblLink
16            If Len(.Caption) = 0 Then
17                .Visible = False
18            Else
19                AutoHeight lblLink, Me.Width, True
20                .Left = (Me.Width - .Width) / 2
21                .Font.Underline = True
22            End If
23            .Top = Below(txtMessage)
24        End With
          
25        With cmdButton3
26            .Top = Below(IIf(lblLink.Visible, lblLink, txtMessage))
27            .Width = FormButtonWidth
28            .Left = LeftOfForm(Me.Width, .Width)
29            .Cancel = False
30        End With
          
31        With cmdButton2
32            .Top = cmdButton3.Top
33            .Width = FormButtonWidth
34            .Left = cmdButton3.Left + IIf(cmdButton3.Visible, -FormSpacing - .Width, 0)
35            .Cancel = False
36        End With
          
37        With cmdButton1
38            .Top = cmdButton2.Top
39            .Width = FormButtonWidth
40            .Left = cmdButton2.Left + IIf(cmdButton2.Visible, -FormSpacing - .Width, 0)
41            .Cancel = False
42        End With
          
          ' Set esc target
43        If cmdButton1.Visible Then
44            cmdButton1.Cancel = True
45        ElseIf cmdButton2.Visible Then
46            cmdButton2.Cancel = True
47        Else
48            cmdButton3.Cancel = True
49        End If
          
50        With cmdMoreDetails
51            .Left = txtMessage.Left
52            .Top = cmdButton3.Top
53            .Width = FormButtonWidth
54        End With
          
55        With cmdReportIssue
56            .Top = cmdButton3.Top
57            .Width = FormButtonWidth
58            .Left = cmdMoreDetails.Left + IIf(cmdMoreDetails.Visible, FormSpacing + .Width, 0)
59        End With
          
60        Me.Height = FormHeight(cmdButton1)
61        Me.Width = Me.Width + FormWindowMargin
          
62        Me.BackColor = FormBackColor
End Sub

Private Sub CenterForm()
1         Me.Top = CenterFormTop(Me.Height)
2         Me.Left = CenterFormLeft(Me.Width)
End Sub
