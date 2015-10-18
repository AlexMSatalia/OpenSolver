VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FMsgBoxEx 
   Caption         =   "OpenSolver - Message"
   ClientHeight    =   4270
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
    Me.Tag = cmdButton1.Tag
    Me.Hide
End Sub

Private Sub cmdButton2_Click()
    Me.Tag = cmdButton2.Tag
    Me.Hide
End Sub

Private Sub cmdButton3_Click()
    Me.Tag = cmdButton3.Tag
    Me.Hide
End Sub

' Make the [x] hide the form rather than unload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' If CloseMode = vbFormControlMenu then we know the user
    ' clicked the [x] close button or Alt+F4 to close the form.
    If CloseMode = vbFormControlMenu Then
        Me.Hide
        Cancel = True
    End If
End Sub

Private Sub cmdMoreDetails_Click()
    OpenFile GetErrorLogFilePath, "No error log was found."
End Sub

Private Sub cmdReportIssue_Click()
    Dim ErrorLogFilePath As String
    ErrorLogFilePath = GetErrorLogFilePath()
    If Not FileOrDirExists(ErrorLogFilePath) Then
        MsgBoxEx ("No error log was found")
        Exit Sub
    End If
    
    ' Read in contents of error log
    Dim ErrorLogContents As String
    Open ErrorLogFilePath For Input As #1
        ErrorLogContents = Input$(LOF(1), 1)
    Close #1
    
    ' Form mailto:// link
    Dim EmailBody As String, EmailSubject As String
    EmailSubject = "OpenSolver Error Report"
    EmailBody = "Please insert any other information you think might be relevant here." & vbNewLine & vbNewLine & _
                "You may want to paste in the content of the solver log file, which you can open " & _
                "by going to the OpenSolver menu." & vbNewLine & vbNewLine & _
                "---------- Error log follows ----------" & vbNewLine & vbNewLine & _
                ErrorLogContents
    
    Dim MailToLink As String
    MailToLink = "mailto:andrew@opensolver.org?cc=jack@opensolver.org" & _
                 "&subject=" & URLEncode(EmailSubject) & _
                 "&body=" & URLEncode(EmailBody)
    
    OpenURL MailToLink
End Sub

Private Sub lblLink_Click()
    Call OpenURL(lblLink.ControlTipText)
End Sub

Private Sub UserForm_Activate()
    CenterForm
End Sub

Public Sub AutoLayout()
    AutoFormat Me.Controls
    
    cmdButton1.Visible = cmdButton1.Caption <> ""
    cmdButton2.Visible = cmdButton2.Caption <> ""
    cmdButton3.Visible = cmdButton3.Caption <> ""
    
    ' Calculate the needed width based on how many buttons are visible
    ' We sum the values of .Visible, where True is -1 and False 0
    Dim NumButtonWidths As Double
    NumButtonWidths = -(cmdMoreDetails.Visible + cmdReportIssue.Visible)
    If NumButtonWidths > 0 Then NumButtonWidths = NumButtonWidths + 0.5  ' Account for spacing between main buttons and extra buttons
    NumButtonWidths = NumButtonWidths - (cmdButton1.Visible + cmdButton2.Visible + cmdButton3.Visible)
    
    ' Calculate the width, and set to minimum constant width if button width is less than this.
    Me.Width = Max(FormWidthMessageBox, NumButtonWidths * FormButtonWidth + (NumButtonWidths - 1) * FormSpacing + 2 * FormMargin)
    
    With txtMessage
        .BackColor = FormBackColor
        .Left = FormMargin
        .Top = FormMargin
        AutoHeight txtMessage, Me.Width - 2 * FormMargin
    End With
    
    With lblLink
        If .Caption = "" Then
            .Visible = False
        Else
            AutoHeight lblLink, Me.Width, True
            .Left = (Me.Width - .Width) / 2
            .Font.Underline = True
        End If
        .Top = Below(txtMessage)
    End With
    
    With cmdButton3
        .Top = Below(IIf(lblLink.Visible, lblLink, txtMessage))
        .Width = FormButtonWidth
        .Left = LeftOfForm(Me.Width, .Width)
        .Cancel = False
    End With
    
    With cmdButton2
        .Top = cmdButton3.Top
        .Width = FormButtonWidth
        .Left = cmdButton3.Left + IIf(cmdButton3.Visible, -FormSpacing - .Width, 0)
        .Cancel = False
    End With
    
    With cmdButton1
        .Top = cmdButton2.Top
        .Width = FormButtonWidth
        .Left = cmdButton2.Left + IIf(cmdButton2.Visible, -FormSpacing - .Width, 0)
        .Cancel = False
    End With
    
    ' Set esc target
    If cmdButton1.Visible Then
        cmdButton1.Cancel = True
    ElseIf cmdButton2.Visible Then
        cmdButton2.Cancel = True
    Else
        cmdButton3.Cancel = True
    End If
    
    With cmdMoreDetails
        .Left = txtMessage.Left
        .Top = cmdButton3.Top
        .Width = FormButtonWidth
    End With
    
    With cmdReportIssue
        .Top = cmdButton3.Top
        .Width = FormButtonWidth
        .Left = cmdMoreDetails.Left + IIf(cmdMoreDetails.Visible, FormSpacing + .Width, 0)
    End With
    
    Me.Height = FormHeight(cmdButton1)
    Me.Width = Me.Width + FormWindowMargin
    
    Me.BackColor = FormBackColor
End Sub

Private Sub CenterForm()
    Me.Top = CenterFormTop(Me.Height)
    Me.Left = CenterFormLeft(Me.Width)
End Sub
