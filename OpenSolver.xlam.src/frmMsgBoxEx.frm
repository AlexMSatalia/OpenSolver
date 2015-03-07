VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMsgBoxEx 
   Caption         =   "OpenSolver - Message"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9720
   OleObjectBlob   =   "frmMsgBoxEx.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMsgBoxEx"
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
    MailToLink = "mailto:a.mason@auckland.ac.nz?cc=jdun087@aucklanduni.ac.nz" & _
                 "&subject=" & URLEncode(EmailSubject) & _
                 "&body=" & URLEncode(EmailBody)
    
    OpenURL MailToLink
End Sub

Private Sub lblLink_Click()
    Call OpenURL(lblLink.ControlTipText)
End Sub

Private Sub UserForm_Activate()
    AutoLayout
End Sub

Private Sub AutoLayout()
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
    Me.width = Max_Double(FormWidthMessageBox, NumButtonWidths * FormButtonWidth + (NumButtonWidths - 1) * FormSpacing + 2 * FormMargin)
    
    With txtMessage
        .BackColor = FormBackColor
        .width = Me.width - 2 * FormMargin
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
        .width = Me.width - 2 * FormMargin
        .left = FormMargin
        .top = FormMargin
    End With
    
    With lblLink
        If .Caption = "" Then
            .Visible = False
        Else
            .width = Me.width
            .AutoSize = False
            .AutoSize = True
            .AutoSize = False
            .left = (Me.width - .width) / 2
            .Font.Underline = True
        End If
        .top = txtMessage.top + txtMessage.height + FormSpacing
    End With
    
    With cmdButton3
        .top = IIf(lblLink.Visible, lblLink.height + FormSpacing, 0) + lblLink.top
        .width = FormButtonWidth
        .left = Me.width - FormMargin - .width
    End With
    
    With cmdButton2
        .top = cmdButton3.top
        .width = FormButtonWidth
        .left = cmdButton3.left + IIf(cmdButton3.Visible, -FormSpacing - .width, 0)
    End With
    
    With cmdButton1
        .top = cmdButton2.top
        .width = FormButtonWidth
        .left = cmdButton2.left + IIf(cmdButton2.Visible, -FormSpacing - .width, 0)
    End With
    
    With cmdMoreDetails
        .left = txtMessage.left
        .top = cmdButton3.top
        .width = FormButtonWidth
    End With
    
    With cmdReportIssue
        .top = cmdButton3.top
        .width = FormButtonWidth
        .left = cmdMoreDetails.left + IIf(cmdMoreDetails.Visible, FormSpacing + .width, 0)
    End With
    
    Me.height = cmdButton1.top + cmdButton1.height + FormMargin + FormTitleHeight
    Me.width = Me.width + FormWindowMargin
    
    Me.BackColor = FormBackColor
End Sub

