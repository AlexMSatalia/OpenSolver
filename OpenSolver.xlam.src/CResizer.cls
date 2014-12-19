VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CResizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MFrameResizer = "FrameResizeGrab"
Private Const MResizer = "ResizeGrab"
Private WithEvents m_objResizer As MSForms.Frame
Attribute m_objResizer.VB_VarHelpID = -1
Private m_sngLeftResizePos As Single
Private m_sngTopResizePos As Single
Private m_blnResizing As Boolean
Attribute m_blnResizing.VB_VarHelpID = -1
Private WithEvents m_frmParent As MSForms.UserForm
Attribute m_frmParent.VB_VarHelpID = -1
Private m_strParentName As String
Attribute m_strParentName.VB_VarHelpID = -1
Private m_objParent As Object
Public width As Long

Private Sub Class_Terminate()
    m_objParent.Controls.Remove MResizer
End Sub

Private Sub m_frmParent_Layout()
    If Not m_blnResizing Then
        With m_objResizer
            .top = m_objParent.InsideHeight - .height
            .left = m_objParent.InsideWidth - .width
        End With
    End If
End Sub

Private Sub m_objResizer_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If Button = 1 Then
        m_sngLeftResizePos = X
        m_sngTopResizePos = Y
        m_blnResizing = True
    End If

End Sub

Private Sub m_objResizer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Dim ChangeY As Single
    
    If m_strParentName = "Model" Then
        If Button = 1 Then
            ChangeY = Y - m_sngTopResizePos
            
            If m_objParent.InsideHeight + ChangeY > frmModel.MinHeight Then
                Call frmModel.MoveItems(m_frmParent, ChangeY)
                
                With m_objResizer
                    .Move .left + X - m_sngLeftResizePos, .top + ChangeY
                    .left = m_objParent.InsideWidth - .width
                    .top = m_objParent.InsideHeight - .height
                End With
            End If
            
        End If
    End If
    
End Sub

Private Sub m_objResizer_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        m_blnResizing = False
    End If
End Sub

Public Function Add(Parent As MSForms.UserForm, ParentName As String) As MSForms.Frame
    Dim labTemp As MSForms.label
    
    Set m_frmParent = Parent
    m_strParentName = ParentName
    Set m_objParent = Parent
    Set m_objResizer = m_objParent.Controls.Add("Forms.Frame.1", MFrameResizer, True)
    Set labTemp = m_objResizer.Add("Forms.label.1", MResizer, True)
    With labTemp
        With .Font
            .Name = "Marlett"
            .Charset = 2
            .Size = 10
            .Bold = True
        End With
    .BackStyle = fmBackStyleTransparent
    .AutoSize = True
    .BorderStyle = fmBorderStyleNone
    .Caption = "o"
    .MousePointer = fmMousePointerSizeNWSE
    .ForeColor = RGB(100, 100, 100)
    .ZOrder
    .top = 1
    .left = 1
    .Enabled = False
    End With
    
    With m_objResizer
        .MousePointer = fmMousePointerSizeNWSE
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .ZOrder
        .Caption = ""
        .width = labTemp.width + 1
        width = .width
        .height = labTemp.height + 1
        .top = m_objParent.InsideHeight - .height
        .left = m_objParent.InsideWidth - .width
    End With
End Function
