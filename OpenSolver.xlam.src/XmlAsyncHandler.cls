VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XmlAsyncHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_xmlHttp As Object ' MSXML2.XMLHTTP
Dim MacroToRun As String

'http://dailydoseofexcel.com/archives/2006/10/09/async-xmlhttp-calls/

Public Sub Initialize(ByRef xmlHttpRequest As Object, MacroName As String)
    Set m_xmlHttp = xmlHttpRequest
    MacroToRun = MacroName
End Sub

Sub OnReadyStateChange()
Attribute OnReadyStateChange.VB_UserMemId = 0
    If m_xmlHttp.readyState = 4 Then
        If m_xmlHttp.status = 200 Then
            Application.Run MacroToRun, m_xmlHttp.responseText
        Else
           'Error happened
        End If
    End If
End Sub
