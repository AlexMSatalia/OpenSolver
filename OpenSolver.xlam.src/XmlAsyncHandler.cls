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
1         Set m_xmlHttp = xmlHttpRequest
2         MacroToRun = MacroName
End Sub

Sub OnReadyStateChange()
Attribute OnReadyStateChange.VB_UserMemId = 0
1         If m_xmlHttp.readyState = 4 Then
2             If m_xmlHttp.status = 200 Then
3                 Application.Run MacroToRun, m_xmlHttp.responseText
4             Else
                 'Error happened
5             End If
6         End If
End Sub
