VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OperatorArgCountStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Stack for counting numbers of operator arguments

Private items As Collection

Private Sub Class_Initialize()
    Set items = New Collection
End Sub

Public Sub PushNewCount()
    items.Add 1
End Sub

Public Function PopCount() As Integer
    Dim Count As Integer
    Count = items(items.Count)
    items.Remove (items.Count)
    PopCount = Count
End Function

Public Sub Increase()
    Dim Count As Integer
    Count = items(items.Count) + 1
    items.Remove (items.Count)
    items.Add Count
End Sub

Public Function Display() As String
    Dim i As Integer, output As String
    For i = 1 To items.Count
        output = output & items(i) & ", "
    Next i
    Display = output
End Function

