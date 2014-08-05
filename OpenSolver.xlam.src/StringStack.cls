VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Stack for Strings

Private items As Collection

Private Sub Class_Initialize()
    Set items = New Collection
End Sub

Public Sub Push(Item As String)
    items.Add Item
End Sub

Public Function Pop() As String
    Dim obj As String
    obj = items(items.Count)
    items.Remove (items.Count)
    Pop = obj
End Function

Public Function Peek() As String
    Dim obj As String
    obj = items(items.Count)
    Peek = obj
End Function

Property Get Count() As Long
    Count = items.Count
End Property
