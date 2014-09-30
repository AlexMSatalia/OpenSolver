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
8980      Set items = New Collection
End Sub

Public Sub Push(Item As String)
8981      items.Add Item
End Sub

Public Function Pop() As String
          Dim obj As String
8982      obj = items(items.Count)
8983      items.Remove (items.Count)
8984      Pop = obj
End Function

Public Function Peek() As String
          Dim obj As String
8985      obj = items(items.Count)
8986      Peek = obj
End Function

Property Get Count() As Long
8987      Count = items.Count
End Property

Public Function Display() As String
          Dim i As Long, output As String
8988      For i = 1 To items.Count
8989          output = output & items(i) & ", "
8990      Next i
8991      Display = output
End Function
