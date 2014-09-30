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
9163      Set items = New Collection
End Sub

Public Sub PushNewCount()
9164      items.Add 1
End Sub

Public Function PopCount() As Long
          Dim Count As Long
9165      Count = items(items.Count)
9166      items.Remove (items.Count)
9167      PopCount = Count
End Function

Public Sub Increase()
          Dim Count As Long
9168      Count = items(items.Count) + 1
9169      items.Remove (items.Count)
9170      items.Add Count
End Sub

Public Function Display() As String
          Dim i As Long, output As String
9171      For i = 1 To items.Count
9172          output = output & items(i) & ", "
9173      Next i
9174      Display = output
End Function

