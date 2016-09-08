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

Private Items As Collection

Private Sub Class_Initialize()
9163      Set Items = New Collection
End Sub

Public Sub PushNewCount()
9164      Items.Add 1
End Sub

Public Function PopCount() As Long
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim Count As Long
9165      Count = Items(Items.Count)
9166      Items.Remove (Items.Count)
9167      PopCount = Count

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("OperatorArgCountStack", "PopCount") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Public Sub Increase()
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim Count As Long
9168      Count = Items(Items.Count) + 1
9169      Items.Remove (Items.Count)
9170      Items.Add Count

ExitSub:
          If RaiseError Then RethrowError
          Exit Sub

ErrorHandler:
          If Not ReportError("OperatorArgCountStack", "Increase") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Public Function Display() As String
          Dim i As Long, output As String
9171      For i = 1 To Items.Count
9172          output = output & Items(i) & ", "
9173      Next i
9174      Display = output
End Function

