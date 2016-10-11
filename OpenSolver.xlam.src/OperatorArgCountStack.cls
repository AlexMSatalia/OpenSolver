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
1         Set Items = New Collection
End Sub

Public Sub PushNewCount()
1         Items.Add 1
End Sub

Public Function PopCount() As Long
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim Count As Long
3         Count = Items(Items.Count)
4         Items.Remove (Items.Count)
5         PopCount = Count

ExitFunction:
6         If RaiseError Then RethrowError
7         Exit Function

ErrorHandler:
8         If Not ReportError("OperatorArgCountStack", "PopCount") Then Resume
9         RaiseError = True
10        GoTo ExitFunction
End Function

Public Sub Increase()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim Count As Long
3         Count = Items(Items.Count) + 1
4         Items.Remove (Items.Count)
5         Items.Add Count

ExitSub:
6         If RaiseError Then RethrowError
7         Exit Sub

ErrorHandler:
8         If Not ReportError("OperatorArgCountStack", "Increase") Then Resume
9         RaiseError = True
10        GoTo ExitSub
End Sub

Public Function Display() As String
          Dim i As Long, output As String
1         For i = 1 To Items.Count
2             output = output & Items(i) & ", "
3         Next i
4         Display = output
End Function

