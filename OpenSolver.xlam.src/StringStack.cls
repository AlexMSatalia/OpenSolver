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

Private Items As Collection

Private Sub Class_Initialize()
1         Set Items = New Collection
End Sub

Public Sub Push(Item As String)
1         Items.Add Item
End Sub

Public Function Pop() As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim obj As String
3         obj = Items(Items.Count)
4         Items.Remove (Items.Count)
5         Pop = obj

ExitFunction:
6         If RaiseError Then RethrowError
7         Exit Function

ErrorHandler:
8         If Not ReportError("StringStack", "Pop") Then Resume
9         RaiseError = True
10        GoTo ExitFunction
End Function

Public Function Peek() As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim obj As String
3         obj = Items(Items.Count)
4         Peek = obj

ExitFunction:
5         If RaiseError Then RethrowError
6         Exit Function

ErrorHandler:
7         If Not ReportError("StringStack", "Peek") Then Resume
8         RaiseError = True
9         GoTo ExitFunction
End Function

Property Get Count() As Long
1         Count = Items.Count
End Property

Public Function Display() As String
          Dim i As Long, output As String
1         For i = 1 To Items.Count
2             output = output & Items(i) & ", "
3         Next i
4         Display = output
End Function
