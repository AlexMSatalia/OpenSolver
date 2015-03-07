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
8980      Set Items = New Collection
End Sub

Public Sub Push(Item As String)
8981      Items.Add Item
End Sub

Public Function Pop() As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim obj As String
8982      obj = Items(Items.Count)
8983      Items.Remove (Items.Count)
8984      Pop = obj

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("StringStack", "Pop") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Public Function Peek() As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim obj As String
8985      obj = Items(Items.Count)
8986      Peek = obj

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("StringStack", "Peek") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Property Get Count() As Long
8987      Count = Items.Count
End Property

Public Function Display() As String
          Dim i As Long, output As String
8988      For i = 1 To Items.Count
8989          output = output & Items(i) & ", "
8990      Next i
8991      Display = output
End Function
