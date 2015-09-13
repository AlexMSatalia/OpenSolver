VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Tag As String
Public Id As String
Public Label As String
Public OnAction As String
Public ScreenTip As String
Public SuperTip As String
Public Image As String
Public Size As String
Public Children As Collection
