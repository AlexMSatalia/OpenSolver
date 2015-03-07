VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------
' CConstraint
'-----------------------------------------

' Defines a constraint This was originally a user-defined type,
' but apparently they can't be added to collections

Public LHS As Range
Public RHS As Range
Public RHSstring As String
Public ConstraintType As String
Public Relation As Range
