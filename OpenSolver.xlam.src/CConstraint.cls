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

Public Function GetKey() As String
          ' Constraints added manually, or added when Model is used on an
          ' existing sheet, don't have a relation.
          ' They use the LHS as a key instead
          'If Relation Is Nothing Then
          '    If RHS Is Nothing Then
          '        GetKey = LHS.Address + ConstraintType + RHSstring
          '    Else
          '        GetKey = LHS.Address + ConstraintType + RHS.Address
          '    End If
          '    'GetKey = LHS.Address
          'Else
          '    ' Added with AutoModel
          '    ' Key is relational operator cell
          '    GetKey = Relation.Address
          'End If
          
          
          ' Disregard above - standardised key introduced 18June2011
3545      If RHS Is Nothing Then
3546          GetKey = LHS.Address + ConstraintType + RHSstring
3547      Else
3548          GetKey = LHS.Address + ConstraintType + RHS.Address
3549      End If
End Function
