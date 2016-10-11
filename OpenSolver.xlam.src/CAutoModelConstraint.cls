VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAutoModelConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public LHS As Range
Public RHS As Range
Public RelationType As RelationConsts
Public RelationCell As Range

Public Sub Init(newLHS As Range, newRelationType As RelationConsts, newRelationCell As Range, Optional newRHS As Range)
1         Set LHS = newLHS
2         Set RHS = newRHS
3         RelationType = newRelationType
4         Set RelationCell = newRelationCell
End Sub

Public Property Get KeyCell() As Range
1        Set KeyCell = IIf(RelationHasRHS(RelationType), RelationCell, LHS)
End Property

Public Property Get Key() As String
1        Key = KeyCell.Address
End Property
