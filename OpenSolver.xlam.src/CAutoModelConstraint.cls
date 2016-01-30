VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAutoModelConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public LHS As Range
Public RHS As Range
Public RelationType As RelationConsts
Public RelationCell As Range

Public Sub Init(newLHS As Range, newRelationType As RelationConsts, newRelationCell As Range, Optional newRHS As Range)
    Set LHS = newLHS
    Set RHS = newRHS
    RelationType = newRelationType
    Set RelationCell = newRelationCell
End Sub

Public Property Get KeyCell() As Range
   Set KeyCell = IIf(RelationHasRHS(RelationType), RelationCell, LHS)
End Property

Public Property Get Key() As String
   Key = KeyCell.Address
End Property
