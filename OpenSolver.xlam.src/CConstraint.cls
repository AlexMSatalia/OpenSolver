VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public LHS As Range
Public RHS As Range
Public RHSstring As String
Public RelationType As RelationConsts
Public RelationCell As Range  ' Used during AutoModel: The cell on the sheet containing the relation information

Public Sub Init(newLHS As Range, newRelationType As RelationConsts, Optional newRelationCell As Range, Optional newRHS As Range, Optional newRHSstring As String)
    Set LHS = newLHS
    Set RHS = newRHS
    RHSstring = newRHSstring
    RelationType = newRelationType
    Set RelationCell = newRelationCell
End Sub

Public Property Get KeyCell() As Range
   Set KeyCell = IIf(RelationHasRHS(RelationType), RelationCell, LHS)
End Property

Public Property Get Key() As String
   Key = KeyCell.Address
End Property


