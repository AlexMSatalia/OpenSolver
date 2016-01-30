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

' Used for tracking the constraints on the model form
' All modification must via `Update` - enforced by only having getters for these private variables

' These are set by the fields on the form
Private pLHSRefersTo As String
Private pRHSRefersTo As String
Private pRHSRefersToLocal As String  ' RefersTo in local format
Private pRelation As RelationConsts

' These are recalculated automatically each time the constraints are updated
Private pLHSRange As Range
Private pRHSRange As Range
Private pRHSFormula As String

Public Property Get LHSRefersTo() As String
    LHSRefersTo = pLHSRefersTo
End Property
Public Property Get RHSRefersTo() As String
    RHSRefersTo = pRHSRefersTo
End Property
Public Property Get RHSRefersToLocal() As String
    RHSRefersToLocal = pRHSRefersToLocal
End Property
Public Property Get Relation() As RelationConsts
    Relation = pRelation
End Property
Public Property Get LHSRange() As Range
    Set LHSRange = pLHSRange
End Property
Public Property Get RHSRange() As Range
    Set RHSRange = pRHSRange
End Property
Public Property Get RHSFormula() As String
    RHSFormula = pRHSFormula
End Property

Public Sub Update(newLHSRefersTo As String, newRelation As RelationConsts, newRHSRefersTo As String, sheet As Worksheet)
    pLHSRefersTo = newLHSRefersTo
    pRelation = newRelation
    pRHSRefersTo = IIf(RelationHasRHS(newRelation), newRHSRefersTo, vbNullString)
    pRHSRefersToLocal = ConvertToCurrentLocale(pRHSRefersTo)
    
    ConvertRefersToConstraint pLHSRefersTo, pRHSRefersTo, pLHSRange, pRHSRange, pRHSFormula
End Sub

Public Function ListDisplayString(sheet As Worksheet, Optional showNamedRanges As Boolean = False) As String
    ListDisplayString = GetDisplayAddress(pLHSRefersTo, sheet, showNamedRanges) & _
                        " " & RelationEnumToString(pRelation) & " " & _
                        GetDisplayAddress(pRHSRefersToLocal, sheet, showNamedRanges)
End Function
