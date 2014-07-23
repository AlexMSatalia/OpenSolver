Attribute VB_Name = "OpenSolverParser"
Option Explicit

'==============================================================================
' ConvertCellToStandardName
' Range's address property always gives a $A$1 style address, but doesn't
' include the sheet. This function removes any nasty characters, and sticks
' the sheet name at the front, thus giving nice unique names for Python and
' VBA collections to use.
Function ConvertCellToStandardName(rngCell As Range, Optional strCleanParentName As String = "") As String
    Dim strCleanAddress As String
    strCleanAddress = rngCell.Address
    If strCleanParentName = "" Then strCleanParentName = Replace(rngCell.Parent.Name, " ", "_")
    strCleanParentName = Replace(strCleanParentName, "-", "_")
    strCleanAddress = Replace(strCleanAddress, "$", "")
    strCleanAddress = Replace(strCleanAddress, ":", "_")
    strCleanAddress = Replace(strCleanAddress, "-", "_")
    ConvertCellToStandardName = strCleanParentName + "_" + strCleanAddress
End Function
'==============================================================================

' Looks up node in collection and returns .strFormulaParsed. If node doesn't exist, returns the supplied default value
Function GetFormulaWithDefault(Formulae As Collection, NodeName As String, Default As String) As String
    If TestKeyExists(Formulae, NodeName) Then
        GetFormulaWithDefault = Formulae(NodeName).strFormulaParsed
    Else
        GetFormulaWithDefault = Default
    End If
End Function

' Shows .strFormulaParsed for all nodes in the Collection
Sub showFormulae(Formulae As Collection)
    Dim f As Variant, showstr As String
    For Each f In Formulae
        showstr = showstr & f.strFormulaParsed & vbNewLine & vbNewLine
    Next f
    MsgBox showstr
End Sub

' Shows all members of a collection
Sub showCollection(c As Collection)
    Dim f As Variant, showstr As String
    For Each f In c
        showstr = showstr & f & vbNewLine & vbNewLine
    Next f
    MsgBox showstr
End Sub
