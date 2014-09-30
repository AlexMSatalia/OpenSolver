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
7438      strCleanAddress = rngCell.Address
7439      If strCleanParentName = "" Then strCleanParentName = Replace(rngCell.Parent.Name, " ", "_")
7440      strCleanParentName = Replace(strCleanParentName, "-", "_")
7441      strCleanAddress = Replace(strCleanAddress, "$", "")
7442      strCleanAddress = Replace(strCleanAddress, ":", "_")
7443      strCleanAddress = Replace(strCleanAddress, "-", "_")
7444      ConvertCellToStandardName = strCleanParentName + "_" + strCleanAddress
End Function
'==============================================================================

' Looks up node in collection and returns .strFormulaParsed. If node doesn't exist, returns the supplied default value
Function GetFormulaWithDefault(Formulae As Collection, NodeName As String, Default As String) As String
7445      If TestKeyExists(Formulae, NodeName) Then
7446          GetFormulaWithDefault = Formulae(NodeName).strFormulaParsed
7447      Else
7448          GetFormulaWithDefault = Default
7449      End If
End Function

' Shows .strFormulaParsed for all nodes in the Collection
Sub showFormulae(Formulae As Collection)
          Dim f As Variant, showstr As String
7450      For Each f In Formulae
7451          showstr = showstr & f.strFormulaParsed & vbNewLine & vbNewLine
7452          Debug.Print f.strFormulaParsed
7453      Next f
          'MsgBox showstr
End Sub

' Shows all members of a collection
Sub showCollection(c As Collection)
          Dim f As Variant, showstr As String
7454      For Each f In c
7455          showstr = showstr & f & vbNewLine & vbNewLine
7456      Next f
7457      MsgBox showstr
End Sub
