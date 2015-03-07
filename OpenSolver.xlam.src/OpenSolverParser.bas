Attribute VB_Name = "OpenSolverParser"
Option Explicit

Public SheetNameMap As Collection          ' Stores a map from sheet name to cleaned name
Public SheetNameMapReverse As Collection   ' Stores a map from cleaned name to sheet name


'==============================================================================
' ConvertCellToStandardName
' Range's address property always gives a $A$1 style address, but doesn't
' include the sheet. This function removes any nasty characters, and sticks
' the sheet name at the front, thus giving nice unique names for Python and
' VBA collections to use.
Function ConvertCellToStandardName(rngCell As Range, Optional strParentName As String = "") As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim strCleanAddress As String
7438      strCleanAddress = rngCell.Address
7441      strCleanAddress = Replace(strCleanAddress, "$", "")
7442      strCleanAddress = Replace(strCleanAddress, ":", "_")
7443      strCleanAddress = Replace(strCleanAddress, "-", "_")

7439      If strParentName = "" Then strParentName = Replace(rngCell.Parent.Name, " ", "_")
          
          Dim strCleanParentName As String
          If TestKeyExists(SheetNameMap, strParentName) Then
              strCleanParentName = SheetNameMap(strParentName)
          Else
7440          strCleanParentName = Replace(strParentName, "-", "_")
              strCleanParentName = Replace(strCleanParentName, "+", "_")
              strCleanParentName = Replace(strCleanParentName, " ", "_")
              strCleanParentName = Replace(strCleanParentName, "(", "_")
              strCleanParentName = Replace(strCleanParentName, ")", "_")
              strCleanParentName = Replace(strCleanParentName, ":", "_")
              strCleanParentName = Replace(strCleanParentName, "*", "_")
              strCleanParentName = Replace(strCleanParentName, "/", "_")
              strCleanParentName = Replace(strCleanParentName, "^", "_")
              ' If the cleaned name already exists, append an extra "1"
              Do While TestKeyExists(SheetNameMapReverse, strCleanParentName)
                  strCleanParentName = strCleanParentName & "1"
              Loop
              SheetNameMap.Add strCleanParentName, strParentName
              SheetNameMapReverse.Add strParentName, strCleanParentName
          End If

7444      ConvertCellToStandardName = strCleanParentName + "_" + strCleanAddress

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverParser", "ConvertCellToStandardName") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function
'==============================================================================

' Looks up node in collection and returns .strFormulaParsed. If node doesn't exist, returns the supplied default value
Function GetFormulaWithDefault(Formulae As Collection, NodeName As String, Default As String) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

7445      If TestKeyExists(Formulae, NodeName) Then
7446          GetFormulaWithDefault = Formulae(NodeName).strFormulaParsed
7447      Else
7448          GetFormulaWithDefault = Default
7449      End If

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverParser", "GetFormulaWithDefault") Then Resume
          RaiseError = True
          GoTo ExitFunction
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
