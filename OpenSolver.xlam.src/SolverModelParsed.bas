Attribute VB_Name = "SolverModelParsed"
Option Explicit

Public SheetNameMap As Collection          ' Stores a map from sheet name to cleaned name
Public SheetNameMapReverse As Collection   ' Stores a map from cleaned name to sheet name

'==============================================================================
' ConvertCellToStandardName
' Range's address property always gives a $A$1 style address, but doesn't
' include the sheet. This function removes any nasty characters, and sticks
' the sheet name at the front, thus giving nice unique names for Python and
' VBA collections to use.
Function ConvertCellToStandardName(rngCell As Range, Optional strParentName As String = vbNullString) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim strCleanAddress As String
3         strCleanAddress = rngCell.Address

          Dim BannedChar As Variant
4         For Each BannedChar In Array("$", ":", "-")
5             strCleanAddress = Replace(strCleanAddress, BannedChar, vbNullString)
6         Next BannedChar

7         If Len(strParentName) = 0 Then strParentName = Replace(rngCell.Parent.Name, " ", "_")
          
          Dim strCleanParentName As String
8         If TestKeyExists(SheetNameMap, strParentName) Then
9             strCleanParentName = SheetNameMap(strParentName)
10        Else
11            strCleanParentName = strParentName
              
12            For Each BannedChar In Array("-", "+", " ", "(", ")", ":", "*", "/", "^", "!")
13                strCleanParentName = Replace(strCleanParentName, BannedChar, "_")
14            Next BannedChar
              
              ' If the cleaned name already exists, append an extra "1"
15            Do While TestKeyExists(SheetNameMapReverse, strCleanParentName)
16                strCleanParentName = strCleanParentName & "1"
17            Loop
18            SheetNameMap.Add strCleanParentName, strParentName
19            SheetNameMapReverse.Add strParentName, strCleanParentName
20        End If

21        ConvertCellToStandardName = strCleanParentName + "_" + strCleanAddress

ExitFunction:
22        If RaiseError Then RethrowError
23        Exit Function

ErrorHandler:
24        If Not ReportError("OpenSolverParser", "ConvertCellToStandardName") Then Resume
25        RaiseError = True
26        GoTo ExitFunction
End Function
'==============================================================================

' Looks up node in collection and returns .strFormulaParsed. If node doesn't exist, returns the supplied default value
Function GetFormulaWithDefault(Formulae As Collection, NodeName As String, Default As String) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If TestKeyExists(Formulae, NodeName) Then
4             GetFormulaWithDefault = Formulae(NodeName).strFormulaParsed
5         Else
6             GetFormulaWithDefault = Default
7         End If

ExitFunction:
8         If RaiseError Then RethrowError
9         Exit Function

ErrorHandler:
10        If Not ReportError("OpenSolverParser", "GetFormulaWithDefault") Then Resume
11        RaiseError = True
12        GoTo ExitFunction
End Function


