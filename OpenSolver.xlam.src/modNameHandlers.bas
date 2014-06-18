Attribute VB_Name = "modNameHandlers"
'==============================================================================
' OpenSolver
' Copyright Andrew Mason, Iain Dunning, 2011
' http://www.opensolver.org
'==============================================================================
' modNameHandlers
' Functions to read and write values from workbook Names
'==============================================================================
Option Explicit

'==============================================================================
Function GetStringFromName(ByRef book As Workbook, ByVal strName As String, ByRef strValue As String) As Boolean
    ' GetStringFromName
    On Error Resume Next
    Dim NM As Name
    Set NM = book.Names(strName)
    If Err.Number <> 0 Then
        GetStringFromName = False
        Exit Function
    End If
    strValue = Mid(book.Names(strName).RefersTo, 2)
    GetStringFromName = True
End Function

'==============================================================================
Function GetIntegerFromName(ByRef book As Workbook, ByVal strName As String, ByRef lngValue As Long) As Boolean
    ' GetIntegerFromName
    ' Gets an integer from a Name.
    ' If the Name does not contain an integer, return False.
    Dim dblValue As Double
    If GetValueFromName(book, strName, dblValue) Then
        lngValue = CLng(dblValue)
        GetIntegerFromName = (lngValue = dblValue)
    Else
        GetIntegerFromName = False
    End If
End Function

'==============================================================================
Function GetValueFromName(ByRef book As Workbook, ByVal strName As String, ByRef dblValue As Double) As Boolean
    ' GetValueFromName
    ' Gets a numeric value from a Name
    ' If the name does not contain a numeric value, return False
    Dim isMissing As Boolean, isRange As Boolean, r As Range
    Dim isFormula As Boolean, isError As Boolean, sRefersTo As String
    
    GetValueOrRangeFromName book, strName, isMissing, isRange, r, isFormula, isError, sRefersTo, dblValue
    GetValueFromName = (Not isMissing) And (Not isRange) And (Not isFormula) And (Not isError)
End Function


'==============================================================================
Sub GetValueOrRangeFromName( _
    ByRef book As Workbook, strName As String, ByRef isMissing As Boolean, _
    ByRef isRange As Boolean, ByRef r As Range, _
    ByRef isFormula As Boolean, ByRef isError As Boolean, _
    ByRef sRefersTo As String, ByRef dblValue As Double)
    ' GetValueOrFromName
    ' Extracts general information from a Name, with lots of fallback
    ' flags to let you know what the Name refers to if not a simple
    ' value or range.
    ' More info about Names:
    '   http://www.cpearson.com/excel/DefinedNames.aspx
    ' There are some issues with internationalisation, see below.
    
    ' Assume the best
    isError = False
    isFormula = False

    Dim NM As Name
    
    ' Test the Name exists
    On Error Resume Next
    Set NM = book.Names(strName)
    If Err.Number <> 0 Then
        isMissing = True
        Exit Sub
    End If
    
    ' Name does exist, test if it is pointing to a range
    isMissing = False
    On Error Resume Next
    Set r = NM.RefersToRange
    isRange = (Err.Number = 0)
    
    If Not isRange Then
        ' Name could be:
        '   Value:   "=5"
        '   Error:   "=Sheet1!#REF!"
        '   Formula: "=Test4!$M$11/4+Test4!$A$3"
        sRefersTo = Mid(NM.RefersTo, 2)
        ' Test for the error case
        If right(sRefersTo, 6) = "!#REF!" Then
            isError = True
        Else
            ' Test for a numeric constant, in US format
            If IsAmericanNumber(sRefersTo) Then
                dblValue = Val(sRefersTo)   ' Force a conversion to a number using Val which uses US settings (no regionalisation)
            Else
                ' Must be a formula then
                isFormula = True
            End If
        End If
    End If
    
End Sub

