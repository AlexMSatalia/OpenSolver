Attribute VB_Name = "OpenSolverStoredNames"
Option Explicit

Const SolverPrefix As String = "solver_"
Const TempPrefix As String = "OpenSolver_Temp"

Sub GetSheetNameAsValueOrRange(sheet As Worksheet, theName As String, IsMissing As Boolean, IsRange As Boolean, r As Range, RefersToFormula As Boolean, RangeRefersToError As Boolean, RefersTo As String, value As Double)
    ' Wrapper for a sheet-prefixed defined name
    GetNameAsValueOrRange sheet.Parent, EscapeSheetName(sheet) & theName, IsMissing, IsRange, r, RefersToFormula, RangeRefersToError, RefersTo, value
End Sub

Sub GetNameAsValueOrRange(book As Workbook, theName As String, IsMissing As Boolean, IsRange As Boolean, r As Range, RefersToFormula As Boolean, RangeRefersToError As Boolean, RefersTo As String, value As Double)
' See http://www.cpearson.com/excel/DefinedNames.aspx, but see below for internationalisation problems with this code
172       RangeRefersToError = False
173       RefersToFormula = False
          ' Dim r As Range
          Dim NM As Name
174       On Error Resume Next
175       Set NM = book.Names(theName)
176       If Err.Number <> 0 Then
177           IsMissing = True
178           Exit Sub
179       End If
180       IsMissing = False
          RefersTo = Mid(NM.RefersTo, 2)
181       On Error Resume Next
182       Set r = GetRefersToRange(RefersTo)
183       If Err.Number = 0 Then
184           IsRange = True
185       Else
186           IsRange = False
187           ' String will be of form: "5", or "Sheet1!#REF!" or "#REF!$A$1" or "Sheet1!$A$1/4+Sheet1!$A$3"
189           If Right(RefersTo, 6) = "!#REF!" Or Left(RefersTo, 5) = "#REF!" Then
191               RangeRefersToError = True
192           Else
                  ' Test for a numeric constant, in US format
193               If IsAmericanNumber(RefersTo) Then
194                   value = Val(RefersTo)   ' Force a conversion to a number using Val which uses US settings (no regionalisation)
195               Else
196                   RefersToFormula = True
197               End If
198           End If
199       End If
End Sub

Function GetNamedDoubleIfExists(sheet As Worksheet, Name As String, DoubleValue As Double) As Boolean
          Dim IsRange As Boolean, r As Range, RefersToFormula As Boolean, RangeRefersToError As Boolean, RefersTo As String, IsMissing As Boolean
154       GetSheetNameAsValueOrRange sheet, Name, IsMissing, IsRange, r, RefersToFormula, RangeRefersToError, RefersTo, DoubleValue
155       GetNamedDoubleIfExists = Not IsMissing And Not IsRange And Not RefersToFormula And Not RangeRefersToError
End Function

Function GetNamedIntegerIfExists(sheet As Worksheet, Name As String, IntegerValue As Long) As Boolean
          Dim DoubleValue As Double
156       If GetNamedDoubleIfExists(sheet, Name, DoubleValue) Then
157           IntegerValue = CLng(DoubleValue)
158           GetNamedIntegerIfExists = (IntegerValue = DoubleValue)
161       End If
End Function

Function GetNamedIntegerAsBooleanIfExists(sheet As Worksheet, Name As String, BooleanValue As Boolean) As Boolean
          Dim IntegerValue As Long
          If GetNamedIntegerIfExists(sheet, Name, IntegerValue) Then
              BooleanValue = (IntegerValue = 1)
              GetNamedIntegerAsBooleanIfExists = True
          End If
End Function

Function GetNamedBooleanIfExists(sheet As Worksheet, Name As String, BooleanValue As Boolean) As Boolean
    Dim IsRange As Boolean, r As Range, RefersToFormula As Boolean, RangeRefersToError As Boolean, RefersTo As String, IsMissing As Boolean, value As Double
    GetSheetNameAsValueOrRange sheet, Name, IsMissing, IsRange, r, RefersToFormula, RangeRefersToError, RefersTo, value

    If Not IsMissing And Not IsRange And Not RangeRefersToError Then
        On Error GoTo NotBoolean
        BooleanValue = CBool(RefersTo)
        GetNamedBooleanIfExists = True
    Else
NotBoolean:
        GetNamedBooleanIfExists = False
    End If
End Function

Function GetNamedStringIfExists(sheet As Worksheet, Name As String, StringValue As String) As Boolean
    Dim IsRange As Boolean, r As Range, RefersToFormula As Boolean, RangeRefersToError As Boolean, value As Double, IsMissing As Boolean
    GetSheetNameAsValueOrRange sheet, Name, IsMissing, IsRange, r, RefersToFormula, RangeRefersToError, StringValue, value
    
    ' Remove any quotes
    If Left(StringValue, 1) = """" Then StringValue = Mid(StringValue, 2, Len(StringValue) - 2)
    
    GetNamedStringIfExists = Not IsMissing And Not IsRange And Not RangeRefersToError
End Function

Function GetNamedDoubleWithDefault(sheet As Worksheet, Name As String, DefaultValue As Double) As Double
    If Not GetNamedDoubleIfExists(sheet, Name, GetNamedDoubleWithDefault) Then
        GetNamedDoubleWithDefault = DefaultValue
        SetDoubleNameOnSheet Name, GetNamedDoubleWithDefault, sheet
    End If
End Function

Function GetNamedIntegerWithDefault(sheet As Worksheet, Name As String, DefaultValue As Long) As Long
    If Not GetNamedIntegerIfExists(sheet, Name, GetNamedIntegerWithDefault) Then
        GetNamedIntegerWithDefault = DefaultValue
        SetIntegerNameOnSheet Name, GetNamedIntegerWithDefault, sheet
    End If
End Function

Function GetNamedIntegerAsBooleanWithDefault(sheet As Worksheet, Name As String, DefaultValue As Boolean) As Boolean
    If Not GetNamedIntegerAsBooleanIfExists(sheet, Name, GetNamedIntegerAsBooleanWithDefault) Then
        GetNamedIntegerAsBooleanWithDefault = DefaultValue
        SetBooleanAsIntegerNameOnSheet Name, GetNamedIntegerAsBooleanWithDefault, sheet
    End If
End Function

Function GetNamedBooleanWithDefault(sheet As Worksheet, Name As String, DefaultValue As Boolean) As Boolean
    If Not GetNamedBooleanIfExists(sheet, Name, GetNamedBooleanWithDefault) Then
        GetNamedBooleanWithDefault = DefaultValue
        SetBooleanNameOnSheet Name, GetNamedBooleanWithDefault, sheet
    End If
End Function

Function GetNamedStringWithDefault(sheet As Worksheet, Name As String, DefaultValue As String) As String
    If Not GetNamedStringIfExists(sheet, Name, GetNamedStringWithDefault) Then
        GetNamedStringWithDefault = DefaultValue
        SetNameOnSheet Name, GetNamedStringWithDefault, sheet
    End If
End Function

Sub DeleteNameOnSheet(Name As String, Optional sheet As Worksheet, Optional SolverName As Boolean = False)
          GetActiveSheetIfMissing sheet
608       Name = EscapeSheetName(sheet) & IIf(SolverName, SolverPrefix, "") & Name
609       On Error Resume Next
610       sheet.Parent.Names(Name).Delete
doesntExist:
End Sub

Sub SetNameOnSheet(Name As String, value As Variant, Optional sheet As Worksheet, Optional SolverName As Boolean = False)
' If a key exists we can just add it (http://www.cpearson.com/Excel/DefinedNames.aspx)
          GetActiveSheetIfMissing sheet
600       Name = EscapeSheetName(sheet) & IIf(SolverName, SolverPrefix, "") & Name
603       sheet.Parent.Names.Add Name, value, False
End Sub

Sub SetNamedRangeIfExists(ByVal Name As String, ByRef RangeToSet As Range, Optional sheet As Worksheet, Optional SolverName As Boolean = False)
    If RangeToSet Is Nothing Then
        DeleteNameOnSheet Name, sheet, SolverName
    Else
        SetNamedRangeOnSheet Name, RangeToSet, sheet, SolverName
    End If
End Sub

Sub SetNamedRangeOnSheet(Name As String, value As Range, Optional sheet As Worksheet, Optional SolverName As Boolean = False)
    SetNameOnSheet Name, value, sheet, SolverName
End Sub

Sub SetIntegerNameOnSheet(Name As String, value As Long, Optional sheet As Worksheet, Optional SolverName As Boolean = False)
    SetDoubleNameOnSheet Name, CDbl(value), sheet, SolverName
End Sub

Sub SetDoubleNameOnSheet(Name As String, value As Double, Optional sheet As Worksheet, Optional SolverName As Boolean = False)
    SetNameOnSheet Name, value, sheet, SolverName
End Sub

Sub SetBooleanNameOnSheet(Name As String, value As Boolean, Optional sheet As Worksheet, Optional SolverName As Boolean = False)
    SetNameOnSheet Name, value, sheet, SolverName
End Sub

Sub SetBooleanAsIntegerNameOnSheet(Name As String, value As Boolean, Optional sheet As Worksheet, Optional SolverName As Boolean = False)
    SetIntegerNameOnSheet Name, IIf(value, 1, 2), sheet, SolverName
End Sub

Sub SetRefersToNameOnSheet(Name As String, value As String, Optional sheet As Worksheet, Optional SolverName As Boolean = False)
    If Len(value) > 0 Then
        SetNameOnSheet Name, AddEquals(value), sheet, SolverName
    Else
        DeleteNameOnSheet Name, sheet, SolverName
    End If
End Sub

Sub SetAnyMissingDefaultSolverOptions(sheet As Worksheet)
          ' We set all the default values, as per Solver in Excel 2007, but with some changes. This ensures Solver does not delete the few values we actually use
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim SolverOptions() As Variant, SolverDefaults() As Variant
          SolverOptions = Array("drv", "est", "nwt", "scl", "cvg", "rlx")
          SolverDefaults = Array("1", "1", "1", "2", "0.0001", "2")
          
          GetActiveSheetIfMissing sheet
          
          Dim i As Long, value As Double
          For i = LBound(SolverOptions) To UBound(SolverOptions)
              If Not GetNamedDoubleIfExists(sheet, "solver_" & SolverOptions(i), value) Then
                  SetNameOnSheet CStr(SolverOptions(i)), "=" & SolverDefaults(i), sheet, True
              End If
          Next i

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("OpenSolverStoredNames", "SetAnyMissingDefaultExcel2007SolverOptions") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Function GetRefersToRange(RefersTo As String) As Range
    If Len(RefersTo) = 0 Then Exit Function
    
    ' Add the text as a name and retrieve the corresponding sanitised range
    ' We can't use .RefersToRange as this is broken in non-US locales
    ' Instead we use sheet.Range()
    On Error GoTo InvalidRange
    Set GetRefersToRange = Application.Range(RefersTo)
    Exit Function
    
InvalidRange:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Function RefEditToRefersTo(RefEditText As String) As String
    If Len(RefEditText) = 0 Then Exit Function
    
    ' Try to evaluate the formula and catch `Error 2015` for an invalid formula
    ' ============
    ' IMPORTANT: We can't catch this error using an error handler!
    ' This is sometimes run inside a RefEdit event, in which the
    ' error handlers don't work properly. We have to use a method of
    ' detecting the invalid formula that never throws any error.
    ' ============
    Dim varReturn As Variant
    varReturn = Application.Evaluate(RefEditText)
    If VarType(varReturn) = vbError Then
        If CLng(varReturn) = 2015 Then
            RefEditToRefersTo = RefEditText
            Exit Function
        End If
    End If

    ' Add the text as a name and retrieve the sanitised RefersTo
    Dim n As Name
    Set n = ActiveWorkbook.Names.Add(TempPrefix, AddEquals(Trim(RefEditText)))
    RefEditToRefersTo = Mid(n.RefersTo, 2)
    n.Delete
End Function

Function RangeToRefersTo(ConvertRange As Range) As String
    If ConvertRange Is Nothing Then Exit Function
    
    ' Add the text as a name and retrieve the sanitised refers to
    Dim n As Name
    Set n = ActiveWorkbook.Names.Add(TempPrefix, ConvertRange)
    RangeToRefersTo = Mid(n.RefersTo, 2)
    n.Delete
End Function
