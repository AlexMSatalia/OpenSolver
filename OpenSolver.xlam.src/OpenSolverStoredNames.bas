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

Function GetNamedBooleanIfExists(sheet As Worksheet, Name As String, BooleanValue As Boolean) As Boolean
    Dim IsRange As Boolean, r As Range, RefersToFormula As Boolean, RangeRefersToError As Boolean, RefersTo As String, IsMissing As Boolean, value As Double
    GetSheetNameAsValueOrRange sheet, Name, IsMissing, IsRange, r, RefersToFormula, RangeRefersToError, RefersTo, value

    If Not IsMissing And Not IsRange And Not RangeRefersToError Then
        If RefersToFormula Then
            ' It's a string, probably "=TRUE" or "=FALSE"
            ' We handle this conversion for legacy reasons, in versions 2.8.2 and earlier OpenSolver bool options
            ' were stored as strings "=TRUE" or "=FALSE". This was changed to 0/1 to resolve locale issues with CBool.
            ' e.g. In French, the strings would be "=VRAI" and "=FAUX", which didn't always work with CBool.
            On Error GoTo NotBoolean
            BooleanValue = CBool(RefersTo)
            GetNamedBooleanIfExists = True
        Else
            ' It's a value
            If CLng(value) = value Then
                ' It's integer, we can interpret that as a bool
                GetNamedBooleanIfExists = True
                BooleanValue = IntToBool(CLng(value))
            Else
                ' It's a double, so not a boolean
                GetNamedBooleanIfExists = False
            End If
        End If
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
          Dim FullName As String ' Don't mangle the value of Name!
600       FullName = EscapeSheetName(sheet) & IIf(SolverName, SolverPrefix, "") & Name
603       sheet.Parent.Names.Add FullName, value, False
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
    SetIntegerNameOnSheet Name, BoolToInt(value), sheet, SolverName
End Sub

Sub SetRefersToNameOnSheet(Name As String, value As String, Optional sheet As Worksheet, Optional SolverName As Boolean = False)
    If Len(value) > 0 Then
        SetNameOnSheet Name, AddEquals(value), sheet, SolverName
    Else
        DeleteNameOnSheet Name, sheet, SolverName
    End If
End Sub

Function BoolToInt(BoolValue As Boolean) As Long
    BoolToInt = IIf(BoolValue, 1, 0)
End Function

Function IntToBool(IntValue As Long) As Boolean
    IntToBool = (IntValue = 1)
End Function

Function SafeCBool(value As Variant, DefaultValue As Boolean) As Boolean
    On Error GoTo NotBoolean
    SafeCBool = CBool(value)
    Exit Function
    
NotBoolean:
    SafeCBool = DefaultValue
End Function

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

Sub TestBooleanConversion()
    Dim sheet As Worksheet, Name As String, BoolValue As Boolean
    Set sheet = ThisWorkbook.ActiveSheet
    Name = "BoolTest"
    
    ' ***** Getting
    ' Checks that boolean is correctly detected or not, and check value if boolean is found
    
    ' Non boolean
    SetNameOnSheet Name, "abc", sheet
    Debug.Assert Not GetNamedBooleanIfExists(sheet, Name, BoolValue)
    SetNameOnSheet Name, 1.23, sheet
    Debug.Assert Not GetNamedBooleanIfExists(sheet, Name, BoolValue)
    
    ' T/F strings
    SetNameOnSheet Name, "false", sheet
    Debug.Assert GetNamedBooleanIfExists(sheet, Name, BoolValue)
    Debug.Assert Not BoolValue
    SetNameOnSheet Name, "true", sheet
    Debug.Assert GetNamedBooleanIfExists(sheet, Name, BoolValue)
    Debug.Assert BoolValue
    SetNameOnSheet Name, "=FALSE", sheet
    Debug.Assert GetNamedBooleanIfExists(sheet, Name, BoolValue)
    Debug.Assert Not BoolValue
    SetNameOnSheet Name, "=TRUE", sheet
    Debug.Assert GetNamedBooleanIfExists(sheet, Name, BoolValue)
    Debug.Assert BoolValue
    
    ' Integer values
    SetNameOnSheet Name, 0, sheet
    Debug.Assert GetNamedBooleanIfExists(sheet, Name, BoolValue)
    Debug.Assert Not BoolValue
    SetNameOnSheet Name, 1, sheet
    Debug.Assert GetNamedBooleanIfExists(sheet, Name, BoolValue)
    Debug.Assert BoolValue
    SetNameOnSheet Name, 2, sheet
    Debug.Assert GetNamedBooleanIfExists(sheet, Name, BoolValue)
    Debug.Assert Not BoolValue
    
    ' ***** Setting
    ' Checks that a bool can be interpreted correctly after setting it
    SetBooleanNameOnSheet Name, False, sheet
    Debug.Assert GetNamedBooleanIfExists(sheet, Name, BoolValue)
    Debug.Assert Not BoolValue
    SetBooleanNameOnSheet Name, True, sheet
    Debug.Assert GetNamedBooleanIfExists(sheet, Name, BoolValue)
    Debug.Assert BoolValue
    
    Dim DefaultBool As Boolean, value As Variant
    For Each value In Array(True, False)
        DefaultBool = CBool(value)
        
        ' ***** Getting with defaults
        ' Should return DefaultBool on any name that can't be interpreted as bool, otherwise returns expected result
        
        ' Non boolean
        SetNameOnSheet Name, "abc", sheet
        Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = DefaultBool
        SetNameOnSheet Name, 1.23, sheet
        Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = DefaultBool
        
        ' T/F strings
        SetNameOnSheet Name, "false", sheet
        Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = False
        SetNameOnSheet Name, "true", sheet
        Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = True
        SetNameOnSheet Name, "=FALSE", sheet
        Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = False
        SetNameOnSheet Name, "=TRUE", sheet
        Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = True
        SetNameOnSheet Name, "=FAUX", sheet
        Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = DefaultBool
        SetNameOnSheet Name, "=VRAI", sheet
        Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = DefaultBool
        
        ' Integer values
        SetNameOnSheet Name, 0, sheet
        Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = False
        SetNameOnSheet Name, 1, sheet
        Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = True
        SetNameOnSheet Name, 2, sheet
        Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = False
        
        ' ***** Test SafeCBool
        ' Should return DefaultBool on any conversion where CBool would fail, otherwise returns output of CBool
        Debug.Assert SafeCBool("abc", DefaultBool) = DefaultBool
        Debug.Assert SafeCBool("1.23", DefaultBool) = True
        Debug.Assert SafeCBool(1.23, DefaultBool) = True
        Debug.Assert SafeCBool(False, DefaultBool) = False
        Debug.Assert SafeCBool(True, DefaultBool) = True
        Debug.Assert SafeCBool("false", DefaultBool) = False
        Debug.Assert SafeCBool("true", DefaultBool) = True
        Debug.Assert SafeCBool("FALSE", DefaultBool) = False
        Debug.Assert SafeCBool("TRUE", DefaultBool) = True
        Debug.Assert SafeCBool("FAUX", DefaultBool) = DefaultBool
        Debug.Assert SafeCBool("VRAI", DefaultBool) = DefaultBool
        Debug.Assert SafeCBool("0", DefaultBool) = False
        Debug.Assert SafeCBool("1", DefaultBool) = True
        Debug.Assert SafeCBool("2", DefaultBool) = True
        Debug.Assert SafeCBool(0, DefaultBool) = False
        Debug.Assert SafeCBool(1, DefaultBool) = True
        Debug.Assert SafeCBool(2, DefaultBool) = True
        Debug.Assert SafeCBool(BoolToInt(False), DefaultBool) = False
        Debug.Assert SafeCBool(BoolToInt(True), DefaultBool) = True
    Next value
End Sub
