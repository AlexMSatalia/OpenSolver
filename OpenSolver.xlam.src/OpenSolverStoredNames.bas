Attribute VB_Name = "OpenSolverStoredNames"
Option Explicit

Const SolverPrefix As String = "solver_"
Const TempPrefix As String = "OpenSolver_Temp"

Sub GetSheetNameAsValueOrRange(sheet As Worksheet, theName As String, IsMissing As Boolean, IsRange As Boolean, r As Range, RefersToFormula As Boolean, RangeRefersToError As Boolean, RefersTo As String, value As Double)
          ' Wrapper for a sheet-prefixed defined name
1         GetNameAsValueOrRange sheet.Parent, EscapeSheetName(sheet) & theName, IsMissing, IsRange, r, RefersToFormula, RangeRefersToError, RefersTo, value
End Sub

Sub GetNameAsValueOrRange(book As Workbook, theName As String, IsMissing As Boolean, IsRange As Boolean, r As Range, RefersToFormula As Boolean, RangeRefersToError As Boolean, RefersTo As String, value As Double)
      ' See http://www.cpearson.com/excel/DefinedNames.aspx, but see below for internationalisation problems with this code
1         RangeRefersToError = False
2         RefersToFormula = False
          ' Dim r As Range
          Dim NM As Name
3         On Error Resume Next
4         Set NM = book.Names(theName)
5         If Err.Number <> 0 Then
6             IsMissing = True
7             Exit Sub
8         End If
9         IsMissing = False
10        RefersTo = Mid(NM.RefersTo, 2)
11        On Error Resume Next
12        Set r = GetRefersToRange(RefersTo)
13        If Err.Number = 0 Then
14            IsRange = True
15        Else
16            IsRange = False
              ' String will be of form: "5", or "Sheet1!#REF!" or "#REF!$A$1" or "Sheet1!$A$1/4+Sheet1!$A$3"
17            If Right(RefersTo, 6) = "!#REF!" Or Left(RefersTo, 5) = "#REF!" Then
18                RangeRefersToError = True
19            Else
                  ' Test for a numeric constant, in US format
20                If IsAmericanNumber(RefersTo) Then
21                    value = Val(RefersTo)   ' Force a conversion to a number using Val which uses US settings (no regionalisation)
22                Else
23                    RefersToFormula = True
24                End If
25            End If
26        End If
End Sub

Function GetNamedDoubleIfExists(sheet As Worksheet, Name As String, DoubleValue As Double) As Boolean
          Dim IsRange As Boolean, r As Range, RefersToFormula As Boolean, RangeRefersToError As Boolean, RefersTo As String, IsMissing As Boolean
1         GetSheetNameAsValueOrRange sheet, Name, IsMissing, IsRange, r, RefersToFormula, RangeRefersToError, RefersTo, DoubleValue
2         GetNamedDoubleIfExists = Not IsMissing And Not IsRange And Not RefersToFormula And Not RangeRefersToError
End Function

Function GetNamedIntegerIfExists(sheet As Worksheet, Name As String, IntegerValue As Long) As Boolean
          Dim DoubleValue As Double
1         If GetNamedDoubleIfExists(sheet, Name, DoubleValue) Then
2             IntegerValue = CLng(DoubleValue)
3             GetNamedIntegerIfExists = (IntegerValue = DoubleValue)
4         End If
End Function

Function GetNamedBooleanIfExists(sheet As Worksheet, Name As String, BooleanValue As Boolean) As Boolean
          Dim IsRange As Boolean, r As Range, RefersToFormula As Boolean, RangeRefersToError As Boolean, RefersTo As String, IsMissing As Boolean, value As Double
1         GetSheetNameAsValueOrRange sheet, Name, IsMissing, IsRange, r, RefersToFormula, RangeRefersToError, RefersTo, value

2         If Not IsMissing And Not IsRange And Not RangeRefersToError Then
3             If RefersToFormula Then
                  ' It's a string, probably "=TRUE" or "=FALSE"
                  ' We handle this conversion for legacy reasons, in versions 2.8.2 and earlier OpenSolver bool options
                  ' were stored as strings "=TRUE" or "=FALSE". This was changed to 0/1 to resolve locale issues with CBool.
                  ' e.g. In French, the strings would be "=VRAI" and "=FAUX", which didn't always work with CBool.
4                 On Error GoTo NotBoolean
5                 BooleanValue = CBool(RefersTo)
6                 GetNamedBooleanIfExists = True
7             Else
                  ' It's a value
8                 If CLng(value) = value Then
                      ' It's integer, we can interpret that as a bool
9                     GetNamedBooleanIfExists = True
10                    BooleanValue = IntToBool(CLng(value))
11                Else
                      ' It's a double, so not a boolean
12                    GetNamedBooleanIfExists = False
13                End If
14            End If
15        Else
NotBoolean:
16            GetNamedBooleanIfExists = False
17        End If
End Function

Function GetNamedStringIfExists(sheet As Worksheet, Name As String, StringValue As String) As Boolean
          Dim IsRange As Boolean, r As Range, RefersToFormula As Boolean, RangeRefersToError As Boolean, value As Double, IsMissing As Boolean
1         GetSheetNameAsValueOrRange sheet, Name, IsMissing, IsRange, r, RefersToFormula, RangeRefersToError, StringValue, value
          
          ' Remove any quotes
2         If Left(StringValue, 1) = """" Then StringValue = Mid(StringValue, 2, Len(StringValue) - 2)
          
3         GetNamedStringIfExists = Not IsMissing And Not IsRange And Not RangeRefersToError
End Function

Function GetNamedDoubleWithDefault(sheet As Worksheet, Name As String, DefaultValue As Double) As Double
1         If Not GetNamedDoubleIfExists(sheet, Name, GetNamedDoubleWithDefault) Then
2             GetNamedDoubleWithDefault = DefaultValue
3             SetDoubleNameOnSheet Name, GetNamedDoubleWithDefault, sheet
4         End If
End Function

Function GetNamedIntegerWithDefault(sheet As Worksheet, Name As String, DefaultValue As Long) As Long
1         If Not GetNamedIntegerIfExists(sheet, Name, GetNamedIntegerWithDefault) Then
2             GetNamedIntegerWithDefault = DefaultValue
3             SetIntegerNameOnSheet Name, GetNamedIntegerWithDefault, sheet
4         End If
End Function

Function GetNamedBooleanWithDefault(sheet As Worksheet, Name As String, DefaultValue As Boolean) As Boolean
1         If Not GetNamedBooleanIfExists(sheet, Name, GetNamedBooleanWithDefault) Then
2             GetNamedBooleanWithDefault = DefaultValue
3             SetBooleanNameOnSheet Name, GetNamedBooleanWithDefault, sheet
4         End If
End Function

Function GetNamedStringWithDefault(sheet As Worksheet, Name As String, DefaultValue As String) As String
1         If Not GetNamedStringIfExists(sheet, Name, GetNamedStringWithDefault) Then
2             GetNamedStringWithDefault = DefaultValue
3             SetNameOnSheet Name, GetNamedStringWithDefault, sheet
4         End If
End Function

Sub DeleteNameOnSheet(Name As String, Optional sheet As Worksheet, Optional SolverName As Boolean = False)
1         GetActiveSheetIfMissing sheet
2         Name = EscapeSheetName(sheet) & IIf(SolverName, SolverPrefix, vbNullString) & Name
3         On Error Resume Next
4         sheet.Parent.Names(Name).Delete
doesntExist:
End Sub

Sub SetNameOnSheet(Name As String, value As Variant, Optional sheet As Worksheet, Optional SolverName As Boolean = False)
      ' If a key exists we can just add it (http://www.cpearson.com/Excel/DefinedNames.aspx)
1         GetActiveSheetIfMissing sheet
          Dim FullName As String ' Don't mangle the value of Name!
2         FullName = EscapeSheetName(sheet) & IIf(SolverName, SolverPrefix, vbNullString) & Name
3         sheet.Parent.Names.Add FullName, value, False
End Sub

Sub SetNamedRangeIfExists(ByVal Name As String, ByRef RangeToSet As Range, Optional sheet As Worksheet, Optional SolverName As Boolean = False)
1         If RangeToSet Is Nothing Then
2             DeleteNameOnSheet Name, sheet, SolverName
3         Else
4             SetNamedRangeOnSheet Name, RangeToSet, sheet, SolverName
5         End If
End Sub

Sub SetNamedRangeOnSheet(Name As String, value As Range, Optional sheet As Worksheet, Optional SolverName As Boolean = False)
1         SetNameOnSheet Name, value, sheet, SolverName
End Sub

Sub SetIntegerNameOnSheet(Name As String, value As Long, Optional sheet As Worksheet, Optional SolverName As Boolean = False)
1         SetDoubleNameOnSheet Name, CDbl(value), sheet, SolverName
End Sub

Sub SetDoubleNameOnSheet(Name As String, value As Double, Optional sheet As Worksheet, Optional SolverName As Boolean = False)
1         SetNameOnSheet Name, value, sheet, SolverName
End Sub

Sub SetBooleanNameOnSheet(Name As String, value As Boolean, Optional sheet As Worksheet, Optional SolverName As Boolean = False)
1         SetIntegerNameOnSheet Name, BoolToInt(value), sheet, SolverName
End Sub

Sub SetRefersToNameOnSheet(Name As String, value As String, Optional sheet As Worksheet, Optional SolverName As Boolean = False)
1         If Len(value) > 0 Then
2             SetNameOnSheet Name, AddEquals(value), sheet, SolverName
3         Else
4             DeleteNameOnSheet Name, sheet, SolverName
5         End If
End Sub

Function BoolToInt(BoolValue As Boolean) As Long
1         BoolToInt = IIf(BoolValue, 1, 0)
End Function

Function IntToBool(IntValue As Long) As Boolean
1         IntToBool = (IntValue = 1)
End Function

Function SafeCBool(value As Variant, DefaultValue As Boolean) As Boolean
1         On Error GoTo NotBoolean
2         SafeCBool = CBool(value)
3         Exit Function
          
NotBoolean:
4         SafeCBool = DefaultValue
End Function

Sub SetAnyMissingDefaultSolverOptions(sheet As Worksheet)
                ' We set all the default values, as per Solver in Excel 2007, but with some changes. This ensures Solver does not delete the few values we actually use
                Dim RaiseError As Boolean
1               RaiseError = False
2               On Error GoTo ErrorHandler

                Dim SolverOptions() As Variant, SolverDefaults() As Variant
3               SolverOptions = Array("drv", "est", "nwt", "scl", "cvg", "rlx")
4               SolverDefaults = Array("1", "1", "1", "2", "0.0001", "2")
                
5               GetActiveSheetIfMissing sheet
                
                Dim i As Long, value As Double
6               For i = LBound(SolverOptions) To UBound(SolverOptions)
7                   If Not GetNamedDoubleIfExists(sheet, "solver_" & SolverOptions(i), value) Then
8                       SetNameOnSheet CStr(SolverOptions(i)), "=" & SolverDefaults(i), sheet, True
9                   End If
10              Next i

ExitSub:
11              If RaiseError Then RethrowError
12              Exit Sub

ErrorHandler:
13              If Not ReportError("OpenSolverStoredNames", "SetAnyMissingDefaultExcel2007SolverOptions") Then Resume
14              RaiseError = True
15              GoTo ExitSub
End Sub

Function GetRefersToRange(RefersTo As String) As Range
1         If Len(RefersTo) = 0 Then Exit Function
          
          ' Add the text as a name and retrieve the corresponding sanitised range
          ' We can't use .RefersToRange as this is broken in non-US locales
          ' Instead we use sheet.Range()
2         On Error GoTo InvalidRange
3         Set GetRefersToRange = Application.Range(RefersTo)
4         Exit Function
          
InvalidRange:
5         RethrowError Err
End Function

Function RefEditToRefersTo(RefEditText As String) As String
1         If Len(RefEditText) = 0 Then Exit Function
          
          ' Try to evaluate the formula and catch `Error 2015` for an invalid formula
          ' ============
          ' IMPORTANT: We can't catch this error using an error handler!
          ' This is sometimes run inside a RefEdit event, in which the
          ' error handlers don't work properly. We have to use a method of
          ' detecting the invalid formula that never throws any error.
          ' ============
          Dim varReturn As Variant
2         varReturn = Application.Evaluate(AddEquals(RefEditText))
3         If VarType(varReturn) = vbError Then
4             If CLng(varReturn) = 2015 Then
5                 RefEditToRefersTo = RefEditText
6                 Exit Function
7             End If
8         End If

          ' Add the text as a name and retrieve the sanitised RefersTo
9         On Error GoTo DeleteName
          Dim n As Name
10        Set n = ActiveWorkbook.Names.Add(TempPrefix, AddEquals(Trim(RefEditText)))
11        RefEditToRefersTo = Mid(n.RefersTo, 2)

DeleteName:
          ' Try deleting name so we don't leave it lying around
12        On Error Resume Next
13        n.Delete
End Function

Function RangeToRefersTo(ConvertRange As Range) As String
1         If ConvertRange Is Nothing Then Exit Function
          
          ' Add the range and retrieve the sanitised refers to
2         On Error GoTo DeleteName
          Dim n As Name
3         Set n = ActiveWorkbook.Names.Add(TempPrefix, ConvertRange)
4         RangeToRefersTo = Mid(n.RefersTo, 2)

DeleteName:
          ' Try deleting name so we don't leave it lying around
5         On Error Resume Next
6         n.Delete
End Function

Sub TestBooleanConversion()
          Dim sheet As Worksheet, Name As String, BoolValue As Boolean
1         Set sheet = ThisWorkbook.ActiveSheet
2         Name = "BoolTest"
          
          ' ***** Getting
          ' Checks that boolean is correctly detected or not, and check value if boolean is found
          
          ' Non boolean
3         SetNameOnSheet Name, "abc", sheet
4         Debug.Assert Not GetNamedBooleanIfExists(sheet, Name, BoolValue)
5         SetNameOnSheet Name, 1.23, sheet
6         Debug.Assert Not GetNamedBooleanIfExists(sheet, Name, BoolValue)
          
          ' T/F strings
7         SetNameOnSheet Name, "false", sheet
8         Debug.Assert GetNamedBooleanIfExists(sheet, Name, BoolValue)
9         Debug.Assert Not BoolValue
10        SetNameOnSheet Name, "true", sheet
11        Debug.Assert GetNamedBooleanIfExists(sheet, Name, BoolValue)
12        Debug.Assert BoolValue
13        SetNameOnSheet Name, "=FALSE", sheet
14        Debug.Assert GetNamedBooleanIfExists(sheet, Name, BoolValue)
15        Debug.Assert Not BoolValue
16        SetNameOnSheet Name, "=TRUE", sheet
17        Debug.Assert GetNamedBooleanIfExists(sheet, Name, BoolValue)
18        Debug.Assert BoolValue
          
          ' Integer values
19        SetNameOnSheet Name, 0, sheet
20        Debug.Assert GetNamedBooleanIfExists(sheet, Name, BoolValue)
21        Debug.Assert Not BoolValue
22        SetNameOnSheet Name, 1, sheet
23        Debug.Assert GetNamedBooleanIfExists(sheet, Name, BoolValue)
24        Debug.Assert BoolValue
25        SetNameOnSheet Name, 2, sheet
26        Debug.Assert GetNamedBooleanIfExists(sheet, Name, BoolValue)
27        Debug.Assert Not BoolValue
          
          ' ***** Setting
          ' Checks that a bool can be interpreted correctly after setting it
28        SetBooleanNameOnSheet Name, False, sheet
29        Debug.Assert GetNamedBooleanIfExists(sheet, Name, BoolValue)
30        Debug.Assert Not BoolValue
31        SetBooleanNameOnSheet Name, True, sheet
32        Debug.Assert GetNamedBooleanIfExists(sheet, Name, BoolValue)
33        Debug.Assert BoolValue
          
          Dim DefaultBool As Boolean, value As Variant
34        For Each value In Array(True, False)
35            DefaultBool = CBool(value)
              
              ' ***** Getting with defaults
              ' Should return DefaultBool on any name that can't be interpreted as bool, otherwise returns expected result
              
              ' Non boolean
36            SetNameOnSheet Name, "abc", sheet
37            Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = DefaultBool
38            SetNameOnSheet Name, 1.23, sheet
39            Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = DefaultBool
              
              ' T/F strings
40            SetNameOnSheet Name, "false", sheet
41            Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = False
42            SetNameOnSheet Name, "true", sheet
43            Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = True
44            SetNameOnSheet Name, "=FALSE", sheet
45            Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = False
46            SetNameOnSheet Name, "=TRUE", sheet
47            Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = True
48            SetNameOnSheet Name, "=FAUX", sheet
49            Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = DefaultBool
50            SetNameOnSheet Name, "=VRAI", sheet
51            Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = DefaultBool
              
              ' Integer values
52            SetNameOnSheet Name, 0, sheet
53            Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = False
54            SetNameOnSheet Name, 1, sheet
55            Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = True
56            SetNameOnSheet Name, 2, sheet
57            Debug.Assert GetNamedBooleanWithDefault(sheet, Name, DefaultBool) = False
              
              ' ***** Test SafeCBool
              ' Should return DefaultBool on any conversion where CBool would fail, otherwise returns output of CBool
58            Debug.Assert SafeCBool("abc", DefaultBool) = DefaultBool
59            Debug.Assert SafeCBool("1.23", DefaultBool) = True
60            Debug.Assert SafeCBool(1.23, DefaultBool) = True
61            Debug.Assert SafeCBool(False, DefaultBool) = False
62            Debug.Assert SafeCBool(True, DefaultBool) = True
63            Debug.Assert SafeCBool("false", DefaultBool) = False
64            Debug.Assert SafeCBool("true", DefaultBool) = True
65            Debug.Assert SafeCBool("FALSE", DefaultBool) = False
66            Debug.Assert SafeCBool("TRUE", DefaultBool) = True
67            Debug.Assert SafeCBool("FAUX", DefaultBool) = DefaultBool
68            Debug.Assert SafeCBool("VRAI", DefaultBool) = DefaultBool
69            Debug.Assert SafeCBool("0", DefaultBool) = False
70            Debug.Assert SafeCBool("1", DefaultBool) = True
71            Debug.Assert SafeCBool("2", DefaultBool) = True
72            Debug.Assert SafeCBool(0, DefaultBool) = False
73            Debug.Assert SafeCBool(1, DefaultBool) = True
74            Debug.Assert SafeCBool(2, DefaultBool) = True
75            Debug.Assert SafeCBool(BoolToInt(False), DefaultBool) = False
76            Debug.Assert SafeCBool(BoolToInt(True), DefaultBool) = True
77        Next value
End Sub
