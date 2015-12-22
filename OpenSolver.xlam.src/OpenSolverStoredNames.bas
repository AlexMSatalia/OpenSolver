Attribute VB_Name = "OpenSolverStoredNames"
Option Explicit

Const SolverPrefix As String = "solver_"

Function SheetNameExistsInWorkbook(sheet As Worksheet, Name As String, Optional o As Object) As Boolean
    SheetNameExistsInWorkbook = NameExistsInWorkbook(sheet.Parent, EscapeSheetName(sheet) & Name, o)
End Function

Function NameExistsInWorkbook(book As Workbook, Name As String, Optional o As Object) As Boolean
          ' WARNING: If the name has a sheet prefix, eg Sheet1!OpenSolverCBCParameters, then this will NOT find the range
          ' if the range has been defined globally (which happens when the user defines a name if that name exists only once)
142       On Error Resume Next
143       Set o = book.Names(Name)
144       NameExistsInWorkbook = (Err.Number = 0)
End Function

Function GetSheetNameValueIfExists(sheet As Worksheet, theName As String, ByRef value As String) As Boolean
    GetSheetNameValueIfExists = GetNameValueIfExists(sheet.Parent, EscapeSheetName(sheet) & theName, value)
End Function

Function GetNameValueIfExists(book As Workbook, theName As String, ByRef value As String) As Boolean
          ' See http://www.cpearson.com/excel/DefinedNames.aspx
          Dim s As String
          Dim HasRef As Boolean
          Dim r As Range
          Dim NM As Name
          
117       On Error Resume Next
118       Set NM = book.Names(theName)
119       If Err.Number <> 0 Then ' Name does not exist
120           value = ""
121           GetNameValueIfExists = False
122           Exit Function
123       End If
          
124       On Error Resume Next
125       Set r = NM.RefersToRange
126       If Err.Number = 0 Then
127           HasRef = True
128       Else
129           HasRef = False
130       End If
131       If HasRef = True Then
132           value = r.value
133       Else
134           s = NM.RefersTo
135           If StrComp(Mid(s, 2, 1), Chr(34), vbBinaryCompare) = 0 Then
                  ' text constant
136               value = Mid(s, 3, Len(s) - 3)
137           Else
                  ' numeric contant (AJM: or Formula)
138               value = Mid(s, 2)
139           End If
140       End If
141       GetNameValueIfExists = True
End Function

Function GetNameRefersToIfExists(book As Workbook, Name As String, RefersTo As String) As Boolean
          ' WARNING: If the name has a sheet prefix, eg Sheet1!OpenSolverCBCParameters, then this will NOT find the range
          ' if the range has been defined globally (which happens when the user defines a name if that name exists only once)
145       On Error Resume Next
146       RefersTo = book.Names(Name).RefersTo
147       GetNameRefersToIfExists = (Err.Number = 0)
End Function

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
182       Set r = NM.RefersToRange
183       If Err.Number = 0 Then
184           IsRange = True
185       Else
186           IsRange = False
187           ' String will be of form: "5", or "Sheet1!#REF!" or "#REF!$A$1" or "Sheet1!$A$1/4+Sheet1!$A$3"
189           If Right(RefersTo, 6) = "!#REF!" Or Left(RefersTo, 5) = "#REF!" Then
191               RangeRefersToError = True
192           Else
              ' If StrComp(Mid(S, 2, 1), Chr(34), vbBinaryCompare) = 0 Then
                  ' text constant
              '    S = Mid(S, 3, Len(S) - 3)
              'Else
                  ' numeric contant (or possibly a string? We ignore strings - Solver rejects them as invalid on entry)
                  ' The following Pearson code FAILS because "Value=RefersTo" applies regional settings, but Names are always stored as strings containing values in US settings (with no regionalisation)
                  ' value = RefersTo
                  ' If Err.Number = 13 Then
                  '    RefersToFormula = True
                  ' End If
                  
                  ' Test for a numeric constant, in US format
193               If IsAmericanNumber(RefersTo) Then
194                   value = Val(RefersTo)   ' Force a conversion to a number using Val which uses US settings (no regionalisation)
195               Else
196                   RefersToFormula = True
197               End If
198           End If
199       End If
End Sub

Function GetNamedRangeIfExists(book As Workbook, Name As String, r As Range) As Boolean
          ' WARNING: If the name has a sheet prefix, eg Sheet1!OpenSolverCBCParameters, then this will NOT find the range
          ' if the range has been defined globally (which happens when the user defines a name if that name exists only once)
148       On Error Resume Next
149       Set r = book.Names(Name).RefersToRange
150       GetNamedRangeIfExists = (Err.Number = 0)
End Function

Function GetNamedRangeIfExistsOnSheet(sheet As Worksheet, Name As String, r As Range) As Boolean
          GetActiveSheetIfMissing sheet
          ' This finds a named range (either local or global) if it exists, and if it refers to the specified sheet.
          ' It will not find a globally defined name
151       On Error Resume Next
152       Set r = sheet.Range(Name)   ' This will return either a local or globally defined named range, that must refer to the specified sheet. OTherwise there is an error
153       GetNamedRangeIfExistsOnSheet = Err.Number = 0
End Function

Function GetNamedNumericValueIfExists(sheet As Worksheet, Name As String, value As Double) As Boolean
          ' Get a named range that must contain a double value or the form "=12.34" or "=12" etc, with no spaces
          Dim IsRange As Boolean, r As Range, RefersToFormula As Boolean, RangeRefersToError As Boolean, RefersTo As String, IsMissing As Boolean
154       GetSheetNameAsValueOrRange sheet, Name, IsMissing, IsRange, r, RefersToFormula, RangeRefersToError, RefersTo, value
155       GetNamedNumericValueIfExists = Not IsMissing And Not IsRange And Not RefersToFormula And Not RangeRefersToError
End Function

Function GetNamedIntegerIfExists(sheet As Worksheet, Name As String, IntegerValue As Long) As Boolean
          ' Get a named range that must contain an integer value
          Dim value As Double
156       If GetNamedNumericValueIfExists(sheet, Name, value) Then
157           IntegerValue = Int(value)
158           GetNamedIntegerIfExists = IntegerValue = value
159       Else
160           GetNamedIntegerIfExists = False
161       End If
End Function

Function GetNamedIntegerWithDefault(Name As String, Optional sheet As Worksheet, Optional DefaultValue As Long = 0) As Long
    GetActiveSheetIfMissing sheet
    
    Dim value As String
    If Not GetSheetNameValueIfExists(sheet, Name, value) Then GoTo SetDefault
    On Error GoTo SetDefault
    GetNamedIntegerWithDefault = CLng(value)
    Exit Function
    
SetDefault:
    GetNamedIntegerWithDefault = DefaultValue
    SetIntegerNameOnSheet Name, GetNamedIntegerWithDefault, sheet
End Function

Function GetNamedDoubleWithDefault(Name As String, Optional sheet As Worksheet, Optional DefaultValue As Double = 0) As Double
    GetActiveSheetIfMissing sheet
    
    Dim value As String
    If Not GetSheetNameValueIfExists(sheet, Name, value) Then GoTo SetDefault
    On Error GoTo SetDefault
    GetNamedDoubleWithDefault = Val(value)
    Exit Function
    
SetDefault:
    GetNamedDoubleWithDefault = DefaultValue
    SetDoubleNameOnSheet Name, GetNamedDoubleWithDefault, sheet
End Function

Function GetNamedBooleanWithDefault(Name As String, Optional sheet As Worksheet, Optional DefaultValue As Boolean = False) As Boolean
    GetActiveSheetIfMissing sheet
    
    Dim value As String
    If Not GetSheetNameValueIfExists(sheet, Name, value) Then GoTo SetDefault
    On Error GoTo SetDefault
    GetNamedBooleanWithDefault = CBool(value)  ' TODO: Check localisation
    Exit Function
    
SetDefault:
    GetNamedBooleanWithDefault = DefaultValue
    SetBooleanNameOnSheet Name, GetNamedBooleanWithDefault, sheet
End Function

Function GetNamedIntegerAsBooleanWithDefault(Name As String, Optional sheet As Worksheet, Optional DefaultValue As Boolean = False) As Boolean
    GetActiveSheetIfMissing sheet
    
    Dim value As Long
    If Not GetNamedIntegerIfExists(sheet, Name, value) Then GoTo SetDefault
    If value <> 1 And value <> 2 Then GoTo SetDefault
    GetNamedIntegerAsBooleanWithDefault = (value = 1)
    Exit Function
    
SetDefault:
    GetNamedIntegerAsBooleanWithDefault = DefaultValue
    SetBooleanAsIntegerNameOnSheet Name, GetNamedIntegerAsBooleanWithDefault, sheet
End Function

Function GetNamedStringIfExists(book As Workbook, Name As String, value As String) As Boolean
' Get a named range that must contain a string value (probably with quotes)
162       If GetNameRefersToIfExists(book, Name, value) Then
163           If Left(value, 2) = "=""" Then ' Remove delimiters and equals in: ="...."
164               value = Mid(value, 3, Len(value) - 3)
165           ElseIf Left(value, 1) = "=" Then
166               value = Mid(value, 2)
167           End If
168           GetNamedStringIfExists = True
169       Else
170           GetNamedStringIfExists = False
171       End If
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

Sub SetAnyMissingDefaultSolverOptions(Optional sheet As Worksheet)
          ' We set all the default values, as per Solver in Excel 2007, but with some changes. This ensures Solver does not delete the few values we actually use
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim SolverOptions() As Variant, SolverDefaults() As Variant
          SolverOptions = Array("drv", "est", "nwt", "scl", "cvg", "rlx")
          SolverDefaults = Array("1", "1", "1", "2", "0.0001", "2")
          
          GetActiveSheetIfMissing sheet
          
          Dim i As Long, s As String
          For i = LBound(SolverOptions) To UBound(SolverOptions)
              If Not GetSheetNameValueIfExists(sheet, "solver_" & SolverOptions(i), s) Then
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

Public Sub ValidateConstraint(LHSRange As Range, Relation As RelationConsts, Optional RHSRange As Range, Optional RHSFormula As String, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    
    If LHSRange.Areas.Count > 1 Then
        Err.Raise OpenSolver_ModelError, Description:="Left-hand-side of constraint must have only one area."
    End If
    
    If RelationHasRHS(Relation) Then
        If Not RHSRange Is Nothing Then
            If RHSRange.Count > 1 And RHSRange.Count <> LHSRange.Count Then
                Err.Raise OpenSolver_ModelError, Description:="Right-hand-side of constraint has more than one cell, and does not match the number of cells on the left-hand-side."
            End If
        Else
            ' Try to convert it to a US locale string internally
            Dim internalRHS As String
            internalRHS = ConvertFromCurrentLocale(RHSFormula)

            ' Can we evaluate this function or constant?
            Dim varReturn As Variant
            varReturn = sheet.Evaluate(internalRHS) ' Must be worksheet.evaluate to get references to names local to the sheet
            If VBA.VarType(varReturn) = vbError Then
                Err.Raise OpenSolver_ModelError, Description:="The formula or value for the RHS is not valid. Please check and try again."
            End If

            ' Convert any cell references to absolute
            If Left(internalRHS, 1) <> "=" Then internalRHS = "=" & internalRHS
            varReturn = Application.ConvertFormula(internalRHS, FromReferenceStyle:=xlA1, ToReferenceStyle:=xlA1, ToAbsolute:=xlAbsolute)

            If (VBA.VarType(varReturn) = vbError) Then
                ' Its valid, but couldn't convert to standard form, probably because not A1... just leave it
            Else
                ' Always comes back with a = at the start
                ' Unfortunately, return value will have wrong locale...
                ' But not much can be done with that?
                internalRHS = Mid(varReturn, 2, Len(varReturn))
            End If
            RHSFormula = internalRHS
        End If
        
    Else
        If Not RHSRange Is Nothing Or _
           (RHSFormula <> "" And RHSFormula <> "integer" And RHSFormula <> "binary" And RHSFormula <> "alldiff") Then
            Err.Raise OpenSolver_ModelError, Description:="No right-hand-side is permitted for this relation"
        End If
    End If
End Sub

Sub ValidateParametersRange(ParametersRange As Range)
    If ParametersRange Is Nothing Then Exit Sub
    If ParametersRange.Areas.Count > 1 Or ParametersRange.Columns.Count <> 2 Then
        Err.Raise OpenSolver_SolveError, Description:="The Extra Solver Parameters range must be a single two-column table of keys and values."
    End If
End Sub

