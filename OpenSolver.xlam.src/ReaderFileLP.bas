Attribute VB_Name = "ReaderFileLP"
Option Explicit

Function GetHeaderArray() As Variant
1         GetHeaderArray = Array( _
              Array("minimize", "minimise", "min", "maximize", "maximise", "max"), _
              Array("subject to", "s.t.", "st", "st.", "such that"), _
              Array("bounds", "bound"), _
              Array("integers", "integer", "int", "generals", "general", "gen", "binary", "binaries", "bin"), _
              Array("semis", "semi", "semi-continuous"), _
              Array("sos"), _
              Array("end") _
              )
End Function

Function GetSenseArray() As Variant
1         GetSenseArray = Array("<", ">", "=")
End Function

Function GetOperatorArray() As Variant
1         GetOperatorArray = Array("+", "-") ' And probably others
End Function

Function ReadLPFile(FileName As String, Optional ByRef sheet As Worksheet, Optional MinimiseUserInteraction As Boolean)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim ModelText As Variant
          Dim Objective As Variant, Constraints As Variant, Bounds As Variant, VariableTypes As Variant, SemiContinuous As Variant, SOS As Variant

2         ModelText = ParseSections(FileName) ' Read in file and clean up whitespace/comments
3         Objective = SetObjectiveFunction(ModelText)
4         Constraints = SetConstraints(ModelText)
5         Bounds = SetBounds(ModelText)
6         VariableTypes = SetVariableTypes(ModelText)
          SemiContinuous = SetSemiCon(ModelText)
          SOS = SetSOS(ModelText)
7         WriteToSheet FileName, Objective, Constraints, Bounds, VariableTypes, SemiContinuous, SOS, sheet
8         WriteToModel Objective, Constraints, VariableTypes

          ShowSolverModel ActiveSheet
          
          If (UBound(SemiContinuous) >= 0 Or UBound(SOS) >= 0) And MinimiseUserInteraction = False Then
              Dim WarningMessage As String
              WarningMessage = "WARNING: The LP file imported contains " & _
                               IIf(UBound(SemiContinuous) >= 0, "semi-continuous variables", "") & _
                               IIf(UBound(SemiContinuous) >= 0 And UBound(SOS) >= 0, " and ", "") & _
                               IIf(UBound(SOS) >= 0, "special ordered sets", "") & _
                               ". These have been printed to the sheet, but have not been written to the model nor supported during model solve."
              MsgBox WarningMessage, , "OpenSolver Warning"
          End If
          
ExitSub:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "ReadLPFile") Then Resume
          RaiseError = True
          GoTo ExitSub
End Function

Function ParseSections(FileName As String) As Variant
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim FileLine As String, ModelStr As String, LineCount As Integer
          Dim NonLinearSense As Variant
          NonLinearSense = Array("*", "^", "[", "]")
          LineCount = 0
          
          Dim fso As Object, InputFile As Variant
          Set fso = CreateObject("Scripting.FileSystemObject")
          
1         Set InputFile = fso.OpenTextFile(FileName, 1)
          
2         Dim EndOfLine As Boolean
          Do While InputFile.AtEndOfStream <> True
          
              LineCount = LineCount + 1
3             EndOfLine = False
4             FileLine = InputFile.ReadLine
LineCheck:
              ' Search and delete comments
5             RemoveComments FileLine
6             FileLine = Trim(FileLine)
              
              Dim NLSense As Variant
              For Each NLSense In NonLinearSense
                  If InStr(FileLine, NLSense) > 0 Then
                      RaiseUserError "Non-linearity detected on line " & LineCount & _
                                     ". This feature does not currently support non-linear LP files. " & _
                                     "The non-linear line is printed below:" & vbCrLf & vbCrLf & FileLine
                  End If
              Next NLSense
              
              ' Line ends occur in two situations
              ' 1: Headers are on their own line
              ' 2: Constraints and bounds are on their own lines
              
              ' Constraints and bounds end with inequalities or are free
              Dim sense As Variant, Operator As Variant
              
7             If InStr(FileLine, "free") > 0 Then
8                 EndOfLine = True
9                 For Each Operator In GetOperatorArray
10                    If InStr(FileLine, Operator) > 0 Then
11                        EndOfLine = False
12                    End If
13                Next Operator
                  If EndOfLine Then GoTo EOL
14            End If
              
15            For Each sense In GetSenseArray
16                If InStr(FileLine, sense) > 0 Then
17                    EndOfLine = True
18                End If
19            Next sense

              ' Special case: Special ordered sets
              If InStr(FileLine, "::") > 0 Then
                  EndOfLine = True
                  GoTo EOL
              End If

              ' Check for headers
              Dim HeaderVariation As Variant, HeaderArray As Variant
20            HeaderArray = GetHeaderArray
              
              Dim i As Long, LongestVariation As String
21            For i = 0 To UBound(HeaderArray, 1)

22                LongestVariation = ""
23                For Each HeaderVariation In HeaderArray(i)
24                    If InStr(LCase(FileLine), HeaderVariation & " ") = 1 Or LCase(FileLine) = HeaderVariation Then
                          ' We look for exactly the header OR header followed by a space
25                        If Len(HeaderVariation) > Len(LongestVariation) Then
26                            LongestVariation = HeaderVariation
27                        End If
28                        EndOfLine = True
29                    End If
30                Next HeaderVariation

31                If Len(LongestVariation) > 0 Then
                      ' Check for a constraints on the same line as the header
                      ModelStr = ModelStr & vbCrLf & LongestVariation & vbCrLf
33                    If Len(FileLine) - Len(LongestVariation) > 0 Then
                          EndOfLine = False
                          FileLine = Right(Trim(FileLine), Len(FileLine) - Len(LongestVariation))
                          GoTo LineCheck
                      End If
                      GoTo Bypass
34                End If

35            Next i

EOL:
36            If EndOfLine Then
37                ModelStr = ModelStr & FileLine & vbCrLf
38            Else
                  ' Variable types require whitespace between variable names
39                If Len(Trim(FileLine)) <> 0 Then FileLine = " " & FileLine
40                ModelStr = ModelStr & FileLine
41            End If
Bypass:
42        Loop
43        InputFile.Close
          
          ' Clean up final string - trim blank lines and unnecessary whitespace
44        While InStr(ModelStr, "  ") > 0
45            ModelStr = Replace(ModelStr, "  ", " ")
46        Wend
          ModelStr = Trim(ModelStr)

          ' A well formatted .lp file should end with "End" - confirm anyway
          If Right(ModelStr, 5) <> "end" + vbCrLf Then
              ModelStr = ModelStr + vbCrLf + "end"
          End If
       
          Dim strLines As Variant, parsedLines() As Variant
47        strLines = Split(ModelStr, vbCrLf)
48        ReDim parsedLines(UBound(strLines))

          ' Remove lines that only contain spaces
          Dim j As Long
49        j = 0
50        For i = 0 To UBound(strLines)
51            If strLines(i) <> "" And strLines(i) <> " " Then
52                parsedLines(j) = strLines(i)
53                j = j + 1
54            End If
55        Next i
       
56        ReDim Preserve parsedLines(j - 1)
57        ParseSections = parsedLines

ExitFunction:
          If RaiseError Then RethrowError
          InputFile.Close
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "ParseSections") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function WriteToSheet(FileName As String, Objective As Variant, Constraints As Variant, Bounds As Variant, VariableTypes As Variant, SemiContinuous As Variant, SOS As Variant, Optional ByRef sheet As Worksheet)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
1         Application.ScreenUpdating = False

          Dim VariableList As Variant
          VariableList = FindAllVariables(Objective(1), Constraints, Bounds, VariableTypes, SemiContinuous, SOS)

          Dim FileNameShort As String
          FileNameShort = Right(FileName, Len(FileName) - InStrRev(FileName, "\"))
          If sheet Is Nothing Then
              Set sheet = MakeNewSheet(FileNameShort, False)
          Else
              sheet.Activate
              sheet.UsedRange.Delete ' Overwrite existing cells, but keep associated references
          End If
          
          Cells.ColumnWidth = 10 ' To fit MAX_DOUBLE in bounds
3         Cells(1, 1) = FileNameShort
4         Cells(2, 1) = FileName
5         Cells(2, 2) = " " ' Prevent long file name overflowing on to next cell

          Dim nVariables As Long, nConstraints As Long, nSets As Long
7         nVariables = UBound(VariableList)
8         nConstraints = UBound(Constraints)
          nSets = UBound(SOS)

          ' Sort all data by variable
          Dim VariableProperties() As Variant
9         ReDim VariableProperties(nVariables)
          Dim i As Long
10        For i = 0 To nVariables
11            VariableProperties(i) = Array( _
                  VariableList(i), _
                  FindObjCoeff(VariableList(i), Objective(1)), _
                  GenerateAColumn(VariableList(i), Constraints), _
                  FindLowerBound(VariableList(i), Bounds), _
                  FindUpperBound(VariableList(i), Bounds), _
                  FindVariableType(VariableList(i), VariableTypes), _
                  FindSemiCon(VariableList(i), SemiContinuous), _
                  FindSOS(VariableList(i), SOS) _
                  )
12        Next i
          
          Dim VariableType As RelationConsts
          Dim ExtendedAMatrix As Variant
13        ReDim ExtendedAMatrix(11 + nConstraints + nSets, nVariables)
          
          If nSets >= 0 Then
              Dim SetsNameArray() As Variant
              ReDim SetsNameArray(nSets, 0)
          End If
          
          ' Make variable columns
          Dim j As Long, k As Long, l As Long
14        For j = 0 To nVariables

15            ExtendedAMatrix(0, j) = VariableProperties(j)(0) ' Variable name
16            ExtendedAMatrix(1, j) = "0" ' Initial value of variable
17            ExtendedAMatrix(3, j) = VariableProperties(j)(1) ' Objective coefficient
18            SetNamedRangeOnSheet "Variable_" & VariableProperties(j)(0), Cells(2, 3 + j)

19            For k = 0 To nConstraints
20                ExtendedAMatrix(5 + k, j) = VariableProperties(j)(2)(k) ' Constraints
21            Next k

22            ExtendedAMatrix(7 + nConstraints, j) = VariableProperties(j)(3) ' Lower bound
23            ExtendedAMatrix(8 + nConstraints, j) = VariableProperties(j)(4) ' Upper bound

24            VariableType = VariableProperties(j)(5)
25            SetNamedRangeOnSheet "Variable_" & VariableProperties(j)(0) & "_Type", Cells(10 + nConstraints, 3 + j)
26            ExtendedAMatrix(9 + nConstraints, j) = RelationEnumToString(VariableType) ' Variable type

              If VariableProperties(j)(6) Then ExtendedAMatrix(10 + nConstraints, j) = "Yes" ' Semi-continuous
              
              For l = 0 To nSets
                  ExtendedAMatrix(11 + nConstraints + l, j) = VariableProperties(j)(7)(l) ' Special ordered sets
              Next l
              
27        Next j

          Dim AMatrixRange As Range
28        Set AMatrixRange = ActiveSheet.Range(Cells(1, 3), Cells(12 + nConstraints + nSets, 3 + nVariables))
29        AMatrixRange.value = ExtendedAMatrix

          If nSets >= 0 Then
              For i = 0 To nSets
                  SetsNameArray(i, 0) = SOS(i)(0) & "(" & SOS(i)(1) & ")"
              Next i
              Dim SetsRange As Range
              Set SetsRange = ActiveSheet.Range(Cells(12 + nConstraints, 2), Cells(12 + nConstraints + nSets, 2))
              SetsRange.value = SetsNameArray
          End If

          ' Write ranges
30        SetNamedRangeOnSheet "Objective_coefficients", Range(Cells(4, 3), Cells(4, 3 + nVariables))
31        SetNamedRangeOnSheet "Decision_variables", Range(Cells(2, 3), Cells(2, 3 + nVariables))
32        SetNamedRangeOnSheet "Lower_bounds", Range(Cells(8 + nConstraints, 3), Cells(8 + nConstraints, 3 + nVariables))
33        SetNamedRangeOnSheet "Upper_bounds", Range(Cells(9 + nConstraints, 3), Cells(9 + nConstraints, 3 + nVariables))
          
          If nConstraints = -1 Then
              GoTo SkipConstraints
          End If
          
          ' Write everything to the right of the A matrix
          Dim Relation As RelationConsts
          Dim SumArray() As Variant, ConstraintNameArray() As Variant
34        ReDim SumArray(nConstraints, 3)
36        ReDim ConstraintNameArray(nConstraints, 0)
          
37        For i = 0 To nConstraints

38            SetNamedRangeOnSheet "Constraint_" & i, Range(Cells(6 + i, 3), Cells(6 + i, 3 + nVariables))
39            ConstraintNameArray(i, 0) = Constraints(i)(3)

40            Relation = Constraints(i)(1)

44            SumArray(i, 0) = "=SUMPRODUCT(Constraint_" & i & ",Decision_variables)"
45            SumArray(i, 1) = RelationEnumToString(Relation)
46            SumArray(i, 2) = Constraints(i)(2)

              SetNamedRangeOnSheet "Constraint_" & i & "_LHS", Cells(6 + i, 5 + nVariables)
              SetNamedRangeOnSheet "Constraint_" & i & "_Sense", Cells(6 + i, 6 + nVariables)
              SetNamedRangeOnSheet "Constraint_" & i & "_RHS", Cells(6 + i, 7 + nVariables)

47        Next i
        
60        SetNamedRangeOnSheet "ConstraintNames", Range(Cells(6, 2), Cells(6 + nConstraints, 2))
61        Range("ConstraintNames").value = ConstraintNameArray
        
62        SetNamedRangeOnSheet "ConstraintCalcs", Range(Cells(6, 5 + nVariables), Cells(6 + nConstraints, 7 + nVariables))
63        Range("ConstraintCalcs").value = SumArray

SkipConstraints:
64        Cells(4, 5 + nVariables).value = "=SUMPRODUCT(Objective_coefficients, Decision_variables)" & Objective(3)
65        SetNamedRangeOnSheet "Objective_Function", Cells(4, 5 + nVariables)

          ' Write descriptors
66        If Objective(0) = 1 Then
67            Cells(4, 1) = "Maximise:"
68        ElseIf Objective(0) = 2 Then
69            Cells(4, 1) = "Minimise:"
70        Else
71            RaiseGeneralError "Unknown objective sense: " & Objective(0)
72        End If
73        Cells(4, 2) = Objective(2)
74        Cells(6, 1) = "Subject to:"
75        Cells(8 + nConstraints, 1) = "Variable bounds: "
76        Cells(8 + nConstraints, 2) = ">="
77        Cells(9 + nConstraints, 2) = "<="
78        Cells(10 + nConstraints, 1) = "Variable type: "
          If UBound(SemiContinuous) >= 0 Then
              Cells(11 + nConstraints, 1) = "Semi-continuous?"
              Cells(11 + nConstraints, 1).AddComment ("Semi-continuous variables are printed from the file, but not added to the model nor supported during solve.")
              Cells(11 + nConstraints, 1).Comment.Shape.TextFrame.AutoSize = True
          End If
          If nSets >= 0 Then
              ' This leaves a one-line gap if SOS exists but not semi-continuous variables.
              Cells(12 + nConstraints, 1) = "Set weighting: "
              Cells(12 + nConstraints, 1).AddComment ("Special ordered sets are printed from the file, but not added to the model nor supported during solve.")
              Cells(12 + nConstraints, 1).Comment.Shape.TextFrame.AutoSize = True
          End If

79        Range(Cells(3, 1), Cells(12 + nConstraints, 1)).Columns.AutoFit
80        Columns(2).EntireColumn.AutoFit
81        Columns(1).Font.Bold = True
82        Rows(1).Font.Bold = True
83        Cells(2, 1).Font.Bold = False
          
84        Application.ScreenUpdating = True
        
ExitFunction:
          If RaiseError Then RethrowError
          Application.ScreenUpdating = True
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "WriteToSheet") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function SetObjectiveFunction(ModelText As Variant) As Variant
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          ' Returns array: [Objective sense (enum), Parsed formulae (array), Objective name, Constant term (string)]
          ' Parsed formulae: [Operator, Coefficient, Variable]
          
          Dim ObjectiveArray(3) As Variant
          
          ' Objective sense should always be line 1
          Dim HeaderVariation As Variant, HeaderArray As Variant
1         HeaderArray = GetHeaderArray
2         For Each HeaderVariation In HeaderArray(0)
3             If InStr(ModelText(0), HeaderVariation) > 0 Then
4                 ObjectiveArray(0) = ObjectiveSenseStringToEnum(HeaderVariation)
5             End If
6         Next HeaderVariation

          If IsEmpty(ObjectiveArray(0)) Then
              RaiseUserError "Could not find objective sense in the LP file."
          End If

          Dim LHSFormula As String, ObjectiveName As String
7         LHSFormula = ModelText(1)
          ObjectiveName = ""
8         RemoveNames LHSFormula, ObjectiveName
          
          Dim SplitFormula As Variant, ConstantTerm As String, ParsedFormula As Variant
9         SplitFormula = SplitEquationLHS(Trim(LHSFormula))
          ConstantTerm = ""
10        ParsedFormula = ParseEquationLHS(SplitFormula, ConstantTerm)
          
11        ObjectiveArray(1) = ParsedFormula
          ObjectiveArray(2) = ObjectiveName
          ObjectiveArray(3) = ConstantTerm
12        SetObjectiveFunction = ObjectiveArray

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "SetObjectiveFunction") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function SetConstraints(ModelText As Variant, Optional SortConstraints As Boolean = False) As Variant
          ' Returns array: [List of constraints]
          ' Each constraint: [Variables (array), Sense, RHS, Constraint name]
          ' Variables: [List of variables (array)]
          ' List of variables: [Operator, Coefficient, variable name]
          ' To access constraint i, variable j coefficient:
          ' coefficient = ArrayName(i)(0)(j)(1)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim i As Long, j As Long, NumConstraints As Long
1         i = FindHeaderLine(1, ModelText)
2         j = FindNextHeaderLine(1, ModelText)
3         NumConstraints = j - i - 1

          If NumConstraints = 0 Then
              SetConstraints = Array()
              GoTo ExitFunction
          End If
          
          Dim Constraints() As Variant, ConstraintText As String, ConstraintName As String
4         ReDim Constraints(NumConstraints - 1)
          Dim k As Long
5         For k = 0 To NumConstraints - 1

              ConstraintName = ""
6             ConstraintText = ModelText(i + k + 1)
7             RemoveNames ConstraintText, ConstraintName

8             Constraints(k) = ConcatArrays(SplitConstraint(Trim(ConstraintText)), Array(ConstraintName))
9             ConstraintText = Constraints(k)(0)
10            Constraints(k)(0) = ParseEquationLHS(SplitEquationLHS(ConstraintText))
              
11        Next k
          
          If SortConstraints = True Then
              Dim SortedConstraints As Variant
              SortedConstraints = SortConstraintsByIneq(Constraints)
              SetConstraints = SortedConstraints
          Else
12            SetConstraints = Constraints
          End If

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "SetConstraints") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function SetBounds(ModelText As Variant) As Variant
          ' Returns array: [Bound number (array)]
          ' Bound number: [Lower bound, Variable, Upper bound]
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim i As Long, j As Long, NumBounds As Long
1         i = FindHeaderLine(2, ModelText)
2         j = FindNextHeaderLine(2, ModelText)
3         NumBounds = j - i - 1
          If i = -1 Or NumBounds = 0 Then
              ' Bounds section is optional
              SetBounds = Array()
              GoTo ExitFunction
          End If
          
          Dim Bounds() As Variant
4         ReDim Bounds(NumBounds - 1)
          Dim k As Long, BoundText As String
5         For k = 0 To NumBounds - 1
6             BoundText = ModelText(i + k + 1)
7             Bounds(k) = SplitBounds(BoundText)
8         Next k
          
9         SetBounds = Bounds

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "SetBounds") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function SetVariableTypes(ModelText As Variant) As Variant
          ' Returns array: [Variables]
          ' Variables: [Variable type, Variable name]
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim i As Long, j As Long
1         i = FindHeaderLine(3, ModelText)
2         j = FindNextHeaderLine(3, ModelText)

          ' No variable types specified
          If i = -1 Then
              SetVariableTypes = Array()
              GoTo ExitFunction
          End If
          
          Dim variables As String, VariableHeader As RelationConsts
          Dim CurrentVariableType As Variant, VariableTypes() As Variant
3         VariableTypes = Array()

4         While (j - i) > 1
              ' Check variable line exists
              If i + 1 <> FindHeaderLine(3, ModelText, i + 1) Then
5                 VariableHeader = RelationStringToEnum(ModelText(i))
6                 variables = ModelText(i + 1)
7                 CurrentVariableType = FormVariableTypeArray(VariableHeader, variables)
8                 VariableTypes = ConcatArrays(VariableTypes, CurrentVariableType)
                  i = i + 1
              End If
9             i = i + 1
10        Wend
          
11        SetVariableTypes = VariableTypes

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "SetVariableTypes") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function SetSemiCon(ModelText As Variant) As Variant
    ' Returns array of semi-continuous variables
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
          
    Dim i As Long, j As Long
    i = FindHeaderLine(4, ModelText)
    j = FindNextHeaderLine(4, ModelText)
    If i = -1 Or (j - i) = 1 Then
        SetSemiCon = Array()
        GoTo ExitFunction
    End If
    SetSemiCon = Split(Trim(ModelText(FindHeaderLine(4, ModelText) + 1)), " ")

ExitFunction:
    If RaiseError Then RethrowError
    Exit Function

ErrorHandler:
    If Not ReportError("ReaderFileLP", "SetSemiCon") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

Function SetSOS(ModelText As Variant) As Variant
    ' Returns array: [Sets (array)]
    ' Sets: [set name, set type, variables (array)]
    ' Variables: [variable name, variable weight]
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
    
    SetSOS = Array()
    
    Dim i As Long, j As Long
    Dim SOSArray() As Variant
    
    i = FindHeaderLine(5, ModelText)
    j = FindNextHeaderLine(5, ModelText)
    
    Dim NumSets As Long
    NumSets = j - i - 1
    
    If i = -1 Or NumSets = 0 Then
        SetSOS = Array()
        GoTo ExitFunction
    End If
    
    ReDim SOSArray(NumSets - 1)

    Dim k As Long
    For k = 0 To NumSets - 1
        SOSArray(k) = SplitSOS(ModelText(i + k + 1))
        SOSArray(k)(2) = ParseSOS(SOSArray(k)(2))
    Next k
    
    SetSOS = SOSArray
    
ExitFunction:
    If RaiseError Then RethrowError
    Exit Function

ErrorHandler:
    If Not ReportError("ReaderFileLP", "SetSOS") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

Function RemoveComments(ByRef LineString As String)
          Dim pos As Long
1         pos = InStr(LineString, "\")
2         If pos <> 0 Then
3             LineString = Left(LineString, pos - 1)
4         End If
End Function

Function RemoveNames(ByRef LineString As String, Optional ByRef Name As String)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
    
          Dim pos As Long
1         pos = InStr(LineString, ":")
2         If pos <> 0 Then
              Name = Trim(Left(LineString, pos - 1))
              If ValidateName(Name) = False Then
                  RaiseUserError "Invalid function name " & Chr(34) & Name & Chr(34) & " in line: " & vbCrLf & vbCrLf & LineString
              End If
3             LineString = Right(LineString, Len(LineString) - pos - 1)
4         End If

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "RemoveNames") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function ValidateName(Name As String) As Boolean
    ValidateName = True
    Dim Char As String, Char2 As String
    Dim pos As Long
    
    ' Cannot start with a number or period
    Char = Mid(Name, 1, 1)
    If Not Char Like "[A-Za-z!" & Chr(34) & "#$%&()/,;?@_`'{}|~]" Then
        ValidateName = False
    End If
    
    ' Cannot start with the letter e followed by another e or a number
    Char2 = Mid(Name, 2, 1)
    If Char Like "[Ee]" And Char2 Like "[Ee0-9]" Then
        ValidateName = False
    End If
    
    ' Contains only allowed symbols
    For pos = 1 To Len(Name)
        Char = Mid(Name, pos, 1)
        If Not Char Like "[A-Za-z0-9!" & Chr(34) & "#$%&()/,.;?@_`'{}|~]" Then
            ValidateName = False
        End If
    Next pos
    
End Function

Function SplitEquationLHS(ByVal LHSFormula As String) As Variant
          ' To split by all operators, first replace them all by the same delimiter
          ' Choose a delimiter that would be invalid for variable names
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim SplitFormula As Variant, SplitFormulaTemp() As Variant, FormulaString As String
          Dim Operator As Variant
          Dim i As Long, j As Long

1         For Each Operator In GetOperatorArray
2             LHSFormula = Replace(LHSFormula, Operator, ".DELIMIT" & Operator)
3         Next Operator
          If InStr(LHSFormula, ".DELIMIT") = 1 Then
              LHSFormula = Right(LHSFormula, Len(LHSFormula) - Len(".DELIMIT"))
          End If

4         SplitFormula = Split(LHSFormula, ".DELIMIT")
          ReDim SplitFormulaTemp(UBound(SplitFormula))
          
          j = 0
          For i = 0 To UBound(SplitFormula)
              FormulaString = SplitFormula(i)
              If IsExponent(FormulaString) Then
                  ' Join exponent to next element as it is a coefficient, not variable
                  SplitFormula(i + 1) = SplitFormula(i) & SplitFormula(i + 1)
              Else
                  SplitFormulaTemp(j) = SplitFormula(i)
                  j = j + 1
              End If
          Next i
          
          ReDim Preserve SplitFormulaTemp(j - 1)
          SplitEquationLHS = SplitFormulaTemp

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "SplitEquationLHS") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function IsExponent(Formula As String) As Boolean
    Dim i As Long
    Formula = Replace(Formula, "+", "")
    Formula = Replace(Formula, "-", "")
    ' Look for something like "[0-9. ]+e"
    If Right(Formula, 1) = "e" Then
        IsExponent = True
        For i = 1 To Len(Formula) - 1
            If Not Mid(Formula, i, 1) Like "[0-9. ]" Then
                IsExponent = False
            End If
        Next i
    End If
End Function


Function ParseEquationLHS(SplitFormula As Variant, Optional ByRef ConstantTerm As String) As Variant
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim ParsedArray() As Variant
          Dim i As Long, j As Long, k As Long
          k = 0
          Dim CoeffAndVariable As String
1         ReDim ParsedArray(UBound(SplitFormula))
          
2         For i = 0 To UBound(SplitFormula)

3             ParsedArray(k) = Array(Left(SplitFormula(i), 1), "", "")
              
4             CoeffAndVariable = Right(SplitFormula(i), Len(SplitFormula(i)) - 1)
5             CoeffAndVariable = Replace(CoeffAndVariable, " ", "")
              
6             If ParsedArray(k)(0) <> "-" And ParsedArray(k)(0) <> "+" Then
7                 CoeffAndVariable = ParsedArray(k)(0) & CoeffAndVariable
8                 ParsedArray(k)(0) = "+" ' This should only occur on the first line
9             End If
              
10            For j = 0 To Len(CoeffAndVariable)
                  ' Find when the coefficient ends and variable name begins
11                  If ValidateName(Right(CoeffAndVariable, Len(CoeffAndVariable) - j)) Then Exit For
12            Next j
              
              If j > Len(CoeffAndVariable) Then
                  ' If we can't find the variable name beginning, it must be a constant term
                  ConstantTerm = ConstantTerm & ParsedArray(k)(0) & CoeffAndVariable
              Else
13                ParsedArray(k)(1) = Left(CoeffAndVariable, j)
14                If ParsedArray(k)(1) = "" Then
15                    ParsedArray(k)(1) = 1
16                End If
17                ParsedArray(k)(2) = Right(CoeffAndVariable, Len(CoeffAndVariable) - j)
                  k = k + 1
              End If
              
18        Next i

          ReDim Preserve ParsedArray(k - 1)
19        ParseEquationLHS = ParsedArray
          
ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "ParseEquationLHS") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function SplitConstraint(ConstraintLine As String) As Variant
          ' Look for sense and split across them
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim SenseArray As Variant
1         SenseArray = GetSenseArray
          
          Dim MinLeftPos As Long, MaxRightPos As Long
2         MinLeftPos = Len(ConstraintLine)
3         MaxRightPos = 0
          Dim LeftPos As Long, RightPos As Long
          
          Dim sense As Variant
4         For Each sense In SenseArray
5             LeftPos = InStr(ConstraintLine, sense)
6             RightPos = InStrRev(ConstraintLine, sense)
7             If LeftPos > 0 And LeftPos < MinLeftPos Then
8                 MinLeftPos = LeftPos
9             End If
10            If RightPos > MaxRightPos Then
11                MaxRightPos = RightPos
12            End If
13        Next sense
          
          Dim Relation As RelationConsts
14        Relation = RelationStringToEnum(Mid(ConstraintLine, MinLeftPos, MaxRightPos - MinLeftPos + 1))
          
15        SplitConstraint = Array(Left(ConstraintLine, MinLeftPos - 1), Relation, Right(ConstraintLine, Len(ConstraintLine) - MaxRightPos))
          
ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "SplitConstraint") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function FindHeaderLine(Header As Integer, ModelText As Variant, Optional StartLine As Long = 0) As Long
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim Found As Boolean
          Dim HeaderVariation As Variant, HeaderArray As Variant, HeaderLine As Long
          FindHeaderLine = -1
1         Found = False
2         HeaderArray = GetHeaderArray
          
3         For Each HeaderVariation In HeaderArray(Header)
4             For HeaderLine = StartLine To UBound(ModelText)
5                 If InStr(ModelText(HeaderLine), HeaderVariation & " ") = 1 Or ModelText(HeaderLine) = HeaderVariation Then
6                     Found = True
                      FindHeaderLine = HeaderLine
7                     Exit For
8                 End If
9             Next HeaderLine
10            If Found Then Exit For
11        Next HeaderVariation

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "FindHeaderLine") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function FindNextHeaderLine(Header As Integer, ModelText As Variant) As Long
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim Found As Boolean
          Dim HeaderVariation As Variant, HeaderArray As Variant, NextHeader As Integer
          NextHeader = Header + 1
1         Found = False
2         HeaderArray = GetHeaderArray
          
          While Not Found
          
3             For FindNextHeaderLine = 0 To UBound(ModelText)
4                 For Each HeaderVariation In HeaderArray(NextHeader)
5                     If InStr(ModelText(FindNextHeaderLine), HeaderVariation & " ") = 1 Or ModelText(FindNextHeaderLine) = HeaderVariation Then
6                         Found = True
7                         Exit For
8                     End If
9                 Next HeaderVariation
10                If Found Then Exit For
11            Next FindNextHeaderLine

              NextHeader = NextHeader + 1
              
          Wend

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "FindNextHeaderLine") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function SplitBounds(BoundString As String) As Variant
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
1         BoundString = Replace(BoundString, " ", "")
          Dim IneqCount As Integer
2         IneqCount = Len(BoundString) - Len(Replace(BoundString, "<", ""))
          
3         BoundString = Replace(BoundString, "infinity", MAX_DOUBLE, , , vbTextCompare)
4         BoundString = Replace(BoundString, "inf", MAX_DOUBLE, , , vbTextCompare)
          
          Dim i As Long, j As Long, freePos As Long
          Dim LowerBound As String, UpperBound As String, VariableName As String
          
          ' Inequalities are always in the form "<="
5         i = InStr(BoundString, "<")
6         j = InStrRev(BoundString, "=")

          If InStr(BoundString, ">") > 0 Or IneqCount > 2 Then
              RaiseUserError "Bound has invalid format: " & vbCrLf & vbCrLf & BoundString
          End If
          
7         If IneqCount = 2 Then
              ' l <= x <= u
8             LowerBound = Left(BoundString, i - 1)
9             UpperBound = Right(BoundString, Len(BoundString) - j)
10            VariableName = Mid(BoundString, i + 2, j - i - 3)
11        ElseIf IneqCount = 1 Then
12            If IsNumeric(Left(BoundString, 1)) Or Left(BoundString, 1) = "+" Or Left(BoundString, 1) = "-" Then
                  ' l <= x
13                LowerBound = Left(BoundString, i - 1)
14                UpperBound = MAX_DOUBLE
15                VariableName = Right(BoundString, Len(BoundString) - i - 1)
16            Else
                  ' x <= u
17                LowerBound = 0
18                UpperBound = Right(BoundString, Len(BoundString) - i - 1)
19                VariableName = Left(BoundString, i - 1)
20            End If
21        ElseIf InStr(BoundString, "=") > 0 Then
              ' x = n
22            LowerBound = Right(BoundString, Len(BoundString) - j - 1)
23            UpperBound = Right(BoundString, Len(BoundString) - j - 1)
24            VariableName = Left(BoundString, j - 1)
25        ElseIf InStr(1, BoundString, "free", vbTextCompare) > 0 Then
              ' x free
26            LowerBound = -MAX_DOUBLE
27            UpperBound = MAX_DOUBLE
              freePos = InStrRev(BoundString, "free", , vbTextCompare)
28            VariableName = Left(BoundString, freePos - 1)
          Else
              RaiseGeneralError "Could not interpret variable bound: " & vbCrLf & vbCrLf & BoundString
29        End If

30        SplitBounds = Array(LowerBound, VariableName, UpperBound)
          
ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "SplitBounds") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function FormVariableTypeArray(ByVal VariableHeader As String, VariableName As String) As Variant
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim VariablesArray As Variant
1         VariableName = Trim(VariableName)
2         VariablesArray = Split(VariableName, " ")
          Dim VariableString As Variant, TempArray() As Variant
3         ReDim TempArray(UBound(VariablesArray))
          
          Dim i As Long
4         i = 0
5         For Each VariableString In VariablesArray
6             TempArray(i) = Array(VariableHeader, VariableString)
7             i = i + 1
8         Next VariableString
          
9         FormVariableTypeArray = TempArray
          
ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "FormVariableTypeArray") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function FindAllVariables(Objective As Variant, Constraints As Variant, Bounds As Variant, VariableTypes As Variant, SemiContinuous As Variant, SOS As Variant) As Variant
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim VariableDictionary As Object
1         Set VariableDictionary = New Dictionary

          Dim i As Long, j As Long

2         For i = 0 To UBound(Objective)
3             VariableDictionary(Objective(i)(2)) = 1
4         Next i
          
5         For i = 0 To UBound(Constraints)
6             For j = 0 To UBound(Constraints(i)(0))
7                 VariableDictionary(Constraints(i)(0)(j)(2)) = 1
8             Next j
9         Next i
          
10        For i = 0 To UBound(Bounds)
11            VariableDictionary(Bounds(i)(1)) = 1
12        Next i
          
13        For i = 0 To UBound(VariableTypes)
14            VariableDictionary(VariableTypes(i)(1)) = 1
15        Next i

          For i = 0 To UBound(SemiContinuous)
              VariableDictionary(SemiContinuous(i)) = 1
          Next i

          For i = 0 To UBound(SOS)
              For j = 0 To UBound(SOS(i)(2))
                  VariableDictionary(SOS(i)(2)(j)(0)) = 1
              Next j
          Next i
          
16        FindAllVariables = VariableDictionary.Keys()
          
ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "FindAllVariables") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function FindObjCoeff(VariableName As Variant, Objective As Variant) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim i As Long
1         For i = 0 To UBound(Objective)
2             If Objective(i)(2) = VariableName Then
3                 FindObjCoeff = Objective(i)(0) & Objective(i)(1)
                  GoTo ExitFunction
4             End If
5         Next i

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "FindObjCoeff") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function GenerateAColumn(VariableName As Variant, Constraints As Variant) As Variant
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim i As Long, j As Long
          Dim AColumn() As Variant
          
          If UBound(Constraints) = -1 Then
              GoTo ExitFunction
          End If
          
1         ReDim AColumn(UBound(Constraints))
2         For i = 0 To UBound(Constraints)
3             For j = 0 To UBound(Constraints(i)(0))
4                 If Constraints(i)(0)(j)(2) = VariableName Then
5                     AColumn(i) = Constraints(i)(0)(j)(0) & Constraints(i)(0)(j)(1)
6                 End If
7             Next j
8         Next i
9         GenerateAColumn = AColumn

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "GenerateAColumn") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function FindLowerBound(VariableName As Variant, Bounds As Variant) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
1         FindLowerBound = 0
          Dim i As Long
2         For i = 0 To UBound(Bounds)
3             If Bounds(i)(1) = VariableName Then
4                 FindLowerBound = Bounds(i)(0)
                  GoTo ExitFunction
5             End If
6         Next i

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "FindLowerBound") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function FindUpperBound(VariableName As Variant, Bounds As Variant) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
1         FindUpperBound = MAX_DOUBLE
          Dim i As Long
2         For i = 0 To UBound(Bounds)
3             If Bounds(i)(1) = VariableName Then
4                 FindUpperBound = Bounds(i)(2)
                  GoTo ExitFunction
5             End If
6         Next i

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "FindUpperBound") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function FindVariableType(VariableName As Variant, VariableTypes As Variant) As RelationConsts
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim i As Long
1         For i = 0 To UBound(VariableTypes)
2             If VariableTypes(i)(1) = VariableName Then
3                 FindVariableType = VariableTypes(i)(0)
                  GoTo ExitFunction
4             End If
5         Next i

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "FindVariableType") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function FindSemiCon(VariableName As Variant, SemiContinuous As Variant) As Boolean
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
          
    FindSemiCon = False
    Dim i As Long
    For i = 0 To UBound(SemiContinuous)
        If SemiContinuous(i) = VariableName Then FindSemiCon = True
        GoTo ExitFunction
    Next i
    
ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "FindSemiCon") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function FindSOS(VariableName As Variant, SOS As Variant) As Variant
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          If UBound(SOS) = -1 Then
              FindSOS = Array()
              GoTo ExitFunction
          End If
          Dim i As Long, j As Long
          Dim SetColumn() As Variant
1         ReDim SetColumn(UBound(SOS))
2         For i = 0 To UBound(SOS)
3             For j = 0 To UBound(SOS(i)(2))
4                 If SOS(i)(2)(j)(0) = VariableName Then
5                     SetColumn(i) = SOS(i)(2)(j)(1)
6                 End If
7             Next j
8         Next i
9         FindSOS = SetColumn

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "FindSOS") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function ConcatArrays(Array1 As Variant, Array2 As Variant) As Variant
          ' Important that this works even if either array is empty (but not both)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim NewArray() As Variant
1         ReDim NewArray(UBound(Array1) + UBound(Array2) + 1)
          Dim i As Long, j As Long
2         For i = 0 To UBound(Array1)
3             NewArray(i) = Array1(i)
4         Next i
5         For j = 0 To UBound(Array2)
6             NewArray(i + j) = Array2(j)
7         Next j
8         ConcatArrays = NewArray

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "ConcatArrays") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function WriteToModel(Objective As Variant, Constraints As Variant, VariableTypes As Variant)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim ObjectiveSense As ObjectiveSenseType
2         ObjectiveSense = Objective(0)
          
3         SetObjectiveSense ObjectiveSense
4         SetObjectiveFunctionCell Range("Objective_Function")
5         SetDecisionVariables Range("Decision_Variables")

          If UBound(Constraints) = -1 Then
              GoTo SkipConstraints
          End If

          ' Set constraints
          Dim i As Long, j As Long
          Dim CurrentBlockLHS As Range, CurrentBlockRHS As Range
          Set CurrentBlockLHS = Range("Constraint_0_LHS")
          Set CurrentBlockRHS = Range("Constraint_0_RHS")
          For i = 0 To UBound(Constraints)
          
              If i <> 0 Then
                  ' If the sense has not changed from the previous constraint
                  If Range("Constraint_" & i & "_Sense") = Range("Constraint_" & i - 1 & "_Sense") Then
                      Set CurrentBlockLHS = ProperUnion(CurrentBlockLHS, Range("Constraint_" & i & "_LHS"))
                      Set CurrentBlockRHS = ProperUnion(CurrentBlockRHS, Range("Constraint_" & i & "_RHS"))
                  End If
              End If
              
              ' If the sense is the last in the block
              If i = UBound(Constraints) Then GoTo LastBlock ' VBA version of short-circuit OR
              If Range("Constraint_" & i & "_Sense") <> Range("Constraint_" & i + 1 & "_Sense") Then
LastBlock:
                  SetNamedRangeOnSheet "ConstraintBlock_" & j & "_LHS", CurrentBlockLHS
                  SetNamedRangeOnSheet "ConstraintBlock_" & j & "_RHS", CurrentBlockRHS
                  AddConstraint Range("ConstraintBlock_" & j & "_LHS"), RelationStringToEnum(Range("Constraint_" & i & "_Sense")), Range("ConstraintBlock_" & j & "_RHS")
                  j = j + 1
                  If i <> UBound(Constraints) Then
                      Set CurrentBlockLHS = Range("Constraint_" & i + 1 & "_LHS")
                      Set CurrentBlockRHS = Range("Constraint_" & i + 1 & "_RHS")
                  End If
              End If
              
          Next i

SkipConstraints:
          ' Set variable bounds
12        AddConstraint Range("Decision_Variables"), RelationGE, Range("Lower_Bounds")
13        AddConstraint Range("Decision_Variables"), RelationLE, Range("Upper_Bounds")
          
          ' Set variable types
          Dim DecisionVariable As Variant
          Dim VariableType As RelationConsts
14        For Each DecisionVariable In VariableTypes
15            If DecisionVariable(0) <> "" Then
16                VariableType = DecisionVariable(0)
17                AddConstraint Range("Variable_" & DecisionVariable(1)), VariableType
18            End If
19        Next DecisionVariable

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "WriteToModel") Then Resume
          RaiseError = True
          GoTo ExitFunction

End Function

Function SortConstraintsByIneq(Constraints As Variant) As Variant
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          ' 3 types of constraints, many potential elements
          ' VBA doesn't have any native sorting functions, so we make our own
          ' This should(?) be fairly efficient as long as number of constraints >> 3
          
          Dim OutputArray() As Variant
1         ReDim OutputArray(UBound(Constraints))
          
          Dim Element As Variant, ElementType As Integer, ElementTypeCount As Variant
2         ElementTypeCount = Array(0, 0, 0, 0)
          
3         For Each Element In Constraints
4             ElementType = Element(1)
5             ElementTypeCount(ElementType) = ElementTypeCount(ElementType) + 1
6         Next Element
              
7         ElementTypeCount(2) = ElementTypeCount(2) + ElementTypeCount(1)
          Dim CurrentTypeCount As Variant
8         CurrentTypeCount = Array(0, 0, 0, 0)
              
          Dim TypePosition As Long
9         For Each Element In Constraints
10            TypePosition = ElementTypeCount(Element(1) - 1) + CurrentTypeCount(Element(1))
11            OutputArray(TypePosition) = Element
12            CurrentTypeCount(Element(1)) = CurrentTypeCount(Element(1)) + 1
13        Next Element
          
14        SortConstraintsByIneq = OutputArray
          
ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "SortConstraintsByIneq") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function SplitSOS(ByVal SOSString As String) As Variant
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    Dim SOSArray As Variant
    Dim SetName As String
    If InStr(SOSString, ":") <> InStr(SOSString, "::") Then
        RemoveNames SOSString, SetName
        SetName = SetName & " "
    Else
        SetName = ""
    End If
    
    Dim SetType As String
    SetType = Left(Trim(SOSString), InStr(Trim(SOSString), "::") - 1)
    
    Dim SetVariables As String
    SetVariables = Right(SOSString, Len(SOSString) - InStr(SOSString, "::") - 1)
    
    SplitSOS = Array(SetName, SetType, SetVariables)
    
ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "SplitSOS") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function ParseSOS(ByVal SOSString As String) As Variant
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
    
    Dim SOSArray As Variant
    ' Turn x1 : 5     x2 : 6
    ' into x1:5 x2:6
    ' so we can split across spaces
    SOSString = Trim(SOSString)
    While InStr(SOSString, "  ") > 0
        SOSString = Replace(SOSString, "  ", " ")
    Wend
    SOSString = Replace(SOSString, " :", ":")
    SOSString = Replace(SOSString, ": ", ":")
    SOSArray = Split(SOSString, " ")
    
    Dim ParsedArray() As Variant
    ReDim ParsedArray(UBound(SOSArray))
    
    Dim i As Long, pos As Long
    
    For i = 0 To UBound(SOSArray)
        pos = InStr(SOSArray(i), ":")
        ParsedArray(i) = Array(Left(SOSArray(i), pos - 1), Right(SOSArray(i), Len(SOSArray(i)) - pos))
    Next i
    
    ParseSOS = ParsedArray
    
ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ReaderFileLP", "ParseSOS") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function
