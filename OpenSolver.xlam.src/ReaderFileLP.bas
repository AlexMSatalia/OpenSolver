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
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim ModelText As Variant
          Dim Objective As Variant, Constraints As Variant, Bounds As Variant, VariableTypes As Variant, SemiContinuous As Variant, SOS As Variant

3         ModelText = ParseSections(FileName) ' Read in file and clean up whitespace/comments
4         Objective = SetObjectiveFunction(ModelText)
5         Constraints = SetConstraints(ModelText)
6         Bounds = SetBounds(ModelText)
7         VariableTypes = SetVariableTypes(ModelText)
8         SemiContinuous = SetSemiCon(ModelText)
9         SOS = SetSOS(ModelText)
10        WriteToSheet FileName, Objective, Constraints, Bounds, VariableTypes, SemiContinuous, SOS, sheet
11        WriteToModel Objective, Constraints, VariableTypes

12        ShowSolverModel ActiveSheet
          
13        If (UBound(SemiContinuous) >= 0 Or UBound(SOS) >= 0) And MinimiseUserInteraction = False Then
              Dim WarningMessage As String
14            WarningMessage = "WARNING: The LP file imported contains " & _
                               IIf(UBound(SemiContinuous) >= 0, "semi-continuous variables", "") & _
                               IIf(UBound(SemiContinuous) >= 0 And UBound(SOS) >= 0, " and ", "") & _
                               IIf(UBound(SOS) >= 0, "special ordered sets", "") & _
                               ". These have been printed to the sheet, but have not been written to the model nor supported during model solve."
15            MsgBox WarningMessage, , "OpenSolver Warning"
16        End If
          
ExitSub:
17        If RaiseError Then RethrowError
18        Exit Function

ErrorHandler:
19        If Not ReportError("ReaderFileLP", "ReadLPFile") Then Resume
20        RaiseError = True
21        GoTo ExitSub
End Function

Function ParseSections(FileName As String) As Variant
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim FileLine As String, ModelStr As String, LineCount As Integer
          Dim NonLinearSense As Variant
3         NonLinearSense = Array("*", "^", "[", "]")
4         LineCount = 0
          
          Dim fso As Object, InputFile As Variant
5         Set fso = CreateObject("Scripting.FileSystemObject")
          
6         Set InputFile = fso.OpenTextFile(FileName, 1)
          
7         Dim EndOfLine As Boolean
8         Do While InputFile.AtEndOfStream <> True
          
9             LineCount = LineCount + 1
10            EndOfLine = False
11            FileLine = InputFile.ReadLine
LineCheck:
              ' Search and delete comments
12            RemoveComments FileLine
13            FileLine = Trim(FileLine)
              
              Dim NLSense As Variant
14            For Each NLSense In NonLinearSense
15                If InStr(FileLine, NLSense) > 0 Then
16                    RaiseUserError "Non-linearity detected on line " & LineCount & _
                                     ". This feature does not currently support non-linear LP files. " & _
                                     "The non-linear line is printed below:" & vbCrLf & vbCrLf & FileLine
17                End If
18            Next NLSense
              
              ' Line ends occur in two situations
              ' 1: Headers are on their own line
              ' 2: Constraints and bounds are on their own lines
              
              ' Constraints and bounds end with inequalities or are free
              Dim sense As Variant, Operator As Variant
              
19            If InStr(FileLine, "free") > 0 Then
20                EndOfLine = True
21                For Each Operator In GetOperatorArray
22                    If InStr(FileLine, Operator) > 0 Then
23                        EndOfLine = False
24                    End If
25                Next Operator
26                If EndOfLine Then GoTo EOL
27            End If
              
28            For Each sense In GetSenseArray
29                If InStr(FileLine, sense) > 0 Then
30                    EndOfLine = True
31                End If
32            Next sense

              ' Special case: Special ordered sets
33            If InStr(FileLine, "::") > 0 Then
34                EndOfLine = True
35                GoTo EOL
36            End If

              ' Check for headers
              Dim HeaderVariation As Variant, HeaderArray As Variant
37            HeaderArray = GetHeaderArray
              
              Dim i As Long, LongestVariation As String
38            For i = 0 To UBound(HeaderArray, 1)

39                LongestVariation = ""
40                For Each HeaderVariation In HeaderArray(i)
41                    If InStr(LCase(FileLine), HeaderVariation & " ") = 1 Or LCase(FileLine) = HeaderVariation Then
                          ' We look for exactly the header OR header followed by a space
42                        If Len(HeaderVariation) > Len(LongestVariation) Then
43                            LongestVariation = HeaderVariation
44                        End If
45                        EndOfLine = True
46                    End If
47                Next HeaderVariation

48                If Len(LongestVariation) > 0 Then
                      ' Check for a constraints on the same line as the header
49                    ModelStr = ModelStr & vbCrLf & LongestVariation & vbCrLf
50                    If Len(FileLine) - Len(LongestVariation) > 0 Then
51                        EndOfLine = False
52                        FileLine = Right(Trim(FileLine), Len(FileLine) - Len(LongestVariation))
53                        GoTo LineCheck
54                    End If
55                    GoTo Bypass
56                End If

57            Next i

EOL:
58            If EndOfLine Then
59                ModelStr = ModelStr & FileLine & vbCrLf
60            Else
                  ' Variable types require whitespace between variable names
61                If Len(Trim(FileLine)) <> 0 Then FileLine = " " & FileLine
62                ModelStr = ModelStr & FileLine
63            End If
Bypass:
64        Loop
65        InputFile.Close
          
          ' Clean up final string - trim blank lines and unnecessary whitespace
66        While InStr(ModelStr, "  ") > 0
67            ModelStr = Replace(ModelStr, "  ", " ")
68        Wend
69        ModelStr = Trim(ModelStr)

          ' A well formatted .lp file should end with "End" - confirm anyway
70        If Right(ModelStr, 5) <> "end" + vbCrLf Then
71            ModelStr = ModelStr + vbCrLf + "end"
72        End If
       
          Dim strLines As Variant, parsedLines() As Variant
73        strLines = Split(ModelStr, vbCrLf)
74        ReDim parsedLines(UBound(strLines))

          ' Remove lines that only contain spaces
          Dim j As Long
75        j = 0
76        For i = 0 To UBound(strLines)
77            If strLines(i) <> "" And strLines(i) <> " " Then
78                parsedLines(j) = strLines(i)
79                j = j + 1
80            End If
81        Next i
       
82        ReDim Preserve parsedLines(j - 1)
83        ParseSections = parsedLines

ExitFunction:
84        If RaiseError Then RethrowError
85        InputFile.Close
86        Exit Function

ErrorHandler:
87        If Not ReportError("ReaderFileLP", "ParseSections") Then Resume
88        RaiseError = True
89        GoTo ExitFunction
End Function

Function WriteToSheet(FileName As String, Objective As Variant, Constraints As Variant, Bounds As Variant, VariableTypes As Variant, SemiContinuous As Variant, SOS As Variant, Optional ByRef sheet As Worksheet)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
3         Application.ScreenUpdating = False

          Dim VariableList As Variant
4         VariableList = FindAllVariables(Objective(1), Constraints, Bounds, VariableTypes, SemiContinuous, SOS)

          Dim FileNameShort As String
5         FileNameShort = Right(FileName, Len(FileName) - InStrRev(FileName, "\"))
6         If sheet Is Nothing Then
7             Set sheet = MakeNewSheet(FileNameShort, False)
8         Else
9             sheet.Activate
10            sheet.UsedRange.Delete ' Overwrite existing cells, but keep associated references
11        End If
          
12        Cells.ColumnWidth = 10 ' To fit MAX_DOUBLE in bounds
13        Cells(1, 1) = FileNameShort
14        Cells(2, 1) = FileName
15        Cells(2, 2) = " " ' Prevent long file name overflowing on to next cell

          Dim nVariables As Long, nConstraints As Long, nSets As Long
16        nVariables = UBound(VariableList)
17        nConstraints = UBound(Constraints)
18        nSets = UBound(SOS)

          ' Sort all data by variable
          Dim VariableProperties() As Variant
19        ReDim VariableProperties(nVariables)
          Dim i As Long
20        For i = 0 To nVariables
21            VariableProperties(i) = Array( _
                  VariableList(i), _
                  FindObjCoeff(VariableList(i), Objective(1)), _
                  GenerateAColumn(VariableList(i), Constraints), _
                  FindLowerBound(VariableList(i), Bounds), _
                  FindUpperBound(VariableList(i), Bounds), _
                  FindVariableType(VariableList(i), VariableTypes), _
                  FindSemiCon(VariableList(i), SemiContinuous), _
                  FindSOS(VariableList(i), SOS) _
                  )
22        Next i
          
          Dim VariableType As RelationConsts
          Dim ExtendedAMatrix As Variant
23        ReDim ExtendedAMatrix(11 + nConstraints + nSets, nVariables)
          
24        If nSets >= 0 Then
              Dim SetsNameArray() As Variant
25            ReDim SetsNameArray(nSets, 0)
26        End If
          
          ' Make variable columns
          Dim j As Long, k As Long, l As Long
27        For j = 0 To nVariables

28            ExtendedAMatrix(0, j) = VariableProperties(j)(0) ' Variable name
29            ExtendedAMatrix(1, j) = "0" ' Initial value of variable
30            ExtendedAMatrix(3, j) = VariableProperties(j)(1) ' Objective coefficient
31            SetNamedRangeOnSheet "Variable_" & VariableProperties(j)(0), Cells(2, 3 + j)

32            For k = 0 To nConstraints
33                ExtendedAMatrix(5 + k, j) = VariableProperties(j)(2)(k) ' Constraints
34            Next k

35            ExtendedAMatrix(7 + nConstraints, j) = VariableProperties(j)(3) ' Lower bound
36            ExtendedAMatrix(8 + nConstraints, j) = VariableProperties(j)(4) ' Upper bound

37            VariableType = VariableProperties(j)(5)
38            SetNamedRangeOnSheet "Variable_" & VariableProperties(j)(0) & "_Type", Cells(10 + nConstraints, 3 + j)
39            ExtendedAMatrix(9 + nConstraints, j) = RelationEnumToString(VariableType) ' Variable type

40            If VariableProperties(j)(6) Then ExtendedAMatrix(10 + nConstraints, j) = "Yes" ' Semi-continuous
              
41            For l = 0 To nSets
42                ExtendedAMatrix(11 + nConstraints + l, j) = VariableProperties(j)(7)(l) ' Special ordered sets
43            Next l
              
44        Next j

          Dim AMatrixRange As Range
45        Set AMatrixRange = ActiveSheet.Range(Cells(1, 3), Cells(12 + nConstraints + nSets, 3 + nVariables))
46        AMatrixRange.value = ExtendedAMatrix

47        If nSets >= 0 Then
48            For i = 0 To nSets
49                SetsNameArray(i, 0) = SOS(i)(0) & "(" & SOS(i)(1) & ")"
50            Next i
              Dim SetsRange As Range
51            Set SetsRange = ActiveSheet.Range(Cells(12 + nConstraints, 2), Cells(12 + nConstraints + nSets, 2))
52            SetsRange.value = SetsNameArray
53        End If

          ' Write ranges
54        SetNamedRangeOnSheet "Objective_coefficients", Range(Cells(4, 3), Cells(4, 3 + nVariables))
55        SetNamedRangeOnSheet "Decision_variables", Range(Cells(2, 3), Cells(2, 3 + nVariables))
56        SetNamedRangeOnSheet "Lower_bounds", Range(Cells(8 + nConstraints, 3), Cells(8 + nConstraints, 3 + nVariables))
57        SetNamedRangeOnSheet "Upper_bounds", Range(Cells(9 + nConstraints, 3), Cells(9 + nConstraints, 3 + nVariables))
          
58        If nConstraints = -1 Then
59            GoTo SkipConstraints
60        End If
          
          ' Write everything to the right of the A matrix
          Dim Relation As RelationConsts
          Dim SumArray() As Variant, ConstraintNameArray() As Variant
61        ReDim SumArray(nConstraints, 3)
62        ReDim ConstraintNameArray(nConstraints, 0)
          
63        For i = 0 To nConstraints

64            SetNamedRangeOnSheet "Constraint_" & i, Range(Cells(6 + i, 3), Cells(6 + i, 3 + nVariables))
65            ConstraintNameArray(i, 0) = Constraints(i)(3)

66            Relation = Constraints(i)(1)

67            SumArray(i, 0) = "=SUMPRODUCT(Constraint_" & i & ",Decision_variables)"
68            SumArray(i, 1) = RelationEnumToString(Relation)
69            SumArray(i, 2) = Constraints(i)(2)

70            SetNamedRangeOnSheet "Constraint_" & i & "_LHS", Cells(6 + i, 5 + nVariables)
71            SetNamedRangeOnSheet "Constraint_" & i & "_Sense", Cells(6 + i, 6 + nVariables)
72            SetNamedRangeOnSheet "Constraint_" & i & "_RHS", Cells(6 + i, 7 + nVariables)

73        Next i
        
74        SetNamedRangeOnSheet "ConstraintNames", Range(Cells(6, 2), Cells(6 + nConstraints, 2))
75        Range("ConstraintNames").value = ConstraintNameArray
        
76        SetNamedRangeOnSheet "ConstraintCalcs", Range(Cells(6, 5 + nVariables), Cells(6 + nConstraints, 7 + nVariables))
77        Range("ConstraintCalcs").value = SumArray

SkipConstraints:
78        Cells(4, 5 + nVariables).value = "=SUMPRODUCT(Objective_coefficients, Decision_variables)" & Objective(3)
79        SetNamedRangeOnSheet "Objective_Function", Cells(4, 5 + nVariables)

          ' Write descriptors
80        If Objective(0) = 1 Then
81            Cells(4, 1) = "Maximise:"
82        ElseIf Objective(0) = 2 Then
83            Cells(4, 1) = "Minimise:"
84        Else
85            RaiseGeneralError "Unknown objective sense: " & Objective(0)
86        End If
87        Cells(4, 2) = Objective(2)
88        Cells(6, 1) = "Subject to:"
89        Cells(8 + nConstraints, 1) = "Variable bounds: "
90        Cells(8 + nConstraints, 2) = ">="
91        Cells(9 + nConstraints, 2) = "<="
92        Cells(10 + nConstraints, 1) = "Variable type: "
93        If UBound(SemiContinuous) >= 0 Then
94            Cells(11 + nConstraints, 1) = "Semi-continuous?"
95            Cells(11 + nConstraints, 1).AddComment ("Semi-continuous variables are printed from the file, but not added to the model nor supported during solve.")
96            Cells(11 + nConstraints, 1).Comment.Shape.TextFrame.AutoSize = True
97        End If
98        If nSets >= 0 Then
              ' This leaves a one-line gap if SOS exists but not semi-continuous variables.
99            Cells(12 + nConstraints, 1) = "Set weighting: "
100           Cells(12 + nConstraints, 1).AddComment ("Special ordered sets are printed from the file, but not added to the model nor supported during solve.")
101           Cells(12 + nConstraints, 1).Comment.Shape.TextFrame.AutoSize = True
102       End If

103       Range(Cells(3, 1), Cells(12 + nConstraints, 1)).Columns.AutoFit
104       Columns(2).EntireColumn.AutoFit
105       Columns(1).Font.Bold = True
106       Rows(1).Font.Bold = True
107       Cells(2, 1).Font.Bold = False
          
108       Application.ScreenUpdating = True
        
ExitFunction:
109       If RaiseError Then RethrowError
110       Application.ScreenUpdating = True
111       Exit Function

ErrorHandler:
112       If Not ReportError("ReaderFileLP", "WriteToSheet") Then Resume
113       RaiseError = True
114       GoTo ExitFunction
End Function

Function SetObjectiveFunction(ModelText As Variant) As Variant
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          ' Returns array: [Objective sense (enum), Parsed formulae (array), Objective name, Constant term (string)]
          ' Parsed formulae: [Operator, Coefficient, Variable]
          
          Dim ObjectiveArray(3) As Variant
          
          ' Objective sense should always be line 1
          Dim HeaderVariation As Variant, HeaderArray As Variant
3         HeaderArray = GetHeaderArray
4         For Each HeaderVariation In HeaderArray(0)
5             If InStr(ModelText(0), HeaderVariation) > 0 Then
6                 ObjectiveArray(0) = ObjectiveSenseStringToEnum(HeaderVariation)
7             End If
8         Next HeaderVariation

9         If IsEmpty(ObjectiveArray(0)) Then
10            RaiseUserError "Could not find objective sense in the LP file."
11        End If

          Dim LHSFormula As String, ObjectiveName As String
12        LHSFormula = ModelText(1)
13        ObjectiveName = ""
14        RemoveNames LHSFormula, ObjectiveName
          
          Dim SplitFormula As Variant, ConstantTerm As String, ParsedFormula As Variant
15        SplitFormula = SplitEquationLHS(Trim(LHSFormula))
16        ConstantTerm = ""
17        ParsedFormula = ParseEquationLHS(SplitFormula, ConstantTerm)
          
18        ObjectiveArray(1) = ParsedFormula
19        ObjectiveArray(2) = ObjectiveName
20        ObjectiveArray(3) = ConstantTerm
21        SetObjectiveFunction = ObjectiveArray

ExitFunction:
22        If RaiseError Then RethrowError
23        Exit Function

ErrorHandler:
24        If Not ReportError("ReaderFileLP", "SetObjectiveFunction") Then Resume
25        RaiseError = True
26        GoTo ExitFunction
End Function

Function SetConstraints(ModelText As Variant, Optional SortConstraints As Boolean = False) As Variant
          ' Returns array: [List of constraints]
          ' Each constraint: [Variables (array), Sense, RHS, Constraint name]
          ' Variables: [List of variables (array)]
          ' List of variables: [Operator, Coefficient, variable name]
          ' To access constraint i, variable j coefficient:
          ' coefficient = ArrayName(i)(0)(j)(1)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim i As Long, j As Long, NumConstraints As Long
3         i = FindHeaderLine(1, ModelText)
4         j = FindNextHeaderLine(1, ModelText)
5         NumConstraints = j - i - 1

6         If NumConstraints = 0 Then
7             SetConstraints = Array()
8             GoTo ExitFunction
9         End If
          
          Dim Constraints() As Variant, ConstraintText As String, ConstraintName As String
10        ReDim Constraints(NumConstraints - 1)
          Dim k As Long
11        For k = 0 To NumConstraints - 1

12            ConstraintName = ""
13            ConstraintText = ModelText(i + k + 1)
14            RemoveNames ConstraintText, ConstraintName

15            Constraints(k) = ConcatArrays(SplitConstraint(Trim(ConstraintText)), Array(ConstraintName))
16            ConstraintText = Constraints(k)(0)
17            Constraints(k)(0) = ParseEquationLHS(SplitEquationLHS(ConstraintText))
              
18        Next k
          
19        If SortConstraints = True Then
              Dim SortedConstraints As Variant
20            SortedConstraints = SortConstraintsByIneq(Constraints)
21            SetConstraints = SortedConstraints
22        Else
23            SetConstraints = Constraints
24        End If

ExitFunction:
25        If RaiseError Then RethrowError
26        Exit Function

ErrorHandler:
27        If Not ReportError("ReaderFileLP", "SetConstraints") Then Resume
28        RaiseError = True
29        GoTo ExitFunction
End Function

Function SetBounds(ModelText As Variant) As Variant
          ' Returns array: [Bound number (array)]
          ' Bound number: [Lower bound, Variable, Upper bound]
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim i As Long, j As Long, NumBounds As Long
3         i = FindHeaderLine(2, ModelText)
4         j = FindNextHeaderLine(2, ModelText)
5         NumBounds = j - i - 1
6         If i = -1 Or NumBounds = 0 Then
              ' Bounds section is optional
7             SetBounds = Array()
8             GoTo ExitFunction
9         End If
          
          Dim Bounds() As Variant
10        ReDim Bounds(NumBounds - 1)
          Dim k As Long, BoundText As String
11        For k = 0 To NumBounds - 1
12            BoundText = ModelText(i + k + 1)
13            Bounds(k) = SplitBounds(BoundText)
14        Next k
          
15        SetBounds = Bounds

ExitFunction:
16        If RaiseError Then RethrowError
17        Exit Function

ErrorHandler:
18        If Not ReportError("ReaderFileLP", "SetBounds") Then Resume
19        RaiseError = True
20        GoTo ExitFunction
End Function

Function SetVariableTypes(ModelText As Variant) As Variant
          ' Returns array: [Variables]
          ' Variables: [Variable type, Variable name]
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim i As Long, j As Long
3         i = FindHeaderLine(3, ModelText)
4         j = FindNextHeaderLine(3, ModelText)

          ' No variable types specified
5         If i = -1 Then
6             SetVariableTypes = Array()
7             GoTo ExitFunction
8         End If
          
          Dim variables As String, VariableHeader As RelationConsts
          Dim CurrentVariableType As Variant, VariableTypes() As Variant
9         VariableTypes = Array()

10        While (j - i) > 1
              ' Check variable line exists
11            If i + 1 <> FindHeaderLine(3, ModelText, i + 1) Then
12                VariableHeader = RelationStringToEnum(ModelText(i))
13                variables = ModelText(i + 1)
14                CurrentVariableType = FormVariableTypeArray(VariableHeader, variables)
15                VariableTypes = ConcatArrays(VariableTypes, CurrentVariableType)
16                i = i + 1
17            End If
18            i = i + 1
19        Wend
          
20        SetVariableTypes = VariableTypes

ExitFunction:
21        If RaiseError Then RethrowError
22        Exit Function

ErrorHandler:
23        If Not ReportError("ReaderFileLP", "SetVariableTypes") Then Resume
24        RaiseError = True
25        GoTo ExitFunction
End Function

Function SetSemiCon(ModelText As Variant) As Variant
          ' Returns array of semi-continuous variables
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
                
          Dim i As Long, j As Long
3         i = FindHeaderLine(4, ModelText)
4         j = FindNextHeaderLine(4, ModelText)
5         If i = -1 Or (j - i) = 1 Then
6             SetSemiCon = Array()
7             GoTo ExitFunction
8         End If
9         SetSemiCon = Split(Trim(ModelText(FindHeaderLine(4, ModelText) + 1)), " ")

ExitFunction:
10        If RaiseError Then RethrowError
11        Exit Function

ErrorHandler:
12        If Not ReportError("ReaderFileLP", "SetSemiCon") Then Resume
13        RaiseError = True
14        GoTo ExitFunction
End Function

Function SetSOS(ModelText As Variant) As Variant
          ' Returns array: [Sets (array)]
          ' Sets: [set name, set type, variables (array)]
          ' Variables: [variable name, variable weight]
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
3         SetSOS = Array()
          
          Dim i As Long, j As Long
          Dim SOSArray() As Variant
          
4         i = FindHeaderLine(5, ModelText)
5         j = FindNextHeaderLine(5, ModelText)
          
          Dim NumSets As Long
6         NumSets = j - i - 1
          
7         If i = -1 Or NumSets = 0 Then
8             SetSOS = Array()
9             GoTo ExitFunction
10        End If
          
11        ReDim SOSArray(NumSets - 1)

          Dim k As Long
12        For k = 0 To NumSets - 1
13            SOSArray(k) = SplitSOS(ModelText(i + k + 1))
14            SOSArray(k)(2) = ParseSOS(SOSArray(k)(2))
15        Next k
          
16        SetSOS = SOSArray
          
ExitFunction:
17        If RaiseError Then RethrowError
18        Exit Function

ErrorHandler:
19        If Not ReportError("ReaderFileLP", "SetSOS") Then Resume
20        RaiseError = True
21        GoTo ExitFunction
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
1         RaiseError = False
2         On Error GoTo ErrorHandler
    
          Dim pos As Long
3         pos = InStr(LineString, ":")
4         If pos <> 0 Then
5             Name = Trim(Left(LineString, pos - 1))
6             If ValidateName(Name) = False Then
7                 RaiseUserError "Invalid function name " & Chr(34) & Name & Chr(34) & " in line: " & vbCrLf & vbCrLf & LineString
8             End If
9             LineString = Right(LineString, Len(LineString) - pos - 1)
10        End If

ExitFunction:
11        If RaiseError Then RethrowError
12        Exit Function

ErrorHandler:
13        If Not ReportError("ReaderFileLP", "RemoveNames") Then Resume
14        RaiseError = True
15        GoTo ExitFunction
End Function

Function ValidateName(Name As String) As Boolean
1         ValidateName = True
          Dim Char As String, Char2 As String
          Dim pos As Long
          
          ' Cannot start with a number or period
2         Char = Mid(Name, 1, 1)
3         If Not Char Like "[A-Za-z!" & Chr(34) & "#$%&()/,;?@_`'{}|~]" Then
4             ValidateName = False
5         End If
          
          ' Cannot start with the letter e followed by another e or a number
6         Char2 = Mid(Name, 2, 1)
7         If Char Like "[Ee]" And Char2 Like "[Ee0-9]" Then
8             ValidateName = False
9         End If
          
          ' Contains only allowed symbols
10        For pos = 1 To Len(Name)
11            Char = Mid(Name, pos, 1)
12            If Not Char Like "[A-Za-z0-9!" & Chr(34) & "#$%&()/,.;?@_`'{}|~]" Then
13                ValidateName = False
14            End If
15        Next pos
          
End Function

Function SplitEquationLHS(ByVal LHSFormula As String) As Variant
          ' To split by all operators, first replace them all by the same delimiter
          ' Choose a delimiter that would be invalid for variable names
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim SplitFormula As Variant, SplitFormulaTemp() As Variant, FormulaString As String
          Dim Operator As Variant
          Dim i As Long, j As Long

3         For Each Operator In GetOperatorArray
4             LHSFormula = Replace(LHSFormula, Operator, ".DELIMIT" & Operator)
5         Next Operator
6         If InStr(LHSFormula, ".DELIMIT") = 1 Then
7             LHSFormula = Right(LHSFormula, Len(LHSFormula) - Len(".DELIMIT"))
8         End If

9         SplitFormula = Split(LHSFormula, ".DELIMIT")
10        ReDim SplitFormulaTemp(UBound(SplitFormula))
          
11        j = 0
12        For i = 0 To UBound(SplitFormula)
13            FormulaString = SplitFormula(i)
14            If IsExponent(FormulaString) Then
                  ' Join exponent to next element as it is a coefficient, not variable
15                SplitFormula(i + 1) = SplitFormula(i) & SplitFormula(i + 1)
16            Else
17                SplitFormulaTemp(j) = SplitFormula(i)
18                j = j + 1
19            End If
20        Next i
          
21        ReDim Preserve SplitFormulaTemp(j - 1)
22        SplitEquationLHS = SplitFormulaTemp

ExitFunction:
23        If RaiseError Then RethrowError
24        Exit Function

ErrorHandler:
25        If Not ReportError("ReaderFileLP", "SplitEquationLHS") Then Resume
26        RaiseError = True
27        GoTo ExitFunction
End Function

Function IsExponent(Formula As String) As Boolean
          Dim i As Long
1         Formula = Replace(Formula, "+", "")
2         Formula = Replace(Formula, "-", "")
          ' Look for something like "[0-9. ]+e"
3         If Right(Formula, 1) = "e" Then
4             IsExponent = True
5             For i = 1 To Len(Formula) - 1
6                 If Not Mid(Formula, i, 1) Like "[0-9. ]" Then
7                     IsExponent = False
8                 End If
9             Next i
10        End If
End Function


Function ParseEquationLHS(SplitFormula As Variant, Optional ByRef ConstantTerm As String) As Variant
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim ParsedArray() As Variant
          Dim i As Long, j As Long, k As Long
3         k = 0
          Dim CoeffAndVariable As String
4         ReDim ParsedArray(UBound(SplitFormula))
          
5         For i = 0 To UBound(SplitFormula)

6             ParsedArray(k) = Array(Left(SplitFormula(i), 1), "", "")
              
7             CoeffAndVariable = Right(SplitFormula(i), Len(SplitFormula(i)) - 1)
8             CoeffAndVariable = Replace(CoeffAndVariable, " ", "")
              
9             If ParsedArray(k)(0) <> "-" And ParsedArray(k)(0) <> "+" Then
10                CoeffAndVariable = ParsedArray(k)(0) & CoeffAndVariable
11                ParsedArray(k)(0) = "+" ' This should only occur on the first line
12            End If
              
13            For j = 0 To Len(CoeffAndVariable)
                  ' Find when the coefficient ends and variable name begins
14                  If ValidateName(Right(CoeffAndVariable, Len(CoeffAndVariable) - j)) Then Exit For
15            Next j
              
16            If j > Len(CoeffAndVariable) Then
                  ' If we can't find the variable name beginning, it must be a constant term
17                ConstantTerm = ConstantTerm & ParsedArray(k)(0) & CoeffAndVariable
18            Else
19                ParsedArray(k)(1) = Left(CoeffAndVariable, j)
20                If ParsedArray(k)(1) = "" Then
21                    ParsedArray(k)(1) = 1
22                End If
23                ParsedArray(k)(2) = Right(CoeffAndVariable, Len(CoeffAndVariable) - j)
24                k = k + 1
25            End If
              
26        Next i

27        ReDim Preserve ParsedArray(k - 1)
28        ParseEquationLHS = ParsedArray
          
ExitFunction:
29        If RaiseError Then RethrowError
30        Exit Function

ErrorHandler:
31        If Not ReportError("ReaderFileLP", "ParseEquationLHS") Then Resume
32        RaiseError = True
33        GoTo ExitFunction
End Function

Function SplitConstraint(ConstraintLine As String) As Variant
          ' Look for sense and split across them
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim SenseArray As Variant
3         SenseArray = GetSenseArray
          
          Dim MinLeftPos As Long, MaxRightPos As Long
4         MinLeftPos = Len(ConstraintLine)
5         MaxRightPos = 0
          Dim LeftPos As Long, RightPos As Long
          
          Dim sense As Variant
6         For Each sense In SenseArray
7             LeftPos = InStr(ConstraintLine, sense)
8             RightPos = InStrRev(ConstraintLine, sense)
9             If LeftPos > 0 And LeftPos < MinLeftPos Then
10                MinLeftPos = LeftPos
11            End If
12            If RightPos > MaxRightPos Then
13                MaxRightPos = RightPos
14            End If
15        Next sense
          
          Dim Relation As RelationConsts
16        Relation = RelationStringToEnum(Mid(ConstraintLine, MinLeftPos, MaxRightPos - MinLeftPos + 1))
          
17        SplitConstraint = Array(Left(ConstraintLine, MinLeftPos - 1), Relation, Right(ConstraintLine, Len(ConstraintLine) - MaxRightPos))
          
ExitFunction:
18        If RaiseError Then RethrowError
19        Exit Function

ErrorHandler:
20        If Not ReportError("ReaderFileLP", "SplitConstraint") Then Resume
21        RaiseError = True
22        GoTo ExitFunction
End Function

Function FindHeaderLine(Header As Integer, ModelText As Variant, Optional StartLine As Long = 0) As Long
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim Found As Boolean
          Dim HeaderVariation As Variant, HeaderArray As Variant, HeaderLine As Long
3         FindHeaderLine = -1
4         Found = False
5         HeaderArray = GetHeaderArray
          
6         For Each HeaderVariation In HeaderArray(Header)
7             For HeaderLine = StartLine To UBound(ModelText)
8                 If InStr(ModelText(HeaderLine), HeaderVariation & " ") = 1 Or ModelText(HeaderLine) = HeaderVariation Then
9                     Found = True
10                    FindHeaderLine = HeaderLine
11                    Exit For
12                End If
13            Next HeaderLine
14            If Found Then Exit For
15        Next HeaderVariation

ExitFunction:
16        If RaiseError Then RethrowError
17        Exit Function

ErrorHandler:
18        If Not ReportError("ReaderFileLP", "FindHeaderLine") Then Resume
19        RaiseError = True
20        GoTo ExitFunction
End Function

Function FindNextHeaderLine(Header As Integer, ModelText As Variant) As Long
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim Found As Boolean
          Dim HeaderVariation As Variant, HeaderArray As Variant, NextHeader As Integer
3         NextHeader = Header + 1
4         Found = False
5         HeaderArray = GetHeaderArray
          
6         While Not Found
          
7             For FindNextHeaderLine = 0 To UBound(ModelText)
8                 For Each HeaderVariation In HeaderArray(NextHeader)
9                     If InStr(ModelText(FindNextHeaderLine), HeaderVariation & " ") = 1 Or ModelText(FindNextHeaderLine) = HeaderVariation Then
10                        Found = True
11                        Exit For
12                    End If
13                Next HeaderVariation
14                If Found Then Exit For
15            Next FindNextHeaderLine

16            NextHeader = NextHeader + 1
              
17        Wend

ExitFunction:
18        If RaiseError Then RethrowError
19        Exit Function

ErrorHandler:
20        If Not ReportError("ReaderFileLP", "FindNextHeaderLine") Then Resume
21        RaiseError = True
22        GoTo ExitFunction
End Function

Function SplitBounds(BoundString As String) As Variant
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
3         BoundString = Replace(BoundString, " ", "")
          Dim IneqCount As Integer
4         IneqCount = Len(BoundString) - Len(Replace(BoundString, "<", ""))
          
5         BoundString = Replace(BoundString, "infinity", MAX_DOUBLE, , , vbTextCompare)
6         BoundString = Replace(BoundString, "inf", MAX_DOUBLE, , , vbTextCompare)
          
          Dim i As Long, j As Long, freePos As Long
          Dim LowerBound As String, UpperBound As String, VariableName As String
          
          ' Inequalities are always in the form "<="
7         i = InStr(BoundString, "<")
8         j = InStrRev(BoundString, "=")

9         If InStr(BoundString, ">") > 0 Or IneqCount > 2 Then
10            RaiseUserError "Bound has invalid format: " & vbCrLf & vbCrLf & BoundString
11        End If
          
12        If IneqCount = 2 Then
              ' l <= x <= u
13            LowerBound = Left(BoundString, i - 1)
14            UpperBound = Right(BoundString, Len(BoundString) - j)
15            VariableName = Mid(BoundString, i + 2, j - i - 3)
16        ElseIf IneqCount = 1 Then
17            If IsNumeric(Left(BoundString, 1)) Or Left(BoundString, 1) = "+" Or Left(BoundString, 1) = "-" Then
                  ' l <= x
18                LowerBound = Left(BoundString, i - 1)
19                UpperBound = MAX_DOUBLE
20                VariableName = Right(BoundString, Len(BoundString) - i - 1)
21            Else
                  ' x <= u
22                LowerBound = 0
23                UpperBound = Right(BoundString, Len(BoundString) - i - 1)
24                VariableName = Left(BoundString, i - 1)
25            End If
26        ElseIf InStr(BoundString, "=") > 0 Then
              ' x = n
27            LowerBound = Right(BoundString, Len(BoundString) - j - 1)
28            UpperBound = Right(BoundString, Len(BoundString) - j - 1)
29            VariableName = Left(BoundString, j - 1)
30        ElseIf InStr(1, BoundString, "free", vbTextCompare) > 0 Then
              ' x free
31            LowerBound = -MAX_DOUBLE
32            UpperBound = MAX_DOUBLE
33            freePos = InStrRev(BoundString, "free", , vbTextCompare)
34            VariableName = Left(BoundString, freePos - 1)
35        Else
36            RaiseGeneralError "Could not interpret variable bound: " & vbCrLf & vbCrLf & BoundString
37        End If

38        SplitBounds = Array(LowerBound, VariableName, UpperBound)
          
ExitFunction:
39        If RaiseError Then RethrowError
40        Exit Function

ErrorHandler:
41        If Not ReportError("ReaderFileLP", "SplitBounds") Then Resume
42        RaiseError = True
43        GoTo ExitFunction
End Function

Function FormVariableTypeArray(ByVal VariableHeader As String, VariableName As String) As Variant
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim VariablesArray As Variant
3         VariableName = Trim(VariableName)
4         VariablesArray = Split(VariableName, " ")
          Dim VariableString As Variant, TempArray() As Variant
5         ReDim TempArray(UBound(VariablesArray))
          
          Dim i As Long
6         i = 0
7         For Each VariableString In VariablesArray
8             TempArray(i) = Array(VariableHeader, VariableString)
9             i = i + 1
10        Next VariableString
          
11        FormVariableTypeArray = TempArray
          
ExitFunction:
12        If RaiseError Then RethrowError
13        Exit Function

ErrorHandler:
14        If Not ReportError("ReaderFileLP", "FormVariableTypeArray") Then Resume
15        RaiseError = True
16        GoTo ExitFunction
End Function

Function FindAllVariables(Objective As Variant, Constraints As Variant, Bounds As Variant, VariableTypes As Variant, SemiContinuous As Variant, SOS As Variant) As Variant
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim VariableDictionary As Object
3         Set VariableDictionary = New Dictionary

          Dim i As Long, j As Long

4         For i = 0 To UBound(Objective)
5             VariableDictionary(Objective(i)(2)) = 1
6         Next i
          
7         For i = 0 To UBound(Constraints)
8             For j = 0 To UBound(Constraints(i)(0))
9                 VariableDictionary(Constraints(i)(0)(j)(2)) = 1
10            Next j
11        Next i
          
12        For i = 0 To UBound(Bounds)
13            VariableDictionary(Bounds(i)(1)) = 1
14        Next i
          
15        For i = 0 To UBound(VariableTypes)
16            VariableDictionary(VariableTypes(i)(1)) = 1
17        Next i

18        For i = 0 To UBound(SemiContinuous)
19            VariableDictionary(SemiContinuous(i)) = 1
20        Next i

21        For i = 0 To UBound(SOS)
22            For j = 0 To UBound(SOS(i)(2))
23                VariableDictionary(SOS(i)(2)(j)(0)) = 1
24            Next j
25        Next i
          
26        FindAllVariables = VariableDictionary.Keys()
          
ExitFunction:
27        If RaiseError Then RethrowError
28        Exit Function

ErrorHandler:
29        If Not ReportError("ReaderFileLP", "FindAllVariables") Then Resume
30        RaiseError = True
31        GoTo ExitFunction
End Function

Function FindObjCoeff(VariableName As Variant, Objective As Variant) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim i As Long
3         For i = 0 To UBound(Objective)
4             If Objective(i)(2) = VariableName Then
5                 FindObjCoeff = Objective(i)(0) & Objective(i)(1)
6                 GoTo ExitFunction
7             End If
8         Next i

ExitFunction:
9         If RaiseError Then RethrowError
10        Exit Function

ErrorHandler:
11        If Not ReportError("ReaderFileLP", "FindObjCoeff") Then Resume
12        RaiseError = True
13        GoTo ExitFunction
End Function

Function GenerateAColumn(VariableName As Variant, Constraints As Variant) As Variant
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim i As Long, j As Long
          Dim AColumn() As Variant
          
3         If UBound(Constraints) = -1 Then
4             GoTo ExitFunction
5         End If
          
6         ReDim AColumn(UBound(Constraints))
7         For i = 0 To UBound(Constraints)
8             For j = 0 To UBound(Constraints(i)(0))
9                 If Constraints(i)(0)(j)(2) = VariableName Then
10                    AColumn(i) = Constraints(i)(0)(j)(0) & Constraints(i)(0)(j)(1)
11                End If
12            Next j
13        Next i
14        GenerateAColumn = AColumn

ExitFunction:
15        If RaiseError Then RethrowError
16        Exit Function

ErrorHandler:
17        If Not ReportError("ReaderFileLP", "GenerateAColumn") Then Resume
18        RaiseError = True
19        GoTo ExitFunction
End Function

Function FindLowerBound(VariableName As Variant, Bounds As Variant) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
3         FindLowerBound = 0
          Dim i As Long
4         For i = 0 To UBound(Bounds)
5             If Bounds(i)(1) = VariableName Then
6                 FindLowerBound = Bounds(i)(0)
7                 GoTo ExitFunction
8             End If
9         Next i

ExitFunction:
10        If RaiseError Then RethrowError
11        Exit Function

ErrorHandler:
12        If Not ReportError("ReaderFileLP", "FindLowerBound") Then Resume
13        RaiseError = True
14        GoTo ExitFunction
End Function

Function FindUpperBound(VariableName As Variant, Bounds As Variant) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
3         FindUpperBound = MAX_DOUBLE
          Dim i As Long
4         For i = 0 To UBound(Bounds)
5             If Bounds(i)(1) = VariableName Then
6                 FindUpperBound = Bounds(i)(2)
7                 GoTo ExitFunction
8             End If
9         Next i

ExitFunction:
10        If RaiseError Then RethrowError
11        Exit Function

ErrorHandler:
12        If Not ReportError("ReaderFileLP", "FindUpperBound") Then Resume
13        RaiseError = True
14        GoTo ExitFunction
End Function

Function FindVariableType(VariableName As Variant, VariableTypes As Variant) As RelationConsts
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim i As Long
3         For i = 0 To UBound(VariableTypes)
4             If VariableTypes(i)(1) = VariableName Then
5                 FindVariableType = VariableTypes(i)(0)
6                 GoTo ExitFunction
7             End If
8         Next i

ExitFunction:
9         If RaiseError Then RethrowError
10        Exit Function

ErrorHandler:
11        If Not ReportError("ReaderFileLP", "FindVariableType") Then Resume
12        RaiseError = True
13        GoTo ExitFunction
End Function

Function FindSemiCon(VariableName As Variant, SemiContinuous As Variant) As Boolean
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
                
3         FindSemiCon = False
          Dim i As Long
4         For i = 0 To UBound(SemiContinuous)
5             If SemiContinuous(i) = VariableName Then FindSemiCon = True
6             GoTo ExitFunction
7         Next i
          
ExitFunction:
8               If RaiseError Then RethrowError
9               Exit Function

ErrorHandler:
10              If Not ReportError("ReaderFileLP", "FindSemiCon") Then Resume
11              RaiseError = True
12              GoTo ExitFunction
End Function

Function FindSOS(VariableName As Variant, SOS As Variant) As Variant
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
3         If UBound(SOS) = -1 Then
4             FindSOS = Array()
5             GoTo ExitFunction
6         End If
          Dim i As Long, j As Long
          Dim SetColumn() As Variant
7         ReDim SetColumn(UBound(SOS))
8         For i = 0 To UBound(SOS)
9             For j = 0 To UBound(SOS(i)(2))
10                If SOS(i)(2)(j)(0) = VariableName Then
11                    SetColumn(i) = SOS(i)(2)(j)(1)
12                End If
13            Next j
14        Next i
15        FindSOS = SetColumn

ExitFunction:
16        If RaiseError Then RethrowError
17        Exit Function

ErrorHandler:
18        If Not ReportError("ReaderFileLP", "FindSOS") Then Resume
19        RaiseError = True
20        GoTo ExitFunction
End Function

Function ConcatArrays(Array1 As Variant, Array2 As Variant) As Variant
          ' Important that this works even if either array is empty (but not both)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim NewArray() As Variant
3         ReDim NewArray(UBound(Array1) + UBound(Array2) + 1)
          Dim i As Long, j As Long
4         For i = 0 To UBound(Array1)
5             NewArray(i) = Array1(i)
6         Next i
7         For j = 0 To UBound(Array2)
8             NewArray(i + j) = Array2(j)
9         Next j
10        ConcatArrays = NewArray

ExitFunction:
11        If RaiseError Then RethrowError
12        Exit Function

ErrorHandler:
13        If Not ReportError("ReaderFileLP", "ConcatArrays") Then Resume
14        RaiseError = True
15        GoTo ExitFunction
End Function

Function WriteToModel(Objective As Variant, Constraints As Variant, VariableTypes As Variant)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim ObjectiveSense As ObjectiveSenseType
3         ObjectiveSense = Objective(0)
          
4         SetObjectiveSense ObjectiveSense
5         SetObjectiveFunctionCell Range("Objective_Function")
6         SetDecisionVariables Range("Decision_Variables")

7         If UBound(Constraints) = -1 Then
8             GoTo SkipConstraints
9         End If

          ' Set constraints
          Dim i As Long, j As Long
          Dim CurrentBlockLHS As Range, CurrentBlockRHS As Range
10        Set CurrentBlockLHS = Range("Constraint_0_LHS")
11        Set CurrentBlockRHS = Range("Constraint_0_RHS")
12        For i = 0 To UBound(Constraints)
          
13            If i <> 0 Then
                  ' If the sense has not changed from the previous constraint
14                If Range("Constraint_" & i & "_Sense") = Range("Constraint_" & i - 1 & "_Sense") Then
15                    Set CurrentBlockLHS = ProperUnion(CurrentBlockLHS, Range("Constraint_" & i & "_LHS"))
16                    Set CurrentBlockRHS = ProperUnion(CurrentBlockRHS, Range("Constraint_" & i & "_RHS"))
17                End If
18            End If
              
              ' If the sense is the last in the block
19            If i = UBound(Constraints) Then GoTo LastBlock ' VBA version of short-circuit OR
20            If Range("Constraint_" & i & "_Sense") <> Range("Constraint_" & i + 1 & "_Sense") Then
LastBlock:
21                SetNamedRangeOnSheet "ConstraintBlock_" & j & "_LHS", CurrentBlockLHS
22                SetNamedRangeOnSheet "ConstraintBlock_" & j & "_RHS", CurrentBlockRHS
23                AddConstraint Range("ConstraintBlock_" & j & "_LHS"), RelationStringToEnum(Range("Constraint_" & i & "_Sense")), Range("ConstraintBlock_" & j & "_RHS")
24                j = j + 1
25                If i <> UBound(Constraints) Then
26                    Set CurrentBlockLHS = Range("Constraint_" & i + 1 & "_LHS")
27                    Set CurrentBlockRHS = Range("Constraint_" & i + 1 & "_RHS")
28                End If
29            End If
              
30        Next i

SkipConstraints:
          ' Set variable bounds
31        AddConstraint Range("Decision_Variables"), RelationGE, Range("Lower_Bounds")
32        AddConstraint Range("Decision_Variables"), RelationLE, Range("Upper_Bounds")
          
          ' Set variable types
          Dim DecisionVariable As Variant
          Dim VariableType As RelationConsts
33        For Each DecisionVariable In VariableTypes
34            If DecisionVariable(0) <> "" Then
35                VariableType = DecisionVariable(0)
36                AddConstraint Range("Variable_" & DecisionVariable(1)), VariableType
37            End If
38        Next DecisionVariable

ExitFunction:
39        If RaiseError Then RethrowError
40        Exit Function

ErrorHandler:
41        If Not ReportError("ReaderFileLP", "WriteToModel") Then Resume
42        RaiseError = True
43        GoTo ExitFunction

End Function

Function SortConstraintsByIneq(Constraints As Variant) As Variant
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          ' 3 types of constraints, many potential elements
          ' VBA doesn't have any native sorting functions, so we make our own
          ' This should(?) be fairly efficient as long as number of constraints >> 3
          
          Dim OutputArray() As Variant
3         ReDim OutputArray(UBound(Constraints))
          
          Dim Element As Variant, ElementType As Integer, ElementTypeCount As Variant
4         ElementTypeCount = Array(0, 0, 0, 0)
          
5         For Each Element In Constraints
6             ElementType = Element(1)
7             ElementTypeCount(ElementType) = ElementTypeCount(ElementType) + 1
8         Next Element
              
9         ElementTypeCount(2) = ElementTypeCount(2) + ElementTypeCount(1)
          Dim CurrentTypeCount As Variant
10        CurrentTypeCount = Array(0, 0, 0, 0)
              
          Dim TypePosition As Long
11        For Each Element In Constraints
12            TypePosition = ElementTypeCount(Element(1) - 1) + CurrentTypeCount(Element(1))
13            OutputArray(TypePosition) = Element
14            CurrentTypeCount(Element(1)) = CurrentTypeCount(Element(1)) + 1
15        Next Element
          
16        SortConstraintsByIneq = OutputArray
          
ExitFunction:
17        If RaiseError Then RethrowError
18        Exit Function

ErrorHandler:
19        If Not ReportError("ReaderFileLP", "SortConstraintsByIneq") Then Resume
20        RaiseError = True
21        GoTo ExitFunction
End Function

Function SplitSOS(ByVal SOSString As String) As Variant
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim SOSArray As Variant
          Dim SetName As String
3         If InStr(SOSString, ":") <> InStr(SOSString, "::") Then
4             RemoveNames SOSString, SetName
5             SetName = SetName & " "
6         Else
7             SetName = ""
8         End If
          
          Dim SetType As String
9         SetType = Left(Trim(SOSString), InStr(Trim(SOSString), "::") - 1)
          
          Dim SetVariables As String
10        SetVariables = Right(SOSString, Len(SOSString) - InStr(SOSString, "::") - 1)
          
11        SplitSOS = Array(SetName, SetType, SetVariables)
          
ExitFunction:
12              If RaiseError Then RethrowError
13              Exit Function

ErrorHandler:
14              If Not ReportError("ReaderFileLP", "SplitSOS") Then Resume
15              RaiseError = True
16              GoTo ExitFunction
End Function

Function ParseSOS(ByVal SOSString As String) As Variant
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim SOSArray As Variant
          ' Turn x1 : 5     x2 : 6
          ' into x1:5 x2:6
          ' so we can split across spaces
3         SOSString = Trim(SOSString)
4         While InStr(SOSString, "  ") > 0
5             SOSString = Replace(SOSString, "  ", " ")
6         Wend
7         SOSString = Replace(SOSString, " :", ":")
8         SOSString = Replace(SOSString, ": ", ":")
9         SOSArray = Split(SOSString, " ")
          
          Dim ParsedArray() As Variant
10        ReDim ParsedArray(UBound(SOSArray))
          
          Dim i As Long, pos As Long
          
11        For i = 0 To UBound(SOSArray)
12            pos = InStr(SOSArray(i), ":")
13            ParsedArray(i) = Array(Left(SOSArray(i), pos - 1), Right(SOSArray(i), Len(SOSArray(i)) - pos))
14        Next i
          
15        ParseSOS = ParsedArray
          
ExitFunction:
16              If RaiseError Then RethrowError
17              Exit Function

ErrorHandler:
18              If Not ReportError("ReaderFileLP", "ParseSOS") Then Resume
19              RaiseError = True
20              GoTo ExitFunction
End Function
