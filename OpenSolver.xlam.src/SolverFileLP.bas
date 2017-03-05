Attribute VB_Name = "SolverFileLP"
Option Explicit

Public Const LPFileName As String = "model.lp"

Function GetLPFilePath(ByRef Path As String) As Boolean
1               GetLPFilePath = GetTempFilePath(LPFileName, Path)
End Function

' Output the model to an LP format text file. See http://lpsolve.sourceforge.net/5.5/CPLEX-format.htm
Sub WriteLPFile_Diff(s As COpenSolver, ModelFilePathName As String)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim i As Long, var As Long, row As Long
          
3         Open ModelFilePathName For Output As #1
4         Print #1, "\ Model solved using the solver: "; DisplayName(s.Solver)

5         Print #1, "\ Model for sheet "; s.sheet.Name
          
6         Print #1, "\ Model has"; s.NumConstraints; " Excel constraints giving";
7         Print #1, s.NumRows; " constraint rows and"; s.NumVars; " variables."
          
8         If s.SolveRelaxation And (s.NumDiscreteVars > 0) Then
9             Print #1, "\ (Formulation for relaxed problem)"
10        End If

11        Print #1, ObjectiveSenseToLPString(s.ObjectiveSense)
12        Print #1, "Obj:";
13        If s.ObjectiveSense = TargetObjective Then
              ' We want the objective to achieve some target value; we have no objective; nothing is output
14            Print #1, ' A new line meaning a blank objective; add a comment to this effect next
15            Print #1, "\ We have no objective function as the objective must achieve a given target value"
16            Print #1, ' New line
17            Print #1, "SUBJECT TO"
18            Print #1, "\ The objective must achieve a given target value; this constraint enforces this."
19        Else
              ' Support for the constant in the objective isn't universal, so we don't add it
              'Print #1, " " & StrEx(s.objectivefunctionconstant);
20        End If

          ' Output objective coeffs
          ' NOTE: We output a non-sparse cost vector to ensure that the variables appear in
          '       the same order in the .lp file, which makes matching the results easier
          Dim CostVector() As Double
21        CostVector = s.CostCoeffs.AsVector(s.NumVars)
22        For var = 1 To s.NumVars
23            Print #1, " "; StrEx(CostVector(var)); _
                        " "; GetLPNameFromVarName(s.VarName(var));
24        Next var

25        If s.ObjectiveSense = TargetObjective Then
26            Print #1, " = "; StrEx(s.ObjectiveTargetValue - s.ObjectiveFunctionConstant)
27        Else
28            Print #1, ' New line
29            Print #1, "SUBJECT TO"
30        End If
          
          Dim constraint As Long, instance As Long
31        For row = 1 To s.NumRows
32            constraint = s.RowToConstraint(row)
33            instance = s.GetConstraintInstance(row, constraint)
34            If instance = 1 Then
                  ' We are outputting the first row of a new Excel constraint block; put a comment in the .lp file
35                Print #1, "\ "; s.ConstraintSummary(constraint)
36            End If
         
37            With s.SparseA(row)
38                For i = 1 To .Count
39                    Print #1, " "; StrEx(.Coefficient(i)); _
                                " "; GetLPNameFromVarName(s.VarName(.Index(i)));
40                Next i
41            End With
              
42            If s.SparseA(row).Count = 0 Then
                  ' This constraint must be trivially satisfied!
                  ' We output the row as a comment
43                Print #1, "\ (A row with all zero coeffs)";
44            End If

45            Print #1, RelationToLPString(s.Relation(constraint)) & StrEx(s.RHS(row))
46        Next row

47        Print #1,   ' New line
48        Print #1, "BOUNDS"
49        Print #1,   ' New line
          
          Dim c As Range
50        If s.SolveRelaxation And s.NumBinVars > 0 Then
51            Print #1, "\ (Upper bounds of 1 on the relaxed binary variables)"
52            For Each c In s.BinaryCellsRange
53                Print #1, GetLPNameFromVarName(GetCellName(c)); " <= 1"
54            Next c
55            Print #1, ' New line
56        End If
          
          ' The LP file assumes >=0 bounds on all variables unless we tell it otherwise.
57        If Not s.AssumeNonNegativeVars Then
              ' We need to make all variables FREE variables (i.e. no lower bounds), except for the Binary variables
58            Print #1, "\'Assume Non Negative' is FALSE, so default lower bounds of zero are removed from all non-binary variables."
              Dim NonBinaryCellsRange As Range
59            Set NonBinaryCellsRange = SetDifference(s.AdjustableCells, s.BinaryCellsRange)
60            If Not NonBinaryCellsRange Is Nothing Then
61                For Each c In NonBinaryCellsRange
62                    Print #1, " "; GetLPNameFromVarName(GetCellName(c)); " FREE"
63                Next c
64            End If
65            Print #1,   ' New line
66        Else
              ' If AssumeNonNegative, then we need to remove the implicit >=0 bounds
              ' for any variables with explicit lower bounds, unless they are binary
              Dim VarName As Variant
67            For Each VarName In s.VarLowerBounds.Keys()
68                If s.VarCategory(s.VarNameToIndex(VarName)) <> VarBinary Then
69                    Print #1, " "; GetLPNameFromVarName(CStr(VarName)); " FREE"
70                End If
71            Next VarName
72        End If
          
          ' Output any integer variables
73        If Not s.SolveRelaxation And s.NumIntVars > 0 Then
74            Print #1, "GENERAL"
75            For Each c In s.IntegerCellsRange
76                Print #1, " "; GetLPNameFromVarName(GetCellName(c));
77            Next c
78            Print #1, ' New line
79        End If

          ' Output binary variables
80        If Not s.SolveRelaxation And s.NumBinVars > 0 Then
81            Print #1, "BINARY"
82            For Each c In s.BinaryCellsRange
83                Print #1, " "; GetLPNameFromVarName(GetCellName(c));
84            Next c
85            Print #1, ' New line
86        End If
          
87        Print #1, "END"
          
ExitSub:
88        Close #1
89        If RaiseError Then RethrowError
90        Exit Sub

ErrorHandler:
91        If Not ReportError("SolverFileLP", "WriteLPFile") Then Resume
92        RaiseError = True
93        GoTo ExitSub
End Sub

Function RelationToLPString(rel As RelationConsts) As String
1         RelationToLPString = " " & RelationEnumToString(rel) & " "
End Function

Function ObjectiveSenseToLPString(ObjSense As ObjectiveSenseType) As String
      ' We use Minimise for both minimisation, and also for seeking a target (TargetObjective)
1         Select Case ObjSense
          Case MinimiseObjective, TargetObjective:
2             ObjectiveSenseToLPString = "MINIMIZE"
3         Case MaximiseObjective:
4             ObjectiveSenseToLPString = "MAXIMIZE"
5         End Select
End Function

Function GetLPNameFromVarName(s As String) As String
      ' http://lpsolve.sourceforge.net/5.5/CPLEX-format.htm
      ' The letter E or e, alone or followed by other valid symbols, or followed by another E or e, should be avoided as this notation is reserved for exponential entries. Thus, variables cannot be named e9, E-24, E8cats, or other names that could be interpreted as an exponent. Even variable names such as eels or example can cause a read error, depending on their placement in an input line.
1         If Left(s, 1) = "E" Then
2             GetLPNameFromVarName = "_" & s
3         Else
4             GetLPNameFromVarName = s
5         End If
End Function

Function GetVarNameFromLPName(s As String) As String
      ' Removes any added "_" from the name
1         GetVarNameFromLPName = s
2         If Left(GetVarNameFromLPName, 1) = "_" Then
3             GetVarNameFromLPName = Mid(GetVarNameFromLPName, 2)
4         End If
End Function
