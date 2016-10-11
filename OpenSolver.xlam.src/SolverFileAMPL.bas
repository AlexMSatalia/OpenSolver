Attribute VB_Name = "SolverFileAMPL"
Option Explicit

Public Const AMPLFileName As String = "model.ampl"

Function GetAMPLFilePath(ByRef Path As String) As Boolean
1               GetAMPLFilePath = GetTempFilePath(AMPLFileName, Path)
End Function

Sub WriteAMPLFile_Diff(s As COpenSolver, ModelFilePathName As String)
           Dim RaiseError As Boolean
1          RaiseError = False
2          On Error GoTo ErrorHandler

3          Open ModelFilePathName For Output As #1 ' supply path with filename
                
           ' Model File - Replace with Data File
4          Print #1, "# Define our sets, parameters and variables (with names matching those"
5          Print #1, "# used in defining the data items)"
           
           ' Intialise Variables
           Dim var As Long
6          For var = 1 To s.NumVars
7              Print #1, "var "; s.VarName(var); ConvertVarTypeAMPL(s.VarCategory(var), s.SolveRelaxation);
8              If s.AssumeNonNegativeVars Then
                   ' If no lower bound has been applied then we need to add >= 0
9                  If Not s.VarLowerBounds.Exists(s.VarName(var)) Then
10                     Print #1, ", >= 0";
11                 End If
12             End If
13             If s.InitialSolutionIsValid Then
14                 Print #1, ", := "; s.VarInitialValue(var);
15             End If
16             Print #1, ";"
17         Next var
           
18         Print #1,   ' New line
           
19         If Not s.ObjRange Is Nothing Then
               ' Objective function replaced with constraint if
20             If s.ObjectiveSense = TargetObjective Then
21                 Print #1, "# We have no objective function as the objective must achieve a given target value"
22                 Print #1,
23                 Print #1, "# The objective must achieve a given target value; this constraint enforces this."
24                 Print #1, "subject to TargetConstr:"
25                 Print #1, "  "; StrEx(s.ObjectiveTargetValue); " == ";
26             Else
27                 Print #1, ObjectiveSenseToAMPLString(s.ObjectiveSense); "Total_Cost:"
28                 Print #1, "  ";
29             End If
           
               ' Parameter - Costs
30             With s.CostCoeffs
                   Dim i As Long
31                 For i = 1 To .Count
32                     Print #1, StrEx(.Coefficient(i)); " * "; s.VarName(.Index(i)); " ";
33                 Next i
34             End With
35             Print #1, StrEx(s.ObjectiveFunctionConstant); ";"
36         End If
           
           ' Subject to Constraints
           Dim row As Long
37         For row = 1 To s.NumRows
               Dim constraint As Long, instance As Long
38             constraint = s.RowToConstraint(row)
39             instance = s.GetConstraintInstance(row, constraint)
40             If instance = 1 Then
                   ' We are outputting the first row of a new Excel constraint block; put a comment in the .lp file
41                 Print #1, ' New line
42                 Print #1, "# "; s.ConstraintSummary(constraint)
43             End If
               
              
44             If s.SparseA(row).Count = 0 Then
                   ' This constraint must be satisfied trivially!
45                 Print #1, "# (A row with all zero coeffs)";
46             Else
                   'Output the constraint header
47                 Print #1, "subject to c"; StrEx(row, False); ": ";
                
                   ' Output variables
48                 With s.SparseA(row)
49                     For i = 1 To .Count
50                         Print #1, StrEx(.Coefficient(i)); " * "; s.VarName(.Index(i)); " ";
51                     Next i
52                 End With
53             End If
54             Print #1, RelationToAMPLString(s.Relation(constraint)); StrEx(s.RHS(row)); ";"
55         Next row
56         Print #1, ' New line
           
           ' Run Commands
           Dim AMPLFileSolver As ISolverFileAMPL
57         Set AMPLFileSolver = s.Solver
           
58         Print #1, "# Solve the problem"
59         Print #1, "option solver "; AMPLFileSolver.AmplSolverName; ";"
60         Print #1, "option "; AMPLFileSolver.AmplSolverName; "_options "; _
                      Quote(ParametersToKwargs(s.SolverParameters)); ";"
61         Print #1, "solve;"

62         Print #1,   ' New line

           ' Display variables
63         For var = 1 To s.NumVars
64             Print #1, "_display "; s.VarName(var); ";"
65         Next var

66         If Not s.ObjRange Is Nothing And Not s.ObjectiveSense = TargetObjective Then
              ' Display objective
67             Print #1, "_display Total_Cost;"
68         Else
               ' We use the keyword "_display" to know where to begin scanning for variable values and also when to stop scanning.
               ' Even if there is not an objective, we still need to display something so we can read in the variables.
69             Print #1, "_display 1;"
70         End If

71         Print #1, "display solve_result_num, solve_result;"
72         Print #1,   ' New line
           
ExitSub:
73         Close #1
74         If RaiseError Then RethrowError
75         Exit Sub

ErrorHandler:
76         If Not ReportError("SolverFileAMPL", "WriteAMPLFile_Diff") Then Resume
77         RaiseError = True
78         GoTo ExitSub
End Sub
Sub WriteAMPLFile_Parsed(s As COpenSolver, ModelFilePathName As String)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim m As CModelParsed
3         Set m = s.ParsedModel

4         Open ModelFilePathName For Output As #1
              
          ' Define useful constants
5         Print #1, "param pi = 4 * atan(1);"
          
6         Print #1, "# 'Sheet="; s.sheet.Name; "'"
              
          ' Vars
          Dim VarName As String, c As Range
7         For Each c In s.AdjustableCells
8             VarName = ConvertCellToStandardName(c)
9             Print #1, "var "; VarName; _
                        ConvertVarTypeAMPL(m.VarTypeMap(VarName), s.SolveRelaxation);
              
10            If s.AssumeNonNegativeVars Then
                  ' If no lower bound has been applied then we need to add >= 0
11                If Not s.VarLowerBounds.Exists(GetCellName(c)) Then
12                    Print #1, ", >= 0";
13                End If
14            End If
15            If s.InitialSolutionIsValid Then
16                Print #1, ", := "; s.VarInitialValue(s.VarNameToIndex(GetCellName(c)));
17            End If
18            Print #1, ";"
19        Next
20        Print #1, ' New line
          
          Dim Formula As Variant
21        For Each Formula In m.Formulae
22            Print #1, "var "; Formula.strAddress; _
                        " := "; StrExNoPlus(Formula.initialValue); ";"
23        Next Formula
24        Print #1, ' New line
          
25        If Not s.ObjRange Is Nothing Then
              Dim objCellName As String
26            objCellName = ConvertCellToStandardName(s.ObjRange)
              
27            If s.ObjectiveSense = TargetObjective Then
                  ' Replace objective function with constraint
28                Print #1, "# We have no objective function as the objective must achieve a given target value"
29                Print #1, "subject to targetObj:"
30                Print #1, "  "; s.ObjectiveTargetValue; " == ";
31            Else
32                Print #1, ObjectiveSenseToAMPLString(s.ObjectiveSense); "Total_Cost:"
33            End If
34            Print #1, objCellName; ";"
35            Print #1, ' New line
36        End If
             
          Dim i As Long
37        For i = 1 To m.LHSKeys.Count
              Dim strLHS As String, strRel As String, strRHS As String
38            strLHS = m.LHSKeys(i)
39            strRel = RelationToAMPLString(m.RELs(i))
40            strRHS = m.RHSKeys(i)
41            Print #1, "# Actual constraint: "; strLHS; strRel; strRHS
42            Print #1, "subject to c"; StrEx(i, False); ":"
43            Print #1, "    "; strLHS; strRel; strRHS; ";"
44        Next i
          
45        For i = 1 To m.Formulae.Count
46            Print #1, "# Parsed formula for "; m.Formulae(i).strAddress
47            Print #1, "subject to f"; StrEx(i, False); ":"
48            Print #1, "    "; m.Formulae(i).strAddress; " == "; m.Formulae(i).strFormulaParsed; ";"
49        Next i
          
          ' Run Commands
          Dim AMPLFileSolver As ISolverFileAMPL
50        Set AMPLFileSolver = s.Solver
          
51        Print #1, "# Solve the problem"
52        Print #1, "option solver "; AMPLFileSolver.AmplSolverName; ";"
53        Print #1, "solve;"
         
          ' Display variables
54        For Each c In s.AdjustableCells
55            Print #1, "_display "; ConvertCellToStandardName(c); ";"
56        Next
          
          ' Display objective
57        If Not s.ObjRange Is Nothing Then
58            Print #1, "_display "; objCellName; ";"
59        Else
              ' We use the keyword "_display" to know where to begin scanning for variable values and also when to stop scanning.
              ' Even if there is not an objective, we still need to display something so we can read in the variables.
60            Print #1, "_display 1;"
61        End If
              
          ' Display solving condition
62        Print #1, "display solve_result_num, solve_result;"

ExitSub:
63        Close #1
64        If RaiseError Then RethrowError
65        Exit Sub

ErrorHandler:
66        If Not ReportError("SolverFileAMPL", "WriteAMPLFile_Parsed") Then Resume
67        RaiseError = True
68        GoTo ExitSub

End Sub

Sub ReadResults_AMPL(s As COpenSolver, solution As String)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim SolutionStatus As String
3         s.SolutionWasLoaded = False

          ' Check logs first
4         s.Solver.CheckLog s

          Dim openingParen As Long, closingParen As Long, result As String
          ' Extract the solve status
5         openingParen = InStr(solution, "solve_result =")
6         SolutionStatus = Mid(solution, openingParen + 1 + Len("solve_result ="))
          
          ' Trim to end of line - marked by line feed char
7         SolutionStatus = Left(SolutionStatus, InStr(SolutionStatus, Chr(10)))

          ' Determine Feasibility
8         If SolutionStatus Like "unbounded*" Then
9             s.SolveStatus = OpenSolverResult.Unbounded
10            s.SolveStatusString = "No Solution Found (Unbounded)"
11        ElseIf SolutionStatus Like "infeasible*" Then
12            s.SolveStatus = OpenSolverResult.Infeasible
13            s.SolveStatusString = "No Feasible Solution"
14        ElseIf SolutionStatus Like "*limit*" Then  ' Stopped on iterations or time
15            s.SolveStatus = OpenSolverResult.LimitedSubOptimal
16            s.SolveStatusString = "Stopped on User Limit (Time/Iterations)"
17            s.SolutionWasLoaded = True
18        ElseIf SolutionStatus Like "solved*" Then
19            s.SolveStatus = OpenSolverResult.Optimal
20            s.SolveStatusString = "Optimal"
21            s.SolutionWasLoaded = True
22        Else
23            s.SolveStatus = OpenSolverResult.ErrorOccurred
24            openingParen = InStr(solution, ">>>")
25            If openingParen = 0 Then
26                openingParen = InStr(solution, "processing commands.")
27                s.SolveStatusString = Mid(solution, openingParen + 1 + Len("processing commands."))
28            Else
29                closingParen = InStr(solution, "<<<")
30                s.SolveStatusString = "Error: " & Mid(solution, openingParen, closingParen - openingParen)
31            End If
32            s.SolveStatusString = "Neos Returned:" & vbNewLine & vbNewLine & SolutionStatus
33        End If

34        If s.SolutionWasLoaded Then
              ' Save the solution values
              Dim c As Range, i As Long, VarName As String
35            i = 1
36            For Each c In s.AdjustableCells
37                If s.Solver.ModelType = Parsed Then
38                    VarName = ConvertCellToStandardName(c)
39                Else
40                    VarName = s.VarName(i)
41                End If
                  
42                openingParen = InStr(solution, VarName)
43                closingParen = openingParen + InStr(Mid(solution, openingParen + 1), "_display")
44                result = Mid(solution, openingParen + Len(VarName) + 1, Max(closingParen - openingParen - Len(VarName) - 1, 0))
          
45                s.VarFinalValue(i) = Val(result)
46                s.VarCellName(i) = s.VarName(i)
47                i = i + 1
48            Next
49        End If

ExitSub:
50        If RaiseError Then RethrowError
51        Exit Sub

ErrorHandler:
52        If Not ReportError("SolverFileAMPL", "ReadResults_AMPL") Then Resume
53        RaiseError = True
54        GoTo ExitSub
End Sub

' Given the value of an OpenSolver RelationConst, pick the equivalent AMPL comparison operator
Function ConvertRelationToAMPL(Relation As RelationConsts) As String
1         Select Case Relation
              Case RelationConsts.RelationLE: ConvertRelationToAMPL = " <= "
2             Case RelationConsts.RelationEQ: ConvertRelationToAMPL = " == "
3             Case RelationConsts.RelationGE: ConvertRelationToAMPL = " >= "
4         End Select
End Function

Function ConvertVarTypeAMPL(intVarType As Long, SolveRelaxation As Boolean) As String
1         If SolveRelaxation Then
2             Select Case intVarType
              Case VarContinuous, VarInteger
3                 ConvertVarTypeAMPL = vbNullString
4             Case VarBinary
5                 ConvertVarTypeAMPL = ", <= 1, >= 0"
6             End Select
7         Else
8             Select Case intVarType
              Case VarContinuous
9                 ConvertVarTypeAMPL = vbNullString
10            Case VarInteger
11                ConvertVarTypeAMPL = ", integer"
12            Case VarBinary
13                ConvertVarTypeAMPL = ", binary"
14            End Select
15        End If
End Function

Function RelationToAMPLString(rel As RelationConsts) As String
1         Select Case rel
              Case RelationLE: RelationToAMPLString = " <= "
2             Case RelationEQ: RelationToAMPLString = " == "
3             Case RelationGE: RelationToAMPLString = " >= "
4         End Select
End Function
Function ObjectiveSenseToAMPLString(ObjSense As ObjectiveSenseType) As String
1               Select Case ObjSense
                Case MaximiseObjective: ObjectiveSenseToAMPLString = "maximize "
2               Case MinimiseObjective: ObjectiveSenseToAMPLString = "minimize "
3               End Select
End Function
