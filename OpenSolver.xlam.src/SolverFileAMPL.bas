Attribute VB_Name = "SolverFileAMPL"
Option Explicit

Public Const AMPLFileName As String = "model.ampl"

Function GetAMPLFilePath(ByRef Path As String) As Boolean
          GetAMPLFilePath = GetTempFilePath(AMPLFileName, Path)
End Function

Sub WriteAMPLFile_Diff(s As COpenSolver, ModelFilePathName As String)
           Dim RaiseError As Boolean
           RaiseError = False
           On Error GoTo ErrorHandler

1777       Open ModelFilePathName For Output As #1 ' supply path with filename
                
           ' Model File - Replace with Data File
1778       Print #1, "# Define our sets, parameters and variables (with names matching those"
1779       Print #1, "# used in defining the data items)"
           
           ' Intialise Variables
           Dim var As Long
1803       For var = 1 To s.NumVars
1804           Print #1, "var "; s.VarName(var); ConvertVarTypeAMPL(s.VarCategory(var), s.SolveRelaxation);
1805           If s.AssumeNonNegativeVars Then
                   ' If no lower bound has been applied then we need to add >= 0
1806               If Not s.VarLowerBounds.Exists(s.VarName(var)) Then
1807                   Print #1, ", >= 0";
1808               End If
1809           End If
1810           Print #1, ";"
1811       Next var
           
1812       Print #1,   ' New line
           
           If Not s.ObjRange Is Nothing Then
               ' Objective function replaced with constraint if
1820           If s.ObjectiveSense = TargetObjective Then
1821               Print #1, "# We have no objective function as the objective must achieve a given target value"
1822               Print #1,
1823               Print #1, "# The objective must achieve a given target value; this constraint enforces this."
1824               Print #1, "subject to TargetConstr:"
                   Print #1, "  "; StrEx(s.ObjectiveTargetValue); " == ";
1826           Else
                   Print #1, ObjectiveSenseToAMPLString(s.ObjectiveSense); "Total_Cost:"
1832               Print #1, "  ";
1834           End If
           
               ' Parameter - Costs
1814           With s.CostCoeffs
                   Dim i As Long
                   For i = 1 To .Count
                       Print #1, StrEx(.Coefficient(i)); " * "; s.VarName(.Index(i)); " ";
                   Next i
               End With
               Print #1, StrEx(s.ObjectiveFunctionConstant); ";"
           End If
           
           ' Subject to Constraints
           Dim row As Long
1835       For row = 1 To s.NumRows
               Dim constraint As Long, instance As Long
               constraint = s.RowToConstraint(row)
1837           instance = s.GetConstraintInstance(row, constraint)
1838           If instance = 1 Then
                   ' We are outputting the first row of a new Excel constraint block; put a comment in the .lp file
1867               Print #1, ' New line
1839               Print #1, "# "; s.ConstraintSummary(constraint)
1840           End If
               
              
1856           If s.SparseA(row).Count = 0 Then
                   ' This constraint must be satisfied trivially!
1858               Print #1, "# (A row with all zero coeffs)";
1862           Else
                   'Output the constraint header
1863               Print #1, "subject to c"; StrEx(row, False); ": ";
                
                   ' Output variables
1841               With s.SparseA(row)
1842                   For i = 1 To .Count
1845                       Print #1, StrEx(.Coefficient(i)); " * "; s.VarName(.Index(i)); " ";
1846                   Next i
1847               End With
               End If
               Print #1, RelationToAMPLString(s.Relation(constraint)); StrEx(s.RHS(row)); ";"
1868       Next row
1869       Print #1, ' New line
           
           ' Run Commands
           Dim AMPLFileSolver As ISolverFileAMPL
           Set AMPLFileSolver = s.Solver
           
1870       Print #1, "# Solve the problem"
1871       Print #1, "option solver "; AMPLFileSolver.AmplSolverName; ";"
           Print #1, "option "; AMPLFileSolver.AmplSolverName; "_options "; _
                      Quote(ParametersToKwargs(s.SolverParameters)); ";"
1872       Print #1, "solve;"

1873       Print #1,   ' New line

           ' Display variables
1874       For var = 1 To s.NumVars
1875           Print #1, "_display "; s.VarName(var); ";"
1876       Next var

           If Not s.ObjRange Is Nothing And Not s.ObjectiveSense = TargetObjective Then
              ' Display objective
               Print #1, "_display Total_Cost;"
           Else
               ' We use the keyword "_display" to know where to begin scanning for variable values and also when to stop scanning.
               ' Even if there is not an objective, we still need to display something so we can read in the variables.
               Print #1, "_display 1;"
           End If

1877       Print #1, "display solve_result_num, solve_result;"
1878       Print #1,   ' New line
           
ExitSub:
           Close #1
           If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
           Exit Sub

ErrorHandler:
           If Not ReportError("SolverFileAMPL", "WriteAMPLFile_Diff") Then Resume
           RaiseError = True
           GoTo ExitSub
End Sub
Sub WriteAMPLFile_Parsed(s As COpenSolver, ModelFilePathName As String)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim m As CModelParsed
          Set m = s.ParsedModel

7188      Open ModelFilePathName For Output As #1
              
          ' Define useful constants
7191      Print #1, "param pi = 4 * atan(1);"
          
7192      Print #1, "# 'Sheet="; s.sheet.Name; "'"
              
          ' Vars
          Dim VarName As String, c As Range
7193      For Each c In s.AdjustableCells
              VarName = ConvertCellToStandardName(c)
7194          Print #1, "var "; VarName; _
                        ConvertVarTypeAMPL(m.VarTypeMap(VarName), s.SolveRelaxation);
              
7195          If s.AssumeNonNegativeVars Then
                  ' If no lower bound has been applied then we need to add >= 0
7196              If Not s.VarLowerBounds.Exists(GetCellName(c)) Then
7197                  Print #1, ", >= 0";
7198              End If
7199          End If
              ' TODO use initial values collection instead
7200          If VarType(c) = vbEmpty Then
7201              Print #1, " := 0";
7202          Else
7203              Print #1, " := " & StrExNoPlus(CDbl(c.Value2));
7204          End If
7205          Print #1, ";"
7206      Next
7207      Print #1, ' New line
          
          Dim Formula As Variant
7208      For Each Formula In m.Formulae
7209          Print #1, "var "; Formula.strAddress; _
                        " := "; StrExNoPlus(Formula.initialValue); ";"
7210      Next Formula
7211      Print #1, ' New line
          
7212      If Not s.ObjRange Is Nothing Then
              Dim objCellName As String
7213          objCellName = ConvertCellToStandardName(s.ObjRange)
              
7214          If s.ObjectiveSense = TargetObjective Then
                  ' Replace objective function with constraint
7215              Print #1, "# We have no objective function as the objective must achieve a given target value"
7216              Print #1, "subject to targetObj:"
7217              Print #1, "  "; s.ObjectiveTargetValue; " == ";
7219          Else
7220              Print #1, ObjectiveSenseToAMPLString(s.ObjectiveSense); "Total_Cost:"
7226          End If
7225          Print #1, objCellName; ";"
              Print #1, ' New line
7227      End If
             
          Dim i As Long
7228      For i = 1 To m.LHSKeys.Count
              Dim strLHS As String, strRel As String, strRHS As String
7229          strLHS = m.LHSKeys(i)
7230          strRel = RelationToAMPLString(m.RELs(i))
7231          strRHS = m.RHSKeys(i)
7232          Print #1, "# Actual constraint: "; strLHS; strRel; strRHS
7233          Print #1, "subject to c"; StrEx(i, False); ":"
7234          Print #1, "    "; strLHS; strRel; strRHS; ";"
7235      Next i
          
7236      For i = 1 To m.Formulae.Count
7237          Print #1, "# Parsed formula for "; m.Formulae(i).strAddress
7238          Print #1, "subject to f"; StrEx(i, False); ":"
7239          Print #1, "    "; m.Formulae(i).strAddress; " == "; m.Formulae(i).strFormulaParsed; ";"
7240      Next i
          
          ' Run Commands
          Dim AMPLFileSolver As ISolverFileAMPL
          Set AMPLFileSolver = s.Solver
          
7241      Print #1, "# Solve the problem"
7242      Print #1, "option solver "; AMPLFileSolver.AmplSolverName; ";"
7243      Print #1, "solve;"
         
          ' Display variables
7244      For Each c In s.AdjustableCells
7246          Print #1, "_display "; ConvertCellToStandardName(c); ";"
7247      Next
          
          ' Display objective
7248      If Not s.ObjRange Is Nothing Then
7249          Print #1, "_display "; objCellName; ";"
7250      Else
              ' We use the keyword "_display" to know where to begin scanning for variable values and also when to stop scanning.
              ' Even if there is not an objective, we still need to display something so we can read in the variables.
7251          Print #1, "_display 1;"
7252      End If
              
          ' Display solving condition
7253      Print #1, "display solve_result_num, solve_result;"

ExitSub:
          Close #1
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("SolverFileAMPL", "WriteAMPLFile_Parsed") Then Resume
          RaiseError = True
          GoTo ExitSub

End Sub

Sub ReadResults_AMPL(s As COpenSolver, solution As String)
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    Dim SolutionStatus As String
    s.SolutionWasLoaded = False

    ' Check logs first
    s.Solver.CheckLog s

    Dim openingParen As Long, closingParen As Long, result As String
    ' Extract the solve status
    openingParen = InStr(solution, "solve_result =")
    SolutionStatus = Mid(solution, openingParen + 1 + Len("solve_result ="))
    
    ' Trim to end of line - marked by line feed char
    SolutionStatus = Left(SolutionStatus, InStr(SolutionStatus, Chr(10)))

    ' Determine Feasibility
    If SolutionStatus Like "unbounded*" Then
        s.SolveStatus = OpenSolverResult.Unbounded
        s.SolveStatusString = "No Solution Found (Unbounded)"
    ElseIf SolutionStatus Like "infeasible*" Then
        s.SolveStatus = OpenSolverResult.Infeasible
        s.SolveStatusString = "No Feasible Solution"
    ElseIf SolutionStatus Like "*limit*" Then  ' Stopped on iterations or time
        s.SolveStatus = OpenSolverResult.LimitedSubOptimal
        s.SolveStatusString = "Stopped on User Limit (Time/Iterations)"
        s.SolutionWasLoaded = True
    ElseIf SolutionStatus Like "solved*" Then
        s.SolveStatus = OpenSolverResult.Optimal
        s.SolveStatusString = "Optimal"
        s.SolutionWasLoaded = True
    Else
        s.SolveStatus = OpenSolverResult.ErrorOccurred
        openingParen = InStr(solution, ">>>")
        If openingParen = 0 Then
            openingParen = InStr(solution, "processing commands.")
            s.SolveStatusString = Mid(solution, openingParen + 1 + Len("processing commands."))
        Else
            closingParen = InStr(solution, "<<<")
            s.SolveStatusString = "Error: " & Mid(solution, openingParen, closingParen - openingParen)
        End If
        s.SolveStatusString = "Neos Returned:" & vbNewLine & vbNewLine & SolutionStatus
    End If

    If s.SolutionWasLoaded Then
        ' Save the solution values
        Dim c As Range, i As Long, VarName As String
        i = 1
        For Each c In s.AdjustableCells
            If s.Solver.ModelType = Parsed Then
                VarName = ConvertCellToStandardName(c)
            Else
                VarName = s.VarName(i)
            End If
            
            openingParen = InStr(solution, VarName)
            closingParen = openingParen + InStr(Mid(solution, openingParen + 1), "_display")
            result = Mid(solution, openingParen + Len(VarName) + 1, Max(closingParen - openingParen - Len(VarName) - 1, 0))
    
            s.VarFinalValue(i) = Val(result)
            s.VarCellName(i) = s.VarName(i)
            i = i + 1
        Next
    End If

ExitSub:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Sub

ErrorHandler:
    If Not ReportError("SolverFileAMPL", "ReadResults_AMPL") Then Resume
    RaiseError = True
    GoTo ExitSub
End Sub

' Given the value of an OpenSolver RelationConst, pick the equivalent AMPL comparison operator
Function ConvertRelationToAMPL(Relation As RelationConsts) As String
7258      Select Case Relation
              Case RelationConsts.RelationLE: ConvertRelationToAMPL = " <= "
7259          Case RelationConsts.RelationEQ: ConvertRelationToAMPL = " == "
7260          Case RelationConsts.RelationGE: ConvertRelationToAMPL = " >= "
7261      End Select
End Function

Function ConvertVarTypeAMPL(intVarType As Long, SolveRelaxation As Boolean) As String
          If SolveRelaxation Then
              Select Case intVarType
              Case VarContinuous, VarInteger
                  ConvertVarTypeAMPL = ""
              Case VarBinary
                  ConvertVarTypeAMPL = ", <= 1, >= 0"
              End Select
          Else
7262          Select Case intVarType
              Case VarContinuous
7263              ConvertVarTypeAMPL = ""
7264          Case VarInteger
7265              ConvertVarTypeAMPL = ", integer"
7266          Case VarBinary
7267              ConvertVarTypeAMPL = ", binary"
7268          End Select
          End If
End Function

Function RelationToAMPLString(rel As RelationConsts) As String
7258      Select Case rel
              Case RelationLE: RelationToAMPLString = " <= "
7259          Case RelationEQ: RelationToAMPLString = " == "
7260          Case RelationGE: RelationToAMPLString = " >= "
7261      End Select
End Function
Function ObjectiveSenseToAMPLString(ObjSense As ObjectiveSenseType) As String
          Select Case ObjSense
          Case MaximiseObjective: ObjectiveSenseToAMPLString = "maximize "
          Case MinimiseObjective: ObjectiveSenseToAMPLString = "minimize "
          End Select
End Function
