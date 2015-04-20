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
           
           Dim AMPLFileSolver As ISolverFileAMPL
           Set AMPLFileSolver = s.Solver

           Dim i As Long, j As Long, var As Long, row As Long, coeff As Double, c As Range, Line As String
           Dim VarDic As Collection
1773       Set VarDic = New Collection

           Dim commentStart As String  'Character for starting comments for chosen solver
1775       commentStart = "#"

1777       Open ModelFilePathName For Output As #1 ' supply path with filename
                
           ' Model File - Replace with Data File
1778       Print #1, "# Define our sets, parameters and variables (with names matching those"
1779       Print #1, "# used in defining the data items)"
           
           ' Variables
1780       For var = 1 To s.numVars
1781           VarDic.Add "", ValidLPFileVarName(s.VarNames(var))
1782       Next var
           
           ' Sets - Vars
1783       If Not s.SolveRelaxation Then
1784           If Not s.IntegerCellsRange Is Nothing Then
1785               For Each c In s.IntegerCellsRange
1786                   VarDic.Remove (ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False)))
1787                   VarDic.Add ", integer", ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False))
1788               Next c
1789           End If
           End If
           If Not s.BinaryCellsRange Is Nothing Then
1791           For Each c In s.BinaryCellsRange
                   VarDic.Remove (ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False)))
1792               If Not s.SolveRelaxation Then
1793                   VarDic.Add ", binary", ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False))
                   Else
                       VarDic.Add ", <= 1, >= 0", ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False))
                   End If
1794           Next c
1796       End If
           
           Dim constraint As Long
           Dim instance As Long
1797       constraint = 1
           
           ' Reindex bounded variables by relative address so that VarNames(var) can search the collection
           Dim BoundedVariables As New Collection
1798       For Each c In s.AdjustableCells
1799           If TestKeyExists(s.VarLowerBounds, c.Address) Then
1800               BoundedVariables.Add c, c.Address(RowAbsolute:=False, ColumnAbsolute:=False)
1801           End If
1802       Next c
           
           ' Intialise Variables
1803       For var = 1 To s.numVars
1804           Line = "var " & ValidLPFileVarName(s.VarNames(var)) & VarDic.Item(ValidLPFileVarName(s.VarNames(var)))
1805           If s.AssumeNonNegativeVars Then
                   ' If no lower bound has been applied then we need to add >= 0
1806               If Not TestKeyExists(BoundedVariables, s.VarNames(var)) Then
1807                   Line = Line & " >= 0"
1808               End If
1809           End If
1810           Print #1, Line & ";"
1811       Next var
           
           ' Parameter - Costs
1812       Print #1,   ' New line
1813       Line = "  "
1814       For var = 1 To s.numVars
               If Abs(s.CostCoeffs(var)) > EPSILON Then
1815               Line = Line & ValidLPFileVarName(s.VarNames(var)) & "*" & StrEx(s.CostCoeffs(var))
1816               If var < s.numVars Then
1817                   Line = Line & " + "
                   End If
1818           End If
1819       Next var
           Line = Line & " " & StrEx(s.objValue)
           
           ' Objective function replaced with constraint if
1820       If s.ObjectiveSense = TargetObjective Then
1821           Print #1, commentStart & " We have no objective function as the objective must achieve a given target value"
1822           Print #1,
               
1823           Print #1, commentStart & " The objective must achieve a given target value; this constraint enforces this."
1824           Print #1, "subject to TargetConstr:"
1825           Print #1, Line & " = " & StrEx(s.ObjectiveTargetValue) & ";"
1826       Else
               ' Determine objective direction
1827           If s.ObjectiveSense = MaximiseObjective Then
1828              Print #1, "maximize Total_Cost:"
1829           Else
1830              Print #1, "minimize Total_Cost:"
1831           End If
               
1832           Print #1, Line & ";"
1833           Print #1,   ' New line
1834       End If
           
           ' Subject to Constraints
1835       For row = 1 To s.NumRows
1836           Line = "   "
1837           s.GetConstraintFromRow row, constraint, instance  ' Which Excel constraint are we in, and which instance?
1838           If instance = 1 Then
                   ' We are outputting the first row of a new Excel constraint block; put a comment in the .lp file
1839               Print #1, commentStart & " " & s.ConstraintSummary(constraint)
1840           End If
               
               ' Gather variables
1841           With s.SparseA(row)
1842               For i = 1 To .Count
1843                   var = .Index(i)
1844                   coeff = .Coefficient(i)
1845                   Line = Line & StrEx(coeff) & " * " & ValidLPFileVarName(s.VarNames(var)) & " "
1846               Next i
1847           End With
              
              ' Check sense
1848          Line = Line & RelationToAMPLString(s.Relation(row))
              
              ' Add RHS
1855          Line = Line & StrEx(s.RHS(row))
              
1856          If s.SparseA(row).Count = 0 Then
                  ' We have a constraint that does not vary with the decision variables; check it is satisfied
1857              If s.CheckConstantConstraintIsSatisfied(row, constraint, instance, i, j) Then
                      'We output the row as a comment
1858                  Print #1, commentStart & " (A row with all zero coeffs)" & Line & ";"
1859              Else
1860                  GoTo ExitSub
1861              End If
1862          Else
                  'Output the constraint
1863              Print #1, "subject to c" & row & ":"
1864              Print #1, Line
1865              Print #1, ";"
1866          End If
             
1867          Print #1,   ' New line
1868       Next row
           
           ' Run Commands
1869       Print #1,   ' New line
1870       Print #1, commentStart & " Solve the problem"
1871       Print #1, "option solver " & AMPLFileSolver.AmplSolverName & ";"
           Print #1, "option " & AMPLFileSolver.AmplSolverName & "_options """ & ParametersToKwargs(s.SolverParameters) & """;"
1872       Print #1, "solve;"
1873       Print #1,   ' New line
           ' Display variables
1874       For var = 1 To s.numVars
1875           Print #1, "_display " & ValidLPFileVarName(s.VarNames(var)) & ";"
1876       Next var

           If Not s.ObjRange Is Nothing And Not s.ObjectiveSense = TargetObjective Then
              ' Display objective
               WriteToFile 1, "_display Total_Cost;" & vbNewLine
           Else
               ' We use the keyword "_display" to know where to begin scanning for variable values and also when to stop scanning.
               ' Even if there is not an objective, we still need to display something so we can read in the variables.
               WriteToFile 1, "_display 1;" & vbNewLine
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
          
          Dim m As CModelParsed, AMPLFileSolver As ISolverFileAMPL
          Set m = s.ParsedModel
          Set AMPLFileSolver = s.Solver

7188      Open ModelFilePathName For Output As #1

          ' Note - We can use the following code on its own to produce a mod file
7189      WriteToFile 1, "# Define our sets, parameters and variables (with names matching those"
7190      WriteToFile 1, "# used in defining the data items)"
              
          ' Define useful constants
7191      WriteToFile 1, "param pi = 4 * atan(1);"
          
7192      WriteToFile 1, "# 'Sheet=" + s.sheetName + "'"
             
          Dim Line As String
              
          ' Vars
          ' Initialise each variable independently
          Dim VarName As String, curVarType As Long, c As Range
7193      For Each c In s.AdjustableCells
              VarName = ConvertCellToStandardName(c)
7194          Line = "var " & VarName

              ' Output variable types if needed
              curVarType = m.VarTypeMap(VarName)
              If s.SolveRelaxation Then
                  If curVarType = VarBinary Then
                      Line = Line & " >= 0 <= 1"
                  End If
              Else
                  Line = Line & ConvertVarTypeAMPL(curVarType)
              End If
              
7195          If s.AssumeNonNegativeVars Then
                  ' If no lower bound has been applied then we need to add >= 0
7196              If Not TestKeyExists(s.VarLowerBounds, c.Address(RowAbsolute:=False, ColumnAbsolute:=False)) Then
7197                  Line = Line & " >= 0"
7198              End If
7199          End If
7200          If VarType(c) = vbEmpty Then
7201              Line = Line & " := 0"
7202          Else
7203              Line = Line & " := " & StrExNoPlus(CDbl(c))
7204          End If
7205          WriteToFile 1, Line & ";"
7206      Next
7207      WriteToFile 1, ""
          
          Dim Formula As Variant
7208      For Each Formula In m.Formulae
7209          WriteToFile 1, "var " & Formula.strAddress & " := " & StrExNoPlus(Formula.initialValue) & ";"
7210      Next Formula
7211      WriteToFile 1, ""
          
7212      If Not s.ObjRange Is Nothing Then
              Dim objCellName As String
7213          objCellName = ConvertCellToStandardName(s.ObjRange)
              
7214          If s.ObjectiveSense = TargetObjective Then
                  ' Replace objective function with constraint
7215              WriteToFile 1, "# We have no objective function as the objective must achieve a given target value"
7216              WriteToFile 1, "subject to targetObj:"
7217              WriteToFile 1, "    " & objCellName & " == " & s.ObjectiveTargetValue & ";"
7218              WriteToFile 1, vbNewLine
7219          Else
                  ' Determine objective direction
7220              If s.ObjectiveSense = MaximiseObjective Then
7221                 WriteToFile 1, "maximize Total_Cost:"
7222              Else
7223                 WriteToFile 1, "minimize Total_Cost:"
7224              End If
                  
7225              WriteToFile 1, "    " & objCellName & ";" & vbNewLine
7226          End If
7227      End If
             
          Dim i As Long
7228      For i = 1 To m.LHSKeys.Count
              Dim strLHS As String, strRel As String, strRHS As String
7229          strLHS = m.LHSKeys(i)
7230          strRel = RelationToAMPLString(m.RELs(i))
7231          strRHS = m.RHSKeys(i)
7232          WriteToFile 1, "# Actual constraint: " & strLHS & strRel & strRHS
7233          WriteToFile 1, "subject to c" & i & ":"
7234          WriteToFile 1, "    " & strLHS & strRel & strRHS & ";" & vbNewLine
7235      Next i
          
7236      For i = 1 To m.Formulae.Count
7237          WriteToFile 1, "# Parsed formula for " & m.Formulae(i).strAddress
7238          WriteToFile 1, "subject to f" & i & ":"
7239          WriteToFile 1, "    " & m.Formulae(i).strAddress & " == " & m.Formulae(i).strFormulaParsed & ";" & vbNewLine
7240      Next i
          
          ' Run Commands
7241      WriteToFile 1, "# Solve the problem"
7242      WriteToFile 1, "option solver " & AMPLFileSolver.AmplSolverName & ";"
 
7243      WriteToFile 1, "solve;" & vbNewLine
         
          Dim cellName As String
          ' Display variables
7244      For Each c In s.AdjustableCells
7245          cellName = ConvertCellToStandardName(c)
7246          WriteToFile 1, "_display " & cellName & ";"
7247      Next
          
7248      If Not s.ObjRange Is Nothing Then
              ' Display objective
7249          WriteToFile 1, "_display " & objCellName & ";" & vbNewLine
7250      Else
              ' We use the keyword "_display" to know where to begin scanning for variable values and also when to stop scanning.
              ' Even if there is not an objective, we still need to display something so we can read in the variables.
7251          WriteToFile 1, "_display 1;" & vbNewLine
7252      End If
              
          ' Display solving condition
7253      WriteToFile 1, "display solve_result_num, solve_result;"

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
                VarName = ValidLPFileVarName(s.VarNames(i))
            End If
            
            openingParen = InStr(solution, VarName)
            closingParen = openingParen + InStr(Mid(solution, openingParen + 1), "_display")
            result = Mid(solution, openingParen + Len(VarName) + 1, Max(closingParen - openingParen - Len(VarName) - 1, 0))
    
            s.FinalVarValue(i) = Val(result)
            s.VarCell(i) = s.VarNames(i)
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

Function ConvertVarTypeAMPL(intVarType As Long) As String
7262      Select Case intVarType
          Case VarContinuous
7263          ConvertVarTypeAMPL = ""
7264      Case VarInteger
7265          ConvertVarTypeAMPL = ", integer"
7266      Case VarBinary
7267          ConvertVarTypeAMPL = ", binary"
7268      End Select
End Function

Function RelationToAMPLString(rel As RelationConsts) As String
7258      Select Case rel
              Case RelationLE: RelationToAMPLString = " <= "
7259          Case RelationEQ: RelationToAMPLString = " == "
7260          Case RelationGE: RelationToAMPLString = " >= "
7261      End Select
End Function
