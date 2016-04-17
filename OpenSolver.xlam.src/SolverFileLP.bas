Attribute VB_Name = "SolverFileLP"
Option Explicit

Public Const LPFileName As String = "model.lp"

Function GetLPFilePath(ByRef Path As String) As Boolean
          GetLPFilePath = GetTempFilePath(LPFileName, Path)
End Function

' Output the model to an LP format text file. See http://lpsolve.sourceforge.net/5.5/CPLEX-format.htm
Sub WriteLPFile_Diff(s As COpenSolver, ModelFilePathName As String)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim i As Long, j As Long, var As Long, row As Long, coeff As Double
          
          ' Track which variables we have printed so far - only needed when AssumeNonNegativeVars is true
          ' This is so we can print all unused variables as FREE for sensitivity analysis
          ' When AssumeNonNegativeVars is false, we print all vars as FREE anyway, so this isn't needed
          Dim UsedVariables As Collection, VarName As String
          If s.AssumeNonNegativeVars Then
              Set UsedVariables = New Collection
          End If
          
1615      Open ModelFilePathName For Output As #1
1616      Print #1, "\ Model solved using the solver: "; DisplayName(s.Solver)

1617      Print #1, "\ Model for sheet "; s.sheet.Name
          
1619      Print #1, "\ Model has"; s.NumConstraints; " Excel constraints giving";
          Print #1, s.NumRows; " constraint rows and"; s.NumVars; " variables."
          
1620      If s.SolveRelaxation And (s.NumDiscreteVars > 0) Then
1621          Print #1, "\ (Formulation for relaxed problem)"
1622      End If

1623      Print #1, ObjectiveSenseToLPString(s.ObjectiveSense)
1624      Print #1, "Obj:";
1625      If s.ObjectiveSense = TargetObjective Then
              ' We want the objective to achieve some target value; we have no objective; nothing is output
1626          Print #1, ' A new line meaning a blank objective; add a comment to this effect next
1628          Print #1, "\ We have no objective function as the objective must achieve a given target value"
1629      Else
1631          For var = 1 To s.NumVars
1632              'If Abs(s.CostCoeffs(var)) > EPSILON Then
                      VarName = GetLPNameFromVarName(s.VarName(var))
                      Print #1, " "; StrEx(s.CostCoeffs(var)); " "; VarName;
                      AddUsedVariable s, VarName, UsedVariables
                  'End If
              Next var
              'If Abs(objValue) > EPSILON Then
              '    Print #1, " " & StrEx(objValue);
              'End If
1634          Print #1, ' New line
1635      End If
1636      Print #1, ' New line
1637      Print #1, "SUBJECT TO"
          
          ' If we are seeking a specific objective value, we add this as a constraint
1638      If s.ObjectiveSense = TargetObjective Then
1639          Print #1, "\ The objective must achieve a given target value; this constraint enforces this."
1640          For var = 1 To s.NumVars
                  VarName = GetLPNameFromVarName(s.VarName(var))
                  Print #1, " "; StrEx(s.CostCoeffs(var)); " "; VarName;
                  AddUsedVariable s, VarName, UsedVariables
1645          Next var
1646          Print #1, " = "; StrEx(s.ObjectiveTargetValue)
1653      End If
          
          Dim constraint As Long, instance As Long
1655      For row = 1 To s.NumRows
              constraint = s.RowToConstraint(row)
              instance = s.GetConstraintInstance(row, constraint)
1657          If instance = 1 Then
                  ' We are outputting the first row of a new Excel constraint block; put a comment in the .lp file
1658              Print #1, "\ "; s.ConstraintSummary(constraint)
1659          End If
         
1660          With s.SparseA(row)
1661              For i = 1 To .Count
                      VarName = GetLPNameFromVarName(s.VarName(.Index(i)))
                      Print #1, " "; StrEx(.Coefficient(i)); " "; VarName;
                      AddUsedVariable s, VarName, UsedVariables
1665              Next i
1666          End With
              
1667          If s.SparseA(row).Count = 0 Then
                  ' This constraint must be trivially satisfied!
                  ' We output the row as a comment
1669              Print #1, "\ (A row with all zero coeffs)";
1673          End If
1674
1681          Print #1, RelationToLPString(s.Relation(constraint)) & StrEx(s.RHS(row))
1682      Next row

1683      Print #1,   ' New line
1692      Print #1, "BOUNDS"
          Print #1,   ' New line
          
          Dim c As Range
1684      If s.SolveRelaxation And s.NumBinVars > 0 Then
1685          Print #1, "\ (Upper bounds of 1 on the relaxed binary variables)"
1686          For Each c In s.BinaryCellsRange
                  VarName = GetLPNameFromVarName(GetCellName(c))
1687              Print #1, VarName; " <= 1"
1688          Next c
1689          Print #1, ' New line
1690      End If
          
          ' The LP file assumes >=0 bounds on all variables unless we tell it otherwise.
1691      If Not s.AssumeNonNegativeVars Then
              ' We need to make all variables FREE variables (i.e. no lower bounds), except for the Binary variables
1693          Print #1, "\'Assume Non Negative' is FALSE, so default lower bounds of zero are removed from all non-binary variables."
              Dim NonBinaryCellsRange As Range
              Set NonBinaryCellsRange = SetDifference(s.AdjustableCells, s.BinaryCellsRange)
1695          If Not NonBinaryCellsRange Is Nothing Then
                  For Each c In NonBinaryCellsRange
1696                  Print #1, " "; GetLPNameFromVarName(GetCellName(c)); " FREE"
1697              Next c
              End If
1705          Print #1,   ' New line
1706      Else
              ' If AssumeNonNegative, then we need to apply lower bounds to any variables without explicit lower bounds.
              ' However, we don't make Binary variables free
1707          For var = 1 To s.NumVars
                  VarName = GetLPNameFromVarName(s.VarName(var))
1708              If TestKeyExists(s.VarLowerBounds, s.VarName(var)) Or _
                     Not TestKeyExists(UsedVariables, VarName) Then
                      If s.VarCategory(var) <> VarBinary Then
                          Print #1, " "; VarName; " FREE"
                      End If
                  End If
1711          Next var
1728      End If
          
          ' Output any integer variables
1729      If Not s.SolveRelaxation And s.NumIntVars > 0 Then
1750          Print #1, "GENERAL"
1751          For Each c In s.IntegerCellsRange
1752              Print #1, " "; GetLPNameFromVarName(GetCellName(c));
1753          Next c
1755          Print #1, ' New line
1754      End If

          ' Output binary variables
1757      If Not s.SolveRelaxation And s.NumBinVars > 0 Then
1758          Print #1, "BINARY"
1759          For Each c In s.BinaryCellsRange
1760              Print #1, " "; GetLPNameFromVarName(GetCellName(c));
1761          Next c
1762          Print #1, ' New line
1763      End If
          
1765      Print #1, "END"
          
ExitSub:
          Close #1
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("SolverFileLP", "WriteLPFile") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Function RelationToLPString(rel As RelationConsts) As String
    RelationToLPString = " " & RelationEnumToString(rel) & " "
End Function

Function ObjectiveSenseToLPString(ObjSense As ObjectiveSenseType) As String
' We use Minimise for both minimisation, and also for seeking a target (TargetObjective)
    Select Case ObjSense
    Case MinimiseObjective, TargetObjective:
        ObjectiveSenseToLPString = "MINIMIZE"
    Case MaximiseObjective:
        ObjectiveSenseToLPString = "MAXIMIZE"
    End Select
End Function

Sub AddUsedVariable(s As COpenSolver, VarName As String, UsedVariables As Collection)
    If s.AssumeNonNegativeVars Then
        If Not TestKeyExists(UsedVariables, VarName) Then UsedVariables.Add VarName, VarName
    End If
End Sub

Function GetLPNameFromVarName(s As String) As String
' http://lpsolve.sourceforge.net/5.5/CPLEX-format.htm
' The letter E or e, alone or followed by other valid symbols, or followed by another E or e, should be avoided as this notation is reserved for exponential entries. Thus, variables cannot be named e9, E-24, E8cats, or other names that could be interpreted as an exponent. Even variable names such as eels or example can cause a read error, depending on their placement in an input line.
338       If Left(s, 1) = "E" Then
339           GetLPNameFromVarName = "_" & s
340       Else
341           GetLPNameFromVarName = s
342       End If
End Function

Function GetVarNameFromLPName(s As String) As String
' Removes any added "_" from the name
    GetVarNameFromLPName = s
    If Left(GetVarNameFromLPName, 1) = "_" Then
        GetVarNameFromLPName = Mid(GetVarNameFromLPName, 2)
    End If
End Function
