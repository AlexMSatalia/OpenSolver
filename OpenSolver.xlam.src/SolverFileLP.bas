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
          Dim commentStart As String  'Character for starting comments for chosen solver
1613      commentStart = "\"
          
          ' Track which variables we have printed so far - only needed when AssumeNonNegativeVars is true
          ' This is so we can print all unused variables as FREE for sensitivity analysis
          ' When AssumeNonNegativeVars is false, we print all vars as FREE anyway, so this isn't needed
          Dim UsedVariables As Collection, VarName As String
          If s.AssumeNonNegativeVars Then
              Set UsedVariables = New Collection
          End If
          
1615      Open ModelFilePathName For Output As #1 ' supply path with filename
1616      Print #1, commentStart & " Model solved using the solver '" & DisplayName(s.Solver) & "'"
1617      Print #1, commentStart & " Model for sheet " & s.sheet.Name
1619      Print #1, commentStart & " Model has " & s.NumConstraints & " Excel constraints giving " & s.NumRows & " constraint rows and " & s.NumVars & " variables."
1620      If s.SolveRelaxation And (s.NumBinVars > 0) Then
1621          Print #1, commentStart & " (Formulation for relaxed problem)"
1622      End If
1623      Print #1, IIf(s.ObjectiveSense = MaximiseObjective, "MAXIMIZE", "MINIMIZE")   ' We use Minimise for both minimisation, and also for seeking a target (TargetObjective)
1624      Print #1, "Obj:";
1625      If s.ObjectiveSense = TargetObjective Then
              ' We want the objective to achieve some target value; we have no objective; nothing is output
1626          Print #1, ' A new line meaning a blank objective; add a comment to this effect next
1628          Print #1, commentStart & " We have no objective function as the objective must achieve a given target value"
1629      Else
1631          For var = 1 To s.NumVars
1632              'If Abs(s.CostCoeffs(var)) > EPSILON Then
                      VarName = ValidLPFileVarName(s.VarNames(var))
                      If s.AssumeNonNegativeVars Then
                          If Not TestKeyExists(UsedVariables, VarName) Then UsedVariables.Add VarName, VarName
                      End If
                      Print #1, " " & StrEx(s.CostCoeffs(var)) & " " & VarName;
                  'End If
              Next var
              'If Abs(objValue) > EPSILON Then
              '    Print #1, " " & StrEx(objValue);
              'End If
1634          Print #1,   ' New line
1635      End If
1636      Print #1,   ' New line
1637      Print #1, "SUBJECT TO"
          
          ' If we are seeking a specific objective value, we add this as a constraint
1638      If s.ObjectiveSense = TargetObjective Then
1639          Print #1, commentStart & " The objective must achieve a given target value; this constraint enforces this."
              Dim NonTrivialObjective As Boolean
1640          For var = 1 To s.NumVars
1641              If Abs(s.CostCoeffs(var)) > EPSILON Then
                      VarName = ValidLPFileVarName(s.VarNames(var))
                      If s.AssumeNonNegativeVars Then
                          If Not TestKeyExists(UsedVariables, VarName) Then UsedVariables.Add VarName, VarName
1642                  End If
                      Print #1, " " & StrEx(s.CostCoeffs(var)) & " " & VarName;
1643                  NonTrivialObjective = True
1644              End If
1645          Next var
1646          Print #1, " = " & StrEx(s.ObjectiveTargetValue)
1647          If Not NonTrivialObjective And s.ObjectiveTargetValue <> 0 Then
1648              s.SolveStatus = OpenSolverResult.Infeasible
1649              s.SolveStatusString = "Infeasible Objective Target"
1650              s.SolveStatusComment = "The model's objective cell does not depend on the decision variables" & _
                                         " and so cannot be adjusted to achieve the target value" & s.ObjectiveTargetValue & "."
1651              GoTo ExitSub
1652          End If
1653      End If
          
          Dim constraint As Long, instance As Long
1655      For row = 1 To s.NumRows
              constraint = s.RowToConstraint(row)
              instance = s.GetConstraintInstance(row, constraint)
1657          If instance = 1 Then
                  ' We are outputting the first row of a new Excel constraint block; put a comment in the .lp file
1658              Print #1, commentStart & " " & s.ConstraintSummary(constraint)
1659          End If
         
1660          With s.SparseA(row)
1661              For i = 1 To .Count
1662                  var = .Index(i)
1663                  coeff = .Coefficient(i)
                      VarName = ValidLPFileVarName(s.VarNames(var))
                      If s.AssumeNonNegativeVars Then
                          If Not TestKeyExists(UsedVariables, VarName) Then UsedVariables.Add VarName, VarName
                      End If
1664                  Print #1, " " & StrEx(coeff) & " " & VarName;
1665              Next i
1666          End With
              
1667          If s.SparseA(row).Count = 0 Then
                  ' This constraint must be trivially satisfied!
                  ' We output the row as a comment
1669              Print #1, commentStart & " (A row with all zero coeffs)";
1673          End If
1674
1681          Print #1, RelationToLPString(s.Relation(constraint)) & StrEx(s.RHS(row))
1682      Next row

1683      Print #1,   ' New line
1692      Print #1, "BOUNDS"
          Print #1,   ' New line
          
          Dim c As Range
1684      If s.SolveRelaxation And Not (s.BinaryCellsRange Is Nothing) Then
1685          Print #1, commentStart & " (Upper bounds of 1 on the relaxed binary variables)"
1686          For Each c In s.BinaryCellsRange
                  VarName = ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False))
                  If s.AssumeNonNegativeVars Then
                      If Not TestKeyExists(UsedVariables, VarName) Then UsedVariables.Add VarName, VarName
                  End If
1687              Print #1, VarName; " <= 1"
1688          Next c
1689          Print #1,   ' New line
1690      End If
          
          ' Output the bounds; this should happen before we output the GENERAL or INTEGER sections (at least for CPLEX .lp files)
          ' See http://lpsolve.sourceforge.net/
          ' The LP file assumes lower bounds on all variables unless we tell it otherwise.
1691      If Not s.AssumeNonNegativeVars Then
              ' We need to make all variables FREE variables (i.e. no lower bounds), except for the Binary variables
1693          Print #1, commentStart & "'Assume Non Negative' is FALSE, so default lower bounds of zero are removed from all non-binary variables."
              Dim NonBinaryCellsRange As Range
              Set NonBinaryCellsRange = SetDifference(s.AdjustableCells, s.BinaryCellsRange)
1695          If Not NonBinaryCellsRange Is Nothing Then
                  For Each c In NonBinaryCellsRange
1696                  Print #1, " " & ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False)) & " FREE"
1697              Next c
              End If
1705          Print #1,   ' New line
1706      Else
              ' If AssumeNonNegative, then we need to apply lower bounds to any variables without explicit lower bounds.
              
              ' Get all bounded variables
              Dim BoundedVariables As Range
              Set BoundedVariables = Nothing
1707          For Each c In s.AdjustableCells
                  VarName = ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False))
1708              If TestKeyExists(s.VarLowerBounds, c.Address) Or Not TestKeyExists(UsedVariables, VarName) Then
                      Set BoundedVariables = ProperUnion(BoundedVariables, c)
                  End If
1711          Next c

              ' We need to mark variables with explicit lower bounds as FREE variables (allowing the possibly negative lower bound to be effective).
              ' However, we don't make Binary variables free
1712          If Not BoundedVariables Is Nothing Then
1714              Print #1, commentStart & "'Assume Non Negative' is TRUE, so default lower bounds of zero are removed only from non-binary variables already given explicit lower bounds."
                  Dim NonBinaryBoundedRange As Range
                  Set NonBinaryBoundedRange = SetDifference(BoundedVariables, s.BinaryCellsRange)
                  If Not NonBinaryBoundedRange Is Nothing Then
1716                  For Each c In NonBinaryBoundedRange
1717                      Print #1, " " & ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False)) & " FREE"
1718                  Next c
                  End If
1726              Print #1,   ' New line
1727          End If
1728      End If
          
          ' Output any integer variables
1729      If Not s.SolveRelaxation And Not (s.IntegerCellsRange Is Nothing) Then
1750          Print #1, "GENERAL"
1751          For Each c In s.IntegerCellsRange
1752              Print #1, " " & ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False));
1753          Next c
1754      End If
1755      Print #1, ' New line

          ' Output binary variables
1757      If Not s.SolveRelaxation And Not (s.BinaryCellsRange Is Nothing) Then
1758          Print #1, "BINARY"
1759          For Each c In s.BinaryCellsRange
1760              Print #1, " " & ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False));
1761          Next c
1762          Print #1,   ' New line
1763      End If
1764      Print #1,   ' New line
          
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

