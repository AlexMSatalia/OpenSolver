Attribute VB_Name = "SolverPuLP"
Option Explicit

Public Const SolutionFile_PuLP = "pulp_sol.txt"

Function SolutionFilePath_PuLP() As String
7314      SolutionFilePath_PuLP = GetTempFilePath(SolutionFile_PuLP)
End Function

Function SolveModelParsed_PuLP(ModelFilePathName As String, m As CModelParsed, s As COpenSolverParsed) As Boolean
          Dim SolutionFilePathName As String
7315      SolutionFilePathName = SolutionFilePath_PuLP
7316      DeleteFileAndVerify SolutionFilePathName

          ' Write PuLP file
7317      WritePuLPFile_Parsed m, ModelFilePathName, SolutionFilePathName
          
          ' Solve model
          Dim ExecutionCompleted As Boolean
          ' Need way of finding "python.exe"
          ' Need to work on implementing this
7318      ExecutionCompleted = RunExternalCommand("C:\Python27\python.exe " & ModelFilePathName, "", WindowStyleType.Hide, True)
          
7319      If Not ExecutionCompleted Then
              ' User pressed escape. Dialogs have already been shown. Just finish
7320          Exit Function 'TODO
7321      End If
          
7322      On Error GoTo ErrHandler
7323      Open SolutionFilePathName For Input As #2
          Dim CurLine As String, SplitLine() As String
          
          ' Check for error
7324      Line Input #2, CurLine
7325      If left(CurLine, 5) = "Error" Then
7326          MsgBox (CurLine)
7327          Close #2
7328          SolveModelParsed_PuLP = False
7329          Exit Function
7330      End If
          
          ' TODO: interpret status code from PuLP
          
7331      While Not EOF(2)
7332          Line Input #2, CurLine
7333          SplitLine = Split(CurLine, " ")
7334          m.AdjCellNameMap(SplitLine(0)).value = Val(SplitLine(1))
7335      Wend
7336      Close #2
7337      SolveModelParsed_PuLP = True
7338      Exit Function
          
ErrHandler:
7339      Close #2
7340      Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Function

Sub WritePuLPFile_Parsed(m As CModelParsed, ModelFilePathName As String, SolutionFilePathName As String)
7341      On Error GoTo ErrorHandler
7342      Open ModelFilePathName For Output As #1
          
          ' Import required libraries
7343      WriteToFile 1, "from pulp import *"
7344      WriteToFile 1, "import math"
7345      WriteToFile 1, ""
          
          ' Write the helper functions needed to easily translate formula to Python
7346      WriteToFile 1, "# Define helper functions"
7347      WriteToFile 1, "def ExSumProduct(R1, R2): return LpAffineExpression(dict(zip(R1, R2)))"
7348      WriteToFile 1, "def ExIf(COND, T, objCurNode):"
7349      WriteToFile 1, vbTab & "if COND: return T"
7350      WriteToFile 1, vbTab & "if not COND: return objCurNode"
7351      WriteToFile 1, "def ExSumIf(RANGE, CRITERIA, SUMRANGE = None):"
7352      WriteToFile 1, vbTab & "if SUMRANGE != None: return lpSum([SUMRANGE[i] for i in range(len(SUMRANGE)) if RANGE[i] == CRITERIA])"
7353      WriteToFile 1, "def ExIfError(VAL, VALIFERR):return VAL"
7354      WriteToFile 1, "def ExRoundDown(VAL, ND): return math.floor(VAL)"
7355      WriteToFile 1, ""
          
          ' Catch any errors
7356      WriteToFile 1, "try:"
          
          ' Get the model started
7357      WriteToFile 1, "# Begin PuLP Model", 4
7358      WriteToFile 1, "# 'Sheet=" + m.SolverModelSheet.Name + "'", 4
          
          ' Add problem definition
7359      WriteToFile 1, "# Initialize problem", 4
          Dim pyObjectiveSense As String
7360      If m.ObjectiveSense = MaximiseObjective Then
7361          pyObjectiveSense = "LpMaximize"
7362      Else
7363          pyObjectiveSense = "LpMinimize"
7364      End If
7365      WriteToFile 1, "prob = LpProblem(""opensolver"", " & pyObjectiveSense & ")", 4
7366      WriteToFile 1, "", 4
          
          ' Add variable definitions
7367      WriteToFile 1, "# Define variables", 4
          Dim c As Range, strVarType As String, cellName As String
7368      For Each c In m.AdjustableCells
7369          cellName = ConvertCellToStandardName(c)
7370          strVarType = ConvertVarTypePuLP(m.VarTypeMap(cellName))
7371          WriteToFile 1, cellName + " = LpVariable(""" + cellName + """, 0, cat=" + strVarType + ")", 4
7372      Next
7373      WriteToFile 1, "", 4
          
          Dim i As Long
          Dim strLHS As String, strRel As String, strRHS As String
          
          ' Add constraint cell definitions
7374      WriteToFile 1, "# Define constraint cells", 4
7375      For i = 1 To m.LHSKeys.Count
7376          strLHS = GetFormulaWithDefault(m.Formulae, m.LHSKeys(i), m.LHSKeys(i))
7377          WriteToFile 1, m.LHSKeys(i) & " = " & strLHS, 4
7378      Next i
7379      WriteToFile 1, "", 4
          
          ' Add formula definitions - needs to be highest depth first so no undefined variables are used
7380      WriteToFile 1, "# Define formulae", 4
          Dim lngIndex As Long, lngCurDepth As Long
7381      For lngCurDepth = m.lngMaxDepth To 0 Step -1
7382              For lngIndex = 1 To m.Formulae.Count
7383                  If m.Formulae(lngIndex).lngDepth = lngCurDepth Then
7384                      WriteToFile 1, m.Formulae(lngIndex).strAddress & " = " & m.Formulae(lngIndex).strFormulaParsed, 4
7385                  End If
7386              Next lngIndex
7387      Next lngCurDepth

          ' Add objective function
7388      WriteToFile 1, "# Add objective", 4
7389      WriteToFile 1, "prob += " + ConvertCellToStandardName(m.ObjectiveCell), 4
          
          ' Add constraint inequalities
7390      WriteToFile 1, "# Add constraints", 4
7391      For i = 1 To m.LHSKeys.Count
7392          strRel = ConvertRelationToPuLP(m.Rels(i))
7393          strRHS = GetFormulaWithDefault(m.Formulae, m.RHSKeys(i), m.RHSKeys(i))
              
7394          WriteToFile 1, "#" & m.LHSKeys(i), 4
7395          WriteToFile 1, "prob += " + m.LHSKeys(i) + strRel + strRHS, 4
7396      Next i
7397      WriteToFile 1, "", 4
          
          ' Solve
7398      WriteToFile 1, "# Solve", 4
7399      WriteToFile 1, "prob.solve()", 4
7400      WriteToFile 1, "", 4
          
          ' Add printing output
7401      WriteToFile 1, "# Output results", 4
7402      WriteToFile 1, "f = open(""" & SolutionFilePathName & """, ""w"")", 4
          
          ' Output solve status
7403      WriteToFile 1, "f.write(""Solve status: "" + str(prob.status) + ""\n"")", 4
          
          ' Output variable values
7404      For Each c In m.AdjustableCells
7405          cellName = ConvertCellToStandardName(c)
7406          WriteToFile 1, "f.write(""" + cellName + " "" + str(value(" + cellName + ")) + ""\n"")", 4
7407      Next
7408      WriteToFile 1, "f.close()", 4
7409      WriteToFile 1, "", 4
          
          ' Error handling - dump any error to file
7410      WriteToFile 1, "except Exception as e:"
7411      WriteToFile 1, "f = open(""" & SolutionFilePathName & """, ""w"")", 4
7412      WriteToFile 1, "f.write(""Error: %s"" % e.message)", 4
7413      WriteToFile 1, "f.close()", 4
7414      Close #1
7415      Exit Sub

ErrorHandler:
7416      Close #1
7417      Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Sub

Function ConvertVarTypePuLP(intVarType As Long) As String
7418      Select Case intVarType
          Case VarContinuous
7419          ConvertVarTypePuLP = "LpContinuous"
7420      Case VarInteger
7421          ConvertVarTypePuLP = "LpInteger"
7422      Case VarBinary
7423          ConvertVarTypePuLP = "LpBinary"
7424      End Select
End Function

Function ConvertRelationToPuLP(Relation As RelationConsts) As String
7425      Select Case Relation
              Case RelationConsts.RelationLE: ConvertRelationToPuLP = " <= "
7426          Case RelationConsts.RelationEQ: ConvertRelationToPuLP = " == "
7427          Case RelationConsts.RelationGE: ConvertRelationToPuLP = " >= "
7428      End Select
End Function

Function ConvertFormula_PuLP(tokenText As String) As String
7429      tokenText = LCase(tokenText)
7430      Select Case tokenText
          Case "sqrt"
7431          ConvertFormula_PuLP = "math." + tokenText + "("
7432      Case Else
7433          ConvertFormula_PuLP = tokenText + "("
7434      End Select
7435      Exit Function

ErrorHandler:
7436      MsgBox tokenText + "not implemented for PuLP yet"
7437      ConvertFormula_PuLP = tokenText + "("
End Function
