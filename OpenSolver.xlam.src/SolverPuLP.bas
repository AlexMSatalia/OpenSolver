Attribute VB_Name = "SolverPuLP"
Option Explicit

Public Const SolutionFile_PuLP = "pulp_sol.txt"

Function SolutionFilePath_PuLP() As String
    SolutionFilePath_PuLP = GetTempFilePath(SolutionFile_PuLP)
End Function

Function SolveModelParsed_PuLP(ModelFilePathName As String, m As CModelParsed) As Boolean
    Dim SolutionFilePathName As String
    SolutionFilePathName = SolutionFilePath_PuLP
    DeleteFileAndVerify SolutionFilePathName, "Writing PuLP file", "Unable to delete the solution file: " & SolutionFilePathName
    
    ' Write PuLP file
    WritePuLPFile_Parsed m, ModelFilePathName, SolutionFilePathName
    
    ' Solve model
    Dim ExecutionCompleted As Boolean
    ' Need way of finding "python.exe"
    ' Need to work on implementing this
    ExecutionCompleted = OSSolveSync("C:\Python27\python.exe " & ModelFilePathName, "", "", "", SW_HIDE, True)
    
    If Not ExecutionCompleted Then
        ' User pressed escape. Dialogs have already been shown. Just finish
        Exit Function 'TODO
    End If
    
    Open SolutionFilePathName For Input As #2
    Dim CurLine As String, SplitLine() As String
    
    ' Check for error
    Line Input #2, CurLine
    If left(CurLine, 5) = "Error" Then
        MsgBox (CurLine)
        Close #2
        SolveModelParsed_PuLP = False
        Exit Function
    End If
    
    ' TODO: interpret status code from PuLP
    
    While Not EOF(2)
        Line Input #2, CurLine
        SplitLine = Split(CurLine, " ")
        m.AdjCellNameMap(SplitLine(0)).value = Val(SplitLine(1))
    Wend
    Close #2
    SolveModelParsed_PuLP = True
End Function

Sub WritePuLPFile_Parsed(m As CModelParsed, ModelFilePathName As String, SolutionFilePathName As String)
    On Error GoTo errorHandler
    Open ModelFilePathName For Output As #1
    
    ' Import required libraries
    WriteToFile 1, "from pulp import *"
    WriteToFile 1, "import math"
    WriteToFile 1, ""
    
    ' Write the helper functions needed to easily translate formula to Python
    WriteToFile 1, "# Define helper functions"
    WriteToFile 1, "def ExSumProduct(R1, R2): return LpAffineExpression(dict(zip(R1, R2)))"
    WriteToFile 1, "def ExIf(COND, T, objCurNode):"
    WriteToFile 1, vbTab & "if COND: return T"
    WriteToFile 1, vbTab & "if not COND: return objCurNode"
    WriteToFile 1, "def ExSumIf(RANGE, CRITERIA, SUMRANGE = None):"
    WriteToFile 1, vbTab & "if SUMRANGE != None: return lpSum([SUMRANGE[i] for i in range(len(SUMRANGE)) if RANGE[i] == CRITERIA])"
    WriteToFile 1, "def ExIfError(VAL, VALIFERR):return VAL"
    WriteToFile 1, "def ExRoundDown(VAL, ND): return math.floor(VAL)"
    WriteToFile 1, ""
    
    ' Catch any errors
    WriteToFile 1, "try:"
    
    ' Get the model started
    WriteToFile 1, "# Begin PuLP Model", 4
    WriteToFile 1, "# 'Sheet=" + m.SolverModelSheet.Name + "'", 4
    
    ' Add problem definition
    WriteToFile 1, "# Initialize problem", 4
    Dim pyObjectiveSense As String
    If m.ObjectiveSense = MaximiseObjective Then
        pyObjectiveSense = "LpMaximize"
    Else
        pyObjectiveSense = "LpMinimize"
    End If
    WriteToFile 1, "prob = LpProblem(""opensolver"", " & pyObjectiveSense & ")", 4
    WriteToFile 1, "", 4
    
    ' Add variable definitions
    WriteToFile 1, "# Define variables", 4
    Dim c As Range, strVarType As String, cellName As String
    For Each c In m.AdjustableCells
        cellName = ConvertCellToStandardName(c)
        strVarType = ConvertVarTypePuLP(m.VarTypeMap(cellName))
        WriteToFile 1, cellName + " = LpVariable(""" + cellName + """, 0, cat=" + strVarType + ")", 4
    Next
    WriteToFile 1, "", 4
    
    ' Add objective definition
    WriteToFile 1, "# Define objective cell", 4
    WriteToFile 1, ConvertCellToStandardName(m.ObjectiveCell) & " = " & ConvertFormulaToPuLP(m.Formulae(ConvertCellToStandardName(m.ObjectiveCell)).strFormulaParsed), 4
    WriteToFile 1, "", 4
    
    Dim i As Integer
    Dim strLHS As String, strRel As String, strRHS As String
    
    ' Add constraint cell definitions
    WriteToFile 1, "# Define constraint cells", 4
    For i = 1 To m.LHSKeys.Count
        strLHS = ConvertFormulaToPuLP(GetFormulaWithDefault(m.Formulae, m.LHSKeys(i), m.LHSKeys(i)))
        WriteToFile 1, m.LHSKeys(i) & " = " & strLHS, 4
    Next i
    WriteToFile 1, "", 4
    
    ' Add objective function
    WriteToFile 1, "# Add objective", 4
    WriteToFile 1, "prob += " + ConvertCellToStandardName(m.ObjectiveCell), 4
    
    ' Add constraint inequalities
    WriteToFile 1, "# Add constraints", 4
    For i = 1 To m.LHSKeys.Count
        strRel = ConvertRelationToPuLP(m.Rels(i))
        strRHS = ConvertFormulaToPuLP(GetFormulaWithDefault(m.Formulae, m.RHSKeys(i), m.RHSKeys(i)))
        
        WriteToFile 1, "#" & m.LHSKeys(i), 4
        WriteToFile 1, "prob += " + m.LHSKeys(i) + strRel + strRHS, 4
    Next i
    WriteToFile 1, "", 4
    
    ' Solve
    WriteToFile 1, "# Solve", 4
    WriteToFile 1, "prob.solve()", 4
    WriteToFile 1, "", 4
    
    ' Add printing output
    WriteToFile 1, "# Output results", 4
    WriteToFile 1, "f = open(""" & SolutionFilePathName & """, ""w"")", 4
    
    ' Output solve status
    WriteToFile 1, "f.write(""Solve status: "" + str(prob.status) + ""\n"")", 4
    
    ' Output variable values
    For Each c In m.AdjustableCells
        cellName = ConvertCellToStandardName(c)
        WriteToFile 1, "f.write(""" + cellName + " "" + str(value(" + cellName + ")) + ""\n"")", 4
    Next
    WriteToFile 1, "f.close()", 4
    WriteToFile 1, "", 4
    
    ' Error handling - dump any error to file
    WriteToFile 1, "except Exception as e:"
    WriteToFile 1, "f = open(""" & SolutionFilePathName & """, ""w"")", 4
    WriteToFile 1, "f.write(""Error: %s"" % e.message)", 4
    WriteToFile 1, "f.close()", 4

errorHandler:
    Close #1
End Sub

Function ConvertVarTypePuLP(intVarType As Integer) As String
    Select Case intVarType
    Case VarContinuous
        ConvertVarTypePuLP = "LpContinuous"
    Case VarInteger
        ConvertVarTypePuLP = "LpInteger"
    Case VarBinary
        ConvertVarTypePuLP = "LpBinary"
    End Select
End Function

Function ConvertRelationToPuLP(Relation As RelationConsts) As String
    Select Case Relation
        Case RelationConsts.RelationLE: ConvertRelationToPuLP = " <= "
        Case RelationConsts.RelationEQ: ConvertRelationToPuLP = " == "
        Case RelationConsts.RelationGE: ConvertRelationToPuLP = " >= "
    End Select
End Function

Function ConvertFormulaToPuLP(strFormula As String) As String
    ' Tokenize the formula
    Dim tksFormula As Tokens
    ' Need to add "=" to convert the expression to a formula for the tokeniser
    Set tksFormula = modTokeniser.ParseFormula("=" & strFormula)
            
    Dim strParsed As String
    strParsed = ""
    
    ' Take a walk through the tokens
    Dim i As Integer, tkn As Token
    For i = 1 To tksFormula.Count
        Set tkn = tksFormula.Item(i)

        ' Decide what to insert based on token type
        Select Case tkn.TokenType
        Case TokenType.FunctionOpen
            tkn.Text = LCase(tkn.Text)
            Select Case tkn.Text
            Case "sqrt"
                strParsed = strParsed + "math.sqrt" + "("
            Case Else
                strParsed = strParsed + tkn.Text + "("
            End Select
        Case TokenType.FunctionClose
            strParsed = strParsed + ")"
        Case Else
            strParsed = strParsed + tkn.Text
        End Select
    Next i
    
    ConvertFormulaToPuLP = strParsed
    Exit Function

errorHandler:
    MsgBox "Not implemented for PuLP yet"
End Function
