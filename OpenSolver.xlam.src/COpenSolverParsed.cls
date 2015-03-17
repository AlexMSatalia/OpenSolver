VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COpenSolverParsed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m As CModelParsed

Public SolveStatus As OpenSolverResult
Public SolveStatusString As String    ' Eg Infeasible, Opimal etc
Public SolveStatusComment As String ' Eg constraint 1 cannot be satisified
Public SolutionWasLoaded As Boolean
Public MinimiseUserInteraction As Boolean

Dim SolveOptions As SolveOptionsType
    
Public Property Get GetShowIterationResults() As Boolean
7134      GetShowIterationResults = SolveOptions.ShowIterationResults
End Property

Public Sub SolveModel(s As COpenSolver, Optional SolveRelaxation As Boolean = False, Optional ShouldMinimiseUserInteraction As Boolean = False)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

7135      Application.Calculation = xlCalculationManual
7136      Application.ScreenUpdating = False

          SolveStatus = OpenSolverResult.Unsolved
          SolveStatusString = "Unsolved"
          SolveStatusComment = ""
          MinimiseUserInteraction = ShouldMinimiseUserInteraction
          
7137      Application.EnableCancelKey = xlErrorHandler
          
7139      If s.ModelStatus <> ModelStatus_Built Then
7140          Err.Raise Number:=OpenSolver_SolveError, Description:="The model cannot be solved as it has not yet been built."
7141      End If

          Application.Cursor = xlWait
          
          ' Parse spreadsheet and build model
7142      Set m = New CModelParsed
7143      m.Setup ActiveWorkbook, ActiveSheet
7144      If Not m.ProcessSolverModel(s) Then
7145          GoTo ExitSub
7146      End If
          
          ' Setup output file
          Dim ModelFilePathName As String
7147      ModelFilePathName = ModelFilePath(m.Solver)
          
7148      DeleteFileAndVerify ModelFilePathName
          
7149      GetSolveOptions m.SolverModelSheet, SolveOptions
          
          Dim ExecutionCompleted As Boolean
7153      Select Case m.Solver
          Case "PuLP"
7154          ExecutionCompleted = SolveModelParsed_PuLP(ModelFilePathName, m, Me)
7155      Case "Couenne", "Bonmin"
7156          ExecutionCompleted = SolveModelParsed_NL(ModelFilePathName, m, Me, SolveOptions, SolveRelaxation)
7157      Case "NeosBon", "NeosCou"
7158          ExecutionCompleted = SolveModelParsed_AMPL(ModelFilePathName, m, Me, SolveRelaxation)
7159      Case Else
7160          ExecutionCompleted = False
7161      End Select
          
          ' If we don't have a solve status
7162      If Not ExecutionCompleted Then
7163          SolveStatus = ErrorOccurred
7164      End If

          ' Copy results to COpenSolver for reporting
7165      s.SolveStatus = SolveStatus
7166      s.SolveStatusComment = SolveStatusComment
7167      s.SolveStatusString = SolveStatusString
          s.SolutionWasLoaded = SolutionWasLoaded
          
ExitSub:
          Application.StatusBar = False
          Application.Calculation = xlCalculationAutomatic
          Application.ScreenUpdating = True
          Application.Cursor = xlDefault
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolverParsed", "SolveModel") Then Resume
          RaiseError = True
          
7175      s.SolveStatus = ErrorOccurred
          GoTo ExitSub
End Sub

Function SolveModelParsed_AMPL(ModelFilePathName As String, m As CModelParsed, s As COpenSolverParsed, SolveRelaxation As Boolean) As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

7181      WriteAMPLFile_Parsed ModelFilePathName, m.Solver, SolveRelaxation
          
7182      If RunsOnNeos(m.Solver) Then
7183          SolveModelParsed_AMPL = SolveModelParsed_NEOS(ModelFilePathName, m.Solver, s)
7184      Else
7185          SolveModelParsed_AMPL = False
7186      End If

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("COpenSolverParsed", "SolveModelParsed_AMPL") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Sub WriteAMPLFile_Parsed(ModelFilePathName As String, SolverType As String, SolveRelaxation As Boolean)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

7188      Open ModelFilePathName For Output As #1

          ' Note - We can use the following code on its own to produce a mod file
7189      WriteToFile 1, "# Define our sets, parameters and variables (with names matching those"
7190      WriteToFile 1, "# used in defining the data items)"
              
          ' Define useful constants
7191      WriteToFile 1, "param pi = 4 * atan(1);"
          
7192      WriteToFile 1, "# 'Sheet=" + m.SolverModelSheet.Name + "'"
             
          Dim Line As String
              
          ' Vars
          ' Initialise each variable independently
          Dim VarName As String, curVarType As Long, c As Range
7193      For Each c In m.AdjustableCells
              VarName = ConvertCellToStandardName(c)
7194          Line = "var " & VarName

              ' Output variable types if needed
              curVarType = m.VarTypeMap(VarName)
              If SolveRelaxation Then
                  If curVarType = VarBinary Then
                      Line = Line & " >= 0 <= 1"
                  End If
              Else
                  Line = Line & ConvertVarTypeAMPL(curVarType)
              End If
              
7195          If m.AssumeNonNegative Then
                  ' If no lower bound has been applied then we need to add >= 0
7196              If Not TestKeyExists(m.VarLowerBounds, c.Address(RowAbsolute:=False, ColumnAbsolute:=False)) Then
7197                  Line = Line & " >= 0"
7198              End If
7199          End If
7200          If VarType(c) = vbEmpty Then
7201              Line = Line & " := 0"
7202          Else
7203              Line = Line & " := " & StrEx_NL(CDbl(c))
7204          End If
7205          WriteToFile 1, Line & ";"
7206      Next
7207      WriteToFile 1, ""
          
          Dim Formula As Variant
7208      For Each Formula In m.Formulae
7209          WriteToFile 1, "var " & Formula.strAddress & " := " & StrEx_NL(Formula.initialValue) & ";"
7210      Next Formula
7211      WriteToFile 1, ""
          
7212      If Not m.ObjectiveCell Is Nothing Then
              Dim objCellName As String
7213          objCellName = ConvertCellToStandardName(m.ObjectiveCell)
              
7214          If m.ObjectiveSense = TargetObjective Then
                  ' Replace objective function with constraint
7215              WriteToFile 1, "# We have no objective function as the objective must achieve a given target value"
7216              WriteToFile 1, "subject to targetObj:"
7217              WriteToFile 1, "    " & objCellName & " == " & m.ObjectiveTargetValue & ";"
7218              WriteToFile 1, vbNewLine
7219          Else
                  ' Determine objective direction
7220              If m.ObjectiveSense = MaximiseObjective Then
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
7230          strRel = ConvertRelationToAMPL(m.Rels(i))
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
7242      WriteToFile 1, "option solver " & GetAmplSolverValues(SolverType) & ";"
          
7243      WriteToFile 1, "solve;" & vbNewLine
         
          Dim cellName As String
          ' Display variables
7244      For Each c In m.AdjustableCells
7245          cellName = ConvertCellToStandardName(c)
7246          WriteToFile 1, "_display " & cellName & ";"
7247      Next
          
7248      If Not m.ObjectiveCell Is Nothing Then
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
          If Not ReportError("COpenSolverParsed", "WriteAMPLFile_Parsed") Then Resume
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

' Returns true if a solve status is loaded
Function SolveModelParsed_NEOS(ModelFilePathName As String, Solver As String, s As COpenSolverParsed) As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim solution As String
7269      solution = CallNEOS(ModelFilePathName, Solver, s.MinimiseUserInteraction)

          Dim openingParen As Long, closingParen As Long, SolveStatusString As String, result As String
          ' Extract the solve status
7276      openingParen = InStr(solution, "solve_result =")
7277      SolveStatusString = Mid(solution, openingParen + 1 + Len("solve_result ="))

          ' Determine Feasibility
7278      If SolveStatusString Like "unbounded*" Then
7279          s.SolveStatus = OpenSolverResult.Unbounded
7280          s.SolveStatusString = "No Solution Found (Unbounded)"
7281          SolveModelParsed_NEOS = True
7282          GoTo ExitFunction
7283      ElseIf SolveStatusString Like "infeasible*" Then ' Stopped on iterations or time
7284          s.SolveStatus = OpenSolverResult.Infeasible
7285          s.SolveStatusString = "No Feasible Solution"
7286          SolveModelParsed_NEOS = True
7287          GoTo ExitFunction
7288      ElseIf Not SolveStatusString Like "solved*" Then
7289          s.SolveStatus = OpenSolverResult.ErrorOccurred
7290          openingParen = InStr(solution, ">>>")
7291          If openingParen = 0 Then
7292              openingParen = InStr(solution, "processing commands.")
7293              s.SolveStatusString = Mid(solution, openingParen + 1 + Len("processing commands."))
7294          Else
7295              closingParen = InStr(solution, "<<<")
7296              s.SolveStatusString = "Error: " & Mid(solution, openingParen, closingParen - openingParen)
7297          End If
7298          s.SolveStatusString = "Neos Returned:" & vbNewLine & vbNewLine & s.SolveStatusString
7299          SolveModelParsed_NEOS = True
7300          GoTo ExitFunction
7301      End If
          
          s.SolveStatus = OpenSolverResult.Optimal
7302      s.SolveStatusString = SolveStatusString

          ' Display results to sheet
          Dim c As Range
7303      For Each c In m.AdjustableCells
7304          openingParen = InStr(solution, ConvertCellToStandardName(c))
7305          closingParen = openingParen + InStr(Mid(solution, openingParen + 1), "_display")
7306          result = Mid(solution, openingParen + Len(ConvertCellToStandardName(c)) + 1, Application.Max(closingParen - openingParen - Len(ConvertCellToStandardName(c)) - 1, 0))

              ' Converting result to number
7307          Range(c.Address) = "=" & result & "*1"

              ' Removing equal sign
7308          Range(c.Address) = Range(c.Address).Value2
7309      Next

          SolutionWasLoaded = True
7310      SolveModelParsed_NEOS = True

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("COpenSolverParsed", "SolveModelParsed_NEOS") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function
