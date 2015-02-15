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

Public SolveStatus As OpenSolverResult ' See Enum in OpenSolverModule; Optimal,Infeasible, etc
Public SolveStatusString As String    ' Eg Infeasible, Opimal etc
Public SolveStatusComment As String ' Eg constraint 1 cannot be satisified
Public SolutionWasLoaded As Boolean

Dim SolveOptions As SolveOptionsType
    
Public Property Get GetShowIterationResults() As Boolean
7134      GetShowIterationResults = SolveOptions.ShowIterationResults
End Property

Public Sub SolveModel(s As COpenSolver)
7135      Application.Calculation = xlCalculationManual
7136      Application.ScreenUpdating = False

          SolveStatus = OpenSolverResult.Unsolved
          SolveStatusString = "Unsolved"
          SolveStatusComment = ""
          
          Dim errorString As String
          
7137      Application.EnableCancelKey = xlErrorHandler
7138      On Error GoTo ErrHandler
          
7139      If s.ModelStatus <> ModelStatus_Built Then
7140          Err.Raise Number:=OpenSolver_SolveError, Source:="Solving Parsed Solver Model", Description:="The model cannot be solved as it has not yet been built."
7141      End If
          
          ' Parse spreadsheet and build model
7142      Set m = New CModelParsed
7143      m.Setup ActiveWorkbook, ActiveSheet
7144      If Not m.ProcessSolverModel(s) Then
7145          GoTo ExitSub
7146      End If
          
          ' Setup output file
          Dim ModelFilePathName As String
7147      ModelFilePathName = ModelFilePath(m.Solver)
          
7148      DeleteFileAndVerify ModelFilePathName, "Solving Parsed Solver Model", "Unable to delete " & ModelFilePathName
          
7149      GetSolveOptions m.SolverModelSheet.Name & "!", SolveOptions, errorString
7150      If errorString <> "" Then
7151          Err.Raise Number:=OpenSolver_SolveError, Source:="Solving Parsed Solver Model", Description:=errorString
7152      End If
          
          Dim ExecutionCompleted As Boolean
7153      Select Case m.Solver
          Case "PuLP"
7154          ExecutionCompleted = SolveModelParsed_PuLP(ModelFilePathName, m, Me)
7155      Case "Couenne", "Bonmin"
7156          ExecutionCompleted = SolveModelParsed_NL(ModelFilePathName, m, Me, SolveOptions)
7157      Case "NeosBon", "NeosCou"
7158          ExecutionCompleted = SolveModelParsed_AMPL(ModelFilePathName, m, Me)
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
          s.LinearSolutionWasLoaded = SolutionWasLoaded
          
7168      Application.StatusBar = False
7169      Application.Calculation = xlCalculationAutomatic
7170      Application.ScreenUpdating = True
          
7171      Exit Sub

ErrHandler:
7172      Application.StatusBar = False
7173      Application.Calculation = xlCalculationAutomatic
7174      Application.ScreenUpdating = True
7175      SolveStatus = ErrorOccurred
          UpdateErrorCache Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
7176      Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
7177      Exit Sub

ExitSub:
7178      Application.StatusBar = False
7179      Application.Calculation = xlCalculationAutomatic
7180      Application.ScreenUpdating = True

End Sub

Function SolveModelParsed_AMPL(ModelFilePathName As String, m As CModelParsed, s As COpenSolverParsed) As Boolean
7181      WriteAMPLFile_Parsed ModelFilePathName, m.Solver
          
7182      If RunsOnNeos(m.Solver) Then
7183          SolveModelParsed_AMPL = SolveModelParsed_NEOS(ModelFilePathName, m.Solver, s)
7184      Else
7185          SolveModelParsed_AMPL = False
7186      End If
End Function

Sub WriteAMPLFile_Parsed(ModelFilePathName As String, SolverType As String)
              Dim c As Range
7187          On Error GoTo FileError
7188          Open ModelFilePathName For Output As #1

              ' Note - We can use the following code on its own to produce a mod file
7189          WriteToFile 1, "# Define our sets, parameters and variables (with names matching those"
7190          WriteToFile 1, "# used in defining the data items)"
              
              ' Define useful constants
7191          WriteToFile 1, "param pi = 4 * atan(1);"
              
7192          WriteToFile 1, "# 'Sheet=" + m.SolverModelSheet.Name + "'"
              
              Dim Line As String
              
              ' Vars
              ' Initialise each variable independently
7193          For Each c In m.AdjustableCells
7194              Line = "var " & ConvertCellToStandardName(c) & ConvertVarTypeAMPL(m.VarTypeMap(ConvertCellToStandardName(c)))
7195              If m.AssumeNonNegative Then
                      ' If no lower bound has been applied then we need to add >= 0
7196                  If Not TestKeyExists(m.VarLowerBounds, c.Address(RowAbsolute:=False, ColumnAbsolute:=False)) Then
7197                      Line = Line & " >= 0"
7198                  End If
7199              End If
7200              If VarType(c) = vbEmpty Then
7201                  Line = Line & " := 0"
7202              Else
7203                  Line = Line & " := " & c
7204              End If
7205              WriteToFile 1, Line & ";"
7206          Next
7207          WriteToFile 1, ""
              
              Dim Formula As Variant
7208          For Each Formula In m.Formulae
7209              WriteToFile 1, "var " & Formula.strAddress & " := " & CStr(Formula.initialValue) & ";"
7210          Next Formula
7211          WriteToFile 1, ""
              
7212          If Not m.ObjectiveCell Is Nothing Then
                  Dim objCellName As String
7213              objCellName = ConvertCellToStandardName(m.ObjectiveCell)
                  
7214              If m.ObjectiveSense = TargetObjective Then
                      ' Replace objective function with constraint
7215                  WriteToFile 1, "# We have no objective function as the objective must achieve a given target value"
7216                  WriteToFile 1, "subject to targetObj:"
7217                  WriteToFile 1, "    " & objCellName & " == " & m.ObjectiveTargetValue & ";"
7218                  WriteToFile 1, vbNewLine
7219              Else
                      ' Determine objective direction
7220                  If m.ObjectiveSense = MaximiseObjective Then
7221                     WriteToFile 1, "maximize Total_Cost:"
7222                  Else
7223                     WriteToFile 1, "minimize Total_Cost:"
7224                  End If
                      
7225                  WriteToFile 1, "    " & objCellName & ";" & vbNewLine
7226              End If
7227          End If
                 
              Dim i As Long
7228          For i = 1 To m.LHSKeys.Count
                  Dim strLHS As String, strRel As String, strRHS As String

7229              strLHS = m.LHSKeys(i)
7230              strRel = ConvertRelationToAMPL(m.Rels(i))
7231              strRHS = m.RHSKeys(i)

7232              WriteToFile 1, "# Actual constraint: " & strLHS & strRel & strRHS
7233              WriteToFile 1, "subject to c" & i & ":"
7234              WriteToFile 1, "    " & strLHS & strRel & strRHS & ";" & vbNewLine
7235          Next i
              
7236          For i = 1 To m.Formulae.Count
7237              WriteToFile 1, "# Parsed formula for " & m.Formulae(i).strAddress
7238              WriteToFile 1, "subject to f" & i & ":"
7239              WriteToFile 1, "    " & m.Formulae(i).strAddress & " == " & m.Formulae(i).strFormulaParsed & ";" & vbNewLine
7240          Next i
              
              ' Run Commands
7241          WriteToFile 1, "# Solve the problem"
7242          WriteToFile 1, "option solver " & GetAmplSolverValues(SolverType) & ";"
              
7243          WriteToFile 1, "solve;" & vbNewLine
              
              Dim cellName As String
              ' Display variables
7244          For Each c In m.AdjustableCells
7245              cellName = ConvertCellToStandardName(c)
7246              WriteToFile 1, "_display " & cellName & ";"
7247          Next
              
7248          If Not m.ObjectiveCell Is Nothing Then
                  ' Display objective
7249              WriteToFile 1, "_display " & objCellName & ";" & vbNewLine
7250          Else
                  ' We use the keyword "_display" to know where to begin scanning for variable values and also when to stop scanning.
                  ' Even if there is not an objective, we still need to display something so we can read in the variables.
7251              WriteToFile 1, "_display 1;" & vbNewLine
7252          End If
              
              ' Display solving condition
7253          WriteToFile 1, "display solve_result_num, solve_result;"
7254          Close #1
7255          Exit Sub
              
FileError:
7256          Close #1
7257          Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")

End Sub
'==============================================================================
' ConvertRelationToAMPL
' Given the value of an OpenSolver RelationConst, pick the equivalent AMPL comparison
' operator
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
          Dim solution As String, errorString As String
7269      solution = CallNEOS(ModelFilePathName, Solver, errorString)
7270      If errorString <> "" Then
7271          MsgBox (errorString)
7272          SolveModelParsed_NEOS = False
7273          Exit Function
7274      End If
          
7275      Application.Calculation = xlCalculationAutomatic

          Dim openingParen As Long, closingParen As Long, SolveStatusString As String, result As String
          ' Extract the solve status
7276      openingParen = InStr(solution, "solve_result =")
7277      SolveStatusString = right(solution, Len(solution) - openingParen - Len("solve_result ="))

          ' Determine Feasibility
7278      If SolveStatusString Like "unbounded*" Then
7279          s.SolveStatus = OpenSolverResult.Unbounded
7280          s.SolveStatusString = "No Solution Found (Unbounded)"
7281          SolveModelParsed_NEOS = True
7282          GoTo NeosReturn
7283      ElseIf SolveStatusString Like "infeasible*" Then ' Stopped on iterations or time
7284          s.SolveStatus = OpenSolverResult.Infeasible
7285          s.SolveStatusString = "No Feasible Solution"
7286          SolveModelParsed_NEOS = True
7287          GoTo NeosReturn
7288      ElseIf Not SolveStatusString Like "solved*" Then
7289          s.SolveStatus = OpenSolverResult.ErrorOccurred
7290          openingParen = InStr(solution, ">>>")
7291          If openingParen = 0 Then
7292              openingParen = InStr(solution, "processing commands.")
7293              s.SolveStatusString = right(solution, Len(solution) - openingParen - Len("processing commands."))
7294          Else
7295              closingParen = InStr(solution, "<<<")
7296              s.SolveStatusString = "Error: " & Mid(solution, openingParen, closingParen - openingParen)
7297          End If
7298          s.SolveStatusString = "Neos Returned:" & vbNewLine & vbNewLine & s.SolveStatusString
7299          SolveModelParsed_NEOS = True
7300          GoTo NeosReturn
7301      End If
          
          s.SolveStatus = OpenSolverResult.Optimal
7302      s.SolveStatusString = SolveStatusString

          ' Display results to sheet
          Dim c As Range
7303      For Each c In m.AdjustableCells
7304          openingParen = InStr(solution, ConvertCellToStandardName(c))
7305          closingParen = openingParen + InStr(right(solution, Len(solution) - openingParen), "_display")
7306          result = Mid(solution, openingParen + Len(ConvertCellToStandardName(c)) + 1, Application.Max(closingParen - openingParen - Len(ConvertCellToStandardName(c)) - 1, 0))

              ' Converting result to number
7307          Range(c.Address) = "=" & result & "*1"

              ' Removing equal sign
7308          Range(c.Address) = Range(c.Address).Value2
7309      Next

          SolutionWasLoaded = True
7310      SolveModelParsed_NEOS = True
7311      Application.Calculation = xlCalculationManual
7312      Exit Function
          
NeosReturn:
7313      Exit Function
End Function
