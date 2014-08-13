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

Dim SolveOptions As SolveOptionsType
    
Public Property Get GetShowIterationResults() As Boolean
    GetShowIterationResults = SolveOptions.ShowIterationResults
End Property

Public Sub SolveModel(s As COpenSolver)
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Dim errorString As String
    
    On Error GoTo ErrHandler
    
    ' Parse spreadsheet and build model
    Set m = New CModelParsed
    m.Setup ActiveWorkbook, ActiveSheet
    If Not m.ProcessSolverModel(s) Then
        GoTo ExitSub
    End If
    
    ' Setup output file
    Dim ModelFilePathName As String
    ModelFilePathName = ModelFilePath(m.Solver)
    
    DeleteFileAndVerify ModelFilePathName, "Solving Parsed Solver Model", "Unable to delete " & ModelFilePathName
    
    GetSolveOptions m.SolverModelSheet.Name & "!", SolveOptions, errorString
    If errorString <> "" Then
        Err.Raise Number:=OpenSolver_SolveError, Source:="Solving Parsed Solver Model", Description:=errorString
    End If
    
    Dim ExecutionCompleted As Boolean
    Select Case m.Solver
    Case "PuLP"
        ExecutionCompleted = SolveModelParsed_PuLP(ModelFilePathName, m, Me)
    Case "Couenne", "Bonmin"
        ExecutionCompleted = SolveModelParsed_NL(ModelFilePathName, m, Me)
    Case "NeosBon", "NeosCou"
        ExecutionCompleted = SolveModelParsed_AMPL(ModelFilePathName, m, Me)
    Case Else
        ExecutionCompleted = False
    End Select
    
    ' If we don't have a solve status
    If Not ExecutionCompleted Then
        SolveStatus = ErrorOccurred
    End If
    
    ' Copy results to COpenSolver for reporting
    s.SolveStatus = SolveStatus
    s.SolveStatusComment = SolveStatusComment
    s.SolveStatusString = SolveStatusString
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    SolveStatus = ErrorOccurred
    Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
    Exit Sub

ExitSub:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

Function SolveModelParsed_AMPL(ModelFilePathName As String, m As CModelParsed, s As COpenSolverParsed) As Boolean
    WriteAMPLFile_Parsed ModelFilePathName, m.Solver
    
    If RunsOnNeos(m.Solver) Then
        SolveModelParsed_AMPL = SolveModelParsed_NEOS(ModelFilePathName, m.Solver, s)
    Else
        SolveModelParsed_AMPL = False
    End If
End Function

Sub WriteAMPLFile_Parsed(ModelFilePathName As String, SolverType As String)
        Dim c As Range
        On Error GoTo FileError
        Open ModelFilePathName For Output As #1

        ' Note - We can use the following code on its own to produce a mod file
        WriteToFile 1, "# Define our sets, parameters and variables (with names matching those"
        WriteToFile 1, "# used in defining the data items)"
        
        ' Define useful constants
        WriteToFile 1, "param pi = 4 * atan(1);"
        
        WriteToFile 1, "# 'Sheet=" + m.SolverModelSheet.Name + "'"
        
        Dim Line As String
        
        ' Vars
        ' Initialise each variable independently
        For Each c In m.AdjustableCells
            Line = "var " & ConvertCellToStandardName(c) & ConvertVarTypeAMPL(m.VarTypeMap(ConvertCellToStandardName(c)))
            If m.AssumeNonNegative Then
                ' If no lower bound has been applied then we need to add >= 0
                If Not TestKeyExists(m.VarLowerBounds, c.Address(RowAbsolute:=False, ColumnAbsolute:=False)) Then
                    Line = Line & " >= 0"
                End If
            End If
            If VarType(c) = vbEmpty Then
                Line = Line & " := 0"
            Else
                Line = Line & " := " & c
            End If
            WriteToFile 1, Line & ";"
        Next
        WriteToFile 1, ""
        
        Dim Formula As Variant
        For Each Formula In m.Formulae
            WriteToFile 1, "var " & Formula.strAddress & ";"
        Next Formula
        WriteToFile 1, ""
        
        If Not m.ObjectiveCell Is Nothing Then
            Dim objCellName As String
            objCellName = ConvertCellToStandardName(m.ObjectiveCell)
            
            If m.ObjectiveSense = TargetObjective Then
                ' Replace objective function with constraint
                WriteToFile 1, "# We have no objective function as the objective must achieve a given target value"
                WriteToFile 1, "subject to targetObj:"
                WriteToFile 1, "    " & objCellName & " == " & m.ObjectiveTargetValue & ";"
                WriteToFile 1, vbNewLine
            Else
                ' Determine objective direction
                If m.ObjectiveSense = MaximiseObjective Then
                   WriteToFile 1, "maximize Total_Cost:"
                Else
                   WriteToFile 1, "minimize Total_Cost:"
                End If
                
                WriteToFile 1, "    " & objCellName & ";" & vbNewLine
            End If
        End If
           
        Dim i As Long
        For i = 1 To m.LHSKeys.Count
            Dim strLHS As String, strRel As String, strRHS As String

            strLHS = m.LHSKeys(i)
            strRel = ConvertRelationToAMPL(m.Rels(i))
            strRHS = m.RHSKeys(i)

            ' Try to make a unique constraint name
            WriteToFile 1, "subject to c" & i & "_" & m.LHSKeys(i) & ":"
            WriteToFile 1, "    " & strLHS & strRel & strRHS & ";" & vbNewLine
        Next i
        
        For i = 1 To m.Formulae.Count
            WriteToFile 1, "subject to f" & i & ":"
            WriteToFile 1, "    " & m.Formulae(i).strAddress & " == " & m.Formulae(i).strFormulaParsed & ";" & vbNewLine
        Next i
        
        ' Run Commands
        WriteToFile 1, "# Solve the problem"
        WriteToFile 1, "option solver " & GetAmplSolverValues(SolverType) & ";"
        
        WriteToFile 1, "solve;" & vbNewLine
        
        Dim cellName As String
        ' Display variables
        For Each c In m.AdjustableCells
            cellName = ConvertCellToStandardName(c)
            WriteToFile 1, "_display " & cellName & ";"
        Next
        
        If Not m.ObjectiveCell Is Nothing Then
            ' Display objective
            WriteToFile 1, "_display " & objCellName & ";" & vbNewLine
        Else
            ' We use the keyword "_display" to know where to begin scanning for variable values and also when to stop scanning.
            ' Even if there is not an objective, we still need to display something so we can read in the variables.
            WriteToFile 1, "_display 1;" & vbNewLine
        End If
        
        ' Display solving condition
        WriteToFile 1, "display solve_result_num, solve_result;"
        Close #1
        Exit Sub
        
FileError:
        Close #1
        Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")

End Sub
'==============================================================================
' ConvertRelationToAMPL
' Given the value of an OpenSolver RelationConst, pick the equivalent AMPL comparison
' operator
Function ConvertRelationToAMPL(Relation As RelationConsts) As String
    Select Case Relation
        Case RelationConsts.RelationLE: ConvertRelationToAMPL = " <= "
        Case RelationConsts.RelationEQ: ConvertRelationToAMPL = " == "
        Case RelationConsts.RelationGE: ConvertRelationToAMPL = " >= "
    End Select
End Function


Function ConvertVarTypeAMPL(intVarType As Long) As String
    Select Case intVarType
    Case VarContinuous
        ConvertVarTypeAMPL = ""
    Case VarInteger
        ConvertVarTypeAMPL = ", integer"
    Case VarBinary
        ConvertVarTypeAMPL = ", binary"
    End Select
End Function

' Returns true if a solve status is loaded
Function SolveModelParsed_NEOS(ModelFilePathName As String, Solver As String, s As COpenSolverParsed) As Boolean
    Dim solution As String, errorString As String
    solution = CallNEOS(ModelFilePathName, Solver, errorString)
    If errorString <> "" Then
        MsgBox (errorString)
        SolveModelParsed_NEOS = False
        Exit Function
    End If
    
    Application.Calculation = xlCalculationAutomatic

    Dim openingParen As Long, closingParen As Long, SolveStatusString As String, result As String
    ' Extract the solve status
    openingParen = InStr(solution, "solve_result =")
    SolveStatusString = right(solution, Len(solution) - openingParen - Len("solve_result ="))

    ' Determine Feasibility
    If SolveStatusString Like "unbounded*" Then
        s.SolveStatus = OpenSolverResult.Unbounded
        s.SolveStatusString = "No Solution Found (Unbounded)"
        SolveModelParsed_NEOS = True
        GoTo NeosReturn
    ElseIf SolveStatusString Like "infeasible*" Then ' Stopped on iterations or time
        s.SolveStatus = OpenSolverResult.Infeasible
        s.SolveStatusString = "No Feasible Solution"
        SolveModelParsed_NEOS = True
        GoTo NeosReturn
    ElseIf Not SolveStatusString Like "solved*" Then
        s.SolveStatus = OpenSolverResult.ErrorOccurred
        openingParen = InStr(solution, ">>>")
        If openingParen = 0 Then
            openingParen = InStr(solution, "processing commands.")
            s.SolveStatusString = right(solution, Len(solution) - openingParen - Len("processing commands."))
        Else
            closingParen = InStr(solution, "<<<")
            s.SolveStatusString = "Error: " & Mid(solution, openingParen, closingParen - openingParen)
        End If
        s.SolveStatusString = "Neos Returned:" & vbNewLine & vbNewLine & s.SolveStatusString
        SolveModelParsed_NEOS = True
        GoTo NeosReturn
    End If
    
    s.SolveStatusString = SolveStatusString

    ' Display results to sheet
    Dim c As Range
    For Each c In m.AdjustableCells
        openingParen = InStr(solution, ConvertCellToStandardName(c))
        closingParen = openingParen + InStr(right(solution, Len(solution) - openingParen), "_display")
        result = Mid(solution, openingParen + Len(ConvertCellToStandardName(c)) + 1, Application.Max(closingParen - openingParen - Len(ConvertCellToStandardName(c)) - 1, 0))

        ' Converting result to number
        Range(c.Address) = "=" & result & "*1"

        ' Removing equal sign
        Range(c.Address) = Range(c.Address).Value2
    Next

    SolveModelParsed_NEOS = True
    Application.Calculation = xlCalculationManual
    Exit Function
    
NeosReturn:
    Exit Function
End Function
