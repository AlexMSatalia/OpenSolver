Attribute VB_Name = "modPuLP"
Option Explicit

Dim dblRootNodeTime As Double
Dim dblWhileLoopTime As Double
Dim dblLHSProcessTime As Double
Dim dblRHSProcessTime As Double
Dim dblConCleanTime As Double


'==============================================================================
' WriteToFile
' Writes a string to the given file number, adds a newline, and can easily
' uncomment debug line to print to Immediate if needed.
Private Sub WriteToFile(intFileNum As Integer, strData As String)
    Print #intFileNum, strData
    'Debug.Print strData
End Sub
'==============================================================================

'==============================================================================
' ConvertCellToStandardName
' Range's address property always gives a $A$1 style address, but doesn't
' include the sheet. This function removes any nasty characters, and sticks
' the sheet name at the front, thus giving nice unique names for Python and
' VBA collections to use.
Private Function ConvertCellToStandardName(rngCell As Range, Optional strCleanParentName As String = "") As String
    Dim strCleanAddress As String
    strCleanAddress = rngCell.Address
    If strCleanParentName = "" Then strCleanParentName = Replace(rngCell.Parent.Name, " ", "_")
    strCleanParentName = Replace(strCleanParentName, "-", "_")
    strCleanAddress = Replace(strCleanAddress, "$", "")
    strCleanAddress = Replace(strCleanAddress, ":", "_")
    strCleanAddress = Replace(strCleanAddress, "-", "_")
    ConvertCellToStandardName = strCleanParentName + "_" + strCleanAddress
End Function
'==============================================================================

'==============================================================================
' ConvertRelationToPython
' Given the value of a solver_relX Name, pick the equivalent Python comparison
' operator
Private Function ConvertRelationToPython(ByVal strNameContents As String) As String
    Select Case Mid(strNameContents, 2)
        Case "1": ConvertRelationToPython = " <= "
        Case "2": ConvertRelationToPython = " == "
        Case "3": ConvertRelationToPython = " >= "
    End Select
End Function
'==============================================================================

'==============================================================================
' ConvertRelationToAMPL
' Given the value of a solver_relX Name, pick the equivalent AMPL comparison
' operator
Private Function ConvertRelationToAMPL(ByVal strNameContents As String) As String
    Select Case Mid(strNameContents, 2)
        Case "1": ConvertRelationToAMPL = " &lt;= "
        Case "2": ConvertRelationToAMPL = " == "
        Case "3": ConvertRelationToAMPL = " &gt;= "
    End Select
End Function
'==============================================================================

'==============================================================================
Public Sub GenerateFile(m As CModel2, SolverType As String, boolOtherSheetsIndependent As Boolean)

    '==========================================================================
    ' STEP 0. Misc Setup
    ' Bit of speed, even though we don't write to the sheet
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    ' General cell holder variable
    Dim c As Range
    ' Create a collection for the spreadsheet's formulae (DAG)
    Dim Formulae As New Collection
    ' Map back from standard decision variable names to ranges
    Dim AdjCellNameMap As New Collection
    ' Remember the greatest depth of the DAG
    Dim lngMaxDepth As Long
    lngMaxDepth = 0
    ' Build up a string for all the root nodes that can be printed last
    Dim strProbPlus As String
    ' Record our start time
    Dim timerTotal As New CTimer: timerTotal.StartTimer
    ' Reset additive timers
    dblRootNodeTime = 0
    dblWhileLoopTime = 0
    dblLHSProcessTime = 0
    dblRHSProcessTime = 0
    dblConCleanTime = 0

    Dim commentStart As String  'Character for starting comments for chosen solver
    commentStart = "#"

    '==========================================================================
    ' STEP 1. Setup output file
    ' Open a file handle
    Close #1
    Dim ModelFileName As String, ModelFilePathName As String
    ModelFileName = GetModelFileName(True)
    ModelFilePathName = GetModelFullPath(True)
    
    ' Delete model file, just in case anything goes wrong (and we get left with an old one)
    Close 1
    If Dir(ModelFilePathName) <> "" Then Kill ModelFilePathName
    If Dir(ModelFilePathName) <> "" Then RaiseOSErr OSErrDeleteFile, ModelFilePathName
    
    Open ModelFilePathName For Output As #1
    
    If SolverType = "PuLP" Then
        ' Import standard libraries
        WriteToFile 1, "from coinor.pulp import *"
        WriteToFile 1, "import math"
        ' Write the helper functions needed to easily translate formula to Python
        'WriteToFile 1, "def ExSumProduct(R1, R2): return lpSum([R1[i] * R2[i] for i in range(len(R1))])"
        WriteToFile 1, "def ExSumProduct(R1, R2): return LpAffineExpression(dict(zip(R1, R2)))"
        WriteToFile 1, "def ExIf(COND, T, objCurNode):"
        WriteToFile 1, vbTab & "if COND: return T"
        WriteToFile 1, vbTab & "if not COND: return objCurNode"
        WriteToFile 1, "def ExSumIf(RANGE, CRITERIA, SUMRANGE = None):"
        WriteToFile 1, vbTab & "if SUMRANGE != None: return lpSum([SUMRANGE[i] for i in range(len(SUMRANGE)) if RANGE[i] == CRITERIA])"
        WriteToFile 1, "def ExIfError(VAL, VALIFERR):return VAL"
        WriteToFile 1, "def ExRoundDown(VAL, ND): return math.floor(VAL)"
        ' Get the model started
        WriteToFile 1, "# Begin PuLP Model"
        WriteToFile 1, "print 'Sheet=" + m.SolverModelSheet.Name + "'"
        If m.ObjectiveSense = MaximiseObjective Then
            WriteToFile 1, "prob = LpProblem(""opensolver"", LpMaximize)"
        Else
            WriteToFile 1, "prob = LpProblem(""opensolver"", LpMinimize)"
        End If
    ElseIf SolverType Like "NEOS*" Then
        ' XML
        WriteToFile 1, "&lt;document&gt;"
        WriteToFile 1, "&lt;category&gt;minco&lt;/category&gt;"
        ' Selected Solver
        If SolverType Like "*Bon" Then
            WriteToFile 1, "&lt;solver&gt;Bonmin&lt;/solver&gt;"
        ElseIf SolverType Like "*Cou" Then
            WriteToFile 1, "&lt;solver&gt;Couenne&lt;/solver&gt;"
        End If
        WriteToFile 1, "&lt;inputType&gt;AMPL&lt;/inputType&gt;"
        WriteToFile 1, "&lt;client&gt;&lt;/client&gt;"
        WriteToFile 1, "&lt;priority&gt;short&lt;/priority&gt;"
        WriteToFile 1, "&lt;email&gt;&lt;/email&gt;"
         
        ' Model File
        ' Note - We can use the following code on its own to produce a mod file
        WriteToFile 1, "&lt;model&gt;&lt;![CDATA[# Define our sets, parameters and variables (with names matching those"
        WriteToFile 1, "# used in defining the data items)"
        
        ' Define useful constants
        WriteToFile 1, "param pi = 4 * atan(1);"
        
        WriteToFile 1, "# 'Sheet=" + m.SolverModelSheet.Name + "'"
        
        ' Determine which adjustable cells are integer and binary
        Dim IntegerType As Collection
        Set IntegerType = New Collection
        ' Intitialise all as continuous
        For Each c In m.AdjustableCells
            IntegerType.Add "", ConvertCellToStandardName(c)
        Next
        
        ' Convert integer variables
        If Not (m.IntegerCells Is Nothing) Then
            For Each c In m.IntegerCells
                IntegerType.Remove ConvertCellToStandardName(c)
                IntegerType.Add ", integer", ConvertCellToStandardName(c)
            Next
        End If
        
        ' Convert binary variables
        If Not (m.BinaryCells Is Nothing) Then
            For Each c In m.BinaryCells
                IntegerType.Remove ConvertCellToStandardName(c)
                IntegerType.Add ", binary", ConvertCellToStandardName(c)
            Next
        End If
        
        Dim Line As String
        ' Vars
        ' Initialise each variable independently
        For Each c In m.AdjustableCells
            Line = "var " & ConvertCellToStandardName(c) & IntegerType(ConvertCellToStandardName(c))
            If m.AssumeNonNegative Then
                Line = Line & " &gt;= 0,"
            End If
            Line = Line & " := " & c
            WriteToFile 1, Line & ";"
        Next
        
        WriteToFile 1, "var " & ConvertCellToStandardName(m.ObjectiveCell) & ";" & vbNewLine
         
        ' Objective function replaced with constraint if
        If m.ObjectiveSense = TargetObjective Then
            WriteToFile 1, commentStart & " We have no objective function as the objective must achieve a given target value"
            WriteToFile 1, vbNewLine
        Else
            ' Determine objective direction
            If m.ObjectiveSense = MaximiseObjective Then
               WriteToFile 1, "maximize Total_Cost:"
            Else
               WriteToFile 1, "minimize Total_Cost:"
            End If
            
            WriteToFile 1, "    " & ConvertCellToStandardName(m.ObjectiveCell) & ";" & vbNewLine
        End If
    Else
        GoTo ExitSub
    End If
    
    '==========================================================================
    ' STEP 2. Adjustable cells
    ' Write out all adjustable cells, taking bounds and tyep into consideration
    ' TODO: Is a PuLP dictionary faster?
    Dim timerAdjSetup As New CTimer: timerAdjSetup.StartTimer
    
    Dim cellName As String, strVarType As String, lngArea As Long
    Dim strAdjCellDefines As String ', strThisAreaList As String
    strAdjCellDefines = ""
    
    If SolverType = "PuLP" Then
        For lngArea = 1 To m.AdjustableCells.Areas.Count
            'strThisAreaList = ""
            For Each c In m.AdjustableCells.Areas(lngArea)
                cellName = ConvertCellToStandardName(c)
                If Not TestKeyExists(AdjCellNameMap, cellName) Then
                    AdjCellNameMap.Add c, cellName
                    strVarType = "LpContinuous"
                    If Not (m.BinaryCells Is Nothing) Then
                        If Not (Intersect(c, m.BinaryCells) Is Nothing) Then strVarType = "LpBinary"
                    ElseIf Not (m.IntegerCells Is Nothing) Then
                        If Not (Intersect(c, m.IntegerCells) Is Nothing) Then strVarType = "LpInteger"
                    End If
                    'WriteToFile 1, cellName + " = LpVariable(""" + cellName + """, 0, cat = " + strVarType + ")"
                    strAdjCellDefines = strAdjCellDefines + cellName + " = LpVariable(""" + cellName + """, 0, cat = " + strVarType + ")" + vbNewLine
                End If
            Next
        Next
    End If
    
    ' The cells dependent on the adjustable cells
    Dim rngAdjDepedents As Range
    Set rngAdjDepedents = m.AdjustableCells.Dependents
    
    timerAdjSetup.StopTimer

    '==========================================================================
    ' STEP 3. Objective Cell
    Dim pystrObjective As String, timerObjCell As New CTimer: timerObjCell.StartTimer
    pystrObjective = ConvertFormulaToPython(m.ObjectiveCell.Formula, m.ObjectiveCell, _
                                            m.AdjustableCells, rngAdjDepedents, Formulae, lngMaxDepth, boolOtherSheetsIndependent, SolverType)
    If pystrObjective Like "*Error*" Then
        Exit Sub
    End If
    CleanFormulae Formulae, 1, Formulae.Count, lngMaxDepth
    If SolverType = "PuLP" Then
        strProbPlus = "prob += " + pystrObjective + vbNewLine
    ElseIf SolverType Like "NEOS*" Then
        strProbPlus = ""
    End If
    timerObjCell.StopTimer
    'Debug.Print "Objective Cell Time: "; CStr(Round(timerObjCell.Time, 2)) + " seconds, Formulae.Count = "; Formulae.Count
    
    '==========================================================================
    ' STEP 4. Constraints
    Dim timerConsTime As New CTimer: timerConsTime.StartTimer
    Dim i As Long
    Dim Count As Integer
    Count = 1
    
    For i = 1 To m.ConstraintCount
        Dim nameLHSi As Name, nameRELi As Name, nameRHSi As Name
        Set nameLHSi = Names(m.strNameRoot + "solver_lhs" + CStr(i))
        Set nameRELi = Names(m.strNameRoot + "solver_rel" + CStr(i))
        Set nameRHSi = Names(m.strNameRoot + "solver_rhs" + CStr(i))
        
        Dim pystrLHS As String, pystrREL As String, pystrRHS As String, amplstrREL As String
        Dim lngFormulaeCountBefore As Long
        Dim cRow As Long, cCol As Long, cRowCount As Long, cColCount As Long
        Dim varLHSFormulae As Variant, rngLHS As Range
        
        If nameRELi.value = "=5" Or nameRELi.value = "=4" Then
            GoTo NextCons
        End If
        
        lngFormulaeCountBefore = Formulae.Count
        
        ' Possibilities
        ' LHS is one cell, RHS is one cell/formula/value
        ' LHS is a range (any shape), RHS is one cell/formula/value
        ' LHS is a range (any shape), RHS is a range of same size, not necessarily same shape
        ' Because mismatched shapes are cruel, we will just throw an error if we find them
        ' OpenSolver1 just threw a general error.
        
        cRowCount = nameLHSi.RefersToRange.Rows.Count
        cColCount = nameLHSi.RefersToRange.Columns.Count
        Set rngLHS = nameLHSi.RefersToRange
        varLHSFormulae = rngLHS.Formula
                
        For cRow = 1 To cRowCount
            For cCol = 1 To cColCount
                Count = Count + 1
                'Application.StatusBar = "Solver Constraint #" + CStr(i) + " - R" + CStr(cRow) + " - C" + CStr(cCol)
                'Debug.Print "Solver Constraint #" + CStr(i) + " - R" + CStr(cRow) + " - C" + CStr(cCol)
                'DoEvents
                
                ' Parse LHS
                Dim timerLHS As New CTimer: timerLHS.StartTimer
                If cRowCount = 1 And cColCount = 1 Then
                    pystrLHS = ConvertFormulaToPython(varLHSFormulae, rngLHS(cRow, cCol), m.AdjustableCells, rngAdjDepedents, Formulae, lngMaxDepth, boolOtherSheetsIndependent, SolverType)
                Else
                    pystrLHS = ConvertFormulaToPython(varLHSFormulae(cRow, cCol), rngLHS(cRow, cCol), m.AdjustableCells, rngAdjDepedents, Formulae, lngMaxDepth, boolOtherSheetsIndependent, SolverType)
                End If
                dblLHSProcessTime = dblLHSProcessTime + timerLHS.StopTimer
                
                ' Determine appropriate relation
                pystrREL = ConvertRelationToPython(nameRELi.value)
                amplstrREL = ConvertRelationToAMPL(nameRELi.value)
                    
                ' Parse RHS
                Dim strRHSFormula As String, rngRHSCell As Range
                ' 1.1 RHS is a single cell
                If m.GetRHSType(i) = ConIsSingleCell Then
                    strRHSFormula = nameRHSi.RefersToRange.Formula
                    Set rngRHSCell = nameRHSi.RefersToRange
                ' 1.2 RHS is multiple cells
                ElseIf m.GetRHSType(i) = ConIsMultipleCell Then
                    strRHSFormula = nameRHSi.RefersToRange.Cells(cRow, cCol).Formula
                    Set rngRHSCell = nameRHSi.RefersToRange.Cells(cRow, cCol)
                ' 1.3 RHS is value or formula
                ElseIf m.GetRHSType(i) = ConIsValueOrFormula Then
                    strRHSFormula = nameRHSi.value
                    Set rngRHSCell = Nothing
                End If
                
                Dim timerRHS As New CTimer: timerRHS.StartTimer
                pystrRHS = ConvertFormulaToPython(strRHSFormula, rngRHSCell, m.AdjustableCells, rngAdjDepedents, Formulae, lngMaxDepth, boolOtherSheetsIndependent, SolverType)
                dblRHSProcessTime = dblRHSProcessTime + timerRHS.StopTimer
                
                ' Output the constraint
                If SolverType = "PuLP" Then
                    strProbPlus = strProbPlus + ("#" + nameLHSi.RefersToRange.Address + vbNewLine)
                    strProbPlus = strProbPlus + ("prob += " + pystrLHS + pystrREL + pystrRHS + vbNewLine)
                ElseIf SolverType Like "NEOS*" Then
                    strProbPlus = strProbPlus + "subject to " & pystrLHS & ":" & vbNewLine
                    strProbPlus = strProbPlus + "    " & Formulae(Count).strFormulaParsed & amplstrREL & pystrRHS & ";" & vbNewLine & vbNewLine
                End If
            Next cCol
        Next cRow
        ' Clean Formulae
        Dim timerCleanCon As New CTimer: timerCleanCon.StartTimer
        CleanFormulae Formulae, lngFormulaeCountBefore + 1, Formulae.Count, lngMaxDepth
        dblConCleanTime = dblConCleanTime + timerCleanCon.StopTimer
        
NextCons:
    Next i
    timerConsTime.StopTimer
    
    '==========================================================================
    ' STEP 5. Write to file
    Dim timerWriteToFile As New CTimer: timerWriteToFile.StartTimer
    ' Write adj cells
    WriteToFile 1, strAdjCellDefines
    ' Write the formulae
    Dim lngCurDepth As Long, lngCurNode As Long
    
    Dim lngIndex As Long, objCurNode As CFormula, strParentAdr As Variant
    If SolverType Like "NEOS*" Then
        For lngCurDepth = lngMaxDepth To 0 Step -1
            For lngIndex = 1 To Formulae.Count
                If Formulae(lngIndex).lngDepth = lngCurDepth Then
                    Set objCurNode = Formulae(lngIndex)
                    
                    For Each strParentAdr In objCurNode.Parents
                        With Formulae(strParentAdr)
                            .strFormulaParsed = Replace(.strFormulaParsed, objCurNode.strAddress, "(" & objCurNode.strFormulaParsed & ")")
                        End With
                        'Formulae(strParentAdr).strFormulaParsed = Replace(Formulae(strParentAdr).strFormulaParsed, objCurNode.strAddress, objCurNode.strFormulaParsed)
                    Next
                End If
            Next lngIndex
        Next lngCurDepth
    End If
    
    
    For lngCurDepth = lngMaxDepth To 0 Step -1
        If SolverType = "PuLP" Then
            For lngCurNode = 1 To Formulae.Count
                If Formulae(lngCurNode).lngDepth = lngCurDepth Then
                    WriteToFile 1, Formulae(lngCurNode).strAddress + _
                                   " = " + _
                                   Formulae(lngCurNode).strFormulaParsed + _
                                   " #" + CStr(Formulae(lngCurNode).lngDepth) + " " + CStr(Formulae(lngCurNode).intRefsAdjCell)
                    MsgBox Formulae(lngCurNode).strFormulaParsed
                    MsgBox Formulae(lngCurNode).strAddress
                End If
            Next lngCurNode
        ElseIf SolverType Like "NEOS*" Then
            For lngCurNode = 1 To 1
                If Formulae(lngCurNode).lngDepth = lngCurDepth Then
                    WriteToFile 1, "subject to objConstraint:"
                    WriteToFile 1, "    " & Formulae(lngCurNode).strFormulaParsed & " == " & Formulae(lngCurNode).strAddress & ";" & vbNewLine
                End If
            Next lngCurNode
        End If
    Next lngCurDepth
    ' Write the objective and constraints themselves
    WriteToFile 1, strProbPlus
    
    If SolverType = "PuLP" Then
        ' Writing solving and printing output
        WriteToFile 1, "prob.solve()"
        WriteToFile 1, "# Output results"
        
        Dim SolutionFilePathName As String, SolutionFileName As String
        SolutionFileName = GetSolutionFileName
        SolutionFilePathName = GetSolutionFullPath
        If Dir(SolutionFilePathName) <> "" Then Kill SolutionFilePathName ' delete solution file
        If Dir(SolutionFilePathName) <> "" Then
            ' TODO MsgBox ErrorPrefix & "Unable to delete the CBC solver solution file: " & SolutionFilePathName & ". The problem cannot be solved.", , "OpenSolver Error"
            'GoTo ExitSub
            Exit Sub
        End If
        WriteToFile 1, "f=open(""" & SolutionFilePathName & """,""w"")"
        For Each c In m.AdjustableCells
            cellName = ConvertCellToStandardName(c)
            WriteToFile 1, "f.write(""" + cellName + " ""+str(value(" + cellName + "))+""\n"")"
        Next
        WriteToFile 1, "f.close()"
    ElseIf SolverType Like "NEOS*" Then
        ' Run Commands
        WriteToFile 1, commentStart & " Solve the problem"
        If SolverType Like "*Bon" Then
            WriteToFile 1, "option solver bonmin;"
        ElseIf SolverType Like "*Cou" Then
            WriteToFile 1, "option solver couenne;"
        End If
        WriteToFile 1, "solve;" & vbNewLine
        
        ' Display variables
        For Each c In m.AdjustableCells
            cellName = ConvertCellToStandardName(c)
            WriteToFile 1, "_display " & cellName & ";"
        Next
        ' Display objective
        WriteToFile 1, "_display " & ConvertCellToStandardName(m.ObjectiveCell) & ";" & vbNewLine
        
        ' Display solving condition
        WriteToFile 1, "display solve_result_num, solve_result;"
        
        WriteToFile 1, "end]]&gt;&lt;/model&gt;"
        
        ' Closing XML
        WriteToFile 1, "&lt;data&gt;&lt;![CDATA[]]&gt;&lt;/data&gt;"
        WriteToFile 1, "&lt;commands&gt;&lt;![CDATA[]]&gt;&lt;/commands&gt;"
        WriteToFile 1, "&lt;comments&gt;&lt;![CDATA[]]&gt;&lt;/comments&gt;"
        WriteToFile 1, "&lt;/document&gt;"
    End If
    
    ' Flush and close file handler
    Close #1
    timerWriteToFile.StopTimer
    
    '==========================================================================
    ' STEP 6. Solve it
    Dim timerSolve As New CTimer: timerSolve.StartTimer
    Dim ExecutionCompleted As Boolean
    
    If SolverType = "PuLP" Then
        ' Need way of finding "python.exe"
        ' Need to work on implementing this
        ' ExecutionCompleted = OSSolveSync("C:\Python27\python.exe " & SolutionFilePathName, SW_HIDE, True)
        Exit Sub
    ElseIf SolverType Like "NEOS*" Then
        On Error GoTo ErrHandler
        Application.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityForceDisable
        ExecutionCompleted = CallNEOS(ModelFilePathName, m)
    End If
    
    If Not ExecutionCompleted Then
        ' User pressed escape. Dialogs have already been shown. Just finish
        Exit Sub 'TODO
    End If
    timerSolve.StopTimer
    
    '==========================================================================
    ' STEP 7. Read back in values
    If SolverType = "PuLP" Then
        Open SolutionFilePathName For Input As #2
        Dim CurLine As String, SplitLine() As String
        While Not EOF(2)
            Line Input #2, CurLine
            SplitLine = split(CurLine, " ")
            AdjCellNameMap(SplitLine(0)).value = Val(SplitLine(1))
        Wend
        Close #2
    End If
    
    '==========================================================================
    ' DONE!
    timerTotal.StopTimer
    Debug.Print "Formulae count:", Formulae.Count
    Debug.Print "Total time: ", , timerTotal.Time
    Debug.Print "=   Adj. Cell Time: ", timerAdjSetup.Time
    Debug.Print "  + Obj. Cell Time: ", timerObjCell.Time
    Debug.Print "  + Contraint Time: ", timerConsTime.Time
    Debug.Print "  + Filewrite Time: ", timerWriteToFile.Time
    Debug.Print "  + Solver    Time: ", timerSolve.Time
    Debug.Print "     (total check): ", timerAdjSetup.Time + timerObjCell.Time + timerConsTime.Time + timerWriteToFile.Time + timerSolve.Time
    Debug.Print "Constraint time:"
    Debug.Print "=   LHS Process Time:", dblLHSProcessTime
    Debug.Print "  + RHS Process Time:", dblRHSProcessTime
    Debug.Print "  + Clean Cons  Time:", dblConCleanTime
    Debug.Print "               Total:", dblLHSProcessTime + dblRHSProcessTime + dblConCleanTime, "vs", timerConsTime.Time
    Debug.Print "ConvertFormulaToPython:"
    Debug.Print "=   Root Node Time: ", dblRootNodeTime
    Debug.Print "  + Whileloop Time: ", dblWhileLoopTime
    Debug.Print "             Total: ", dblRootNodeTime + dblWhileLoopTime, "vs", timerObjCell.Time + dblLHSProcessTime + dblRHSProcessTime
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Exit Sub

ExitSub:
    MsgBox "Unknown Solver Type"

ErrHandler:
    MsgBox "Uh Oh"

End Sub


'==============================================================================
' ConvertFormulaToPython
Private Function ConvertFormulaToPython(ByVal strFormula As String, _
                                        ByRef rngSourceCell As Range, _
                                        ByRef rngAdjCells As Range, _
                                        ByRef rngAdjDepedents As Range, _
                                        ByRef Formulae As Collection, _
                                        ByRef lngMaxDepth As Long, _
                                        ByRef boolOtherSheetsIndependent As Boolean, _
                                        ByRef SolverType As String) As String
    '==========================================================================
    ' STEP 0. Misc Setup
    ' The parsed string, this is what we store at the node
    Dim strParsed As String
    ' The current node we are considering
    Dim objCurNode As CFormula
    ' The sheet the next reference token is on
    Dim strSheetPrefix As String
    ' The sheet the adjustable cells are on
    Dim strAdjCellSheet As String
    strAdjCellSheet = rngAdjCells.Parent.Name
    ' The return from an Evaluate call
    Dim varReturn As Variant
    ' Avoid unnecessary evaluations of AddNodeIfNew by checking if whole ranges
    ' are decision variables
    Dim rngBuildingRef As Range
    Dim boolRangeIsAllDecision As Boolean
    Dim lngCurAdjCellArea As Long
    Dim rngIntersect As Range
    ' Avoid re-cleaning a sheet name in every ConvertCellToStandard
    Dim strCleanParentName As String
    
    ' The following are used for converting SumProduct to AMPL
    Dim FunctionName As String
    Dim FunctionCount As Integer
    Dim Count As Integer
    FunctionName = ""
    FunctionCount = 0
    
    '==========================================================================
    ' STEP 1. Setup root node for search
    Dim timerRootNode As New CTimer: timerRootNode.StartTimer
    Dim strRootAddress As String
    If Not (rngSourceCell Is Nothing) Then
        strRootAddress = ConvertCellToStandardName(rngSourceCell)
        ' If root node exists already, don't need to do this again
        If TestKeyExists(Formulae, strRootAddress) Then
            ConvertFormulaToPython = strRootAddress
            Exit Function
        End If
        ' Check that the source cell (e.g. objective, LHS, RHS) is not an
        ' adjustable cell itself
        If rngSourceCell.Parent.Name = strAdjCellSheet Then
            If Not (Intersect(rngSourceCell, rngAdjCells) Is Nothing) Then
                ' Its an adjustable cell, just return standard name and stop
                ConvertFormulaToPython = strRootAddress
                Exit Function
            End If
        End If
    End If
    ' It doesn't exist, it isn't an adjustable cell
    ' Is it even a formula?
    If Len(strFormula) = 0 Then
        ConvertFormulaToPython = "0": Exit Function
    End If
    If left(strFormula, 1) <> "=" Then
        If IsAmericanNumber(strFormula) Then
            ConvertFormulaToPython = strFormula
        Else
            ConvertFormulaToPython = "'" + strFormula + "'"
        End If
        Exit Function
    End If
    ' Store the last index of the list-representation of the DAG
    Dim lngBaseIndex As Long
    lngBaseIndex = Formulae.Count
    ' Create the root node and add it
    Set objCurNode = New CFormula
    If Not (rngSourceCell Is Nothing) Then
        objCurNode.strAddress = strRootAddress
        Set objCurNode.rngAddress = rngSourceCell
        ConvertFormulaToPython = strRootAddress
    Else
        ' Its a constant or formula - return the parsed formula directly
        objCurNode.strAddress = "!" + CStr(Formulae.Count)
        Set objCurNode.rngAddress = Nothing
    End If
    objCurNode.strFormula = strFormula  ' Could also get from rngSourceCell
    objCurNode.lngDepth = 0             ' Root node depth is 0
    objCurNode.boolIsRoot = True
    Formulae.Add objCurNode, objCurNode.strAddress ' Add it to the DAG
    ' Start at the root node
    Dim lngIndex As Long
    lngIndex = lngBaseIndex + 1
    dblRootNodeTime = dblRootNodeTime + timerRootNode.StopTimer
    
    '==========================================================================
    ' STEP 2. Keep processing nodes until nothing interesting left
    Dim timerWhile As New CTimer: timerWhile.StartTimer
    Do While lngIndex <= Formulae.Count
        'DoEvents
        'Debug.Print lngIndex, Formulae.Count
        'If lngIndex = 96 Then
        '    'Debug.Print 96
        'End If
        
        Set objCurNode = Formulae(lngIndex)
        strParsed = ""
        strSheetPrefix = ""
    
        '======================================================================
        ' STEP 2.C. Tokenise if we can't evaluate
        If objCurNode.boolCanEval Then GoTo skipTokenising
        
        ' Tokenize the formula
        Dim tksFormula As Tokens
        Set tksFormula = modTokeniser.ParseFormula(objCurNode.strFormula)
    
        ' If we find a reference to a cell name, it could be the beginning
        ' of a multi-cell range. So always assume thats going to happen, and
        ' look for another reference (LookingForEndOfRef) after a colon :
        ' (CheckForColon). If a colon isn't there, stop looking.
        Dim LookingForEndOfRef As Boolean, CheckForColon As Boolean
        LookingForEndOfRef = False
        CheckForColon = False
        
        ' The candidate multicell range we are building
        Dim BuildingRef As String
        BuildingRef = ""
                
        ' Take a walk through the tokens
        Dim i As Integer, c As Range, tkn As Token
        For i = 1 To tksFormula.Count
            Set tkn = tksFormula.Item(i)
            
            ' CheckForColon means we hit a cell reference, and want to see if
            ' its actually referring to multiple cells
            If CheckForColon = True And tkn.Text <> ":" Then
                ' We don't hit colon right after the reference, we abandon
                ' search for another reference
                LookingForEndOfRef = False
                CheckForColon = False
                If strSheetPrefix <> "" Then BuildingRef = strSheetPrefix + "!" + BuildingRef
                Set rngBuildingRef = Range(BuildingRef)
                If rngBuildingRef.Count > 1 Then
                    strParsed = strParsed + "["
                    For Each c In rngBuildingRef
                        strParsed = strParsed + AddNodeIfNew(objCurNode, c, rngAdjCells, Formulae, lngMaxDepth, rngAdjDepedents) + ","
                    Next c
                    strParsed = left(strParsed, Len(strParsed) - 1) + "]"
                Else
                    strParsed = strParsed + AddNodeIfNew(objCurNode, rngBuildingRef, rngAdjCells, Formulae, lngMaxDepth, rngAdjDepedents)
                End If
                ' Reset the sheet we are tracking
                strSheetPrefix = ""
            End If
    
            ' Decide what to insert based on token type
            Select Case tkn.TokenType
                Case TokenType.Text
                    ' Output with quotes, tokenizer turns "=""test""" -> test
                    strParsed = strParsed + "'" + tkn.Text + "'"
                        
                Case TokenType.Number
                    ' TODO: Scientific notation for small/large numbers?
                    strParsed = strParsed + tkn.Text
                    
                Case TokenType.Bool
                    strParsed = strParsed + IIf(tkn.Text = "TRUE", "True", "False")
                    
                Case TokenType.ErrorText
                    ' TODO: Can't handle that, throw error
                    RaiseOSErr OSErrPulpTokenErrText
                    
                Case TokenType.Reference
                    ' Are we trying to complete a range?
                    If LookingForEndOfRef Then
                        ' We were, so finish building it...
                        BuildingRef = BuildingRef + tkn.Text
                        ' ... and stop trying to build it more
                        LookingForEndOfRef = False
                        CheckForColon = False
                        
                        ' AMPL does not require square bracketing
                        If SolverType Like "PuLP" Then
                            ' Multicell range = Python list
                            strParsed = strParsed + "["
                        End If
                        ' Make sure this range is on the right sheet
                        If strSheetPrefix <> "" Then BuildingRef = strSheetPrefix + "!" + BuildingRef
                        ' Save some time if the cells are all adjustable cells
                        Set rngBuildingRef = Range(BuildingRef)
                        boolRangeIsAllDecision = False
                        ' If on same sheet...
                        If rngBuildingRef.Parent.Name = strAdjCellSheet Then
                            ' For each chuck of adjustable cells
                            For lngCurAdjCellArea = 1 To rngAdjCells.Areas.Count
                                ' Find the crossover
                                Set rngIntersect = Intersect(rngBuildingRef, rngAdjCells.Areas(lngCurAdjCellArea))
                                ' If the crossover = the range we built, then the range we built is all adjustable
                                If Not (rngIntersect Is Nothing) Then
                                    If rngIntersect.Address = rngBuildingRef.Address Then
                                        boolRangeIsAllDecision = True
                                        Exit For
                                    End If
                                End If
                            Next lngCurAdjCellArea
                        End If
                        
                        'Replace preceding comma with + if summing a second argument
                        If FunctionName = "sum" And FunctionCount > 0 Then
                            strParsed = left(strParsed, Len(strParsed) - 1) + "+"
                        End If
                        
                        ' If it wasn't all adjustable, have to do it manually
                        If Not boolRangeIsAllDecision Then
                            Count = 0
                            For Each c In rngBuildingRef
                                ' AMPL does not have a sumproduct function so needs to be set up
                                If FunctionName = "sumproduct" And Not FunctionCount = 0 Then
                                    Count = Count + 1
                                    If FunctionCount > 1 And Count < rngBuildingRef.Count Then
                                        strParsed = Replace(strParsed, "+", "*" & AddNodeIfNew(objCurNode, c, rngAdjCells, Formulae, lngMaxDepth, rngAdjDepedents) & "+", FindPosition(strParsed, "+", Count), 1)
                                    Else
                                        strParsed = Replace(strParsed, ",", "*" & AddNodeIfNew(objCurNode, c, rngAdjCells, Formulae, lngMaxDepth, rngAdjDepedents) & "+", 1, 1)
                                    End If
                                ElseIf FunctionName = "sum" Then
                                    strParsed = strParsed + AddNodeIfNew(objCurNode, c, rngAdjCells, Formulae, lngMaxDepth, rngAdjDepedents) + "+"
                                Else
                                    ' Check if we need a new node, add token to parsed string
                                    strParsed = strParsed + AddNodeIfNew(objCurNode, c, rngAdjCells, Formulae, lngMaxDepth, rngAdjDepedents) + ","
                                End If
                            Next c
                            FunctionCount = FunctionCount + 1
                        Else
                            ' Just put all the adj cell names in
                            strCleanParentName = Replace(rngBuildingRef.Parent.Name, " ", "_")
                            Count = 0
                            For Each c In rngBuildingRef
                                If FunctionName = "sumproduct" And Not FunctionCount = 0 Then
                                    Count = Count + 1
                                    If FunctionCount > 1 And Count < rngBuildingRef.Count Then
                                        strParsed = Replace(strParsed, ",", "*" & ConvertCellToStandardName(c, strCleanParentName) & "+", FindPosition(strParsed, "+", Count), 1)
                                    Else
                                        strParsed = Replace(strParsed, ",", "*" & ConvertCellToStandardName(c, strCleanParentName) & "+", 1, 1)
                                    End If
                                ElseIf FunctionName = "sum" Then
                                    strParsed = strParsed + AddNodeIfNew(objCurNode, c, rngAdjCells, Formulae, lngMaxDepth, rngAdjDepedents) + "+"
                                Else
                                    strParsed = strParsed + ConvertCellToStandardName(c, strCleanParentName) + ","
                                End If
                            Next c
                            FunctionCount = FunctionCount + 1
                        End If
                        ' Reset the sheet we are tracking
                        strSheetPrefix = ""
                        ' Kill the extra comma added by last cell in range
                        strParsed = left(strParsed, Len(strParsed) - 1)
                        
                        ' AMPL does not require square bracketing
                        If SolverType Like "PuLP" Then
                            ' Close the Python list
                            strParsed = strParsed + "]"
                        End If
                    Else
                        ' This is a new range, look for : and another cell ref
                        BuildingRef = tkn.Text
                        CheckForColon = True
                        LookingForEndOfRef = True
                    End If
                    
                Case TokenType.whitespace
                    ' Do nothing
                    
                Case TokenType.UnaryOperator
                    If tkn.Text = "+" Then
                        strParsed = strParsed + "+"
                    ElseIf tkn.Text = "-" Then
                        strParsed = strParsed + "-"
                    Else
                        MsgBox "Unary Operator that isn't + or -! Develop code to handle this: " + tkn.Text
                    End If
                    
                Case TokenType.ArithmeticOperator
                    strParsed = strParsed + tkn.Text
                    
                Case TokenType.ComparisonOperator
                    If tkn.Text = "=" Then
                        strParsed = strParsed + "=="
                    Else 'TODO (maybe): <>?
                        strParsed = strParsed + tkn.Text
                    End If
                    
                Case TokenType.TextOperator
                    ' Text concatenation &
                    strParsed = strParsed + " + "
                    
                Case TokenType.RangeOperator
                    ' Colon. The range we were building is about to be completed
                    ' and should come up after this
                    BuildingRef = BuildingRef + ":"
                    CheckForColon = False
                    
                Case TokenType.ReferenceQualifier
                    strSheetPrefix = "'" + tkn.Text + "'"
                    
                Case TokenType.ExternalReferenceOperator
                    ' We'll insert the ! ourselves manually later
                    
                Case TokenType.PostfixOperator
                    ' Only percentage sign it seems
                    strParsed = strParsed + "/100.0"
                    
                Case TokenType.FunctionOpen
                    If SolverType Like "*NEOS*" Then
                        FunctionName = ""
                        FunctionCount = 0
                        
                        ' AMPL works with lower case expressions
                        tkn.Text = LCase(tkn.Text)
                    
                        ' SUMPRODUCT
                        If tkn.Text = "sumproduct" Then
                            strParsed = strParsed + "("
                            FunctionName = "sumproduct"
                        ' SUMPRODUCT
                        ElseIf tkn.Text = "sum" Then
                            strParsed = strParsed + "("
                            FunctionName = "sum"
                        ' RADIANS
                        ElseIf tkn.Text = "radians" Then
                            strParsed = strParsed + "pi/180*("
                        ' TODO: Unhandled yet
                        ElseIf Not tkn.Text = "min" And Not tkn.Text = "max" Then
                            strParsed = strParsed + tkn.Text + "("
                        Else
                            GoTo ErrHandlerNeos
                        End If
                    ElseIf SolverType Like "PuLP" Then
                        ' SUMPRODUCT - Map to custom function called ExSumProduct
                        If tkn.Text = "SUMPRODUCT" Then
                            strParsed = strParsed + "ExSumProduct("
                        ' SUM - Map to Pulp's lpSum
                        ElseIf tkn.Text = "SUM" Then
                            strParsed = strParsed + "lpSum("
                        ' SUMIF - Map to custom function called ExSumIf
                        ElseIf tkn.Text = "SUMIF" Then
                            strParsed = strParsed + "ExSumIf("
                        ' IF - Map to custom function called ExIf
                        ElseIf tkn.Text = "IF" Then
                            strParsed = strParsed + "ExIf("
                        ' IFERROR - Map to custom function called ExIfError
                        ElseIf tkn.Text = "IFERROR" Then
                            strParsed = strParsed + "ExIfError("
                        ' ROUNDDOWN - Map to custom function called ExRoundDown
                        ElseIf tkn.Text = "ROUNDDOWN" Then
                            strParsed = strParsed + "ExRoundDown("
                        ' TODO: Unhandled yet
                        Else
                            strParsed = strParsed + tkn.Text + "("
                        End If
                    End If
                    
                Case TokenType.ParameterSeparator
                    strParsed = strParsed + tkn.Text
                    
                Case TokenType.FunctionClose
                    strParsed = strParsed + ")"
                    
                Case TokenType.SubExpressionOpen
                    strParsed = strParsed + tkn.Text
                
                Case TokenType.SubExpressionClose
                    strParsed = strParsed + tkn.Text
                    
                Case Else
                    ' TODO: table things
            End Select
        Next i

        ' Are we still looking for another reference?
        If LookingForEndOfRef = True Then
            ' It never came - probably a solitary cell reference at the end
            ' of a formula. Check if we need a new node, add token to parsed string
            If strSheetPrefix <> "" Then BuildingRef = strSheetPrefix + "!" + BuildingRef
            Set rngBuildingRef = Range(BuildingRef)
            If rngBuildingRef.Count > 1 Then
                strParsed = strParsed + "["
                For Each c In rngBuildingRef
                    strParsed = strParsed + AddNodeIfNew(objCurNode, c, rngAdjCells, Formulae, lngMaxDepth, rngAdjDepedents) + ","
                Next c
                strParsed = left(strParsed, Len(strParsed) - 1) + "]"
            Else
                strParsed = strParsed + AddNodeIfNew(objCurNode, rngBuildingRef, rngAdjCells, Formulae, lngMaxDepth, rngAdjDepedents)
            End If
            strSheetPrefix = ""
        End If
        
        
        '======================================================================
        ' STEP 2.C. Store parsed formula in the DAG
skipTokenising:
        If objCurNode.boolIsConstant = False Then objCurNode.strFormulaParsed = strParsed
        If left(objCurNode.strAddress, 1) = "!" Then
            ' Root node is actually a formula or constant, so return the parsed formula
            ConvertFormulaToPython = strParsed
        End If
        If objCurNode.boolIsRoot And objCurNode.boolIsConstant Then
            ' Was just a simple constant, no need to create anything in Formulae
            ConvertFormulaToPython = objCurNode.strFormulaParsed
            Formulae.Remove lngIndex
            Exit Function
        End If
        
        '======================================================================
        ' On to the next node
        lngIndex = lngIndex + 1
    Loop
    dblWhileLoopTime = dblWhileLoopTime + timerWhile.StopTimer
    
    Exit Function

ErrHandlerNeos:
    MsgBox "Min and Max functions are not compatible with NEOS through OpenSolver."
    ConvertFormulaToPython = "Function Error"
    
End Function

Function FindPosition(Word As String, FindString As String, Count As Integer) As Integer
    Dim It As Integer
    
    ' Define iterator
    It = 1
    
    ' Define initial position
    FindPosition = 1
    
    ' Loop until count reached
    While It < Count
        FindPosition = FindPosition + InStr(1, Word, FindString)
        It = It + 1
    Wend
End Function

Sub CleanFormulae(ByRef Formulae As Collection, ByVal lngFStart As Long, ByVal lngFEnd As Long, ByRef lngMaxDepth As Long)

    Dim objCurNode As CFormula, lngIndex As Long, varReturn As Variant


    ' Any AdjCellUnknown can now be decided one way or the other
    Dim lngCurDepth As Long, strChildAddr As Variant
    For lngCurDepth = lngMaxDepth To 0 Step -1
        For lngIndex = lngFStart To lngFEnd
            Set objCurNode = Formulae(lngIndex)
            If objCurNode.lngDepth = lngCurDepth Then
                If objCurNode.intRefsAdjCell = AdjCellUnknown And Not objCurNode.boolIsRoot Then
                    ' Look through children - if any are adj cell depedendant,
                    ' this is too. Otherwise, its not!
                    
                    objCurNode.intRefsAdjCell = AdjCellIndependent
                    For Each strChildAddr In objCurNode.Children
                        If Formulae(strChildAddr).intRefsAdjCell = AdjCellDependent Then
                            objCurNode.intRefsAdjCell = AdjCellDependent
                        End If
                    Next
                    ' If its AdjCellIndependent, eval it!
                    If objCurNode.intRefsAdjCell = AdjCellIndependent Then
                        objCurNode.boolCanEval = True
                        ' NOTE: Application.Evaluate has a 255 character limit
                        ' See also: http://dutchgemini.wordpress.com/2009/08/07/error-2015-using-application-evaluate-in-excel-vba/
                        varReturn = Application.Evaluate(objCurNode.strFormula)
                        If (VBA.VarType(varReturn) = vbError) Then
                            ' Fall back to calculating the cell
                            objCurNode.rngAddress.Calculate
                            objCurNode.strFormulaParsed = CStr(objCurNode.rngAddress.value)
                            objCurNode.boolEvaledWithCalculate = True
                        Else
                            objCurNode.strFormulaParsed = varReturn
                        End If
                        objCurNode.boolIsConstant = True
                    End If
                End If
            End If
        Next lngIndex
    Next lngCurDepth
    
    ' Pull up constants where possible
    Dim lngParent As Long, strParentAdr As Variant
    For lngCurDepth = lngMaxDepth To 0 Step -1
        For lngIndex = lngFStart To lngFEnd
            If Formulae(lngIndex).lngDepth = lngCurDepth Then
                If Formulae(lngIndex).boolIsConstant Then
                    ' Constant - so can pull up into its parents
                    Set objCurNode = Formulae(lngIndex)
                    
                    For Each strParentAdr In objCurNode.Parents
                        With Formulae(strParentAdr)
                            .strFormulaParsed = Replace(.strFormulaParsed, objCurNode.strAddress, objCurNode.strFormulaParsed)
                        End With
                        'Formulae(strParentAdr).strFormulaParsed = Replace(Formulae(strParentAdr).strFormulaParsed, objCurNode.strAddress, objCurNode.strFormulaParsed)
                    Next
                    
                    objCurNode.boolCanBeRemoved = Not objCurNode.boolIsRoot
                End If
            End If
        Next lngIndex
    Next lngCurDepth
    
    ' Remove any unneeded nodes
    For lngIndex = lngFEnd To lngFStart Step -1
        If Formulae(lngIndex).boolCanBeRemoved Then Formulae.Remove lngIndex
    Next lngIndex
End Sub

Sub PUSHDOWN(ByRef objNode As CFormula, ByVal lngNewDepth As Long, ByRef Formulae As Collection, ByRef lngMaxDepth As Long)
    If lngNewDepth > lngMaxDepth Then lngMaxDepth = lngNewDepth
    If lngNewDepth > objNode.lngDepth Then objNode.lngDepth = lngNewDepth
    'Dim child As Long
    'For child = 1 To objNode.Children.Count
    '    PUSHDOWN Formulae(objNode.Children(child)), objNode.lngDepth + 1, Formulae, lngMaxDepth
    'Next
    Dim child As Variant
    For Each child In objNode.Children
        If TestKeyExists(Formulae, CStr(child)) Then PUSHDOWN Formulae(child), objNode.lngDepth + 1, Formulae, lngMaxDepth
    Next
End Sub

Function AddNodeIfNewOLD(ByRef objCurNode As CFormula, ByRef c As Range, ByRef rngAdjCells As Range, ByRef Formulae As Collection, ByRef lngMaxDepth As Long, ByRef rngAdjDepedents As Range) As String
    
    '==========================================================================
    ' 0. Is Simple Cell?
    ' 0.1 Get the standard name of this cell
    Dim strCFormula As String
    strCFormula = c.Formula
    ' 0.2 Is blank cell (TODO: breaks some IF statements)
    If Len(strCFormula) = 0 Then
        AddNodeIfNew = "0"
        GoTo finishedNode
    End If
    ' 0.3 Is constant?
    If left(strCFormula, 1) <> "=" Then
        ' Number?
        If modUtilities.IsAmericanNumber(strCFormula) Then
            AddNodeIfNew = strCFormula
        Else
            ' Not a number, assume text
            AddNodeIfNew = "'" + strCFormula + "'"
        End If
        GoTo finishedNode
    End If
            
    '==========================================================================
    ' 1. Is Existing Node?
    ' 1.1 Get the standard name of this cell
    Dim strStdName As String
    strStdName = ConvertCellToStandardName(c)
    ' 1.2 Does the node exist yet?
    If modUtilities.TestKeyExists(Formulae, strStdName) Then
        ' Node already exists
        Dim objExistNode As CFormula
        Set objExistNode = Formulae(strStdName)
        ' Has it been evaluated?
        If objExistNode.boolIsConstant Then
            ' Its a constant, so we can just fold it straight in
            AddNodeIfNew = objExistNode.strFormulaParsed
            GoTo finishedNode
        End If
        ' Is it dependent on a decision variable?
        If objExistNode.intRefsAdjCell = AdjCellDependent Then
            ' That means this cell is also dependent on adj. cells, which
            ' we may not have known
            objCurNode.intRefsAdjCell = AdjCellDependent
        End If
        ' Add this node to the current nodes Children list
        If Not modUtilities.TestKeyExists(objCurNode.Children, strStdName) Then
            objCurNode.Children.Add strStdName, strStdName
        End If
        ' Tell it about its new parent
        If Not modUtilities.TestKeyExists(objExistNode.Parents, objCurNode.strAddress) Then
            objExistNode.Parents.Add objCurNode.strAddress, objCurNode.strAddress
        End If
        ' Ensure it is at the correct depth
        PUSHDOWN objExistNode, objCurNode.lngDepth + 1, Formulae, lngMaxDepth
        ' Return the standardised name of this existing node
        AddNodeIfNew = strStdName
        GoTo finishedNode
    End If
    
    '==========================================================================
    ' 2. Is Adjustable Cell?
    Dim varReturn As Variant
    Dim objNewNode As New CFormula
    Set objNewNode.rngAddress = c
    objNewNode.strAddress = strStdName
    ' Is on same sheet?
    If objNewNode.GetSheet = rngAdjCells.Parent.Name Then
        ' This formula is on the same sheet as the adjacent cells
        ' Thus we can safely check if this formula's cell is depedent on
        ' a decision variable
        If objNewNode.IsDependentOn(rngAdjDepedents) Then
            ' This formula depends on the value of the adjustable cells
            ' This means we can NOT evaluate it
            objNewNode.boolCanEval = False
            objNewNode.boolIsConstant = False
            objNewNode.intRefsAdjCell = AdjCellDependent
        Else
            ' This formula is NOT dependent on an adjustable cell, so we
            ' can just evaluate its formula (TODO: or take its value?)
            ' We know its not a simple constant, or we would of got
            ' it earlier
            ' NOTE: Application.Evaluate has a 255 character limit
            
            varReturn = Application.Evaluate(strCFormula)
            If (VBA.VarType(varReturn) = vbError) Then
                ' Fall back to calculating the cell
                c.Calculate
                AddNodeIfNew = CStr(c.value)
                GoTo finishedNode
            Else
                AddNodeIfNew = varReturn
                GoTo finishedNode
            End If
        End If
    Else
        ' This formula is on a different sheet to the adjacent cells.
        ' This means we can't check if its depedent on an adjacent cell,
        ' because Range. Dependents only returns the cells for the same
        ' sheet as Range.Parent
        'objNewNode.boolCanEval = False
        'objNewNode.boolIsConstant = False
        'objNewNode.intRefsAdjCell = AdjCellUnknown
        varReturn = Application.Evaluate(strCFormula)
            If (VBA.VarType(varReturn) = vbError) Then
                ' Fall back to calculating the cell
                c.Calculate
                AddNodeIfNew = CStr(c.value)
                GoTo finishedNode
            Else
                AddNodeIfNew = varReturn
                GoTo finishedNode
            End If
    End If
    
    '==========================================================================
    ' 3. Create new node
    With objNewNode
        .strFormula = strCFormula
        .lngDepth = objCurNode.lngDepth + 1
        If .lngDepth > lngMaxDepth Then lngMaxDepth = .lngDepth
        .Parents.Add objCurNode.strAddress, objCurNode.strAddress
    End With
    ' Add this new node to the child list of current node
    objCurNode.Children.Add strStdName, strStdName
    ' Add new node to the DAG
    Formulae.Add objNewNode, strStdName
    ' Return the standardised name of this new node
    AddNodeIfNew = strStdName

finishedNode:
End Function

Function AddNodeIfNew(ByRef objCurNode As CFormula, ByRef c As Range, ByRef rngAdjCells As Range, ByRef Formulae As Collection, ByRef lngMaxDepth As Long, ByRef rngAdjDepedents As Range) As String
    Dim cell As Range
    
    ' 1.1 Get the standard name of this cell
    Dim strStdName As String
    strStdName = ConvertCellToStandardName(c)
    
    ' Check if cell is adjustable cell, otherwise use value
    For Each cell In rngAdjCells
        If ConvertCellToStandardName(cell) = strStdName Then
            AddNodeIfNew = strStdName
            GoTo finishedNode
        End If
    Next
    
    '==========================================================================
    ' 0. Is Simple Cell?
    ' 0.1 Get the standard name of this cell
    Dim strCFormula As String
    strCFormula = c.Formula
    ' 0.2 Is blank cell (TODO: breaks some IF statements)
    If Len(strCFormula) = 0 Then
        AddNodeIfNew = "0"
        GoTo finishedNode
    End If
    ' 0.3 Is constant?
    If left(strCFormula, 1) <> "=" Then
        ' Number?
        If IsAmericanNumber(strCFormula) Then
            AddNodeIfNew = strCFormula
        Else
            ' Not a number, assume text
            AddNodeIfNew = "'" + strCFormula + "'"
        End If
        GoTo finishedNode
    End If
            
    '==========================================================================
    ' 1. Is Existing Node?
    ' 1.2 Does the node exist yet?
    If TestKeyExists(Formulae, strStdName) Then
        ' Node already exists
        Dim objExistNode As CFormula
        Set objExistNode = Formulae(strStdName)
        ' Has it been evaluated?
        If objExistNode.boolIsConstant Then
            ' Its a constant, so we can just fold it straight in
            AddNodeIfNew = objExistNode.strFormulaParsed
            GoTo finishedNode
        End If
        ' Is it dependent on a decision variable?
        If objExistNode.intRefsAdjCell = AdjCellDependent Then
            ' That means this cell is also dependent on adj. cells, which
            ' we may not have known
            objCurNode.intRefsAdjCell = AdjCellDependent
        End If
        ' Add this node to the current nodes Children list
        If Not TestKeyExists(objCurNode.Children, strStdName) Then
            objCurNode.Children.Add strStdName, strStdName
        End If
        ' Tell it about its new parent
        If Not TestKeyExists(objExistNode.Parents, objCurNode.strAddress) Then
            objExistNode.Parents.Add objCurNode.strAddress, objCurNode.strAddress
        End If
        ' Ensure it is at the correct depth
        PUSHDOWN objExistNode, objCurNode.lngDepth + 1, Formulae, lngMaxDepth
        ' Return the standardised name of this existing node
        AddNodeIfNew = strStdName
        GoTo finishedNode
    End If
    
    '==========================================================================
    ' 2. Is Adjustable Cell?
    Dim varReturn As Variant
    Dim objNewNode As New CFormula
    Set objNewNode.rngAddress = c
    objNewNode.strAddress = strStdName
    ' Is on same sheet?
    If objNewNode.GetSheet = rngAdjCells.Parent.Name Then
        ' This formula is on the same sheet as the adjacent cells
        ' Thus we can safely check if this formula's cell is depedent on
        ' a decision variable
        If objNewNode.IsDependentOn(rngAdjDepedents) Then
            ' This formula depends on the value of the adjustable cells
            ' This means we can NOT evaluate it
            objNewNode.boolCanEval = False
            objNewNode.boolIsConstant = False
            objNewNode.intRefsAdjCell = AdjCellDependent
        Else
            ' This formula is NOT dependent on an adjustable cell, so we
            ' can just evaluate its formula (TODO: or take its value?)
            ' We know its not a simple constant, or we would of got
            ' it earlier
            ' NOTE: Application.Evaluate has a 255 character limit
            
            varReturn = Application.Evaluate(strCFormula)
            If (VBA.VarType(varReturn) = vbError) Then
                ' Fall back to calculating the cell
                c.Calculate
                AddNodeIfNew = CStr(c.value)
                GoTo finishedNode
            Else
                AddNodeIfNew = varReturn
                GoTo finishedNode
            End If
        End If
    Else
        ' This formula is on a different sheet to the adjacent cells.
        ' This means we can't check if its depedent on an adjacent cell,
        ' because Range. Dependents only returns the cells for the same
        ' sheet as Range.Parent
        'objNewNode.boolCanEval = False
        'objNewNode.boolIsConstant = False
        'objNewNode.intRefsAdjCell = AdjCellUnknown
        varReturn = Application.Evaluate(strCFormula)
            If (VBA.VarType(varReturn) = vbError) Then
                ' Fall back to calculating the cell
                c.Calculate
                AddNodeIfNew = CStr(c.value)
                GoTo finishedNode
            Else
                AddNodeIfNew = varReturn
                GoTo finishedNode
            End If
    End If
    
    '==========================================================================
    ' 3. Create new node
    With objNewNode
        .strFormula = strCFormula
        .lngDepth = objCurNode.lngDepth + 1
        If .lngDepth > lngMaxDepth Then lngMaxDepth = .lngDepth
        .Parents.Add objCurNode.strAddress, objCurNode.strAddress
    End With
    ' Add this new node to the child list of current node
    objCurNode.Children.Add strStdName, strStdName
    ' Add new node to the DAG
    Formulae.Add objNewNode, strStdName
    ' Return the standardised name of this new node
    AddNodeIfNew = strStdName

finishedNode:
End Function

Function CallNEOS(ModelFilePathName As String, m As CModel2) As Boolean
    On Error GoTo HELPG
    Dim objSvrHTTP As MSXML2.ServerXMLHTTP60, message As String, txtURL As String
    Dim Done As Boolean, result As String
    Dim openingParen As String, closingParen As String, jobNumber As String, Password As String, solutionFile As String, solution As String
    Dim i As Integer, LinearSolveStatusString As String
    
    On Error GoTo errorHandler
    
    ' Server name
    txtURL = "http://www.neos-server.org:3332"
    Set objSvrHTTP = New MSXML2.ServerXMLHTTP60
    
    ' Set up obj for a POST request
    objSvrHTTP.Open "POST", txtURL, False
    
    ' Import file as continuous string
    Open ModelFilePathName For Input As #1
        message = Input$(LOF(1), 1)
    Close #1
    
    ' Clean message up
    message = Replace(message, "<", "&lt;")
    message = Replace(message, ">", "&gt;")
    
    ' Set up message as XML
    message = "<methodCall><methodName>submitJob</methodName><params><param><value><string>" _
       & message & "</string></value></param></params></methodCall>"
    
    ' Send Message to NEOS
    objSvrHTTP.send message
    
    ' Extract Job Number
    openingParen = InStr(objSvrHTTP.responseText, "<int>")
    closingParen = InStr(objSvrHTTP.responseText, "</int>")
    jobNumber = Mid(objSvrHTTP.responseText, openingParen + 5, closingParen - openingParen - 5)
    
    If jobNumber = 0 Then
        MsgBox "An error occured when sending file to NEOS."
        GoTo ExitSub
    End If
    
    ' Extract Password
    openingParen = InStr(objSvrHTTP.responseText, "<string>")
    closingParen = InStr(objSvrHTTP.responseText, "</string>")
    Password = Mid(objSvrHTTP.responseText, openingParen + 8, closingParen - openingParen - 8)
    
    ' Set up Job Status message for XML
    message = "<methodCall><methodName>getJobStatus</methodName><params><param><value><int>" _
       & jobNumber & "</int></value><value><string>" & Password & _
       "</string></value></param></params></methodCall>"
    Done = False
    
    CallingNeos.Show False
    
    ' Loop until job is done
    While Done = False
        DoEvents
        
        ' Reset obj
        Set objSvrHTTP = New MSXML2.ServerXMLHTTP60
        objSvrHTTP.Open "POST", txtURL, False
        
        ' Send message
        objSvrHTTP.send message
        
        ' Extract answer
        openingParen = InStr(objSvrHTTP.responseText, "<string>")
        closingParen = InStr(objSvrHTTP.responseText, "</string>")
        result = Mid(objSvrHTTP.responseText, openingParen + 8, closingParen - openingParen - 8)
        
        ' Evaluate result
        If result = "Done" Then
            Done = True
        ElseIf result <> "Waiting" And result <> "Running" Then
            MsgBox "An error occured when sending file to NEOS. Neos returned: " & result
            GoTo ExitSub
        Else
            Application.Wait (Now + TimeValue("0:00:01"))
        End If
    Wend
    
    CallingNeos.Hide
    
    ' Set up final message for XML
    message = "<methodCall><methodName>getFinalResults</methodName><params><param><value><int>" _
       & jobNumber & "</int></value></param><param><value><string>" & Password & _
       "</string></value></param></params></methodCall>"
    
    ' Reset obj
    Set objSvrHTTP = New MSXML2.ServerXMLHTTP60
    objSvrHTTP.Open "POST", txtURL, False
    
    objSvrHTTP.send message
    
    ' Extract Result
    openingParen = InStr(objSvrHTTP.responseText, "<base64>")
    closingParen = InStr(objSvrHTTP.responseText, "</base64>")
    result = Mid(objSvrHTTP.responseText, openingParen + 8, closingParen - openingParen - 8)
    
    ' The message returned from NEOS is encoded in base 64
    solution = DecodeBase64(result)
    
    Application.Calculation = xlCalculationAutomatic
    
    ' Extract the solve status
    openingParen = InStr(solution, "solve_result =")
    LinearSolveStatusString = right(solution, Len(solution) - openingParen - Len("solve_result ="))
    
    ' Determine Feasibility
    If LinearSolveStatusString Like "unbounded*" Then
        GoTo NeosReturn
        '
    ElseIf LinearSolveStatusString Like "infeasible*" Then ' Stopped on iterations or time
        GoTo NeosReturn
    ElseIf Not LinearSolveStatusString Like "solved*" Then
        openingParen = InStr(solution, ">>>")
        If openingParen = 0 Then
            openingParen = InStr(solution, "processing commands.")
            LinearSolveStatusString = right(solution, Len(solution) - openingParen - Len("processing commands."))
        Else
            closingParen = InStr(solution, "<<<")
            LinearSolveStatusString = "Error: " & Mid(solution, openingParen, closingParen - openingParen)
        End If
        GoTo NeosReturn
    End If
    
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
    
    Application.Calculation = xlCalculationManual
    
    CallNEOS = True
    
    Exit Function
          
ExitSub:
    CallNEOS = False
    Exit Function

NeosReturn:
    CallNEOS = False
    MsgBox "OpenSolver could not solve this problem." & vbNewLine & vbNewLine & "Neos Returned:" & vbNewLine & vbNewLine & LinearSolveStatusString
    Exit Function

errorHandler:
    MsgBox "Sorry we have failed to contact NEOS."
    CallNEOS = False

HELPG:
    MsgBox "Uh Oh"

End Function

' Code by Tim Hastings
Private Function DecodeBase64(ByVal strData As String) As String
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement
  
    ' Help from MSXML
    Set objXML = New MSXML2.DOMDocument60
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.Text = strData
    DecodeBase64 = Stream_BinaryToString(objNode.nodeTypedValue)
  
    ' Clean up
    Set objNode = Nothing
    Set objXML = Nothing
End Function

' Code by Tim Hastings
Function Stream_BinaryToString(Binary)
     Const adTypeText = 2
     Const adTypeBinary = 1
     
     'Create Stream object
     Dim BinaryStream 'As New Stream
     Set BinaryStream = CreateObject("ADODB.Stream")
     
     'Specify stream type - we want To save binary data.
     BinaryStream.Type = adTypeBinary
     
     'Open the stream And write binary data To the object
     BinaryStream.Open
     BinaryStream.Write Binary
     
     'Change stream type To text/string
     BinaryStream.Position = 0
     BinaryStream.Type = adTypeText
     
     'Specify charset For the output text (unicode) data.
     BinaryStream.Charset = "us-ascii"
     
     'Open the stream And get text/string data from the object
     Stream_BinaryToString = BinaryStream.ReadText
     Set BinaryStream = Nothing
End Function

