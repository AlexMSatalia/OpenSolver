Attribute VB_Name = "SolverGurobi"

Option Explicit
Public OpenSolver_Gurobi As COpenSolver 'Access to model
Public SparseA_Gurobi() As CIndexedCoeffs 'Access to sparse A matrix

Public Const SolverTitle_Gurobi = "Gurobi (Linear Solver)"
Public Const SolverDesc_Gurobi = "Gurobi is a solver for linear programming (LP), quadratic and quadratically constrained programming (QP and QCP), and mixed-integer programming (MILP, MIQP, and MIQCP). It requires the user to download and install a version of the Gurobi and to have GurobiOSRun.py in the OpenSolver directory."
Public Const SolverLink_Gurobi = "http://www.gurobi.com/resources/documentation"
Public Const SolverType_Gurobi = OpenSolver_SolverType.Linear

#If Mac Then
Public Const SolverName_Gurobi = "gurobi_cl"
#Else
Public Const SolverName_Gurobi = "gurobi_cl.exe"
#End If

Public Const SolverScript_Gurobi = "gurobi_tmp" & ScriptExtension
Public Const SolverPythonScript_Gurobi = "gurobiOSRun.py"
Public Const Solver_Gurobi = "gurobi" & ScriptExtension

Public Const SolutionFile_Gurobi = "modelsolution.sol"
Public Const SensitivityFile_Gurobi = "sensitivityData.sol"

Function SolutionFilePath_Gurobi() As String
    SolutionFilePath_Gurobi = GetTempFilePath(SolutionFile_Gurobi)
End Function

Function SolverPythonScriptPath_Gurobi() As String
    GetExistingFilePathName ThisWorkbook.Path, SolverPythonScript_Gurobi, SolverPythonScriptPath_Gurobi
End Function

Function ScriptFilePath_Gurobi() As String
    ScriptFilePath_Gurobi = GetTempFilePath(SolverScript_Gurobi)
End Function

Function SensitivityFilePath_Gurobi() As String
    SensitivityFilePath_Gurobi = GetTempFilePath(SensitivityFile_Gurobi)
End Function

Sub CleanFiles_Gurobi(errorPrefix As String)
    ' Solution file
    DeleteFileAndVerify SolutionFilePath_Gurobi(), errorPrefix, "Unable to delete the Gurobi solver solution file: " & SolutionFilePath_Gurobi()
    ' Cost Range file
    DeleteFileAndVerify SensitivityFilePath_Gurobi(), errorPrefix, "Unable to delete the Gurobi solver sensitivity data file: " & SensitivityFilePath_Gurobi()
End Sub

Function About_Gurobi() As String
' Return string for "About" form
    Dim SolverPath As String, errorString As String
    If Not SolverAvailable_Gurobi(SolverPath, errorString) Then
        About_Gurobi = errorString
        Exit Function
    End If
    
    SolverPath = GetExternalSolver_Gurobi()
    
    ' Assemble version info
    About_Gurobi = "Gurobi " & SolverBitness_Gurobi & "-bit" & _
                     " v" & SolverVersion_Gurobi & _
                     " at " & SolverPath
End Function

Function GetGurobiBinFolder() As String
#If Mac Then
    GetGurobiBinFolder = "Macintosh HD:usr:local:bin:"
#Else
    GetExistingFilePathName Environ("GUROBI_HOME"), "bin", GetGurobiBinFolder
#End If
End Function

Function GetExternalSolver_Gurobi() As String
#If Mac Then
    ' On Mac, using the gurobi interactive shell causes errors when there are spaces in the filepath.
    ' The mac gurobi.sh script, unlike windows, doesn't have a check for a gurobi install, thus it doesn't do anything for us here and is safe to skip.
    ' We can just run python by itself. We need to use the default system python (pre-installed on mac) and not any other version (e.g. a version from homebrew)
    GetExternalSolver_Gurobi = "Macintosh HD:usr:bin:python"
#Else
    GetExistingFilePathName GetGurobiBinFolder(), Solver_Gurobi, GetExternalSolver_Gurobi
#End If
End Function

Function SolverAvailable_Gurobi(Optional SolverPath As String, Optional errorString As String) As Boolean
' Returns true if Gurobi is available and sets SolverPath as path to gurobi_cl
    If GetExistingFilePathName(GetGurobiBinFolder, SolverName_Gurobi, SolverPath) And _
       FileOrDirExists(SolverPythonScriptPath_Gurobi()) Then
        SolverAvailable_Gurobi = True
    Else
        SolverPath = ""
        errorString = "No Gurobi installation was detected."
        SolverAvailable_Gurobi = False
    End If
End Function

Function SolverVersion_Gurobi() As String
' Get Gurobi version by running 'gurobi_cl -v' at command line
    Dim SolverPath As String
    If Not SolverAvailable_Gurobi(SolverPath) Then
        SolverVersion_Gurobi = ""
        Exit Function
    End If
    
    ' Set up Gurobi to write version info to text file
    Dim logFile As String
    logFile = GetTempFolder & "gurobiversion.txt"
    If FileOrDirExists(logFile) Then Kill logFile
    
    Dim RunPath As String, FileContents As String
    RunPath = ScriptFilePath_Gurobi()
    If FileOrDirExists(RunPath) Then Kill RunPath
    FileContents = QuotePath(ConvertHfsPath(SolverPath)) & " -v" & " > " & QuotePath(ConvertHfsPath(logFile))
    CreateScriptFile RunPath, FileContents
    
    ' Run Gurobi
    Dim completed As Boolean
    completed = OSSolveSync(ConvertHfsPath(RunPath), "", "", "", SW_HIDE, True)
    
    ' Read version info back from output file
    ' Output like 'Gurobi Optimizer version 5.6.3 (win64)'
    Dim Line As String
    If FileOrDirExists(logFile) Then
        Open logFile For Input As 1
        Line Input #1, Line
        Close #1
        SolverVersion_Gurobi = right(Line, Len(Line) - 25)
        SolverVersion_Gurobi = left(SolverVersion_Gurobi, 5)
    Else
        SolverVersion_Gurobi = ""
    End If
End Function

Function SolverBitness_Gurobi() As String
' Get Gurobi bitness by running 'gurobi_cl -v' at command line
    Dim SolverPath As String
    If Not SolverAvailable_Gurobi(SolverPath) Then
        SolverBitness_Gurobi = ""
        Exit Function
    End If
    
    ' Set up Gurobi to write version info to text file
    Dim logFile As String
    logFile = GetTempFolder & "gurobiversion.txt"
    If FileOrDirExists(logFile) Then Kill logFile
    
    Dim RunPath As String, FileContents As String
    RunPath = ScriptFilePath_Gurobi()
    If FileOrDirExists(RunPath) Then Kill RunPath
    FileContents = QuotePath(ConvertHfsPath(SolverPath)) & " -v" & " > " & QuotePath(ConvertHfsPath(logFile))
    CreateScriptFile RunPath, FileContents
    
    ' Run Gurobi
    Dim completed As Boolean
    completed = OSSolveSync(ConvertHfsPath(RunPath), "", "", "", SW_HIDE, True)
    
    ' Read bitness info back from output file
    ' Output like 'Gurobi Optimizer version 5.6.3 (win64)'
    Dim Line As String
    If FileOrDirExists(logFile) Then
        Open logFile For Input As 1
        Line Input #1, Line
        Close #1
        If right(Line, 3) = "64)" Then
            SolverBitness_Gurobi = "64"
        Else
            SolverBitness_Gurobi = "32"
        End If
    End If
End Function
Function CreateSolveScript_Gurobi(SolutionFilePathName As String, ExtraParametersString As String, SolveOptions As SolveOptionsType) As String
    Dim SolverString As String, CommandLineRunString As String, PrintingOptionString As String
    SolverString = QuotePath(ConvertHfsPath(GetExternalSolver_Gurobi()))

    CommandLineRunString = QuotePath(ConvertHfsPath(SolverPythonScriptPath_Gurobi()))
    '========================================================================================
    'Gurobi can also be run at the command line using gurobi_cl with the following commands
    '
    'Solver="gurobi_cl.exe"
    'CommandLineRunString = " Threads=1" & " TimeLimit=" & Replace(Str(SolveOptions.maxTime), " ", "") & " ResultFile="
    '             & GetTempFolder & Replace(SolutionFileName, ".txt", ".sol") _
    '             & " " & ModelFilePathName
    '
    '========================================================================================
    PrintingOptionString = ""
    
    Dim scriptFile As String, scriptFileContents As String
    scriptFile = ScriptFilePath_Gurobi()
    scriptFileContents = SolverString & " " & CommandLineRunString & PrintingOptionString
    CreateScriptFile scriptFile, scriptFileContents
    
    CreateSolveScript_Gurobi = scriptFile
End Function


Function ReadModel_Gurobi(SolutionFilePathName As String, errorString As String) As Boolean
          
19570     ReadModel_Gurobi = False
          Dim Line As String, index As Integer
19580     On Error GoTo readError
          Dim solutionExpected As Boolean
19590     solutionExpected = True
          
19600     Open SolutionFilePathName For Input As 1 ' supply path with filename
19610     Line Input #1, Line
          ' Check for python exception while running Gurobi
          Dim GurobiError As String ' The string that identifies a gurobi error in the model file
          GurobiError = "Gurobi Error: "
          If left(Line, Len(GurobiError)) = GurobiError Then
              errorString = Line
              GoTo exitFunction
          End If
          'Get the returned status code from gurobi.
          'List of return codes can be seen at - http://www.gurobi.com/documentation/5.1/reference-manual/node865#sec:StatusCodes
19620     If Line = GurobiResult.Optimal Then
19630         OpenSolver_Gurobi.SolveStatus = OpenSolverResult.Optimal
19640         OpenSolver_Gurobi.SolveStatusString = "Optimal"
19650         OpenSolver_Gurobi.LinearSolveStatus = LinearSolveResult.Optimal
19660     ElseIf Line = GurobiResult.Infeasible Then
19670         OpenSolver_Gurobi.SolveStatus = OpenSolverResult.Infeasible
19680         OpenSolver_Gurobi.SolveStatusString = "No Feasible Solution"
19690         solutionExpected = False
19700         OpenSolver_Gurobi.LinearSolveStatus = LinearSolveResult.Infeasible
19710     ElseIf Line = GurobiResult.InfOrUnbound Then
19720         OpenSolver_Gurobi.SolveStatus = OpenSolverResult.Unbounded
19730         OpenSolver_Gurobi.SolveStatusString = "No Solution Found (Infeasible or Unbounded)"
19740         solutionExpected = False
19750         OpenSolver_Gurobi.LinearSolveStatus = LinearSolveResult.Unbounded
19760     ElseIf Line = GurobiResult.Unbounded Then
19770         OpenSolver_Gurobi.SolveStatus = OpenSolverResult.Unbounded
19780         OpenSolver_Gurobi.SolveStatusString = "No Solution Found (Unbounded)"
19790         solutionExpected = False
19800         OpenSolver_Gurobi.LinearSolveStatus = LinearSolveResult.Unbounded
19810     ElseIf Line = GurobiResult.SolveStoppedTime Then
19820         OpenSolver_Gurobi.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
19830         OpenSolver_Gurobi.SolveStatusString = "Stopped on Time Limit"
19840         OpenSolver_Gurobi.LinearSolveStatus = LinearSolveResult.SolveStopped
19850     ElseIf Line = GurobiResult.SolveStoppedIter Then
19860         OpenSolver_Gurobi.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
19870         OpenSolver_Gurobi.SolveStatusString = "Stopped on Iteration Limit"
19880         OpenSolver_Gurobi.LinearSolveStatus = LinearSolveResult.SolveStopped
19890     ElseIf Line = GurobiResult.SolveStoppedUser Then
19900         OpenSolver_Gurobi.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
19910         OpenSolver_Gurobi.SolveStatusString = "Stopped on Ctrl-C"
19920         OpenSolver_Gurobi.LinearSolveStatus = LinearSolveResult.SolveStopped
19930     ElseIf Line = GurobiResult.Unsolved Then
19940         OpenSolver_Gurobi.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
19950         OpenSolver_Gurobi.SolveStatusString = "Stopped on Gurobi Numerical difficulties"
19960         OpenSolver_Gurobi.LinearSolveStatus = LinearSolveResult.SolveStopped
19970     ElseIf Line = GurobiResult.SubOptimal Then
19980         OpenSolver_Gurobi.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
19990         OpenSolver_Gurobi.SolveStatusString = "Unable to satisfy optimality tolerances; a sub-optimal solution is available."
20000         OpenSolver_Gurobi.LinearSolveStatus = LinearSolveResult.SolveStopped
20010     Else
20020         errorString = "The response from the Gurobi solver is not recognised. The response was: " & Line
20030         GoTo readError
20040     End If
          
20050     If solutionExpected Then
20060         Dim NumVar As Integer
              Line Input #1, Line  ' Optimal - objective value              22
20070         If Line <> "" Then
20080             index = InStr(Line, "=")
                  Dim ObjectiveValue As Double
20090             ObjectiveValue = Val(Mid(Line, index + 2))
                  Dim i As Integer
20100             i = 1
20110             While Not EOF(1)
20120                 Line Input #1, Line
20130                 index = InStr(Line, " ")
20160                 OpenSolver_Gurobi.FinalVarValueP(i) = Val(right(Line, Len(Line) - index))
                      'Get the variable name
20170                 OpenSolver_Gurobi.VarCellP(i) = left(Line, index - 1)
20180                 If left(OpenSolver_Gurobi.VarCellP(i), 1) = "_" Then
                          ' Strip any _ character added to make a valid name
                          OpenSolver_Gurobi.VarCellP(i) = Mid(OpenSolver_Gurobi.VarCellP(i), 2)
20190                 End If
                      ' Save number of vars read
                      NumVar = i
                      i = i + 1
20200             Wend
20210         End If
20220         OpenSolver_Gurobi.AdjustableCells.Value2 = 0
              Dim j As Integer
20240         For i = 1 To NumVar
                  ' Need to make sure number is in US locale when Value2 is set
20250             OpenSolver_Gurobi.AdjustableCells.Worksheet.Range(OpenSolver_Gurobi.VarCellP(i)).Value2 = ConvertFromCurrentLocale(OpenSolver_Gurobi.FinalVarValueP(i))
20260         Next i
              
20270         If OpenSolver_Gurobi.bGetDuals Then
20350             Open Replace(SolutionFilePathName, "modelsolution", "sensitivityData") For Input As 2
                  Dim index2 As Integer
                  Dim Stuff() As String
20360             ReDim Stuff(3)
20370             For i = 1 To NumVar
20380                 Line Input #2, Line
20390                 For j = 1 To 3
20400                     index2 = InStr(Line, ",")
20410                     If index2 <> 0 Then
20420                         Stuff(j) = left(Line, index2 - 1)
20430                     Else
20440                         Stuff(j) = Line
20450                     End If
20460                     Line = Mid(Line, index2 + 1)
20470                 Next j
20480                 OpenSolver_Gurobi.ReducedCostsP(i) = Stuff(1)
20490                 OpenSolver_Gurobi.IncreaseVarP(i) = Stuff(3) - OpenSolver_Gurobi.CostCoeffsP(i)
20500                 OpenSolver_Gurobi.DecreaseVarP(i) = OpenSolver_Gurobi.CostCoeffsP(i) - Stuff(2)
20510             Next i
20520             ReDim Stuff(5)
20530             For i = 1 To OpenSolver_Gurobi.NumRows
20540                 Line Input #2, Line
20550                 For j = 1 To 5
20560                     index2 = InStr(Line, ",")
20570                     If index2 <> 0 Then
20580                         Stuff(j) = left(Line, index2 - 1)
20590                     Else
20600                         Stuff(j) = Line
20610                     End If
20620                     Line = Mid(Line, index2 + 1)
20630                 Next j
20640                 OpenSolver_Gurobi.ShadowPriceP(i) = Stuff(1)
20650                 OpenSolver_Gurobi.IncreaseConP(i) = Stuff(5) - Stuff(2)
20660                 OpenSolver_Gurobi.DecreaseConP(i) = Stuff(2) - Stuff(4)
20670                 OpenSolver_Gurobi.FinalValueP(i) = Stuff(2) - Stuff(3)
20680             Next i
20690         End If
20700         ReadModel_Gurobi = True
20710     End If

exitFunction:
20720     Close #1
20730     Close #2
20740     Exit Function
          
readError:
20750     Close #1
20760     Close #2
End Function
