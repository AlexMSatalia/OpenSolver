Attribute VB_Name = "SolverGurobi"
Option Explicit

Public Const SolverTitle_Gurobi = "Gurobi (Linear solver)"
Public Const SolverDesc_Gurobi = "Gurobi is a solver for linear programming (LP), quadratic and quadratically constrained programming (QP and QCP), and mixed-integer programming (MILP, MIQP, and MIQCP). It requires the user to download and install a version of the Gurobi and to have GurobiOSRun.py in the OpenSolver directory."
Public Const SolverLink_Gurobi = "http://www.gurobi.com/resources/documentation"
Public Const SolverType_Gurobi = OpenSolver_SolverType.Linear

Public Const SolverName_Gurobi = "Gurobi"

#If Mac Then
Public Const SolverExec_Gurobi = "gurobi_cl"
#Else
Public Const SolverExec_Gurobi = "gurobi_cl.exe"
#End If

Public Const SolverScript_Gurobi = "gurobi_tmp" & ScriptExtension
Public Const SolverPythonScript_Gurobi = "gurobiOSRun.py"
Public Const Solver_Gurobi = "gurobi" & ScriptExtension

Public Const SolutionFile_Gurobi = "modelsolution.sol"
Public Const SensitivityFile_Gurobi = "sensitivityData.sol"

Public Const UsesPrecision_Gurobi = False
Public Const UsesIterationLimit_Gurobi = False
Public Const UsesTolerance_Gurobi = True
Public Const UsesTimeLimit_Gurobi = True

'Gurobi return status codes
Public Enum GurobiResult
    Optimal = 2
    Infeasible = 3
    InfOrUnbound = 4
    Unbounded = 5
    SolveStoppedIter = 7
    SolveStoppedTime = 9
    SolveStoppedUser = 11
    Unsolved = 12
    SubOptimal = 13
End Enum

Function SolutionFilePath_Gurobi() As String
6369      SolutionFilePath_Gurobi = GetTempFilePath(SolutionFile_Gurobi)
End Function

Function SolverPythonScriptPath_Gurobi() As String
6370      GetExistingFilePathName JoinPaths(ThisWorkbook.Path, SolverDir), SolverPythonScript_Gurobi, SolverPythonScriptPath_Gurobi
End Function

Function ScriptFilePath_Gurobi() As String
6371      ScriptFilePath_Gurobi = GetTempFilePath(SolverScript_Gurobi)
End Function

Function SensitivityFilePath_Gurobi() As String
6372      SensitivityFilePath_Gurobi = GetTempFilePath(SensitivityFile_Gurobi)
End Function

Sub CleanFiles_Gurobi(errorPrefix As String)
          ' Solution file
6373      DeleteFileAndVerify SolutionFilePath_Gurobi(), errorPrefix, "Unable to delete the Gurobi solver solution file: " & SolutionFilePath_Gurobi()
          ' Cost Range file
6374      DeleteFileAndVerify SensitivityFilePath_Gurobi(), errorPrefix, "Unable to delete the Gurobi solver sensitivity data file: " & SensitivityFilePath_Gurobi()
End Sub

Function About_Gurobi() As String
      ' Return string for "About" form
          Dim SolverPath As String, errorString As String
6375      If Not SolverAvailable_Gurobi(SolverPath, errorString) Then
6376          About_Gurobi = errorString
6377          Exit Function
6378      End If
          
          ' Assemble version info
6380      About_Gurobi = "Gurobi " & SolverBitness_Gurobi & "-bit" & _
                           " v" & SolverVersion_Gurobi & _
                           " detected at " & MakeSpacesNonBreaking(ConvertHfsPath(SolverPath))
End Function

Function GetGurobiBinFolder() As String
#If Mac Then
6381      GetGurobiBinFolder = GetDriveName() & ":usr:local:bin:"
#Else
6382      GetExistingFilePathName Environ("GUROBI_HOME"), "bin", GetGurobiBinFolder
#End If
End Function

Function SolverFilePath_Gurobi() As String
#If Mac Then
          ' On Mac, using the gurobi interactive shell causes errors when there are spaces in the filepath.
          ' The mac gurobi.sh script, unlike windows, doesn't have a check for a gurobi install, thus it doesn't do anything for us here and is safe to skip.
          ' We can just run python by itself. We need to use the default system python (pre-installed on mac) and not any other version (e.g. a version from homebrew)
          ' We also need to launch it without going via /Volumes/.../
6383      SolverFilePath_Gurobi = "/usr/bin/python"
#Else
6384      GetExistingFilePathName GetGurobiBinFolder(), Solver_Gurobi, SolverFilePath_Gurobi
#End If
End Function

Function SolverAvailable_Gurobi(Optional SolverPath As String, Optional errorString As String) As Boolean
      ' Returns true if Gurobi is available and sets SolverPath as path to gurobi_cl
6385      If GetExistingFilePathName(GetGurobiBinFolder, SolverExec_Gurobi, SolverPath) And _
             FileOrDirExists(SolverPythonScriptPath_Gurobi()) Then
6386          SolverAvailable_Gurobi = True
6387      Else
6388          SolverPath = ""
6389          errorString = "No Gurobi installation was detected."
6390          SolverAvailable_Gurobi = False
6391      End If
End Function

Function SolverVersion_Gurobi() As String
      ' Get Gurobi version by running 'gurobi_cl -v' at command line
          Dim SolverPath As String
6392      If Not SolverAvailable_Gurobi(SolverPath) Then
6393          SolverVersion_Gurobi = ""
6394          Exit Function
6395      End If
          
          ' Set up Gurobi to write version info to text file
          Dim logFile As String
6396      logFile = GetTempFilePath("gurobiversion.txt")
6397      If FileOrDirExists(logFile) Then Kill logFile
          
          Dim RunPath As String, FileContents As String
6398      RunPath = ScriptFilePath_Gurobi()
6399      If FileOrDirExists(RunPath) Then Kill RunPath
6400      FileContents = MakePathSafe(SolverPath) & " -v"
6401      CreateScriptFile RunPath, FileContents
          
          ' Run Gurobi
          Dim completed As Boolean
6402      completed = RunExternalCommand(MakePathSafe(RunPath), MakePathSafe(logFile), SW_HIDE, True)
          
          ' Read version info back from output file
          ' Output like 'Gurobi Optimizer version 5.6.3 (win64)'
          Dim Line As String
6403      If FileOrDirExists(logFile) Then
6404          On Error GoTo ErrHandler
6405          Open logFile For Input As 1
6406          Line Input #1, Line
6407          Close #1
6408          SolverVersion_Gurobi = Mid(Line, 26, 5)
6410      Else
6411          SolverVersion_Gurobi = ""
6412      End If
6413      Exit Function
          
ErrHandler:
6414      Close #1
6415      Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Function

Function SolverBitness_Gurobi() As String
      ' Get Gurobi bitness by running 'gurobi_cl -v' at command line
          Dim SolverPath As String
6416      If Not SolverAvailable_Gurobi(SolverPath) Then
6417          SolverBitness_Gurobi = ""
6418          Exit Function
6419      End If
          
          ' Set up Gurobi to write version info to text file
          Dim logFile As String
6420      logFile = GetTempFilePath("gurobiversion.txt")
6421      If FileOrDirExists(logFile) Then Kill logFile
          
          Dim RunPath As String, FileContents As String
6422      RunPath = ScriptFilePath_Gurobi()
6423      If FileOrDirExists(RunPath) Then Kill RunPath
6424      FileContents = MakePathSafe(SolverPath) & " -v"
6425      CreateScriptFile RunPath, FileContents
          
          ' Run Gurobi
          Dim completed As Boolean
6426      completed = RunExternalCommand(MakePathSafe(RunPath), MakePathSafe(logFile), SW_HIDE, True)
          
          ' Read bitness info back from output file
          ' Output like 'Gurobi Optimizer version 5.6.3 (win64)'
          Dim Line As String
6427      If FileOrDirExists(logFile) Then
6428          On Error GoTo ErrHandler
6429          Open logFile For Input As 1
6430          Line Input #1, Line
6431          Close #1
6432          If right(Line, 3) = "64)" Then
6433              SolverBitness_Gurobi = "64"
6434          Else
6435              SolverBitness_Gurobi = "32"
6436          End If
6437      End If
6438      Exit Function
          
ErrHandler:
6439      Close #1
6440      Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Function
Function CreateSolveScript_Gurobi(SolutionFilePathName As String, ExtraParameters As Dictionary, SolveOptions As SolveOptionsType) As String
          Dim SolverString As String, CommandLineRunString As String, PrintingOptionString As String, ExtraParametersString As String
6441      SolverString = MakePathSafe(SolverFilePath_Gurobi())

6442      CommandLineRunString = MakePathSafe(SolverPythonScriptPath_Gurobi())
          '========================================================================================
          'Gurobi can also be run at the command line using gurobi_cl with the following commands
          '
          'Solver="gurobi_cl.exe"
          'CommandLineRunString = " Threads=1" & " TimeLimit=" & Replace(Str(SolveOptions.maxTime), " ", "") & " ResultFile="
          '             & GetTempFolder & Replace(SolutionFileName, ".txt", ".sol") _
          '             & " " & ModelFilePathName
          '
          '========================================================================================
6443      PrintingOptionString = "TimeLimit=" & Trim(str(SolveOptions.MaxTime)) & " " & _
                                 "MIPGap=" & Trim(str(SolveOptions.Tolerance))
          
          ExtraParametersString = ParametersToString_Gurobi(ExtraParameters)
          
          Dim scriptFile As String, scriptFileContents As String
6444      scriptFile = ScriptFilePath_Gurobi()
6445      scriptFileContents = SolverString & " " & CommandLineRunString & " " & PrintingOptionString & " " & ExtraParametersString
6446      CreateScriptFile scriptFile, scriptFileContents
          
6447      CreateSolveScript_Gurobi = scriptFile
End Function

Function ParametersToString_Gurobi(ExtraParameters As Dictionary) As String
          Dim ParamPair As KeyValuePair
          For Each ParamPair In ExtraParameters.KeyValuePairs
              ParametersToString_Gurobi = ParametersToString_Gurobi & ParamPair.Key & "=" & ParamPair.value & " "
          Next
          ParametersToString_Gurobi = Trim(ParametersToString_Gurobi)
End Function


Function ReadModel_Gurobi(SolutionFilePathName As String, errorString As String, s As COpenSolver) As Boolean
          
6448      ReadModel_Gurobi = False
          Dim Line As String, Index As Long
6449      On Error GoTo readError
          Dim solutionExpected As Boolean
6450      solutionExpected = True
          
6451      Open SolutionFilePathName For Input As 1 ' supply path with filename
6452      Line Input #1, Line
          ' Check for python exception while running Gurobi
          Dim GurobiError As String ' The string that identifies a gurobi error in the model file
6453      GurobiError = "Gurobi Error: "
6454      If left(Line, Len(GurobiError)) = GurobiError Then
6455          errorString = Line
6456          GoTo exitFunction
6457      End If
          'Get the returned status code from gurobi.
          'List of return codes can be seen at - http://www.gurobi.com/documentation/5.1/reference-manual/node865#sec:StatusCodes
6458      If Line = GurobiResult.Optimal Then
6459          s.SolveStatus = OpenSolverResult.Optimal
6460          s.SolveStatusString = "Optimal"
6462      ElseIf Line = GurobiResult.Infeasible Then
6463          s.SolveStatus = OpenSolverResult.Infeasible
6464          s.SolveStatusString = "No Feasible Solution"
6465          solutionExpected = False
6467      ElseIf Line = GurobiResult.InfOrUnbound Then
6468          s.SolveStatus = OpenSolverResult.Unbounded
6469          s.SolveStatusString = "No Solution Found (Infeasible or Unbounded)"
6470          solutionExpected = False
6472      ElseIf Line = GurobiResult.Unbounded Then
6473          s.SolveStatus = OpenSolverResult.Unbounded
6474          s.SolveStatusString = "No Solution Found (Unbounded)"
6475          solutionExpected = False
6477      ElseIf Line = GurobiResult.SolveStoppedTime Then
6478          s.SolveStatus = OpenSolverResult.LimitedSubOptimal
6479          s.SolveStatusString = "Stopped on Time Limit"
6481      ElseIf Line = GurobiResult.SolveStoppedIter Then
6482          s.SolveStatus = OpenSolverResult.LimitedSubOptimal
6483          s.SolveStatusString = "Stopped on Iteration Limit"
6485      ElseIf Line = GurobiResult.SolveStoppedUser Then
6486          s.SolveStatus = OpenSolverResult.LimitedSubOptimal
6487          s.SolveStatusString = "Stopped on Ctrl-C"
6489      ElseIf Line = GurobiResult.Unsolved Then
6490          s.SolveStatus = OpenSolverResult.LimitedSubOptimal
6491          s.SolveStatusString = "Stopped on Gurobi Numerical difficulties"
6493      ElseIf Line = GurobiResult.SubOptimal Then
6494          s.SolveStatus = OpenSolverResult.LimitedSubOptimal
6495          s.SolveStatusString = "Unable to satisfy optimality tolerances; a sub-optimal solution is available."
6497      Else
6498          errorString = "The response from the Gurobi solver is not recognised. The response was: " & Line
6499          GoTo readError
6500      End If
          
6501      If solutionExpected Then
6502          Application.StatusBar = "OpenSolver: Loading Solution... " & s.SolveStatusString
              Dim NumVar As Long
6503          Line Input #1, Line  ' Optimal - objective value              22
6504          If Line <> "" Then
6505              Index = InStr(Line, "=")
                  Dim ObjectiveValue As Double
6506              ObjectiveValue = Val(Mid(Line, Index + 2))
                  Dim i As Long
6507              i = 1
6508              While Not EOF(1)
6509                  Line Input #1, Line
6510                  Index = InStr(Line, " ")
6511                  s.FinalVarValueP(i) = Val(Mid(Line, Index + 1))
                      'Get the variable name
6512                  s.VarCellP(i) = left(Line, Index - 1)
6513                  If left(s.VarCellP(i), 1) = "_" Then
                          ' Strip any _ character added to make a valid name
6514                      s.VarCellP(i) = Mid(s.VarCellP(i), 2)
6515                  End If
                      ' Save number of vars read
6516                  NumVar = i
6517                  i = i + 1
6518              Wend
6519          End If
6520          s.AdjustableCells.Value2 = 0
              Dim j As Long
6521          For i = 1 To NumVar
                  ' Need to make sure number is in US locale when Value2 is set
6522              s.AdjustableCells.Worksheet.Range(s.VarCellP(i)).Value2 = ConvertFromCurrentLocale(s.FinalVarValueP(i))
6523          Next i
              
6524          If s.bGetDuals Then
6525              Open Replace(SolutionFilePathName, "modelsolution", "sensitivityData") For Input As 2
                  Dim index2 As Long
                  Dim Stuff() As String
6526              ReDim Stuff(3)
6527              For i = 1 To NumVar
6528                  Line Input #2, Line
6529                  For j = 1 To 3
6530                      index2 = InStr(Line, ",")
6531                      If index2 <> 0 Then
6532                          Stuff(j) = left(Line, index2 - 1)
6533                      Else
6534                          Stuff(j) = Line
6535                      End If
6536                      Line = Mid(Line, index2 + 1)
6537                  Next j
6538                  s.ReducedCostsP(i) = Stuff(1)
6539                  s.IncreaseVarP(i) = Stuff(3) - s.CostCoeffsP(i)
6540                  s.DecreaseVarP(i) = s.CostCoeffsP(i) - Stuff(2)
6541              Next i
6542              ReDim Stuff(5)
6543              For i = 1 To s.NumRows
6544                  Line Input #2, Line
6545                  For j = 1 To 5
6546                      index2 = InStr(Line, ",")
6547                      If index2 <> 0 Then
6548                          Stuff(j) = left(Line, index2 - 1)
6549                      Else
6550                          Stuff(j) = Line
6551                      End If
6552                      Line = Mid(Line, index2 + 1)
6553                  Next j
6554                  s.ShadowPriceP(i) = Stuff(1)
6555                  s.IncreaseConP(i) = Stuff(5) - Stuff(2)
6556                  s.DecreaseConP(i) = Stuff(2) - Stuff(4)
6557                  s.FinalValueP(i) = Stuff(2) - Stuff(3)
6558              Next i
6559          End If
6560          ReadModel_Gurobi = True
6561      End If

exitFunction:
6562      Application.StatusBar = False
6563      Close #1
6564      Close #2
6565      Exit Function
          
readError:
6566      Application.StatusBar = False
6567      Close #1
6568      Close #2
6569      Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Function
