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

Public Const ToleranceName_Gurobi = "MIPGap"
Public Const TimeLimitName_Gurobi = "TimeLimit"
Public Const IterationLimitName_Gurobi = "IterationLimit"

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
6369      GetTempFilePath SolutionFile_Gurobi, SolutionFilePath_Gurobi
End Function

Function SolverPythonScriptPath_Gurobi() As String
6370      GetExistingFilePathName JoinPaths(ThisWorkbook.Path, SolverDir), SolverPythonScript_Gurobi, SolverPythonScriptPath_Gurobi
End Function

Function ScriptFilePath_Gurobi() As String
6371      GetTempFilePath SolverScript_Gurobi, ScriptFilePath_Gurobi
End Function

Function SensitivityFilePath_Gurobi() As String
6372      GetTempFilePath SensitivityFile_Gurobi, SensitivityFilePath_Gurobi
End Function

Sub CleanFiles_Gurobi()
6373      DeleteFileAndVerify SolutionFilePath_Gurobi()
          DeleteFileAndVerify ScriptFilePath_Gurobi()
6374      DeleteFileAndVerify SensitivityFilePath_Gurobi()
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
                           " detected at " & MakeSpacesNonBreaking(MakePathSafe(SolverPath))
End Function

Function GetGurobiBinFolder() As String
#If Mac Then
6381      GetGurobiBinFolder = JoinPaths(GetRootDriveName(), "usr", "local", "bin")
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
6383      SolverFilePath_Gurobi = ConvertHfsPathToPosix(JoinPaths(GetRootDriveName(), "usr", "bin", "python"))
#Else
6384      GetExistingFilePathName GetGurobiBinFolder(), Solver_Gurobi, SolverFilePath_Gurobi
#End If
End Function

Function SolverAvailable_Gurobi(Optional SolverPath As String, Optional errorString As String) As Boolean
' Returns true if Gurobi is available and sets SolverPath as path to gurobi_cl
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          If Not FileOrDirExists(SolverPythonScriptPath_Gurobi()) Then
              errorString = "Unable to find OpenSolver Gurobi script ('" & SolverPythonScript_Gurobi & "'). Folders searched:" & _
                            vbNewLine & MakePathSafe(JoinPaths(ThisWorkbook.Path, SolverDir))
6385      ElseIf Not GetExistingFilePathName(GetGurobiBinFolder, SolverExec_Gurobi, SolverPath) Then
6389          errorString = "No Gurobi installation was detected."
6391      End If

          If errorString <> "" Then
6388          SolverPath = ""
6390          SolverAvailable_Gurobi = False
          Else
              SolverAvailable_Gurobi = True
          End If

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("SolverGurobi", "SolverAvailable_Gurobi") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function SolverVersion_Gurobi() As String
' Get Gurobi version by running 'gurobi_cl -v' at command line
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim SolverPath As String
6392      If Not SolverAvailable_Gurobi(SolverPath) Then
6393          SolverVersion_Gurobi = ""
6394          GoTo ExitFunction
6395      End If
          
          Dim result As String
6402      result = ReadExternalCommandOutput(MakePathSafe(SolverPath) & " -v")
          SolverVersion_Gurobi = Mid(result, 26, 5)

ExitFunction:
          Close #1
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("SolverGurobi", "SolverVersion_Gurobi") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function SolverBitness_Gurobi() As String
' Get Gurobi bitness by running 'gurobi_cl -v' at command line
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim SolverPath As String
6416      If Not SolverAvailable_Gurobi(SolverPath) Then
6417          SolverBitness_Gurobi = ""
6418          GoTo ExitFunction
6419      End If

          Dim result As String
          result = ReadExternalCommandOutput(MakePathSafe(SolverPath) & " -v")
          SolverBitness_Gurobi = IIf(InStr(result, "64)") > 0, "64", "32")

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("SolverGurobi", "SolverBitness_Gurobi") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function
Function CreateSolveScript_Gurobi(SolutionFilePathName As String, SolverParameters As Dictionary) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim SolverString As String, CommandLineRunString As String, SolverParametersString As String
6441      SolverString = MakePathSafe(SolverFilePath_Gurobi())

6442      CommandLineRunString = MakePathSafe(SolverPythonScriptPath_Gurobi())
          
          SolverParametersString = ParametersToKwargs(SolverParameters)
          
          Dim scriptFile As String, scriptFileContents As String
6444      scriptFile = ScriptFilePath_Gurobi()
6445      scriptFileContents = SolverString & " " & CommandLineRunString & " " & SolverParametersString
6446      CreateScriptFile scriptFile, scriptFileContents
          
6447      CreateSolveScript_Gurobi = scriptFile

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("SolverGurobi", "CreateSolveScript_Gurobi") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function


Function ReadModel_Gurobi(SolutionFilePathName As String, s As COpenSolver) As Boolean
          
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

6448      ReadModel_Gurobi = False
          
          ' Check logs for invalid parameter values
          Dim logFile As String, message As String
7100      If Not GetTempFilePath("log1.tmp", logFile) Then GoTo ExitFunction
7105      Open logFile For Input As #1
7106      message = Input$(LOF(1), 1)
7107      Close #1

          Dim Key As Variant, SolverParameters As Dictionary
          s.CopySolverParameters SolverParameters
          For Each Key In SolverParameters.Keys
              If InStrText(message, "No parameters matching '" & Key & "' found") Then
                  s.SolveStatus = OpenSolverResult.ErrorOccurred
                  s.SolveStatusString = "The parameter '" & Key & "' was not recognised by Gurobi. Please check the parameter name you have specified, or consult the Gurobi documentation for more information."
                  GoTo ExitFunction
              End If
          Next Key
          
6450      s.SolutionWasLoaded = True
          
6451      Open SolutionFilePathName For Input As #1 ' supply path with filename
          Dim Line As String, Index As Long
6452      Line Input #1, Line
          ' Check for python exception while running Gurobi
          Dim GurobiError As String ' The string that identifies a gurobi error in the model file
6453      GurobiError = "Gurobi Error: "
6454      If left(Line, Len(GurobiError)) = GurobiError Then
6455          Err.Raise OpenSolver_GurobiError, Description:=Line
6457      End If
          'Get the returned status code from gurobi.
          'List of return codes can be seen at - http://www.gurobi.com/documentation/5.1/reference-manual/node865#sec:StatusCodes
6458      If Line = GurobiResult.Optimal Then
6459          s.SolveStatus = OpenSolverResult.Optimal
6460          s.SolveStatusString = "Optimal"
6462      ElseIf Line = GurobiResult.Infeasible Then
6463          s.SolveStatus = OpenSolverResult.Infeasible
6464          s.SolveStatusString = "No Feasible Solution"
6465          s.SolutionWasLoaded = False
6467      ElseIf Line = GurobiResult.InfOrUnbound Then
6468          s.SolveStatus = OpenSolverResult.Unbounded
6469          s.SolveStatusString = "No Solution Found (Infeasible or Unbounded)"
6470          s.SolutionWasLoaded = False
6472      ElseIf Line = GurobiResult.Unbounded Then
6473          s.SolveStatus = OpenSolverResult.Unbounded
6474          s.SolveStatusString = "No Solution Found (Unbounded)"
6475          s.SolutionWasLoaded = False
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
6498          Err.Raise OpenSolver_GurobiError = "The response from the Gurobi solver is not recognised. The response was: " & Line
6500      End If
          
6501      If s.SolutionWasLoaded Then
6502          UpdateStatusBar "OpenSolver: Loading Solution... " & s.SolveStatusString, True
              Dim NumVar As Long, SplitLine() As String
6503          Line Input #1, Line  ' Optimal - objective value              22
6504          If Line <> "" Then
6505              Index = InStr(Line, "=")
                  Dim ObjectiveValue As Double
6506              ObjectiveValue = ConvertToCurrentLocale(Mid(Line, Index + 2))
                  Dim i As Long
6507              i = 1
6508              While Not EOF(1)
6509                  Line Input #1, Line
6510                  SplitLine = SplitWithoutRepeats(Line, " ")
                      Dim FinalValue As String
                      FinalValue = SplitLine(1)
                      
                      ' Check for an exponent that is too large, rougly >300
                      ' Only do it for -ve exponents, since e-30 ~= e-300
                      Index = InStrText(FinalValue, "e-")
                      If Index > 0 Then
                          ' Trim the final digit if the exponent has 3 digits
                          If Len(FinalValue) - Index - 1 > 2 Then
                              FinalValue = left(FinalValue, Len(FinalValue) - 1)
                          End If
                      End If
                      
6511                  s.FinalVarValue(i) = ConvertToCurrentLocale(FinalValue)
                      s.VarCell(i) = SplitLine(0)
6513                  If left(s.VarCell(i), 1) = "_" Then
                          ' Strip any _ character added to make a valid name
6514                      s.VarCell(i) = Mid(s.VarCell(i), 2)
6515                  End If
                      ' Save number of vars read
6516                  NumVar = i
6517                  i = i + 1
6518              Wend
6519          End If
              
6524          If s.bGetDuals Then
6525              Open Replace(SolutionFilePathName, "modelsolution", "sensitivityData") For Input As 2
6527              For i = 1 To NumVar
6528                  Line Input #2, Line
6529                  SplitLine = SplitWithoutRepeats(Line, ",")
6538                  s.ReducedCosts(i) = ConvertToCurrentLocale(SplitLine(0))
6540                  s.DecreaseVar(i) = s.CostCoeffs(i) - ConvertToCurrentLocale(SplitLine(1))
6539                  s.IncreaseVar(i) = ConvertToCurrentLocale(SplitLine(2)) - s.CostCoeffs(i)
6541              Next i

6543              For i = 1 To s.NumRows
6544                  Line Input #2, Line
6545                  SplitLine = SplitWithoutRepeats(Line, ",")
6554                  s.ShadowPrice(i) = ConvertToCurrentLocale(SplitLine(0))
                      Dim RHSValue As Double
                      RHSValue = ConvertToCurrentLocale(SplitLine(1))
6555                  s.IncreaseCon(i) = ConvertToCurrentLocale(SplitLine(4)) - RHSValue
6556                  s.DecreaseCon(i) = RHSValue - ConvertToCurrentLocale(SplitLine(3))
6557                  s.FinalValue(i) = RHSValue - ConvertToCurrentLocale(SplitLine(2))
6558              Next i
6559          End If
6560          ReadModel_Gurobi = True
6561      End If

ExitFunction:
6562      Application.StatusBar = False
6563      Close #1
6564      Close #2
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("SolverGurobi", "ReadModel_Gurobi") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function
