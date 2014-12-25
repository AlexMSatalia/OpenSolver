Attribute VB_Name = "SolverBonmin"
Option Explicit

Public Const SolverTitle_Bonmin = "COIN-OR Bonmin (Non-linear solver)"
Public Const SolverDesc_Bonmin = "Bonmin (Basic Open-source Nonlinear Mixed INteger programming) is an experimental open-source C++ code for solving general MINLPs (Mixed Integer NonLinear Programming). Finds globally optimal solutions to convex nonlinear problems in continuous and discrete variables, and may be applied heuristically to nonconvex problems. Bonmin uses the COIN-OR solvers CBC and IPOPT while solving. For more info on these, see www.coin-or.org/projects. This solver will fail if your spreadsheet uses functions OpenSolver cannot interpret."
Public Const SolverLink_Bonmin = "https://projects.coin-or.org/Bonmin"
Public Const SolverType_Bonmin = OpenSolver_SolverType.NonLinear

#If Mac Then
Public Const SolverName_Bonmin = "bonmin"
#Else
Public Const SolverName_Bonmin = "bonmin.exe"
#End If

Public Const SolverScript_Bonmin = "bonmin" & ScriptExtension

Public Const SolutionFile_Bonmin = "model.sol"

Function ScriptFilePath_Bonmin() As String
9021      ScriptFilePath_Bonmin = GetTempFilePath(SolverScript_Bonmin)
End Function

Function SolutionFilePath_Bonmin() As String
9022      SolutionFilePath_Bonmin = GetTempFilePath(SolutionFile_Bonmin)
End Function

Sub CleanFiles_Bonmin(errorPrefix As String)
          ' Solution file
9023      DeleteFileAndVerify SolutionFilePath_Bonmin(), errorPrefix, "Unable to delete the Bonmin solver solution file: " & SolutionFilePath_Bonmin()
          ' Script file
9024      DeleteFileAndVerify ScriptFilePath_Bonmin(), errorPrefix, "Unable to delete the Bonmin solver script file: " & ScriptFilePath_Bonmin()
End Sub

Function About_Bonmin() As String
      ' Return string for "About" form
          Dim SolverPath As String, errorString As String
9025      If Not SolverAvailable_Bonmin(SolverPath, errorString) Then
9026          About_Bonmin = errorString
9027          Exit Function
9028      End If
          ' Assemble version info
9029      About_Bonmin = "Bonmin " & SolverBitness_Bonmin & "-bit" & _
                          " v" & SolverVersion_Bonmin & _
                          " at " & MakeSpacesNonBreaking(SolverPath)
End Function

Function SolverFilePath_Bonmin(Optional errorString As String) As String
9030      SolverFilePath_Bonmin = SolverFilePath_Default("Bonmin", errorString)
End Function

Function SolverAvailable_Bonmin(Optional SolverPath As String, Optional errorString As String) As Boolean
      ' Returns true if Bonmin is available and sets SolverPath
9031      SolverPath = SolverFilePath_Bonmin(errorString)
9032      If SolverPath = "" Then
9033          SolverAvailable_Bonmin = False
9034      Else
9035          SolverAvailable_Bonmin = True

#If Mac Then
              ' Make sure Bonmin is executable on Mac
9036          system ("chmod +x " & MakePathSafe(SolverPath))
#End If
          
9037      End If
End Function

Function SolverVersion_Bonmin() As String
      ' Get Bonmin version by running 'bonmin -v' at command line
          Dim SolverPath As String
9038      If Not SolverAvailable_Bonmin(SolverPath) Then
9039          SolverVersion_Bonmin = ""
9040          Exit Function
9041      End If
          
          ' Set up Bonmin to write version info to text file
          Dim logFile As String
9042      logFile = GetTempFilePath("bonminversion.txt")
9043      If FileOrDirExists(logFile) Then Kill logFile
          
          Dim RunPath As String, FileContents As String
9044      RunPath = ScriptFilePath_Bonmin()
9045      If FileOrDirExists(RunPath) Then Kill RunPath
9046      FileContents = MakePathSafe(SolverPath) & " -v"
9047      CreateScriptFile RunPath, FileContents
          
          ' Run Bonmin
          Dim completed As Boolean
9048      completed = RunExternalCommand(MakePathSafe(RunPath), MakePathSafe(logFile), SW_HIDE, True)
          
          ' Read version info back from output file
          Dim Line As String
9049      If FileOrDirExists(logFile) Then
9050          On Error GoTo ErrHandler
9051          Open logFile For Input As 1
9052          Line Input #1, Line
9053          Close #1
9054          SolverVersion_Bonmin = right(Line, Len(Line) - 7)
9055          SolverVersion_Bonmin = left(SolverVersion_Bonmin, 5)
9056      Else
9057          SolverVersion_Bonmin = ""
9058      End If
9059      Exit Function
          
ErrHandler:
9060      Close #1
9061      Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Function

Function SolverBitness_Bonmin() As String
      ' Get Bitness of Bonmin solver
          Dim SolverPath As String
9062      If Not SolverAvailable_Bonmin(SolverPath) Then
9063          SolverBitness_Bonmin = ""
9064          Exit Function
9065      End If
          
          ' All Macs are 64-bit so we only provide 64-bit binaries
#If Mac Then
9066      SolverBitness_Bonmin = "64"
#Else
9067      If right(SolverPath, 13) = "64\bonmin.exe" Then
9068          SolverBitness_Bonmin = "64"
9069      Else
9070          SolverBitness_Bonmin = "32"
9071      End If
#End If
End Function

Function CreateSolveScript_Bonmin(ModelFilePathName As String) As String
          ' Create a script to run "/path/to/bonmin.exe /path/to/<ModelFilePathName>"

          Dim SolverString As String, CommandLineRunString As String, PrintingOptionString As String
9072      SolverString = MakePathSafe(SolverFilePath_Bonmin())

9073      CommandLineRunString = MakePathSafe(ModelFilePathName)
          
          Dim scriptFile As String, scriptFileContents As String
9075      scriptFile = ScriptFilePath_Bonmin()
9076      scriptFileContents = SolverString & " " & CommandLineRunString
9077      CreateScriptFile scriptFile, scriptFileContents
          
9078      CreateSolveScript_Bonmin = scriptFile
End Function

Function ReadModel_Bonmin(SolutionFilePathName As String, errorString As String, m As CModelParsed, s As COpenSolverParsed) As Boolean
9079      ReadModel_Bonmin = False
          Dim Line As String, index As Long
9080      On Error GoTo readError
          Dim solutionExpected As Boolean
9081      solutionExpected = True
          
9082      If Not FileOrDirExists(SolutionFilePathName) Then
9083          solutionExpected = False
9084          If Not TryParseLogs(s) Then
9085              errorString = "The solver did not create a solution file. No new solution is available."
9086              GoTo exitFunction
9087          End If
9088      Else
9089          Open SolutionFilePathName For Input As 1 ' supply path with filename
9090          Line Input #1, Line ' Skip empty line at start of file
9091          Line Input #1, Line
9092          Line = Mid(Line, 9)
              
              'Get the returned status code from Bonmin.
9093          If Line Like "Optimal*" Then
9094              s.SolveStatus = OpenSolverResult.Optimal
9095              s.SolveStatusString = "Optimal"
9096          ElseIf Line Like "Infeasible*" Then
9097              s.SolveStatus = OpenSolverResult.Infeasible
9098              s.SolveStatusString = "No Feasible Solution"
9099              solutionExpected = False
9100          ElseIf Line Like "*unbounded*" Then
9101              s.SolveStatus = OpenSolverResult.Unbounded
9102              s.SolveStatusString = "No Solution Found (Unbounded)"
9103              solutionExpected = False
9104          ElseIf Line Like "Error encountered in optimization*" Then
                  ' Try to get status from logs
9105              If Not TryParseLogs(s) Then
9106                  errorString = "Bonmin did not solve the problem, suggesting there was an error in the input parameters. The response was: " & vbCrLf & _
                                    Line & vbCrLf & _
                                    "The Bonmin command line can be found at:" & vbCrLf & _
                                    ScriptFilePath_Bonmin()
9107                  GoTo exitFunction
9108              End If
9109              solutionExpected = False
9110          Else
9111              errorString = "The response from the Bonmin solver is not recognised. The response was: " & _
                                Line & vbCrLf & _
                                "The Bonmin command line can be found at:" & vbCrLf & _
                                ScriptFilePath_Bonmin()
9112              GoTo exitFunction
9113          End If
9114      End If
          
9115      If solutionExpected Then
9116          Application.StatusBar = "OpenSolver: Loading Solution... " & s.SolveStatusString
              
9117          Line Input #1, Line ' Throw away blank line
9118          Line Input #1, Line ' Throw away "Options"
              
              Dim i As Long
9119          For i = 1 To 8
9120              Line Input #1, Line ' Skip all options lines
9121          Next i
              
              ' Note that the variable values are written to file in .nl format
              ' We need to read in the values and the extract the correct values for the adjustable cells
              
              ' Read in all variable values
              Dim VariableValues As New Collection
9122          While Not EOF(1)
9123              Line Input #1, Line
9124              VariableValues.Add CDbl(Line)
9125          Wend
              
              ' Loop through variable cells and find the corresponding value from VariableValues
9126          i = 1
              Dim c As Range, VariableIndex As Long
9127          For Each c In m.AdjustableCells
                  ' Extract the correct variable value
9128              VariableIndex = GetVariableNLIndex(i) + 1
                  
                  ' Need to make sure number is in US locale when Value2 is set
9129              Range(c.Address).Value2 = ConvertFromCurrentLocale(VariableValues(VariableIndex))
9130              i = i + 1
9131          Next c
9132      End If
9133      ReadModel_Bonmin = True

exitFunction:
9134      Application.StatusBar = False
9135      Close #1
9136      Close #2
9137      Exit Function
          
readError:
9138      Application.StatusBar = False
9139      Close #1
9140      Close #2
9141      Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Function

Function TryParseLogs(s As COpenSolverParsed) As Boolean
      ' We examine the log file if it exists to try to find more info about the solve
          
          ' Check if log exists
          Dim logFile As String
9142      logFile = GetTempFilePath("log1.tmp")
          
9143      If Not FileOrDirExists(logFile) Then
9144          TryParseLogs = False
9145          Exit Function
9146      End If
          
          Dim message As String
9147      On Error GoTo ErrHandler
9148      Open logFile For Input As 3
9149      message = Input$(LOF(3), 3)
9150      Close #3
          
9151      If Not left(message, 6) = "Bonmin" Then
             ' Not dealing with a Bonmin log, abort
9152          TryParseLogs = False
9153          Exit Function
9154      End If
          
          ' Scan for information
          
          ' 1 - scan for infeasible
9155      If message Like "*infeasible*" Then
9156          s.SolveStatus = OpenSolverResult.Infeasible
9157          s.SolveStatusString = "No Feasible Solution"
9158          TryParseLogs = True
9159          Exit Function
9160      End If
          
ErrHandler:
9161      Close #3
9162      TryParseLogs = False
End Function
