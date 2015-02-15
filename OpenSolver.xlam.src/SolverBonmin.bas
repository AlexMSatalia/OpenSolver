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

Public Const UsesPrecision_Bonmin = False
Public Const UsesIterationLimit_Bonmin = True
Public Const UsesTolerance_Bonmin = True
Public Const UsesTimeLimit_Bonmin = True

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
