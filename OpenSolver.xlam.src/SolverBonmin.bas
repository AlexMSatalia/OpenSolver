Attribute VB_Name = "SolverBonmin"
Option Explicit

Public Const SolverTitle_Bonmin = "COIN-OR Bonmin (Non-linear solver)"
Public Const SolverDesc_Bonmin = "Bonmin (Basic Open-source Nonlinear Mixed INteger programming) is an experimental open-source C++ code for solving general MINLPs (Mixed Integer NonLinear Programming). Finds globally optimal solutions to convex nonlinear problems in continuous and discrete variables, and may be applied heuristically to nonconvex problems. Bonmin uses the COIN-OR solvers CBC and IPOPT while solving. For more info on these, see www.coin-or.org/projects. This solver will fail if your spreadsheet uses functions OpenSolver cannot interpret."
Public Const SolverLink_Bonmin = "https://projects.coin-or.org/Bonmin"
Public Const SolverType_Bonmin = OpenSolver_SolverType.NonLinear

Public Const SolverName_Bonmin = "Bonmin"

#If Mac Then
Public Const SolverExec_Bonmin = "bonmin"
#Else
Public Const SolverExec_Bonmin = "bonmin.exe"
#End If

Public Const SolverScript_Bonmin = "bonmin" & ScriptExtension
Public Const OptionsFile_Bonmin = "bonmin.opt"

Public Const UsesPrecision_Bonmin = False
Public Const UsesIterationLimit_Bonmin = True
Public Const UsesTolerance_Bonmin = True
Public Const UsesTimeLimit_Bonmin = True

Function ScriptFilePath_Bonmin() As String
9021      GetTempFilePath SolverScript_Bonmin, ScriptFilePath_Bonmin
End Function

Function SolutionFilePath_Bonmin() As String
9022      GetTempFilePath NLSolutionFileName, SolutionFilePath_Bonmin
End Function

Function OptionsFilePath_Bonmin() As String
          GetTempFilePath OptionsFile_Bonmin, OptionsFilePath_Bonmin
End Function

Sub CleanFiles_Bonmin()
9023      DeleteFileAndVerify SolutionFilePath_Bonmin()
9024      DeleteFileAndVerify ScriptFilePath_Bonmin()
          DeleteFileAndVerify OptionsFilePath_Bonmin()
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
                          " at " & MakeSpacesNonBreaking(MakePathSafe(SolverPath))
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
9036          RunExternalCommand "chmod +x " & MakePathSafe(SolverPath)
#End If
          
9037      End If
End Function

Function SolverVersion_Bonmin() As String
      ' Get Bonmin version by running 'bonmin -v' at command line
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim SolverPath As String
9038      If Not SolverAvailable_Bonmin(SolverPath) Then
9039          SolverVersion_Bonmin = ""
9040          GoTo ExitFunction
9041      End If
          
          Dim result As String
9048      result = ReadExternalCommandOutput(MakePathSafe(SolverPath) & " -v")
          SolverVersion_Bonmin = Mid(result, 8, 5)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("SolverBonmin", "SolverVersion_Bonmin") Then Resume
          RaiseError = True
          GoTo ExitFunction
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

Function CreateSolveScript_Bonmin(ModelFilePathName As String, SolveOptions As SolveOptionsType) As String
    CreateSolveScript_Bonmin = CreateSolveScript_NL(ModelFilePathName, SolveOptions)
End Function
