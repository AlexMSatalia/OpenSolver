Attribute VB_Name = "SolverCouenne"
Option Explicit

Public Const SolverTitle_Couenne = "COIN-OR Couenne (Non-linear, non-convex solver)"
Public Const SolverDesc_Couenne = "Couenne (Convex Over and Under ENvelopes for Nonlinear Estimation) is a branch & bound algorithm to solve Mixed-Integer Nonlinear Programming (MINLP) problems of specific forms. Couenne aims at finding global optima of nonconvex MINLPs. It implements linearization, bound reduction, and branching methods within a branch-and-bound framework. This solver will fail if your spreadsheet uses functions OpenSolver cannot interpret."
Public Const SolverLink_Couenne = "https://projects.coin-or.org/Couenne"
Public Const SolverType_Couenne = OpenSolver_SolverType.NonLinear

Public Const SolverName_Couenne = "Couenne"

#If Mac Then
Public Const SolverExec_Couenne = "couenne"
#Else
Public Const SolverExec_Couenne = "couenne.exe"
#End If

Public Const SolverScript_Couenne = "couenne" & ScriptExtension
Public Const OptionsFile_Couenne = "couenne.opt"

Public Const UsesPrecision_Couenne = False
Public Const UsesIterationLimit_Couenne = True
Public Const UsesTolerance_Couenne = True
Public Const UsesTimeLimit_Couenne = True

Function ScriptFilePath_Couenne() As String
8358      GetTempFilePath SolverScript_Couenne, ScriptFilePath_Couenne
End Function

Function SolutionFilePath_Couenne() As String
8359      GetTempFilePath NLSolutionFileName, SolutionFilePath_Couenne
End Function

Function OptionsFilePath_Couenne() As String
          GetTempFilePath OptionsFile_Couenne, OptionsFilePath_Couenne
End Function

Sub CleanFiles_Couenne()
8360      DeleteFileAndVerify SolutionFilePath_Couenne()
8361      DeleteFileAndVerify ScriptFilePath_Couenne()
          DeleteFileAndVerify OptionsFilePath_Couenne()
End Sub

Function About_Couenne() As String
      ' Return string for "About" form
          Dim SolverPath As String, errorString As String
8362      If Not SolverAvailable_Couenne(SolverPath, errorString) Then
8363          About_Couenne = errorString
8364          Exit Function
8365      End If
          ' Assemble version info
8366      About_Couenne = "Couenne " & SolverBitness_Couenne & "-bit" & _
                          " v" & SolverVersion_Couenne & _
                          " at " & MakeSpacesNonBreaking(MakePathSafe(SolverPath))
End Function

Function SolverFilePath_Couenne(Optional errorString As String) As String
8367      SolverFilePath_Couenne = SolverFilePath_Default("Couenne", errorString)
End Function

Function SolverAvailable_Couenne(Optional SolverPath As String, Optional errorString As String) As Boolean
      ' Returns true if Couenne is available and sets SolverPath
8368      SolverPath = SolverFilePath_Couenne(errorString)
8369      If SolverPath = "" Then
8370          SolverAvailable_Couenne = False
8371      Else
8372          SolverAvailable_Couenne = True
8373          errorString = "WARNING: Couenne is EXPERIMENTAL and is not guaranteed to give optimal or even good solutions. Proceed with caution." & vbCrLf & vbCrLf & errorString

#If Mac Then
              ' Make sure couenne is executable on Mac
8374          RunExternalCommand "chmod +x " & MakePathSafe(SolverPath)
#End If
          
8375      End If
End Function

Function SolverVersion_Couenne() As String
' Get Couenne version by running 'couenne -v' at command line
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim SolverPath As String
8376      If Not SolverAvailable_Couenne(SolverPath) Then
8377          SolverVersion_Couenne = ""
8378          GoTo ExitFunction
8379      End If
          
          Dim result As String
8386      result = ReadExternalCommandOutput(MakePathSafe(SolverPath) & " -v")
          SolverVersion_Couenne = Mid(result, 9, 5)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("SolverCouenne", "SolverVersion_Couenne") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function SolverBitness_Couenne() As String
      ' Get Bitness of Couenne solver
          Dim SolverPath As String
8400      If Not SolverAvailable_Couenne(SolverPath) Then
8401          SolverBitness_Couenne = ""
8402          Exit Function
8403      End If
              
          ' All Macs are 64-bit so we only provide 64-bit binaries
#If Mac Then
8404      SolverBitness_Couenne = "64"
#Else
8405      If right(SolverPath, 14) = "64\couenne.exe" Then
8406          SolverBitness_Couenne = "64"
8407      Else
8408          SolverBitness_Couenne = "32"
8409      End If
#End If
End Function

Function CreateSolveScript_Couenne(ModelFilePathName As String, SolveOptions As SolveOptionsType) As String
    CreateSolveScript_Couenne = CreateSolveScript_NL(ModelFilePathName, SolveOptions)
End Function
