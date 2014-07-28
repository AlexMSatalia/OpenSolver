Attribute VB_Name = "SolverCouenne"
Option Explicit

Public Const SolverTitle_Couenne = "COIN-OR Couenne (Non-linear Solver)"
Public Const SolverDesc_Couenne = "Couenne (Convex Over and Under ENvelopes for Nonlinear Estimation) is a branch & bound algorithm to solve Mixed-Integer Nonlinear Programming (MINLP) problems of specific forms. Couenne aims at finding global optima of nonconvex MINLPs. It implements linearization, bound reduction, and branching methods within a branch-and-bound framework."
Public Const SolverLink_Couenne = "https://projects.coin-or.org/Couenne"
Public Const SolverType_Couenne = OpenSolver_SolverType.NonLinear

Public Const SolverName_Couenne = "couenne.exe"
Public Const SolverScript_Couenne = "couenne" & ScriptExtension

Public Const SolutionFile_Couenne = "model.sol"

Function ScriptFilePath_Couenne() As String
    ScriptFilePath_Couenne = GetTempFilePath(SolverScript_Couenne)
End Function

Function SolutionFilePath_Couenne() As String
    SolutionFilePath_Couenne = GetTempFilePath(SolutionFile_Couenne)
End Function

Sub CleanFiles_Couenne(errorPrefix As String)
    ' Solution file
    DeleteFileAndVerify SolutionFilePath_Couenne(), errorPrefix, "Unable to delete the Couenne solver solution file: " & SolutionFilePath_Couenne()
    ' Script file
    DeleteFileAndVerify ScriptFilePath_Couenne(), errorPrefix, "Unable to delete the Couenne solver script file: " & ScriptFilePath_Couenne()
End Sub

Function About_Couenne() As String
' Return string for "About" form
    Dim SolverPath As String, errorString As String
    If Not SolverAvailable_Couenne(SolverPath, errorString) Then
        About_Couenne = errorString
        Exit Function
    End If
    ' Assemble version info
    About_Couenne = "Couenne " & SolverBitness_Couenne & "-bit" & _
                    " v" & SolverVersion_Couenne & _
                    " at " & SolverPath
End Function

Function SolverFilePath_Couenne(errorString As String) As String
#If Mac Then
    If GetExistingFilePathName(ThisWorkbook.Path, left(SolverName_Couenne, Len(SolverName_Couenne) - 4), SolverFilePath_Couenne) Then Exit Function ' Found a mac solver
    errorString = "Unable to find Mac version of Couenne ('couenne') in the folder that contains 'OpenSolver.xlam'"
    Exit Function
#Else
    ' Look for the 64 bit version
    If SystemIs64Bit Then
        If GetExistingFilePathName(ThisWorkbook.Path, Replace(SolverName_Couenne, ".exe", "64.exe"), SolverFilePath_Couenne) Then Exit Function ' Found a 64 bit solver
    End If
    ' Look for the 32 bit version
    If GetExistingFilePathName(ThisWorkbook.Path, SolverName_Couenne, SolverFilePath_Couenne) Then
        If SystemIs64Bit Then
            errorString = "Unable to find 64-bit Couenne (couenne64.exe) in the 'OpenSolver.xlam' folder. 32-bit Couenne will be used instead."
        End If
        Exit Function
    End If
    ' Fail
    SolverFilePath_Couenne = ""
    errorString = "Unable to find 32-bit Couenne (couenne.exe) or 64-bit Couenne (couenne64.exe) in the 'OpenSolver.xlam' folder."
#End If
End Function

Function SolverAvailable_Couenne(Optional SolverPath As String, Optional errorString As String) As Boolean
' Returns true if Couenne is available and sets SolverPath
    SolverPath = SolverFilePath_Couenne(errorString)
    If SolverPath = "" Then
        SolverAvailable_Couenne = False
    Else
        SolverAvailable_Couenne = True

#If Mac Then
        ' Make sure couenne is executable on Mac
        system ("chmod +x " & ConvertHfsPath(SolverPath))
#End If
    
    End If
End Function

Function SolverVersion_Couenne() As String
' Get Couenne version by running 'couenne -v' at command line
    Dim SolverPath As String
    If Not SolverAvailable_Couenne(SolverPath) Then
        SolverVersion_Couenne = ""
        Exit Function
    End If
    
    ' Set up Couenne to write version info to text file
    Dim logFile As String
    logFile = GetTempFolder & "couenneversion.txt"
    If FileOrDirExists(logFile) Then Kill logFile
    
    Dim RunPath As String, FileContents As String
    RunPath = ScriptFilePath_Couenne()
    If FileOrDirExists(RunPath) Then Kill RunPath
    FileContents = """" & ConvertHfsPath(SolverPath) & """" & " -v" & " > """ & ConvertHfsPath(logFile) & """"
    CreateScriptFile RunPath, FileContents
    
    ' Run Couenne
    Dim completed As Boolean
    completed = OSSolveSync(ConvertHfsPath(RunPath), "", "", "", SW_HIDE, True)
    
    ' Read version info back from output file
    Dim Line As String
    If FileOrDirExists(logFile) Then
        Open logFile For Input As 1
        Line Input #1, Line
        Close #1
        SolverVersion_Couenne = right(Line, Len(Line) - 8)
        SolverVersion_Couenne = left(SolverVersion_Couenne, 5)
    Else
        SolverVersion_Couenne = ""
    End If
End Function

Function SolverBitness_Couenne() As String
' Get Bitness of Couenne solver
    Dim SolverPath As String
    If Not SolverAvailable_Couenne(SolverPath) Then
        SolverBitness_Couenne = ""
        Exit Function
    End If
    
    If right(SolverPath, 6) = "64.exe" Or right(SolverPath, 2) = "64" Then
        SolverBitness_Couenne = "64"
    Else
        SolverBitness_Couenne = "32"
    End If
End Function


