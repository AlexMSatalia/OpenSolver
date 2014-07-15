Attribute VB_Name = "SolverGurobi"

Option Explicit
Public Const SolverName_Gurobi = "gurobi_cl.exe"
Public Const Solver_Gurobi = "gurobi" & ScriptExtension
Public Const SolverType_Gurobi = OpenSolver_SolverType.Linear

Public Const SolutionFile_Gurobi = "modelsolution.sol"
Public Const SensitivityFile_Gurobi = "sensitivityData.sol"

Function SolutionFilePath_Gurobi() As String
    SolutionFilePath_Gurobi = GetTempFilePath(SolutionFile_Gurobi)
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
    Dim SolverPath As String
    If Not SolverAvailable_Gurobi(SolverPath) Then
        About_Gurobi = "Gurobi not found"
        Exit Function
    End If
    
    SolverPath = left(SolverPath, InStrRev(SolverPath, PathDelimeter)) & Solver_Gurobi
    
    ' Assemble version info
    About_Gurobi = "Gurobi " & SolverBitness_Gurobi & "-bit" & _
                     " v" & SolverVersion_Gurobi & _
                     " at " & SolverPath
End Function

Function SolverAvailable_Gurobi(ByRef SolverPath As String) As Boolean
' Returns true if Gurobi is available and sets SolverPath as path to gurobi_cl
#If Mac Then
    SolverPath = "Macintosh HD:usr:local:bin:" & left(SolverName_Gurobi, Len(SolverName_Gurobi) - 4)
    If FileOrDirExists(SolverPath) Then
        SolverAvailable_Gurobi = True
    Else
        SolverPath = ""
        SolverAvailable_Gurobi = False
    End If
#Else
    If Environ("GUROBI_HOME") <> "" Then
        If GetExistingFilePathName(Environ("GUROBI_HOME"), "bin\" & SolverName_Gurobi, SolverPath) Then
            SolverAvailable_Gurobi = True
            Exit Function
        End If
    End If
    SolverPath = ""
    SolverAvailable_Gurobi = False
#End If
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
    RunPath = GetTempFolder & "gurobi_script"
    If FileOrDirExists(RunPath) Then Kill RunPath
    FileContents = """" & ConvertHfsPath(SolverPath) & """" & " -v" & " > """ & ConvertHfsPath(logFile) & """"
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
    RunPath = GetTempFolder & "gurobi_script"
    If FileOrDirExists(RunPath) Then Kill RunPath
    FileContents = """" & ConvertHfsPath(SolverPath) & """" & " -v" & " > """ & ConvertHfsPath(logFile) & """"
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


