Attribute VB_Name = "SolverGurobi"

Option Explicit
Public Const SOLVERNAME_GUROBI = "gurobi_cl.exe"
Public Const SOLVER_GUROBI = "gurobi" & ScriptExtension

Function About_Gurobi() As String
' Return string for "About" form
    Dim SolverPath As String
    If Not SolverAvailable_Gurobi(SolverPath) Then
        About_Gurobi = "Gurobi not found"
        Exit Function
    End If
    
    SolverPath = left(SolverPath, InStrRev(SolverPath, PathDelimeter)) & SOLVER_GUROBI
    
    ' Assemble version info
    About_Gurobi = "Gurobi " & SolverBitness_Gurobi & "-bit" & _
                     " v" & SolverVersion_Gurobi & _
                     " at " & SolverPath
End Function

Function SolverAvailable_Gurobi(ByRef SolverPath As String) As Boolean
' Returns true if Gurobi is available and sets SolverPath as path to gurobi_cl
#If Mac Then
    SolverPath = "Macintosh HD:usr:local:bin:" & left(SOLVERNAME_GUROBI, Len(SOLVERNAME_GUROBI) - 4)
    If FileOrDirExists(SolverPath) Then
        SolverAvailable_Gurobi = True
    Else
        SolverPath = ""
        SolverAvailable_Gurobi = False
    End If
#Else
    If Environ("GUROBI_HOME") <> "" Then
        If GetExistingFilePathName(Environ("GUROBI_HOME"), "bin\" & SOLVERNAME_GUROBI, SolverPath) Then
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


