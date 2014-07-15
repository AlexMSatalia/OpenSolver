Attribute VB_Name = "SolverCBC"

Option Explicit
Public Const SolverName_CBC = "cbc.exe"
Public Const SolverScript_CBC = "cbc" & ScriptExtension
Public Const SolverType_CBC = OpenSolver_SolverType.Linear

Public Const SolutionFile_CBC = "modelsolution.txt"
Public Const CostRangesFile_CBC = "costranges.txt"
Public Const RHSRangesFile_CBC = "rhsranges.txt"

Function ScriptFilePath_CBC() As String
    ScriptFilePath_CBC = GetTempFilePath(SolverScript_CBC)
End Function

Function SolutionFilePath_CBC() As String
    SolutionFilePath_CBC = GetTempFilePath(SolutionFile_CBC)
End Function

Function CostRangesFilePath_CBC() As String
    CostRangesFilePath_CBC = GetTempFilePath(CostRangesFile_CBC)
End Function

Function RHSRangesFilePath_CBC() As String
    RHSRangesFilePath_CBC = GetTempFilePath(RHSRangesFile_CBC)
End Function

Sub CleanFiles_CBC(errorPrefix As String)
    ' Solution file
    DeleteFileAndVerify SolutionFilePath_CBC(), errorPrefix, "Unable to delete the CBC solver solution file: " & SolutionFilePath_CBC()
    ' Cost Range file
    DeleteFileAndVerify CostRangesFilePath_CBC(), errorPrefix, "Unable to delete the CBC solver sensitivity data file: " & CostRangesFilePath_CBC()
    ' RHS Range file
    DeleteFileAndVerify RHSRangesFilePath_CBC(), errorPrefix, "Unable to delete the CBC solver sensitivity data file: " & RHSRangesFilePath_CBC()
    ' Script file
    DeleteFileAndVerify ScriptFilePath_CBC(), errorPrefix, "Unable to delete the CBC solver script file: " & ScriptFilePath_CBC()
End Sub

Function About_CBC() As String
' Return string for "About" form
    Dim SolverPath As String
    If Not SolverAvailable_CBC(SolverPath) Then
        About_CBC = "CBC not found"
        Exit Function
    End If
    ' Assemble version info
    About_CBC = "CBC " & SolverBitness_CBC & "-bit" & _
                     " v" & SolverVersion_CBC & _
                     " at " & SolverPath
End Function

Function SolverAvailable_CBC(ByRef SolverPath As String) As Boolean
' Returns true if CBC is available and sets SolverPath
    On Error GoTo NotFound
    GetExternalSolverPathName SolverPath, SolverName_CBC
    SolverAvailable_CBC = True
    Exit Function
NotFound:
    SolverPath = ""
    SolverAvailable_CBC = False
End Function

Function SolverVersion_CBC() As String
' Get CBC version by running 'cbc -exit' at command line
    Dim SolverPath As String
    If Not SolverAvailable_CBC(SolverPath) Then
        SolverVersion_CBC = ""
        Exit Function
    End If
    
    ' Set up cbc to write version info to text file
    Dim logFile As String
    logFile = GetTempFolder & "cbcversion.txt"
    If FileOrDirExists(logFile) Then Kill logFile
    
    Dim RunPath As String, FileContents As String
    RunPath = GetTempFolder & "cbc"
    If FileOrDirExists(RunPath) Then Kill RunPath
    FileContents = """" & ConvertHfsPath(SolverPath) & """" & " -exit" & " > """ & ConvertHfsPath(logFile) & """"
    CreateScriptFile RunPath, FileContents
    
    ' Run cbc
    Dim completed As Boolean
    completed = OSSolveSync(ConvertHfsPath(RunPath), "", "", "", SW_HIDE, True)
    
    ' Read version info back from output file
    Dim Line As String
    If FileOrDirExists(logFile) Then
        Open logFile For Input As 1
        Line Input #1, Line
        Line Input #1, Line
        Close #1
        SolverVersion_CBC = right(Line, Len(Line) - 9)
        SolverVersion_CBC = left(SolverVersion_CBC, Len(SolverVersion_CBC) - 1)
    Else
        SolverVersion_CBC = ""
    End If
End Function

Function SolverBitness_CBC() As String
' Get Bitness of CBC solver
    Dim SolverPath As String
    If Not SolverAvailable_CBC(SolverPath) Then
        SolverBitness_CBC = ""
        Exit Function
    End If
    
    If right(SolverPath, 6) = "64.exe" Or right(SolverPath, 2) = "64" Then
        SolverBitness_CBC = "64"
    Else
        SolverBitness_CBC = "32"
    End If
End Function
