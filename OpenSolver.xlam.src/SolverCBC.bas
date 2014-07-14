Attribute VB_Name = "SolverCBC"

Option Explicit
Public Const SOLVERNAME_CBC = "cbc.exe"

Function About_CBC() As String
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
    On Error GoTo NotFound
    GetExternalSolverPathName SolverPath, SOLVERNAME_CBC
    SolverAvailable_CBC = True
    Exit Function
NotFound:
    SolverPath = ""
    SolverAvailable_CBC = False
End Function

Function SolverVersion_CBC() As String
    Dim SolverPath As String
    If Not SolverAvailable_CBC(SolverPath) Then
        SolverVersion_CBC = ""
        Exit Function
    End If
    
    ' Set up cbc to write version info to text file
    Dim logCommand As String, logFile As String, RunPath As String, completed As Boolean
    logFile = GetTempFolder & "cbcversion.txt"
    If FileOrDirExists(logFile) Then Kill logFile
    logCommand = " > " & """" & logFile & """"
    
    RunPath = GetTempFolder & "cbc.bat"
    If FileOrDirExists(RunPath) Then Kill RunPath
    Open GetTempFolder & "cbc.bat" For Output As 1
    Print #1, "@echo off" & vbCrLf & """" & SolverPath & """" & " -exit" & logCommand
    Close #1
    
    ' Run cbc
    completed = OSSolveSync(RunPath, "", "", "", SW_HIDE, True)
    
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
