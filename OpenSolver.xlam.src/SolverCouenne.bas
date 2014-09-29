Attribute VB_Name = "SolverCouenne"
Option Explicit

Public Const SolverTitle_Couenne = "COIN-OR Couenne (Non-linear, non-convex solver)"
Public Const SolverDesc_Couenne = "Couenne (Convex Over and Under ENvelopes for Nonlinear Estimation) is a branch & bound algorithm to solve Mixed-Integer Nonlinear Programming (MINLP) problems of specific forms. Couenne aims at finding global optima of nonconvex MINLPs. It implements linearization, bound reduction, and branching methods within a branch-and-bound framework."
Public Const SolverLink_Couenne = "https://projects.coin-or.org/Couenne"
Public Const SolverType_Couenne = OpenSolver_SolverType.NonLinear

#If Mac Then
Public Const SolverName_Couenne = "couenne"
#Else
Public Const SolverName_Couenne = "couenne.exe"
#End If

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
                    " at " & MakeSpacesNonBreaking(SolverPath)
End Function

Function SolverFilePath_Couenne(Optional errorString As String) As String
    SolverFilePath_Couenne = SolverFilePath_Default("Couenne", errorString)
End Function

Function SolverAvailable_Couenne(Optional SolverPath As String, Optional errorString As String) As Boolean
' Returns true if Couenne is available and sets SolverPath
    SolverPath = SolverFilePath_Couenne(errorString)
    If SolverPath = "" Then
        SolverAvailable_Couenne = False
    Else
        SolverAvailable_Couenne = True
        errorString = "WARNING: Couenne is EXPERIMENTAL and is not guaranteed to give optimal or even good solutions. Proceed with caution." & vbCrLf & vbCrLf & errorString

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
    logFile = GetTempFilePath("couenneversion.txt")
    If FileOrDirExists(logFile) Then Kill logFile
    
    Dim RunPath As String, FileContents As String
    RunPath = ScriptFilePath_Couenne()
    If FileOrDirExists(RunPath) Then Kill RunPath
    FileContents = QuotePath(ConvertHfsPath(SolverPath)) & " -v" & " > " & QuotePath(ConvertHfsPath(logFile))
    CreateScriptFile RunPath, FileContents
    
    ' Run Couenne
    Dim completed As Boolean
    completed = OSSolveSync(ConvertHfsPath(RunPath), "", "", "", SW_HIDE, True)
    
    ' Read version info back from output file
    Dim Line As String
    If FileOrDirExists(logFile) Then
        On Error GoTo ErrHandler
        Open logFile For Input As 1
        Line Input #1, Line
        Close #1
        SolverVersion_Couenne = right(Line, Len(Line) - 8)
        SolverVersion_Couenne = left(SolverVersion_Couenne, 5)
    Else
        SolverVersion_Couenne = ""
    End If
    Exit Function
    
ErrHandler:
    Close #1
    Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Function

Function SolverBitness_Couenne() As String
' Get Bitness of Couenne solver
    Dim SolverPath As String
    If Not SolverAvailable_Couenne(SolverPath) Then
        SolverBitness_Couenne = ""
        Exit Function
    End If
        
    ' All Macs are 64-bit so we only provide 64-bit binaries
#If Mac Then
    SolverBitness_Couenne = "64"
#Else
    If right(SolverPath, 14) = "64\couenne.exe" Then
        SolverBitness_Couenne = "64"
    Else
        SolverBitness_Couenne = "32"
    End If
#End If
End Function

Function CreateSolveScript_Couenne(ModelFilePathName As String) As String
    ' Create a script to run "/path/to/couenne.exe /path/to/<ModelFilePathName>"

    Dim SolverString As String, CommandLineRunString As String, PrintingOptionString As String
    SolverString = QuotePath(ConvertHfsPath(SolverFilePath_Couenne()))

    CommandLineRunString = QuotePath(ConvertHfsPath(ModelFilePathName))
    PrintingOptionString = ""
    
    Dim scriptFile As String, scriptFileContents As String
    scriptFile = ScriptFilePath_Couenne()
    scriptFileContents = SolverString & " " & CommandLineRunString & PrintingOptionString
    CreateScriptFile scriptFile, scriptFileContents
    
    CreateSolveScript_Couenne = scriptFile
End Function

Function ReadModel_Couenne(SolutionFilePathName As String, errorString As String, m As CModelParsed, s As COpenSolverParsed) As Boolean
    ReadModel_Couenne = False
    Dim Line As String, index As Long
    On Error GoTo readError
    Dim solutionExpected As Boolean
    solutionExpected = True
    
    If Not FileOrDirExists(SolutionFilePathName) Then
        solutionExpected = False
        If Not TryParseLogs(s) Then
            errorString = "The solver did not create a solution file. No new solution is available."
            GoTo exitFunction
        End If
    Else
        Open SolutionFilePathName For Input As 1 ' supply path with filename
        Line Input #1, Line ' Skip empty line at start of file
        Line Input #1, Line
        Line = Mid(Line, 10)
        
        'Get the returned status code from couenne.
        If Line Like "Optimal*" Then
            s.SolveStatus = OpenSolverResult.Optimal
            s.SolveStatusString = "Optimal"
        ElseIf Line Like "Infeasible*" Then
            s.SolveStatus = OpenSolverResult.Infeasible
            s.SolveStatusString = "No Feasible Solution"
            solutionExpected = False
        ElseIf Line Like "Unbounded*" Then
            s.SolveStatus = OpenSolverResult.Unbounded
            s.SolveStatusString = "No Solution Found (Unbounded)"
            solutionExpected = False
        Else
            If Not TryParseLogs(s) Then
                errorString = "The response from the Couenne solver is not recognised. The response was: " & vbCrLf & _
                              Line & vbCrLf & _
                              "The Couenne command line can be found at:" & vbCrLf & _
                              ScriptFilePath_Couenne()
                GoTo exitFunction
            End If
            solutionExpected = False
        End If
    End If
    
    If solutionExpected Then
        Application.StatusBar = "OpenSolver: Loading Solution... " & s.SolveStatusString
        
        Line Input #1, Line ' Throw away blank line
        Line Input #1, Line ' Throw away "Options"
        
        Dim i As Long
        For i = 1 To 8
            Line Input #1, Line ' Skip all options lines
        Next i
        
        ' Note that the variable values are written to file in .nl format
        ' We need to read in the values and the extract the correct values for the adjustable cells
        
        ' Read in all variable values
        Dim VariableValues As New Collection
        While Not EOF(1)
            Line Input #1, Line
            VariableValues.Add CDbl(Line)
        Wend
        
        ' Loop through variable cells and find the corresponding value from VariableValues
        i = 1
        Dim c As Range, VariableIndex As Long
        For Each c In m.AdjustableCells
            ' Extract the correct variable value
            VariableIndex = GetVariableNLIndex(i) + 1
            
            ' Need to make sure number is in US locale when Value2 is set
            Range(c.Address).Value2 = ConvertFromCurrentLocale(VariableValues(VariableIndex))
            i = i + 1
        Next c
    End If

    ReadModel_Couenne = True

exitFunction:
    Application.StatusBar = False
    Close #1
    Close #2
    Exit Function
    
readError:
    Application.StatusBar = False
    Close #1
    Close #2
    Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Function

Function TryParseLogs(s As COpenSolverParsed) As Boolean
' We examine the log file if it exists to try to find more info about the solve
    
    ' Check if log exists
    Dim logFile As String
    logFile = GetTempFilePath("log1.tmp")
    
    If Not FileOrDirExists(logFile) Then
        TryParseLogs = False
        Exit Function
    End If
    
    Dim message As String
    On Error GoTo ErrHandler
    Open logFile For Input As 3
    message = Input$(LOF(3), 3)
    Close #3
    
    If Not left(message, 7) = "Couenne" Then
       ' Not dealing with a Couenne log, abort
        TryParseLogs = False
        Exit Function
    End If
    
    ' Scan for information
    
    ' 1 - scan for infeasible
    If message Like "*infeasible*" Then
        s.SolveStatus = OpenSolverResult.Infeasible
        s.SolveStatusString = "No Feasible Solution"
        TryParseLogs = True
        Exit Function
    End If
    
ErrHandler:
    Close #3
    TryParseLogs = False
End Function
