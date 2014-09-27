Attribute VB_Name = "SolverBonmin"
Option Explicit

Public Const SolverTitle_Bonmin = "COIN-OR Bonmin (Non-linear solver)"
Public Const SolverDesc_Bonmin = "Bonmin (Basic Open-source Nonlinear Mixed INteger programming) is an experimental open-source C++ code for solving general MINLPs (Mixed Integer NonLinear Programming). Finds globally optimal solutions to convex nonlinear problems in continuous and discrete variables, and may be applied heuristically to nonconvex problems. Bonmin uses the COIN-OR solvers CBC and IPOPT while solving. For more info on these, see www.coin-or.org/projects"
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
    ScriptFilePath_Bonmin = GetTempFilePath(SolverScript_Bonmin)
End Function

Function SolutionFilePath_Bonmin() As String
    SolutionFilePath_Bonmin = GetTempFilePath(SolutionFile_Bonmin)
End Function

Sub CleanFiles_Bonmin(errorPrefix As String)
    ' Solution file
    DeleteFileAndVerify SolutionFilePath_Bonmin(), errorPrefix, "Unable to delete the Bonmin solver solution file: " & SolutionFilePath_Bonmin()
    ' Script file
    DeleteFileAndVerify ScriptFilePath_Bonmin(), errorPrefix, "Unable to delete the Bonmin solver script file: " & ScriptFilePath_Bonmin()
End Sub

Function About_Bonmin() As String
' Return string for "About" form
    Dim SolverPath As String, errorString As String
    If Not SolverAvailable_Bonmin(SolverPath, errorString) Then
        About_Bonmin = errorString
        Exit Function
    End If
    ' Assemble version info
    About_Bonmin = "Bonmin " & SolverBitness_Bonmin & "-bit" & _
                    " v" & SolverVersion_Bonmin & _
                    " at " & MakeSpacesNonBreaking(SolverPath)
End Function

Function SolverFilePath_Bonmin(Optional errorString As String) As String
    SolverFilePath_Bonmin = SolverFilePath_Default("Bonmin", errorString)
End Function

Function SolverAvailable_Bonmin(Optional SolverPath As String, Optional errorString As String) As Boolean
' Returns true if Bonmin is available and sets SolverPath
    SolverPath = SolverFilePath_Bonmin(errorString)
    If SolverPath = "" Then
        SolverAvailable_Bonmin = False
    Else
        SolverAvailable_Bonmin = True

#If Mac Then
        ' Make sure Bonmin is executable on Mac
        system ("chmod +x " & ConvertHfsPath(SolverPath))
#End If
    
    End If
End Function

Function SolverVersion_Bonmin() As String
' Get Bonmin version by running 'bonmin -v' at command line
    Dim SolverPath As String
    If Not SolverAvailable_Bonmin(SolverPath) Then
        SolverVersion_Bonmin = ""
        Exit Function
    End If
    
    ' Set up Bonmin to write version info to text file
    Dim logFile As String
    logFile = GetTempFolder & "bonminversion.txt"
    If FileOrDirExists(logFile) Then Kill logFile
    
    Dim RunPath As String, FileContents As String
    RunPath = ScriptFilePath_Bonmin()
    If FileOrDirExists(RunPath) Then Kill RunPath
    FileContents = """" & ConvertHfsPath(SolverPath) & """" & " -v" & " > """ & ConvertHfsPath(logFile) & """"
    CreateScriptFile RunPath, FileContents
    
    ' Run Bonmin
    Dim completed As Boolean
    completed = OSSolveSync(ConvertHfsPath(RunPath), "", "", "", SW_HIDE, True)
    
    ' Read version info back from output file
    Dim Line As String
    If FileOrDirExists(logFile) Then
        On Error GoTo ErrHandler
        Open logFile For Input As 1
        Line Input #1, Line
        Close #1
        SolverVersion_Bonmin = right(Line, Len(Line) - 7)
        SolverVersion_Bonmin = left(SolverVersion_Bonmin, 5)
    Else
        SolverVersion_Bonmin = ""
    End If
    Exit Function
    
ErrHandler:
    Close #1
    Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Function

Function SolverBitness_Bonmin() As String
' Get Bitness of Bonmin solver
    Dim SolverPath As String
    If Not SolverAvailable_Bonmin(SolverPath) Then
        SolverBitness_Bonmin = ""
        Exit Function
    End If
    
    ' All Macs are 64-bit so we only provide 64-bit binaries
#If Mac Then
    SolverBitness_Bonmin = "64"
#Else
    If right(SolverPath, 13) = "64\bonmin.exe" Then
        SolverBitness_Bonmin = "64"
    Else
        SolverBitness_Bonmin = "32"
    End If
#End If
End Function

Function CreateSolveScript_Bonmin(ModelFilePathName As String) As String
    ' Create a script to run "/path/to/bonmin.exe /path/to/<ModelFilePathName>"

    Dim SolverString As String, CommandLineRunString As String, PrintingOptionString As String
    SolverString = QuotePath(ConvertHfsPath(SolverFilePath_Bonmin()))

    CommandLineRunString = QuotePath(ConvertHfsPath(ModelFilePathName))
    PrintingOptionString = ""
    
    Dim scriptFile As String, scriptFileContents As String
    scriptFile = ScriptFilePath_Bonmin()
    scriptFileContents = SolverString & " " & CommandLineRunString & PrintingOptionString
    CreateScriptFile scriptFile, scriptFileContents
    
    CreateSolveScript_Bonmin = scriptFile
End Function

Function ReadModel_Bonmin(SolutionFilePathName As String, errorString As String, m As CModelParsed, s As COpenSolverParsed) As Boolean
    ReadModel_Bonmin = False
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
        Line = Mid(Line, 9)
        
        'Get the returned status code from Bonmin.
        If Line Like "Optimal*" Then
            s.SolveStatus = OpenSolverResult.Optimal
            s.SolveStatusString = "Optimal"
        ElseIf Line Like "Infeasible*" Then
            s.SolveStatus = OpenSolverResult.Infeasible
            s.SolveStatusString = "No Feasible Solution"
            solutionExpected = False
        ElseIf Line Like "*unbounded*" Then
            s.SolveStatus = OpenSolverResult.Unbounded
            s.SolveStatusString = "No Solution Found (Unbounded)"
            solutionExpected = False
        ElseIf Line Like "Error encountered in optimization*" Then
            ' Try to get status from logs
            If Not TryParseLogs(s) Then
                errorString = "Bonmin did not solve the problem, suggesting there was an error in the input parameters. The response was: " & vbCrLf & _
                              Line & vbCrLf & _
                              "The Bonmin command line can be found at:" & vbCrLf & _
                              ScriptFilePath_Bonmin()
                GoTo exitFunction
            End If
            solutionExpected = False
        Else
            errorString = "The response from the Bonmin solver is not recognised. The response was: " & _
                          Line & vbCrLf & _
                          "The Bonmin command line can be found at:" & vbCrLf & _
                          ScriptFilePath_Bonmin()
            GoTo exitFunction
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
    ReadModel_Bonmin = True

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
    
    If Not left(message, 6) = "Bonmin" Then
       ' Not dealing with a Bonmin log, abort
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
