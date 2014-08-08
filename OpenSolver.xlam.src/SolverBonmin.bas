Attribute VB_Name = "SolverBonmin"
Option Explicit

Public Const SolverTitle_Bonmin = "COIN-OR Bonmin (Non-linear Solver)"
Public Const SolverDesc_Bonmin = "Bonmin (Basic Open-source Nonlinear Mixed INteger programming) is an experimental open-source C++ code for solving general MINLP (Mixed Integer NonLinear Programming). Finds globally optimal solutions to convex nonlinear problems in continuous and discrete variables, and may be applied heuristically to nonconvex problems."
Public Const SolverLink_Bonmin = "https://projects.coin-or.org/Bonmin"
Public Const SolverType_Bonmin = OpenSolver_SolverType.NonLinear

Public Const SolverName_Bonmin = "bonmin.exe"
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
                    " at " & SolverPath
End Function

Function SolverFilePath_Bonmin(Optional errorString As String) As String
#If Mac Then
    If GetExistingFilePathName(ThisWorkbook.Path, left(SolverName_Bonmin, Len(SolverName_Bonmin) - 4), SolverFilePath_Bonmin) Then Exit Function ' Found a mac solver
    errorString = "Unable to find Mac version of Bonmin ('bonmin') in the folder that contains 'OpenSolver.xlam'"
    Exit Function
#Else
    ' Look for the 64 bit version
    If SystemIs64Bit Then
        If GetExistingFilePathName(ThisWorkbook.Path, Replace(SolverName_Bonmin, ".exe", "64.exe"), SolverFilePath_Bonmin) Then Exit Function ' Found a 64 bit solver
    End If
    ' Look for the 32 bit version
    If GetExistingFilePathName(ThisWorkbook.Path, SolverName_Bonmin, SolverFilePath_Bonmin) Then
        If SystemIs64Bit Then
            errorString = "Unable to find 64-bit Bonmin (bonmin64.exe) in the 'OpenSolver.xlam' folder. 32-bit Bonmin will be used instead."
        End If
        Exit Function
    End If
    ' Fail
    SolverFilePath_Bonmin = ""
    errorString = "Unable to find 32-bit Bonmin (bonmin.exe) or 64-bit Bonmin (bonmin64.exe) in the 'OpenSolver.xlam' folder."
#End If
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
        Open logFile For Input As 1
        Line Input #1, Line
        Close #1
        SolverVersion_Bonmin = right(Line, Len(Line) - 7)
        SolverVersion_Bonmin = left(SolverVersion_Bonmin, 5)
    Else
        SolverVersion_Bonmin = ""
    End If
End Function

Function SolverBitness_Bonmin() As String
' Get Bitness of Bonmin solver
    Dim SolverPath As String
    If Not SolverAvailable_Bonmin(SolverPath) Then
        SolverBitness_Bonmin = ""
        Exit Function
    End If
    
    If right(SolverPath, 6) = "64.exe" Or right(SolverPath, 2) = "64" Then
        SolverBitness_Bonmin = "64"
    Else
        SolverBitness_Bonmin = "32"
    End If
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
    Dim Line As String, index As Integer
    On Error GoTo readError
    Dim solutionExpected As Boolean
    solutionExpected = True
    
    Open SolutionFilePathName For Input As 1 ' supply path with filename
    Line Input #1, Line ' Skip empty line at start of file
    Line Input #1, Line
    Line = Mid(Line, 9)
    
    'Get the returned status code from Bonmin.
    ' These are currently just using the outputs from CBC
    ' TODO Bonmin doesn't seem to write a solution file unless the solve is successful. We need to extract the solve status from the log file
    If Line Like "Optimal*" Then
        s.SolveStatus = OpenSolverResult.Optimal
        s.SolveStatusString = "Optimal"
    ElseIf Line Like "Infeasible*" Then
        s.SolveStatus = OpenSolverResult.Infeasible
        s.SolveStatusString = "No Feasible Solution"
    ElseIf Line Like "Integer infeasible*" Then
        s.SolveStatus = OpenSolverResult.Infeasible
        s.SolveStatusString = "No Feasible Integer Solution"
    ElseIf Line Like "*unbounded*" Then
        s.SolveStatus = OpenSolverResult.Unbounded
        s.SolveStatusString = "No Solution Found (Unbounded)"
        solutionExpected = False
    ElseIf Line Like "Stopped on time *" Then ' Stopped on iterations or time
        s.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
        s.SolveStatusString = "Stopped on Time Limit"
    ElseIf Line Like "Stopped on iterations*" Then ' Stopped on iterations or time
        s.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
        s.SolveStatusString = "Stopped on Iteration Limit"
    ElseIf Line Like "Stopped on difficulties*" Then ' Stopped on iterations or time
        s.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
        s.SolveStatusString = "Stopped on difficulties"
    ElseIf Line Like "Stopped on ctrl-c*" Then ' Stopped on iterations or time
        s.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
        s.SolveStatusString = "Stopped on Ctrl-C"
    ElseIf Line Like "Status unknown*" Then
        errorString = "Coueene did not solve the problem, suggesting there was an error in the input parameters. The response was: " & vbCrLf _
               & Line _
               & vbCrLf & "The Bonmin command line can be found at:" _
               & vbCrLf & ScriptFilePath_Bonmin()
        GoTo exitFunction
    Else
        errorString = "The response from the Bonmin solver is not recognised. The response was: " & Line
        GoTo exitFunction
    End If
    
    If solutionExpected Then
        Line Input #1, Line ' Throw away blank line
        Line Input #1, Line ' Throw away "Options"
        
        Dim i As Integer
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
        Dim c As Range, VariableIndex As Integer
        For Each c In m.AdjustableCells
            ' Extract the correct variable value
            VariableIndex = GetVariableNLIndex(i) + 1
            
            ' Need to make sure number is in US locale when Value2 is set
            Range(c.Address).Value2 = ConvertFromCurrentLocale(VariableValues(VariableIndex))
            i = i + 1
        Next c

        ReadModel_Bonmin = True
    End If

exitFunction:
    Close #1
    Close #2
    Exit Function
    
readError:
    Close #1
    Close #2
    Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Function


