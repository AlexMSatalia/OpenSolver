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
8358      ScriptFilePath_Couenne = GetTempFilePath(SolverScript_Couenne)
End Function

Function SolutionFilePath_Couenne() As String
8359      SolutionFilePath_Couenne = GetTempFilePath(SolutionFile_Couenne)
End Function

Sub CleanFiles_Couenne(errorPrefix As String)
          ' Solution file
8360      DeleteFileAndVerify SolutionFilePath_Couenne(), errorPrefix, "Unable to delete the Couenne solver solution file: " & SolutionFilePath_Couenne()
          ' Script file
8361      DeleteFileAndVerify ScriptFilePath_Couenne(), errorPrefix, "Unable to delete the Couenne solver script file: " & ScriptFilePath_Couenne()
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
                          " at " & MakeSpacesNonBreaking(SolverPath)
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
8374          system ("chmod +x " & ConvertHfsPath(SolverPath))
#End If
          
8375      End If
End Function

Function SolverVersion_Couenne() As String
      ' Get Couenne version by running 'couenne -v' at command line
          Dim SolverPath As String
8376      If Not SolverAvailable_Couenne(SolverPath) Then
8377          SolverVersion_Couenne = ""
8378          Exit Function
8379      End If
          
          ' Set up Couenne to write version info to text file
          Dim logFile As String
8380      logFile = GetTempFilePath("couenneversion.txt")
8381      If FileOrDirExists(logFile) Then Kill logFile
          
          Dim RunPath As String, FileContents As String
8382      RunPath = ScriptFilePath_Couenne()
8383      If FileOrDirExists(RunPath) Then Kill RunPath
8384      FileContents = QuotePath(ConvertHfsPath(SolverPath)) & " -v" & " > " & QuotePath(ConvertHfsPath(logFile))
8385      CreateScriptFile RunPath, FileContents
          
          ' Run Couenne
          Dim completed As Boolean
8386      completed = OSSolveSync(ConvertHfsPath(RunPath), "", "", "", SW_HIDE, True)
          
          ' Read version info back from output file
          Dim Line As String
8387      If FileOrDirExists(logFile) Then
8388          On Error GoTo ErrHandler
8389          Open logFile For Input As 1
8390          Line Input #1, Line
8391          Close #1
8392          SolverVersion_Couenne = right(Line, Len(Line) - 8)
8393          SolverVersion_Couenne = left(SolverVersion_Couenne, 5)
8394      Else
8395          SolverVersion_Couenne = ""
8396      End If
8397      Exit Function
          
ErrHandler:
8398      Close #1
8399      Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
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

Function CreateSolveScript_Couenne(ModelFilePathName As String) As String
          ' Create a script to run "/path/to/couenne.exe /path/to/<ModelFilePathName>"

          Dim SolverString As String, CommandLineRunString As String, PrintingOptionString As String
8410      SolverString = QuotePath(ConvertHfsPath(SolverFilePath_Couenne()))

8411      CommandLineRunString = QuotePath(ConvertHfsPath(ModelFilePathName))
8412      PrintingOptionString = ""
          
          Dim scriptFile As String, scriptFileContents As String
8413      scriptFile = ScriptFilePath_Couenne()
8414      scriptFileContents = SolverString & " " & CommandLineRunString & PrintingOptionString
8415      CreateScriptFile scriptFile, scriptFileContents
          
8416      CreateSolveScript_Couenne = scriptFile
End Function

Function ReadModel_Couenne(SolutionFilePathName As String, errorString As String, m As CModelParsed, s As COpenSolverParsed) As Boolean
8417      ReadModel_Couenne = False
          Dim Line As String, index As Long
8418      On Error GoTo readError
          Dim solutionExpected As Boolean
8419      solutionExpected = True
          
8420      If Not FileOrDirExists(SolutionFilePathName) Then
8421          solutionExpected = False
8422          If Not TryParseLogs(s) Then
8423              errorString = "The solver did not create a solution file. No new solution is available."
8424              GoTo exitFunction
8425          End If
8426      Else
8427          Open SolutionFilePathName For Input As 1 ' supply path with filename
8428          Line Input #1, Line ' Skip empty line at start of file
8429          Line Input #1, Line
8430          Line = Mid(Line, 10)
              
              'Get the returned status code from couenne.
8431          If Line Like "Optimal*" Then
8432              s.SolveStatus = OpenSolverResult.Optimal
8433              s.SolveStatusString = "Optimal"
8434          ElseIf Line Like "Infeasible*" Then
8435              s.SolveStatus = OpenSolverResult.Infeasible
8436              s.SolveStatusString = "No Feasible Solution"
8437              solutionExpected = False
8438          ElseIf Line Like "Unbounded*" Then
8439              s.SolveStatus = OpenSolverResult.Unbounded
8440              s.SolveStatusString = "No Solution Found (Unbounded)"
8441              solutionExpected = False
8442          Else
8443              If Not TryParseLogs(s) Then
8444                  errorString = "The response from the Couenne solver is not recognised. The response was: " & vbCrLf & _
                                    Line & vbCrLf & _
                                    "The Couenne command line can be found at:" & vbCrLf & _
                                    ScriptFilePath_Couenne()
8445                  GoTo exitFunction
8446              End If
8447              solutionExpected = False
8448          End If
8449      End If
          
8450      If solutionExpected Then
8451          Application.StatusBar = "OpenSolver: Loading Solution... " & s.SolveStatusString
              
8452          Line Input #1, Line ' Throw away blank line
8453          Line Input #1, Line ' Throw away "Options"
              
              Dim i As Long
8454          For i = 1 To 8
8455              Line Input #1, Line ' Skip all options lines
8456          Next i
              
              ' Note that the variable values are written to file in .nl format
              ' We need to read in the values and the extract the correct values for the adjustable cells
              
              ' Read in all variable values
              Dim VariableValues As New Collection
8457          While Not EOF(1)
8458              Line Input #1, Line
8459              VariableValues.Add CDbl(Line)
8460          Wend
              
              ' Loop through variable cells and find the corresponding value from VariableValues
8461          i = 1
              Dim c As Range, VariableIndex As Long
8462          For Each c In m.AdjustableCells
                  ' Extract the correct variable value
8463              VariableIndex = GetVariableNLIndex(i) + 1
                  
                  ' Need to make sure number is in US locale when Value2 is set
8464              Range(c.Address).Value2 = ConvertFromCurrentLocale(VariableValues(VariableIndex))
8465              i = i + 1
8466          Next c
8467      End If

8468      ReadModel_Couenne = True

exitFunction:
8469      Application.StatusBar = False
8470      Close #1
8471      Close #2
8472      Exit Function
          
readError:
8473      Application.StatusBar = False
8474      Close #1
8475      Close #2
8476      Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Function

Function TryParseLogs(s As COpenSolverParsed) As Boolean
      ' We examine the log file if it exists to try to find more info about the solve
          
          ' Check if log exists
          Dim logFile As String
8477      logFile = GetTempFilePath("log1.tmp")
          
8478      If Not FileOrDirExists(logFile) Then
8479          TryParseLogs = False
8480          Exit Function
8481      End If
          
          Dim message As String
8482      On Error GoTo ErrHandler
8483      Open logFile For Input As 3
8484      message = Input$(LOF(3), 3)
8485      Close #3
          
8486      If Not left(message, 7) = "Couenne" Then
             ' Not dealing with a Couenne log, abort
8487          TryParseLogs = False
8488          Exit Function
8489      End If
          
          ' Scan for information
          
          ' 1 - scan for infeasible
8490      If message Like "*infeasible*" Then
8491          s.SolveStatus = OpenSolverResult.Infeasible
8492          s.SolveStatusString = "No Feasible Solution"
8493          TryParseLogs = True
8494          Exit Function
8495      End If
          
ErrHandler:
8496      Close #3
8497      TryParseLogs = False
End Function
