Attribute VB_Name = "SolverNOMAD"
Public OS As COpenSolver
Dim IterationCount As Long

Public Const SolverTitle_NOMAD = "NOMAD (Non-linear solver)"
Public Const SolverDesc_NOMAD = "Nomad (Nonsmooth Optimization by Mesh Adaptive Direct search) is a C++ implementation of the Mesh Adaptive Direct Search (Mads) algorithm that solves non-linear problems. It works by updating the values on the sheet and passing them to the C++ solver. Like many non-linear solvers NOMAD cannot guarantee optimality of its solutions."
Public Const SolverLink_NOMAD = "http://www.gerad.ca/nomad/Project/Home.html"
Public Const SolverType_NOMAD = OpenSolver_SolverType.NonLinear

Public Const SolverName_NOMAD = "NOMAD"

Public Const UsesPrecision_NOMAD = True
Public Const UsesIterationLimit_NOMAD = True
Public Const UsesTolerance_NOMAD = False
Public Const UsesTimeLimit_NOMAD = True

Public Const NomadDllName = "OpenSolverNomad.dll"

' Don't forget we need to chdir to the directory containing the DLL before calling any of the functions
#If VBA7 Then
    Public Declare PtrSafe Function NomadMain Lib "OpenSolverNomad.dll" (ByVal SolveRelaxation As Boolean) As Long
    Private Declare PtrSafe Function NomadVersion Lib "OpenSolverNomad.dll" () As String
    Private Declare PtrSafe Function NomadDllVersion Lib "OpenSolverNomad.dll" Alias "NomadDLLVersion" () As String
#Else
    Public Declare Function NomadMain Lib "OpenSolverNomad.dll" (ByVal SolveRelaxation As Boolean) As Long
    Private Declare Function NomadVersion Lib "OpenSolverNomad.dll" () As String
    Private Declare Function NomadDllVersion Lib "OpenSolverNomad.dll" Alias "NomadDLLVersion" () As String
#End If


Function About_NOMAD() As String
          Dim errorString As String
6968      If Not SolverAvailable_NOMAD(errorString) Then
6969          About_NOMAD = errorString
6970          Exit Function
6971      End If
          ' Assemble version info
6972      About_NOMAD = "NOMAD " & SolverBitness_NOMAD & "-bit v" & SolverVersion_NOMAD() & _
                        " using OpenSolverNomadDLL v" & DllVersion_NOMAD() & " at " & MakeSpacesNonBreaking(DllPath_NOMAD())
End Function

Function NomadDir() As String
6973      NomadDir = JoinPaths(ThisWorkbook.Path, SolverDir)
#If Win64 Then
6974      NomadDir = JoinPaths(NomadDir, SolverDirWin64)
#Else
6975      NomadDir = JoinPaths(NomadDir, SolverDirWin32)
#End If
End Function

Function SolverAvailable_NOMAD(Optional errorString As String) As Boolean
#If Mac Then
6976      errorString = "NOMAD for OpenSolver is not currently supported on Mac"
6977      SolverAvailable_NOMAD = False
6978      Exit Function
#Else
          ' Set current dir for finding the DLL
          Dim currentDir As String
6979      currentDir = CurDir
6980      SetCurrentDirectory NomadDir()
          
          ' Try to access DLL - throws error if not found
6981      On Error GoTo NotFound
6982      NomadVersion
          
6983      SetCurrentDirectory currentDir
6984      SolverAvailable_NOMAD = True
6985      Exit Function

NotFound:
6986      SetCurrentDirectory currentDir
6987      SolverAvailable_NOMAD = False
6988      errorString = "Unable to find NOMAD (" & NomadDllName & ") in the `Solvers` folder (" & NomadDir() & ")"
6989      Exit Function
#End If
End Function

Function SolverVersion_NOMAD() As String
6990      If Not SolverAvailable_NOMAD() Then
6991          SolverVersion_NOMAD = ""
6992          Exit Function
6993      End If
          
          Dim currentDir As String, sNomadVersion As String
          
          ' Set current dir for finding the DLL
6994      currentDir = CurDir
6995      SetCurrentDirectory NomadDir()
          
          ' Get version info from DLL
6996      sNomadVersion = NomadVersion()
6997      sNomadVersion = left(sNomadVersion, InStr(sNomadVersion, vbNullChar) - 1)
          
6998      SetCurrentDirectory currentDir
          
6999      SolverVersion_NOMAD = sNomadVersion
End Function

Function DllVersion_NOMAD() As String
7000      If Not SolverAvailable_NOMAD() Then
7001          DllVersion_NOMAD = ""
7002          Exit Function
7003      End If
          
          Dim currentDir As String, sDllVersion As String
          
          ' Set current dir for finding the DLL
7004      currentDir = CurDir
7005      SetCurrentDirectory NomadDir()
          
          ' Get version info from DLL
7006      sDllVersion = NomadDllVersion()
7007      sDllVersion = left(sDllVersion, InStr(sDllVersion, vbNullChar) - 1)
          
7008      SetCurrentDirectory currentDir
          
7009      DllVersion_NOMAD = sDllVersion
End Function

Function DllPath_NOMAD() As String
7010      GetExistingFilePathName ThisWorkbook.Path, NomadDllName, DllPath_NOMAD
End Function

Function SolverBitness_NOMAD() As String
      ' Get Bitness of NOMAD solver
7011      If Not SolverAvailable_NOMAD() Then
7012          SolverBitness_NOMAD = ""
7013          Exit Function
7014      End If
          
#If Win64 Then
7015          SolverBitness_NOMAD = "64"
#Else
7016          SolverBitness_NOMAD = "32"
#End If
End Function

Function SolveModel_Nomad(SolveRelaxation As Boolean, s As COpenSolver) As Long
          Dim ScreenStatus As Boolean
7017      ScreenStatus = Application.ScreenUpdating
          Dim Show As String
7018      If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_sho", Show) Then
7019          If Show <> 1 Then Application.ScreenUpdating = False
7020      End If

          ' Trap Escape key
7021      Application.EnableCancelKey = xlErrorHandler
          
7022      On Error GoTo errorHandler
          Dim errorPrefix As String
7023      errorPrefix = "OpenSolver Nomad Model Solving"
7024      If s.ModelStatus <> ModelStatus_Built Then
7025          Err.Raise Number:=OpenSolver_NomadError, Source:=errorPrefix, Description:="The model cannot be solved as it has not yet been built."
7026      End If
          
          ' Loop through all decision vars and set their values
          ' This is to try and catch any protected cells as we can't catch VBA errors that occur while NOMAD calls back into VBA
          Dim c As Range
7027      For Each c In s.AdjustableCells
7028          c.Value2 = c.Value2
7029      Next c
          
          ' Set OS for the calls back into Excel from NOMAD
7030      Set OS = s
          
          Dim oldCalculationMode As Long
7031      oldCalculationMode = Application.Calculation
7032      Application.Calculation = xlCalculationManual
          
          Dim currentDir As String
7033      currentDir = CurDir
          
7034      SetCurrentDirectory NomadDir()

7035      IterationCount = 0
          
          ' We need to call NomadMain directly rather than use Application.Run .
          ' Using Application.Run causes the API calls inside the DLL to fail on 64 bit Office
          Dim NomadRetVal As Long
7036      NomadRetVal = NomadMain(SolveRelaxation)
          
          'Catch any errors that occured while Nomad was solving
7037      If NomadRetVal = 1 Then
              Dim errorString As String
7038          errorString = "There was an error while Nomad was solving. No solution has been loaded into the sheet."
              ' Check logs for more info
7039          CheckNomadLogs errorString
              
7040          Err.Raise Number:=OpenSolver_NomadError, Source:=errorPrefix, Description:=errorString
7041          s.SolveStatus = OpenSolverResult.ErrorOccurred
7042      ElseIf NomadRetVal = 2 Then
7043          s.SolveStatusComment = "Nomad reached the maximum number of iterations and returned the best feasible solution it found. This solution is not guaranteed to be an optimal solution." & vbCrLf & vbCrLf & _
                                     "You can increase the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
7044          s.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
7045          s.SolveStatusString = "Stopped on Iteration Limit"
7046          s.SolutionWasLoaded = True
7047      ElseIf NomadRetVal = 3 Then
7048          s.SolveStatusComment = "Nomad reached the maximum time and returned the best feasible solution it found. This solution is not guaranteed to be an optimal solution." & vbCrLf & vbCrLf & _
                                     "You can increase the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
7049          s.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
7050          s.SolveStatusString = "Stopped on Time Limit"
7051          s.SolutionWasLoaded = True
7052      ElseIf NomadRetVal = 4 Then
7053          s.SolveStatusComment = "Nomad reached the maximum time or number of iterations without finding a feasible solution. The best infeasible solution has been returned to the sheet." & vbCrLf & vbCrLf & _
                                     "You can increase the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
7054          s.SolveStatus = OpenSolverResult.Infeasible
7055          s.SolveStatusString = "No Feasible Solution"
7056          s.SolutionWasLoaded = True
7057      ElseIf NomadRetVal = 10 Then
7058          s.SolveStatusComment = "Nomad could not find a feasible solution. The best infeasible solution has been returned to the sheet." & vbCrLf & vbCrLf & _
                                     "Try resolving at a different start point or check whether your model is feasible or relax some of your constraints."
7059          s.SolveStatus = OpenSolverResult.Infeasible
7060          s.SolveStatusString = "No Feasible Solution"
7061          s.SolutionWasLoaded = True
7062      ElseIf NomadRetVal = -3 Then
7063          Err.Raise OpenSolver_UserCancelledError, "Running NOMAD", "Model solve cancelled by user."
7064      Else
7065          s.SolveStatus = NomadRetVal 'optimal
7066          s.SolveStatusString = "Optimal"
7067      End If
          
ExitSub:
          ' We can fall thru to here
7068      SetCurrentDirectory currentDir
7069      Application.Cursor = xlDefault
7070      Application.StatusBar = False ' Resume normal status bar behaviour
7071      Application.ScreenUpdating = True
7072      Application.Calculation = oldCalculationMode
7073      Application.Calculate
7074      Application.ScreenUpdating = ScreenStatus
7075      Close #1 ' Close any open file; this does not seem to ever give errors
7076      SolveModel_Nomad = s.SolveStatus    ' Return the main result
7077      Set OS = Nothing
7078      Exit Function
          
errorHandler:
          ' We only trap Escape (Err.Number=18) here; all other errors are passed back to the caller.
          ' Save error message
          Dim ErrorNumber As Long, ErrorDescription As String, ErrorSource As String
7079      ErrorNumber = Err.Number
7080      ErrorDescription = Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
7081      ErrorSource = Err.Source
7082      If Err.Number = 18 Then
7083          If MsgBox("You have pressed the Escape key. Do you wish to cancel?", _
                         vbCritical + vbYesNo + vbDefaultButton1, _
                         "OpenSolver: User Interrupt Occured...") = vbNo Then
7084              Resume 'continue on from where error occured
7085          Else
                  ' Raise a "user cancelled" error. We cannot use Raise, as that exits immediately without going thru our code below
7086              ErrorNumber = OpenSolver_UserCancelledError
7087              ErrorSource = errorPrefix
7088              ErrorDescription = "Model solve cancelled by user."
7089          End If
7090      End If

ErrorExit:
          ' Exit, raising an error. None of the following actions change the Err.Number etc, but we saved them above just in case...
7091      SetCurrentDirectory currentDir
          'Application.DefaultFilePath = currentExcelDir
7092      Application.Cursor = xlDefault
7093      Application.StatusBar = False ' Resume normal status bar behaviour
7094      Application.ScreenUpdating = True
7095      Application.Calculation = oldCalculationMode
7096      Application.Calculate
7097      Close #1 ' Close any open file; this does not seem to ever give errors
7098      Set OS = Nothing
7099      Err.Raise ErrorNumber, ErrorSource, ErrorDescription

End Function

Sub CheckNomadLogs(errorString As String)
      ' If NOMAD encounters an error, we dump the exception to the log file. We can use this to deduce what went wrong
          Dim logFile As String
7100      logFile = GetTempFilePath("log1.tmp")
          
7101      If Not FileOrDirExists(logFile) Then
7102          Exit Sub
7103      End If
          
          Dim message As String
7104      On Error GoTo ErrHandler
7105      Open logFile For Input As 3
7106      message = Input$(LOF(3), 3)
7107      Close #3
          
7108      If Not message Like "*NOMAD*" Then
7109         Exit Sub
7110      End If
          
7111      If message Like "*invalid parameter*" Then
7112          errorString = "One of the parameters supplied to Nomad was invalid. This usually happens if the precision is too large. Try adjusting the values in the Solve Options dialog box."
7113      End If
          
ErrHandler:
7114      Close #3
End Sub

Function updateVar(X As Variant, Optional BestSolution As Variant = Nothing, Optional Infeasible As Boolean = False)
7115      IterationCount = IterationCount + 1

          ' Update solution
7116      If IterationCount Mod 5 = 1 Then
              Dim status As String
7117          status = "OpenSolver: Running NOMAD. Iteration " & IterationCount & "."
              ' Check for BestSolution = Nothing
7118          If Not VarType(BestSolution) = 9 Then
                  ' Flip solution if maximisation
7119              If OS.ObjectiveSense = MaximiseObjective Then BestSolution = -BestSolution

7120              status = status & " Best solution so far: " & BestSolution
7121              If Infeasible Then
7122                  status = status & " (infeasible)"
7123              End If
7124          End If
7125          Application.StatusBar = status
7126      End If
          
7127      OS.updateVarOS (X)
End Function

Function getValues() As Variant
7128      getValues = OS.getValuesOS()
End Function

Sub RecalculateValues()
7129      Sheets(ActiveSheet.Name).Calculate
End Sub

Function getNumVariables() As Variant
7130      getNumVariables = OS.getNumVariablesOS
End Function

Function getNumConstraints() As Variant
7131      getNumConstraints = OS.getNumConstraintsOS
End Function

Function getVariableData() As Variant
7132      getVariableData = OS.getVariableDataOS()
End Function

Function getOptionData() As Variant
7133      getOptionData = OS.getOptionDataOS()
End Function

