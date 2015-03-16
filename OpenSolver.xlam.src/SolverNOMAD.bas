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

'NOMAD return status codes
Public Enum NomadResult
    UserCancelled = -3
    Optimal = 0
    ErrorOccured = 1
    SolveStoppedIter = 2
    SolveStoppedTime = 3
    Infeasible = 4
    SolveStoppedNoSolution = 10
End Enum

Function About_NOMAD() As String
          Dim errorString As String
6968      If Not SolverAvailable_NOMAD(errorString) Then
6969          About_NOMAD = errorString
6970          Exit Function
6971      End If
          ' Assemble version info
6972      About_NOMAD = "NOMAD " & SolverBitness_NOMAD & "-bit v" & SolverVersion_NOMAD() & _
                        " using OpenSolverNomadDLL v" & DllVersion_NOMAD() & _
                        " at " & MakeSpacesNonBreaking(MakePathSafe(DllPath_NOMAD()))
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
          ' Save to a new string first - modifying the string from the DLL can sometimes crash Excel
          sNomadVersion = NomadVersion()
6996      sNomadVersion = left(Replace(sNomadVersion, vbNullChar, ""), 5)
          
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
          ' Save to a new string first - modifying the string from the DLL can sometimes crash Excel
          sDllVersion = NomadDllVersion()
7006      sDllVersion = left(Replace(sDllVersion, vbNullChar, ""), 5)
          
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
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim ScreenStatus As Boolean
7017      ScreenStatus = Application.ScreenUpdating

          If GetShowSolverProgress() Then Application.ScreenUpdating = False

7024      If s.ModelStatus <> ModelStatus_Built Then
7025          Err.Raise Number:=OpenSolver_NomadError, Description:="The model cannot be solved as it has not yet been built."
7026      End If
          
          ' Set OS for the calls back into Excel from NOMAD
7030      Set OS = s

          ' Check precision is not 0
          Dim SolveOptions As SolveOptionsType
          GetSolveOptions OS.sheet, SolveOptions
          
          If SolveOptions.Precision <= 0 Then
              Err.Raise Number:=OpenSolver_NomadError, Description:="The current level of precision (" & CStr(SolveOptions.Precision) & ") is invalid. Please set the precision to a small positive (non-zero) value and try again."
          End If
          
          Dim oldCalculationMode As Long
7031      oldCalculationMode = Application.Calculation
7032      Application.Calculation = xlCalculationManual

          ' Loop through all decision vars and set their values
          ' This is to try and catch any protected cells as we can't catch VBA errors that occur while NOMAD calls back into VBA
          ' Do this after setting calculation mode to manual!
          Dim c As Range
          For Each c In s.AdjustableCells
              c.Value2 = c.Value2
          Next c
          
          Dim currentDir As String
7033      currentDir = CurDir
          
7034      SetCurrentDirectory NomadDir()

7035      IterationCount = 0
          
          ' We need to call NomadMain directly rather than use Application.Run .
          ' Using Application.Run causes the API calls inside the DLL to fail on 64 bit Office
          Dim NomadRetVal As Long
7036      NomadRetVal = NomadMain(SolveRelaxation)
          
          'Catch any errors that occured while Nomad was solving
7037      Select Case NomadRetVal
          Case NomadResult.ErrorOccured
7041          s.SolveStatus = OpenSolverResult.ErrorOccurred

              ' Check logs for more info
7039          CheckNomadLogs
              
7040          Err.Raise Number:=OpenSolver_NomadError, Description:="There was an error while Nomad was solving. No solution has been loaded into the sheet."
7042      Case NomadResult.SolveStoppedIter
7043          s.SolveStatusComment = "Nomad reached the maximum number of iterations and returned the best feasible solution it found. This solution is not guaranteed to be an optimal solution." & vbCrLf & vbCrLf & _
                                     "You can increase the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
7044          s.SolveStatus = OpenSolverResult.LimitedSubOptimal
7045          s.SolveStatusString = "Stopped on Iteration Limit"
7046          s.SolutionWasLoaded = True
7047      Case NomadResult.SolveStoppedTime
7048          s.SolveStatusComment = "Nomad reached the maximum time and returned the best feasible solution it found. This solution is not guaranteed to be an optimal solution." & vbCrLf & vbCrLf & _
                                     "You can increase the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
7049          s.SolveStatus = OpenSolverResult.LimitedSubOptimal
7050          s.SolveStatusString = "Stopped on Time Limit"
7051          s.SolutionWasLoaded = True
7052      Case NomadResult.Infeasible
7053          s.SolveStatusComment = "Nomad reached the maximum time or number of iterations without finding a feasible solution. The best infeasible solution has been returned to the sheet." & vbCrLf & vbCrLf & _
                                     "You can increase the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
7054          s.SolveStatus = OpenSolverResult.Infeasible
7055          s.SolveStatusString = "No Feasible Solution"
7056          s.SolutionWasLoaded = True
7057      Case NomadResult.SolveStoppedNoSolution
7058          s.SolveStatusComment = "Nomad could not find a feasible solution. The best infeasible solution has been returned to the sheet." & vbCrLf & vbCrLf & _
                                     "Try resolving at a different start point or check whether your model is feasible or relax some of your constraints."
7059          s.SolveStatus = OpenSolverResult.Infeasible
7060          s.SolveStatusString = "No Feasible Solution"
7061          s.SolutionWasLoaded = True
7062      Case NomadResult.UserCancelled
7063          Err.Raise OpenSolver_UserCancelledError, "Running NOMAD", "Model solve cancelled by user."
7064      Case NomadResult.Optimal
7065          s.SolveStatus = OpenSolverResult.Optimal
7066          s.SolveStatusString = "Optimal"
7067      End Select

          SolveModel_Nomad = s.SolveStatus
          
ExitFunction:
7091      SetCurrentDirectory currentDir
7092      Application.Cursor = xlDefault
7093      Application.StatusBar = False
7094      Application.ScreenUpdating = ScreenStatus
7095      Application.Calculation = oldCalculationMode
7096      Application.Calculate
7097      Close #1
7098      Set OS = Nothing
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("SolverNOMAD", "SolveModel_Nomad") Then Resume
          RaiseError = True
          GoTo ExitFunction

End Function

Sub CheckNomadLogs()
' If NOMAD encounters an error, we dump the exception to the log file. We can use this to deduce what went wrong
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim logFile As String
7100      If Not GetTempFilePath("log1.tmp", logFile) Then GoTo ExitSub

          
          Dim message As String
7105      Open logFile For Input As #3
7106      message = Input$(LOF(3), 3)
7107      Close #3
          
7108      If Not message Like "*NOMAD*" Then GoTo ExitSub

          If InStrText(message, "invalid parameter: DIMENSION") Then
              Dim MaxSize As Long, Position As Long
              Position = InStrRev(message, " ")
              MaxSize = CInt(Mid(message, Position + 1, InStrRev(message, ")") - Position - 1))
              Err.Raise OpenSolver_NomadError, Description:="This model contains too many variables for NOMAD to solve. NOMAD is only capable of solving models with up to " & MaxSize & " variables."
          ElseIf message Like "*invalid parameter*" Then
7112          Err.Raise OpenSolver_NomadError, Description:="One of the parameters supplied to NOMAD was invalid. This usually happens if the precision is too large. Try adjusting the values in the Solve Options dialog box."
7113      End If

ExitSub:
          Close #3
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("SolverNOMAD", "CheckNomadLogs") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Sub NOMAD_UpdateVar(X As Variant, Optional BestSolution As Variant = Nothing, Optional Infeasible As Boolean = False)
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
7125          UpdateStatusBar status
7126      End If
          
7127      OS.updateVarOS (X)
End Sub

Function NOMAD_GetValues() As Variant
7128      NOMAD_GetValues = OS.getValuesOS()
End Function

Sub NOMAD_RecalculateValues()
7129      If Not ForceCalculate("Warning: The worksheet calculation did not complete, and so the iteration may not be calculated correctly. Would you like to retry?") Then Exit Sub
End Sub

Function NOMAD_GetNumVariables() As Variant
7130      NOMAD_GetNumVariables = OS.getNumVariablesOS
End Function

Function NOMAD_GetNumConstraints() As Variant
7131      NOMAD_GetNumConstraints = OS.getNumConstraintsOS
End Function

Function NOMAD_GetVariableData() As Variant
7132      NOMAD_GetVariableData = OS.getVariableDataOS()
End Function

Function NOMAD_GetOptionData() As Variant
7133      NOMAD_GetOptionData = OS.getOptionDataOS()
End Function

