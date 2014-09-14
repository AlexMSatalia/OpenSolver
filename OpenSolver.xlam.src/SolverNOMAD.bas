Attribute VB_Name = "SolverNOMAD"
Public OS As COpenSolver
Dim IterationCount As Long

Public Const SolverTitle_NOMAD = "NOMAD (Non-linear solver)"
Public Const SolverDesc_NOMAD = "Nomad (Nonsmooth Optimization by Mesh Adaptive Direct search) is a C++ implementation of the Mesh Adaptive Direct Search (Mads) algorithm that solves non-linear problems. It works by updating the values on the sheet and passing them to the C++ solver. Like many non-linear solvers NOMAD cannot guarantee optimality of its solutions."
Public Const SolverLink_NOMAD = "http://www.gerad.ca/nomad/Project/Home.html"
Public Const SolverType_NOMAD = OpenSolver_SolverType.NonLinear

#If Win64 Then
Public Const NomadDllName = "OpenSolverNomadDll64.dll"
#Else
Public Const NomadDllName = "OpenSolverNomadDll.dll"
#End If

' Don't forget we need to chdir to the directory containing the DLL before calling any of the functions
#If VBA7 Then
    #If Win64 Then
        Public Declare PtrSafe Function NomadMain Lib "OpenSolverNomadDll64.dll" (ByVal SolveRelaxation As Boolean) As Long
        Private Declare PtrSafe Function NomadVersion Lib "OpenSolverNomadDll64.dll" () As String
        Private Declare PtrSafe Function NomadDllVersion Lib "OpenSolverNomadDll64.dll" Alias "NomadDLLVersion" () As String
    #Else
        Public Declare PtrSafe Function NomadMain Lib "OpenSolverNomadDll.dll" (ByVal SolveRelaxation As Boolean) As Long
        Private Declare PtrSafe Function NomadVersion Lib "OpenSolverNomadDll.dll" () As String
        Private Declare PtrSafe Function NomadDllVersion Lib "OpenSolverNomadDll.dll" Alias "NomadDLLVersion" () As String
    #End If
#Else ' All VBA6 is 32 bit
    Public Declare Function NomadMain Lib "OpenSolverNomadDll.dll" (ByVal SolveRelaxation As Boolean) As Long
    Private Declare Function NomadVersion Lib "OpenSolverNomadDll.dll" () As String
    Private Declare Function NomadDllVersion Lib "OpenSolverNomadDll.dll" Alias "NomadDLLVersion" () As String
#End If


Function About_NOMAD() As String
    Dim errorString As String
    If Not SolverAvailable_NOMAD(errorString) Then
        About_NOMAD = errorString
        Exit Function
    End If
    ' Assemble version info
    About_NOMAD = "NOMAD " & SolverBitness_NOMAD & "-bit v" & SolverVersion_NOMAD() & _
                  " using OpenSolverNomadDLL v" & DllVersion_NOMAD() & " at " & MakeSpacesNonBreaking(DllPath_NOMAD())
End Function

Function SolverAvailable_NOMAD(Optional errorString As String) As Boolean
#If Mac Then
    errorString = "NOMAD for OpenSolver is not currently supported on Mac"
    SolverAvailable_NOMAD = False
    Exit Function
#Else
    ' Set current dir for finding the DLL
    Dim currentDir As String
    currentDir = CurDir
    SetCurrentDirectory ThisWorkbook.Path
    
    ' Try to access DLL - throws error if not found
    On Error GoTo NotFound
    NomadVersion
    
    SetCurrentDirectory currentDir
    SolverAvailable_NOMAD = True
    Exit Function

NotFound:
    SetCurrentDirectory currentDir
    SolverAvailable_NOMAD = False
    errorString = "Unable to find the Nomad DLL file `" & NomadDllName & "' in the folder that contains `OpenSolver.xlam'"
    Exit Function
#End If
End Function

Function SolverVersion_NOMAD() As String
    If Not SolverAvailable_NOMAD() Then
        SolverVersion_NOMAD = ""
        Exit Function
    End If
    
    Dim currentDir As String, sNomadVersion As String
    
    ' Set current dir for finding the DLL
    currentDir = CurDir
    SetCurrentDirectory ThisWorkbook.Path
    
    ' Get version info from DLL
    sNomadVersion = NomadVersion()
    sNomadVersion = left(sNomadVersion, InStr(sNomadVersion, vbNullChar) - 1)
    
    SetCurrentDirectory currentDir
    
    SolverVersion_NOMAD = sNomadVersion
End Function

Function DllVersion_NOMAD() As String
    If Not SolverAvailable_NOMAD() Then
        DllVersion_NOMAD = ""
        Exit Function
    End If
    
    Dim currentDir As String, sDllVersion As String
    
    ' Set current dir for finding the DLL
    currentDir = CurDir
    SetCurrentDirectory ThisWorkbook.Path
    
    ' Get version info from DLL
    sDllVersion = NomadDllVersion()
    sDllVersion = left(sDllVersion, InStr(sDllVersion, vbNullChar) - 1)
    
    SetCurrentDirectory currentDir
    
    DllVersion_NOMAD = sDllVersion
End Function

Function DllPath_NOMAD() As String
    GetExistingFilePathName ThisWorkbook.Path, NomadDllName, DllPath_NOMAD
End Function

Function SolverBitness_NOMAD() As String
' Get Bitness of NOMAD solver
    If Not SolverAvailable_NOMAD() Then
        SolverBitness_NOMAD = ""
        Exit Function
    End If
    
    If right(NomadDllName, 6) = "64.dll" Then
        SolverBitness_NOMAD = "64"
    Else
        SolverBitness_NOMAD = "32"
    End If
End Function

Function SolveModel_Nomad(SolveRelaxation As Boolean, s As COpenSolver) As Long
          Dim ScreenStatus As Boolean
48140     ScreenStatus = Application.ScreenUpdating
          Dim Show As String
48150     If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_sho", Show) Then
48160         If Show <> 1 Then Application.ScreenUpdating = False
48170     End If

          ' Trap Escape key
48210     Application.EnableCancelKey = xlErrorHandler
          
48220     On Error GoTo errorHandler
          Dim errorPrefix As String
48230     errorPrefix = "OpenSolver Nomad Model Solving"
48240     If s.ModelStatus <> ModelStatus_Built Then
48250         Err.Raise Number:=OpenSolver_NomadError, Source:=errorPrefix, Description:="The model cannot be solved as it has not yet been built."
48260     End If
          
          ' Loop through all decision vars and set their values
          ' This is to try and catch any protected cells as we can't catch VBA errors that occur while NOMAD calls back into VBA
          Dim c As Range
          For Each c In s.AdjustableCells
              c.Value2 = c.Value2
          Next c
          
          ' Set OS for the calls back into Excel from NOMAD
          Set OS = s
          
          Dim oldCalculationMode As Long
48180     oldCalculationMode = Application.Calculation
48190     Application.Calculation = xlCalculationManual
          
          Dim currentDir As String
48200     currentDir = CurDir
          
48270     SetCurrentDirectory ThisWorkbook.Path

          IterationCount = 0
          
          ' We need to call NomadMain directly rather than use Application.Run .
          ' Using Application.Run causes the API calls inside the DLL to fail on 64 bit Office
          Dim NomadRetVal As Long
48330     NomadRetVal = NomadMain(SolveRelaxation)
          
          'Catch any errors that occured while Nomad was solving
48370     If NomadRetVal = 1 Then
48380         Err.Raise Number:=OpenSolver_NomadError, Source:=errorPrefix, Description:="There " _
                & "was an error while Nomad was solving. No solution has been loaded into the sheet."
48390         s.SolveStatus = OpenSolverResult.ErrorOccurred
48400     ElseIf NomadRetVal = 2 Then
48410         s.SolveStatusComment = "Nomad reached the maximum number of iterations and returned the best feasible solution it found. This solution is not guaranteed to be an optimal solution." & vbCrLf & vbCrLf & _
                                     "You can increase the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
48420         s.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
              s.SolveStatusString = "Stopped on Iteration Limit"
              s.LinearSolutionWasLoaded = True
48430     ElseIf NomadRetVal = 3 Then
48440         s.SolveStatusComment = "Nomad reached the maximum time and returned the best feasible solution it found. This solution is not guaranteed to be an optimal solution." & vbCrLf & vbCrLf & _
                                     "You can increase the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
48450         s.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
              s.SolveStatusString = "Stopped on Time Limit"
              s.LinearSolutionWasLoaded = True
48460     ElseIf NomadRetVal = 4 Then
48470         s.SolveStatusComment = "Nomad reached the maximum time or number of iterations without finding a feasible solution. The best infeasible solution has been returned to the sheet." & vbCrLf & vbCrLf & _
                                     "You can increase the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
48480         s.SolveStatus = OpenSolverResult.Infeasible
              s.SolveStatusString = "No Feasible Solution"
              s.LinearSolutionWasLoaded = True
48490     ElseIf NomadRetVal = 10 Then
48500         s.SolveStatusComment = "Nomad could not find a feasible solution. The best infeasible solution has been returned to the sheet." & vbCrLf & vbCrLf & _
                                     "Try resolving at a different start point or check whether your model is feasible or relax some of your constraints."
48510         s.SolveStatus = OpenSolverResult.Infeasible
              s.SolveStatusString = "No Feasible Solution"
              s.LinearSolutionWasLoaded = True
          ElseIf NomadRetVal = -3 Then
              Err.Raise OpenSolver_UserCancelledError, "Running NOMAD", "Model solve cancelled by user."
48520     Else
48530         s.SolveStatus = NomadRetVal 'optimal
20830         s.SolveStatusString = "Optimal"
48540     End If
          
ExitSub:
          ' We can fall thru to here
48550     SetCurrentDirectory currentDir
48560     Application.Cursor = xlDefault
48570     Application.StatusBar = False ' Resume normal status bar behaviour
48580     Application.ScreenUpdating = True
48590     Application.Calculation = oldCalculationMode
48600     Application.Calculate
48610     Application.ScreenUpdating = ScreenStatus
48620     Close #1 ' Close any open file; this does not seem to ever give errors
48630     SolveModel_Nomad = s.SolveStatus    ' Return the main result
          Set OS = Nothing
48650     Exit Function
          
errorHandler:
          ' We only trap Escape (Err.Number=18) here; all other errors are passed back to the caller.
          ' Save error message
          Dim ErrorNumber As Long, ErrorDescription As String, ErrorSource As String
48660     ErrorNumber = Err.Number
48670     ErrorDescription = Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
48680     ErrorSource = Err.Source
48690     If Err.Number = 18 Then
48700         If MsgBox("You have pressed the Escape key. Do you wish to cancel?", _
                         vbCritical + vbYesNo + vbDefaultButton1, _
                         "OpenSolver: User Interrupt Occured...") = vbNo Then
48710             Resume 'continue on from where error occured
48720         Else
                  ' Raise a "user cancelled" error. We cannot use Raise, as that exits immediately without going thru our code below
48730             ErrorNumber = OpenSolver_UserCancelledError
48740             ErrorSource = errorPrefix
48750             ErrorDescription = "Model solve cancelled by user."
48760         End If
48770     End If

ErrorExit:
          ' Exit, raising an error. None of the following actions change the Err.Number etc, but we saved them above just in case...
48790     SetCurrentDirectory currentDir
          'Application.DefaultFilePath = currentExcelDir
48800     Application.Cursor = xlDefault
48810     Application.StatusBar = False ' Resume normal status bar behaviour
48820     Application.ScreenUpdating = True
48830     Application.Calculation = oldCalculationMode
48840     Application.Calculate
48850     Close #1 ' Close any open file; this does not seem to ever give errors
          Set OS = Nothing
48860     Err.Raise ErrorNumber, ErrorSource, ErrorDescription

End Function

Function updateVar(X As Variant, Optional BestSolution As Variant = Nothing, Optional Infeasible As Boolean = False)
          IterationCount = IterationCount + 1

          ' Update solution
          If IterationCount Mod 5 = 0 Then
              Dim status As String
              status = "OpenSolver: Running NOMAD. Iteration " & IterationCount & "."
              ' Check for BestSolution = Nothing
              If Not VarType(BestSolution) = 9 Then
                  ' Flip solution if maximisation
                  If OS.ObjectiveSense = MaximiseObjective Then BestSolution = -BestSolution

                  status = status & " Best solution so far: " & BestSolution
                  If Infeasible Then
                      status = status & " (infeasible)"
                  End If
              End If
              Application.StatusBar = status
          End If
          
48870     OS.updateVarOS (X)
End Function

Function getValues() As Variant
48880     getValues = OS.getValuesOS()
End Function

Sub RecalculateValues()
48890     Sheets(ActiveSheet.Name).Calculate
End Sub

Function getNumVariables() As Variant
48900     getNumVariables = OS.getNumVariablesOS
End Function

Function getNumConstraints() As Variant
48910     getNumConstraints = OS.getNumConstraintsOS
End Function

Function getVariableData() As Variant
48920     getVariableData = OS.getVariableDataOS()
End Function

Function getOptionData() As Variant
48930     getOptionData = OS.getOptionDataOS()
End Function

