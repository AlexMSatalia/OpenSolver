Attribute VB_Name = "SolverNOMAD"
Public OpenSolver_NOMAD As COpenSolver

Public Const SolverTitle_NOMAD = "NOMAD (Non-linear Solver)"
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
    About_NOMAD = "NOMAD v" & SolverVersion_NOMAD() & " using OpenSolverNomadDLL v" & DllVersion_NOMAD() & " at " & DllPath_NOMAD()
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

Function SolveModel_Nomad(SolveRelaxation As Boolean) As Integer
          Dim ScreenStatus As Boolean
48140     ScreenStatus = Application.ScreenUpdating
          Dim s As String
48150     If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_sho", s) Then
48160         If s <> 1 Then Application.ScreenUpdating = False
48170     End If
          
          Dim oldCalculationMode As Integer
48180     oldCalculationMode = Application.Calculation
48190     Application.Calculation = xlCalculationManual
          
          Dim currentDir As String, currentExcelDir As String
48200     currentDir = CurDir
          
          ' Trap Escape key
48210     Application.EnableCancelKey = xlErrorHandler
          
48220     On Error GoTo errorHandler
          Dim errorPrefix As String
48230     errorPrefix = "OpenSolver Nomad Model Solving"
48240     If ModelStatus <> ModelStatus_Built Then
48250         Err.Raise Number:=OpenSolver_NomadError, Source:=errorPrefix, Description:="The model cannot be solved as it has not yet been built."
48260     End If
          
          Dim pathName As String
          Dim NomadRetVal As Long

48270     SetCurrentDirectory ThisWorkbook.Path
          
          ' We need to call NomadMain directly rather than use Application.Run .
          ' Using Application.Run causes the API calls inside the DLL to fail on 64 bit Office
48330     NomadRetVal = NomadMain(SolveRelaxation)
          
          'Catch any errors that occured while Nomad was solving
48370     If NomadRetVal = 1 Then
48380         Err.Raise Number:=OpenSolver_NomadError, Source:=errorPrefix, Description:="There " _
                & "was an error while Nomad was solving. No solution has been loaded into the sheet."
48390         OpenSolver_NOMAD.SolveStatus = ErrorOccurred
48400     ElseIf NomadRetVal = 2 Then
48410         Err.Raise Number:=OpenSolver_NomadError, Source:=errorPrefix, Description:="Nomad reached " _
                & "the maximum number of iterations and returned the best feasible solution it found. This " _
                & "solution is not guaranteed to be an optimal solution." & vbCrLf & vbCrLf & "You can increase " _
                & "the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
48420         OpenSolver_NOMAD.SolveStatus = -1 'Unsolved
48430     ElseIf NomadRetVal = 3 Then
48440         Err.Raise Number:=OpenSolver_NomadError, Source:=errorPrefix, Description:="Nomad reached the " _
                & "maximum time and returned the best feasible solution it found. This solution is not " _
                & "guaranteed to be an optimal solution." & vbCrLf & vbCrLf & "You can increase the maximum " _
                & "time and iterations under the options in the model dialogue or check whether your model is feasible."
48450         OpenSolver_NOMAD.SolveStatus = TimeLimitedSubOptimal
48460     ElseIf NomadRetVal = 4 Then
48470         Err.Raise Number:=OpenSolver_NomadError, Source:=errorPrefix, Description:="Nomad reached the maximum time " _
                & "or number of iterations without finding a feasible solution. The best infeasible solution has been returned " _
                & "to the sheet." & vbCrLf & vbCrLf & "You can increase the maximum time and iterations under the options in the " _
                & "model dialogue or check whether your model is feasible."
48480         OpenSolver_NOMAD.SolveStatus = 5 'infeasible
48490     ElseIf NomadRetVal = 10 Then
48500         Err.Raise Number:=OpenSolver_NomadError, Source:=errorPrefix, Description:="Nomad could not find a feasible solution. " _
                & "The best infeasible solution has been returned to the sheet." & vbCrLf & vbCrLf & "Try resolving at a different start point or check whether your model " _
                & "is feasible or relax some of your constraints."
48510         OpenSolver_NOMAD.SolveStatus = 5 'infeasible
48520     Else
48530         OpenSolver_NOMAD.SolveStatus = NomadRetVal 'optimal
48540     End If
          
ExitSub:
48550     SetCurrentDirectory currentDir
          'Application.DefaultFilePath = currentExcelDir
          ' We can fall thru to here, or jump here if the problem is shown to be infeasible before we run the CBC Solver
48560     Application.Cursor = xlDefault
48570     Application.StatusBar = False ' Resume normal status bar behaviour
48580     Application.ScreenUpdating = True
48590     Application.Calculation = oldCalculationMode
48600     Application.Calculate
48610     Application.ScreenUpdating = ScreenStatus
48620     Close #1 ' Close any open file; this does not seem to ever give errors
48630     SolveModel_Nomad = OpenSolver_NOMAD.SolveStatus    ' Return the main result
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
48860     Err.Raise ErrorNumber, ErrorSource, ErrorDescription

End Function

Function updateVar(X As Variant)
48870     OpenSolver_NOMAD.updateVarOS (X)
End Function

Function getValues() As Variant
48880     getValues = OpenSolver_NOMAD.getValuesOS()
End Function

Sub RecalculateValues()
48890     Sheets(ActiveSheet.Name).Calculate
End Sub

Function getNumVariables() As Variant
48900     getNumVariables = OpenSolver_NOMAD.getNumVariablesOS
End Function

Function getNumConstraints() As Variant
48910     getNumConstraints = OpenSolver_NOMAD.getNumConstraintsOS
End Function

Function getVariableData() As Variant
48920     getVariableData = OpenSolver_NOMAD.getVariableDataOS()
End Function

Function getOptionData() As Variant
48930     getOptionData = OpenSolver_NOMAD.getOptionDataOS()
End Function
