Attribute VB_Name = "OpenSolverNonLinear"
'====================================================================
'********************NON-LINEAR**************************************
'====================================================================
Public OpenSolver As COpenSolver

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
          'currentExcelDir = Application.DefaultFilePath
          
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
          Dim NomadDllFileName As String

48270     SetCurrentDirectory ThisWorkbook.Path
          'Application.DefaultFilePath = ThisWorkbook.Path
          
          'Check they are not running 64 bit office
          #If Win64 Then
              Dim res As VbMsgBoxResult
48280         NomadDllFileName = "OpenSolverNomadDll64.dll"
48290         res = MsgBox("OpenSolver Warning:" & vbCrLf & "The NOMAD optimizer is unstable on 64 bit office and it is very likely to crash excel. However it may " _
                & "solve before it crashes and then the solution can be viewed through the nomad log file 'log1.tmp' that is saved in the temp folder. This can be " _
                & "viewed from under the OpenSolver menu." & vbCrLf & vbCrLf & "Would you like to continue solving anyway?" & vbCrLf & vbCrLf & "Note: Any input on the " _
                & "errors you recieve or how to make this work in future would be much appreciated." & vbNewLine & vbNewLine _
                & "You may wish to change to one of the NEOS Non-Linear Solvers.", vbYesNo, "OpenSolver NOMAD Error")
48300         If res = vbNo Then GoTo ExitSub
          #Else
48310         NomadDllFileName = "OpenSolverNomadDll.dll"
          #End If
          
48320     If GetExistingFilePathName(ThisWorkbook.Path, NomadDllFileName, pathName) Then
              NomadDllFileName = "WhatEva.dll"
48330         NomadRetVal = Application.Run("NomadMain", SolveRelaxation)
48340     Else
48350         Err.Raise Number:=OpenSolver_NomadError, Source:=errorPrefix, Description:="Unable to find the Nomad Xll file`'" & NomadDllFileName & "' at" _
                & "any of the following location(s):" & vbCrLf & vbCrLf & pathName & vbCrLf & vbCrLf _
                & "Please ensure the file `" & NomadDllFileName & "' is in the folder:" + vbCrLf + ThisWorkbook.Path _
                & vbCrLf & "that contains the `OpenSolver.xlam' file." & vbCrLf & vbCrLf _
                & "Notes: If you do not have `" & NomadDllFileName & "' installed then try downloading it from the OpenSolver website." & vbCrLf _
                & vbCrLf & "Running OpenSolver from within the zipped file will not work."
48360     End If
          
          'Catch any errors that occured while Nomad was solving
48370     If NomadRetVal = 1 Then
48380         Err.Raise Number:=OpenSolver_NomadError, Source:=errorPrefix, Description:="There " _
                & "was an error while Nomad was solving. No solution has been loaded into the sheet."
48390         OpenSolver.SolveStatus = ErrorOccurred
48400     ElseIf NomadRetVal = 2 Then
48410         Err.Raise Number:=OpenSolver_NomadError, Source:=errorPrefix, Description:="Nomad reached " _
                & "the maximum number of iterations and returned the best feasible solution it found. This " _
                & "solution is not guaranteed to be an optimal solution." & vbCrLf & vbCrLf & "You can increase " _
                & "the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
48420         OpenSolver.SolveStatus = -1 'Unsolved
48430     ElseIf NomadRetVal = 3 Then
48440         Err.Raise Number:=OpenSolver_NomadError, Source:=errorPrefix, Description:="Nomad reached the " _
                & "maximum time and returned the best feasible solution it found. This solution is not " _
                & "guaranteed to be an optimal solution." & vbCrLf & vbCrLf & "You can increase the maximum " _
                & "time and iterations under the options in the model dialogue or check whether your model is feasible."
48450         OpenSolver.SolveStatus = TimeLimitedSubOptimal
48460     ElseIf NomadRetVal = 4 Then
48470         Err.Raise Number:=OpenSolver_NomadError, Source:=errorPrefix, Description:="Nomad reached the maximum time " _
                & "or number of iterations without finding a feasible solution. The best infeasible solution has been returned " _
                & "to the sheet." & vbCrLf & vbCrLf & "You can increase the maximum time and iterations under the options in the " _
                & "model dialogue or check whether your model is feasible."
48480         OpenSolver.SolveStatus = 5 'infeasible
48490     ElseIf NomadRetVal = 10 Then
48500         Err.Raise Number:=OpenSolver_NomadError, Source:=errorPrefix, Description:="Nomad could not find a feasible solution. " _
                & "The best infeasible solution has been returned to the sheet." & vbCrLf & vbCrLf & "Try resolving at a different start point or check whether your model " _
                & "is feasible or relax some of your constraints."
48510         OpenSolver.SolveStatus = 5 'infeasible
48520     Else
48530         OpenSolver.SolveStatus = NomadRetVal 'optimal
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
48630     SolveModel_Nomad = OpenSolver.SolveStatus    ' Return the main result
48640     Set OpenSolver = Nothing
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
48780     Set OpenSolver = Nothing
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
48870     OpenSolver.updateVarOS (X)
End Function

Function getValues() As Variant
48880     getValues = OpenSolver.getValuesOS()
End Function

Sub RecalculateValues()
48890     Sheets(ActiveSheet.Name).Calculate
End Sub

Function getNumVariables() As Variant
48900     getNumVariables = OpenSolver.getNumVariablesOS
End Function

Function getNumConstraints() As Variant
48910     getNumConstraints = OpenSolver.getNumConstraintsOS
End Function

Function getVariableData() As Variant
48920     getVariableData = OpenSolver.getVariableDataOS()
End Function

Function getOptionData() As Variant
48930     getOptionData = OpenSolver.getOptionDataOS()
End Function

