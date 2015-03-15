Attribute VB_Name = "SolverCBC"
Option Explicit

Public Const SolverTitle_CBC = "COIN-OR CBC (Linear solver)"
Public Const SolverDesc_CBC = "The COIN Branch and Cut solver (CBC) is the default solver for OpenSolver and is an open-source mixed-integer program (MIP) solver written in C++. CBC is an active open-source project led by John Forrest at www.coin-or.org."
Public Const SolverLink_CBC = "http://www.coin-or.org/Cbc/cbcuserguide.html"
Public Const SolverType_CBC = OpenSolver_SolverType.Linear

Public Const SolverName_CBC = "CBC"

#If Mac Then
Public Const SolverExec_CBC = "cbc"
#Else
Public Const SolverExec_CBC = "cbc.exe"
#End If

Public Const SolverScript_CBC = "cbc" & ScriptExtension

Public Const SolutionFile_CBC = "modelsolution.txt"
Public Const CostRangesFile_CBC = "costranges.txt"
Public Const RHSRangesFile_CBC = "rhsranges.txt"

Public Const UsesPrecision_CBC = False
Public Const UsesIterationLimit_CBC = False
Public Const UsesTolerance_CBC = True
Public Const UsesTimeLimit_CBC = True

Function ScriptFilePath_CBC() As String
6047      ScriptFilePath_CBC = GetTempFilePath(SolverScript_CBC)
End Function

Function SolutionFilePath_CBC() As String
6048      SolutionFilePath_CBC = GetTempFilePath(SolutionFile_CBC)
End Function

Function CostRangesFilePath_CBC() As String
6049      CostRangesFilePath_CBC = GetTempFilePath(CostRangesFile_CBC)
End Function

Function RHSRangesFilePath_CBC() As String
6050      RHSRangesFilePath_CBC = GetTempFilePath(RHSRangesFile_CBC)
End Function

Sub CleanFiles_CBC()
6051      DeleteFileAndVerify SolutionFilePath_CBC()
6052      DeleteFileAndVerify CostRangesFilePath_CBC()
6053      DeleteFileAndVerify RHSRangesFilePath_CBC()
6054      DeleteFileAndVerify ScriptFilePath_CBC()
End Sub

Function About_CBC() As String
          Dim SolverPath As String, errorString As String
6055      If Not SolverAvailable_CBC(SolverPath, errorString) Then
6056          About_CBC = errorString
6057          Exit Function
6058      End If
          ' Assemble version info
6059      About_CBC = "CBC " & SolverBitness_CBC & "-bit" & _
                           " v" & SolverVersion_CBC & _
                           " at " & MakeSpacesNonBreaking(MakePathSafe(SolverPath))
End Function

Function SolverFilePath_CBC(errorString As String) As String
6060      SolverFilePath_CBC = SolverFilePath_Default("CBC", errorString)
End Function

Function SolverAvailable_CBC(Optional SolverPath As String, Optional errorString As String) As Boolean
      ' Returns true if CBC is available and sets SolverPath
6061      SolverPath = SolverFilePath_CBC(errorString)
6062      If SolverPath = "" Then
6063          SolverAvailable_CBC = False
6064      Else
6065          SolverAvailable_CBC = True
#If Mac Then
              ' Make sure cbc is executable on Mac
6066          RunExternalCommand "chmod +x " & MakePathSafe(SolverPath)
#End If
6067      End If
End Function

Function SolverVersion_CBC() As String
' Get CBC version by running 'cbc -exit' at command line
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim SolverPath As String
6068      If Not SolverAvailable_CBC(SolverPath) Then
6069          SolverVersion_CBC = ""
6070          GoTo ExitFunction
6071      End If
          
          Dim result As String
6078      result = ReadExternalCommandOutput(MakePathSafe(SolverPath) & " -exit")
          SolverVersion_CBC = Mid(result, InStr(result, "Version:") + 9, 5)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("SolverCBC", "SolverVersion_CBC") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function SolverBitness_CBC() As String
      ' Get Bitness of CBC solver
          Dim SolverPath As String
6093      If Not SolverAvailable_CBC(SolverPath) Then
6094          SolverBitness_CBC = ""
6095          Exit Function
6096      End If
          
          ' All Macs are 64-bit so we only provide 64-bit binaries
#If Mac Then
6097      SolverBitness_CBC = "64"
#Else
6098      If right(SolverPath, 10) = "64\cbc.exe" Then
6099          SolverBitness_CBC = "64"
6100      Else
6101          SolverBitness_CBC = "32"
6102      End If
#End If
End Function

Function CreateSolveScript_CBC(SolutionFilePathName As String, ExtraParameters As Dictionary, SolveOptions As SolveOptionsType, s As COpenSolver) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim CommandLineRunString As String, PrintingOptionString As String, ExtraParametersString As String
          ExtraParametersString = ParametersToString_CBC(ExtraParameters)
          
          ' have to split up the command line as excel couldn't have a string longer than 255 characters??
6119      CommandLineRunString = " -directory " & MakePathSafe(left(GetTempFolder, Len(GetTempFolder) - 1)) _
                               & " -import " & MakePathSafe(s.ModelFilePathName) _
                               & " -ratioGap " & str(SolveOptions.Tolerance) _
                               & " -seconds " & str(SolveOptions.MaxTime) _
                               & " " & ExtraParametersString _
                               & " -solve " _
                               & IIf(s.bGetDuals, " -printingOptions all ", "") _
                               & " -solution " & MakePathSafe(SolutionFilePathName)

6120      PrintingOptionString = IIf(s.bGetDuals, " -printingOptions rhs  -solution " & RHSRangesFile_CBC & " -printingOptions objective -solution " & CostRangesFile_CBC, "")
                        
          Dim scriptFile As String, scriptFileContents As String
6121      scriptFile = ScriptFilePath_CBC()
6122      scriptFileContents = MakePathSafe(s.ExternalSolverPathName) & CommandLineRunString & PrintingOptionString
6123      CreateScriptFile scriptFile, scriptFileContents
          
6124      CreateSolveScript_CBC = scriptFile

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("SolverCBC", "CreateSolveScript_CBC") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function ParametersToString_CBC(ExtraParameters As Dictionary) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim Key As Variant
          For Each Key In ExtraParameters.Keys
              ParametersToString_CBC = ParametersToString_CBC & IIf(left(Key, 1) <> "-", "-", "") & Key & " " & ExtraParameters.Item(Key) & " "
          Next Key
          ParametersToString_CBC = Trim(ParametersToString_CBC)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("SolverCBC", "ParametersToString_CBC") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function ReadModel_CBC(SolutionFilePathName As String, s As COpenSolver) As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim Response As String
6125      ReadModel_CBC = False
6127      Open SolutionFilePathName For Input As #1 ' supply path with filename
6128      Line Input #1, Response  ' Optimal - objective value              22

6129      s.SolutionWasLoaded = True
6130      If Response Like "Optimal*" Then
6131          s.SolveStatus = OpenSolverResult.Optimal
6132          s.SolveStatusString = "Optimal"
              '
6134      ElseIf Response Like "Infeasible*" Then
6135          s.SolveStatus = OpenSolverResult.Infeasible
6136          s.SolveStatusString = "No Feasible Solution"
              '
6138      ElseIf Response Like "Integer infeasible*" Then
6139          s.SolveStatus = OpenSolverResult.Infeasible
6140          s.SolveStatusString = "No Feasible Integer Solution"
              '
6142      ElseIf Response Like "Unbounded*" Then
6143          s.SolveStatus = OpenSolverResult.Unbounded
6144          s.SolveStatusString = "No Solution Found (Unbounded)"
6146          s.SolutionWasLoaded = False
              '
6147      ElseIf Response Like "Stopped on time *" Then ' Stopped on iterations or time
6148          s.SolveStatus = OpenSolverResult.LimitedSubOptimal
6149          s.SolveStatusString = "Stopped on Time Limit"
6150          If Response Like "*(no integer solution - continuous used)*" Then
6151              s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
6152          End If
              '
6154      ElseIf Response Like "Stopped on iterations*" Then ' Stopped on iterations or time
6155          s.SolveStatus = OpenSolverResult.LimitedSubOptimal
6156          s.SolveStatusString = "Stopped on Iteration Limit"
6157          If Response Like "*(no integer solution - continuous used)*" Then
6158              s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
6159          End If
              '
6161      ElseIf Response Like "Stopped on difficulties*" Then ' Stopped on iterations or time
6162          s.SolveStatus = OpenSolverResult.LimitedSubOptimal
6163          s.SolveStatusString = "Stopped on CBC difficulties"
6164          If Response Like "*(no integer solution - continuous used)*" Then
6165              s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
6166          End If
              '
6168      ElseIf Response Like "Stopped on ctrl-c*" Then ' Stopped on iterations or time
6169          s.SolveStatus = OpenSolverResult.LimitedSubOptimal
6170          s.SolveStatusString = "Stopped on Ctrl-C"
6171          If Response Like "*(no integer solution - continuous used)*" Then
6172              s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
6173          End If
              '
6175      ElseIf Response Like "Status unknown*" Then
6176          Err.Raise OpenSolver_CBCError, Description:="CBC solver did not solve the problem, suggesting there was an error in the CBC input parameters. The response was: " & vbCrLf _
               & Response _
               & vbCrLf & "The CBC command line can be found at:" _
               & vbCrLf & ScriptFilePath_CBC()
6178      Else
6179          Err.Raise OpenSolver_CBCError, Description:="The response from the CBC solver is not recognised. The response was: " & Response
6181      End If
          
          ' Remove the double spaces from Response
6182      Response = Replace(Response, "    ", " ")
6183      Response = Replace(Response, "   ", " ")
6184      Response = Replace(Response, "  ", " ")
6185      Response = Replace(Response, "  ", " ")
6186      Response = Replace(Response, "  ", " ")
6187      Response = Replace(Response, "  ", " ")

6188      If s.SolutionWasLoaded Then
              ' We read in whatever solution CBC returned
6189          UpdateStatusBar "OpenSolver: Loading Solution... " & Response
          
              Dim Line As String, SplitLine() As String, Index As Double, NameValue As String, value As Double, CBCConstraintIndex As Long, StartOffset As Long
6191          If s.bGetDuals Then
                  ' Read in the Solution File
                  ' Line format: Index ConstraintName Value ShadowPrice
                  
                  Dim j As Long, row As Long
6192              CBCConstraintIndex = 0
                  
                  ' Throw away first constraint if it was from a seek objective model
6193              If s.ObjectiveSense = TargetObjective Then
6194                  Line Input #1, Line
6195                  CBCConstraintIndex = CBCConstraintIndex + 1
6196              End If

6197              j = 1
6198              For row = 1 To s.NumRows
6199                  If s.GetSparseACount(row) = 0 Then
                          ' This constraint was not written to the model, as it had no coefficients. Just ignore it.
6200                      s.rConstraintList.Cells(row, 2).ClearContents
6201                  Else
6202                      Line Input #1, Line
6203                      SplitLine = SplitWithoutRepeats(Line, " ")

                          ' In the case of LpStatusInfeasible, we can get lines that start **. We strip the **
                          StartOffset = 0
                          If SplitLine(StartOffset) = "**" Then StartOffset = 1

                          ' Check the index of the row
6212                      If CInt(SplitLine(StartOffset)) <> CBCConstraintIndex Then
6213                          Err.Raise OpenSolver_CBCError, Description:="While reading the CBC solution file, OpenSolver found an unexpected constraint row."
6215                      End If
6216
6220                      NameValue = SplitLine(StartOffset + 1)
6225                      s.FinalValue(j) = ConvertToCurrentLocale(SplitLine(StartOffset + 2))
                          value = ConvertToCurrentLocale(SplitLine(StartOffset + 3))
6230                      If s.ObjectiveSense = MaximiseObjective Then value = -value
6235                      s.ShadowPrice(j) = value
6239                      CBCConstraintIndex = CBCConstraintIndex + 1
6240                      j = j + 1
6241                  End If
6242              Next row
6243              ReadSensitivityData_CBC SolutionFilePathName, s
6244          End If
            
              ' Now we read in the decision variable values
              ' Line format: Index VariableName Value ReducedCost
6245          j = 1
6246          While Not EOF(1)
6247              Line Input #1, Line
6248              SplitLine = SplitWithoutRepeats(Line, " ")

                  ' In the case of LpStatusInfeasible, we can get lines that start **. We strip the **
6253              StartOffset = 0
                  If SplitLine(StartOffset) = "**" Then StartOffset = 1
                  
6257              Index = CInt(SplitLine(StartOffset))
6258              NameValue = SplitLine(StartOffset + 1)
6263              If left(NameValue, 1) = "_" Then NameValue = Mid(NameValue, 2) ' Strip any _ character added to make a valid name
                  s.VarCell(j) = NameValue
6268              s.FinalVarValue(j) = ConvertToCurrentLocale(SplitLine(StartOffset + 2))
                 
                  If s.bGetDuals Then
6271                  value = ConvertToCurrentLocale(SplitLine(StartOffset + 3))
6275                  If s.ObjectiveSense = MaximiseObjective Then value = -value
6280                  s.ReducedCosts(j) = value
6285              End If
6286              j = j + 1
6287          Wend
              s.SolutionWasLoaded = True

6288      End If

6290      ReadModel_CBC = True

ExitFunction:
          Close #1
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("SolverCBC", "ReadModel_CBC") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Sub ReadSensitivityData_CBC(SolutionFilePathName As String, s As COpenSolver)
'Reads the two files with the limits on the bounds of shadow prices and reduced costs
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim RangeFilePathName As String, LineData() As String
          Dim Line As String, row As Long, j As Long
          
          'Find the ranges on the constraints
6295      RangeFilePathName = left(SolutionFilePathName, InStrRev(SolutionFilePathName, Application.PathSeparator)) & RHSRangesFile_CBC

6297      Open RangeFilePathName For Input As #2 ' supply path with filename
6298      Line Input #2, Line 'Dont want first line
6299      j = 1
6300      While Not EOF(2)
6301          Line Input #2, Line
6302          LineData() = Split(Line, ",")
6312          s.IncreaseCon(j) = ConvertToCurrentLocale(LineData(2))
6313          s.DecreaseCon(j) = ConvertToCurrentLocale(LineData(4))
6314          j = j + 1
6315      Wend
6316      Close 2
          
6317      j = 1
          'Find the ranges on the variables
6318      RangeFilePathName = left(SolutionFilePathName, InStrRev(SolutionFilePathName, Application.PathSeparator)) & CostRangesFile_CBC
6319      Open RangeFilePathName For Input As #2 ' supply path with filename
6320      Line Input #2, Line 'Dont want first line
6321      row = s.NumRows + 2
6322      While Not EOF(2)
6323          Line Input #2, Line
              LineData() = Split(Line, ",")
6334          If s.ObjectiveSense = MaximiseObjective Then
6335              s.IncreaseVar(j) = ConvertToCurrentLocale(LineData(4))
6336              s.DecreaseVar(j) = ConvertToCurrentLocale(LineData(2))
6337          Else
6338              s.IncreaseVar(j) = ConvertToCurrentLocale(LineData(2))
6339              s.DecreaseVar(j) = ConvertToCurrentLocale(LineData(4))
6340          End If
6341          j = j + 1
6342      Wend

ExitSub:
          Close #2
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("SolverCBC", "ReadSensitivityData_CBC") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Sub LaunchCommandLine_CBC()
' Open the CBC solver with our last model loaded.
' If we have a worksheet open with a model, then we pass the solver options (max runtime etc) from this model to CBC. Otherwise, we don't pass any options.
6347      On Error GoTo ErrorHandler
            
          Dim WorksheetAvailable As Boolean
6349      WorksheetAvailable = CheckWorksheetAvailable(SuppressDialogs:=True)
          
          Dim SolverPath As String, errorString As String
6350      If Not SolverAvailable_CBC(SolverPath, errorString) Then
6351          Err.Raise OpenSolver_CBCError, Description:=errorString
6352      End If
          
          Dim ModelFilePathName As String
6353      ModelFilePathName = ModelFilePath("CBC")
          
          Dim SolveOptions As SolveOptionsType, SolveOptionsString As String
6354      If WorksheetAvailable Then
6355          GetSolveOptions ActiveSheet, SolveOptions
6356          If errorString = "" Then
6357             SolveOptionsString = " -ratioGap " & CStr(SolveOptions.Tolerance) & " -seconds " & CStr(SolveOptions.MaxTime)
6358          End If
6359      End If
          
          Dim ExtraParametersString As String, ExtraParameters As New Dictionary
6360      If WorksheetAvailable Then
              GetExtraParameters "CBC", ActiveSheet, ExtraParameters
              If errorString <> "" Then
                  ExtraParametersString = ""
              Else
6361              ExtraParametersString = ParametersToString_CBC(ExtraParameters)
6362          End If
6363      End If
             
          Dim CBCRunString As String
6364      CBCRunString = " -directory " & MakePathSafe(left(GetTempFolder, Len(GetTempFolder) - 1)) _
                           & " -import " & MakePathSafe(ModelFilePathName) _
                           & SolveOptionsString _
                           & " " & ExtraParametersString _
                           & " -" ' Force CBC to accept commands from the command line
6365      RunExternalCommand MakePathSafe(SolverPath) & CBCRunString, "", Normal, False

ExitSub:
          Exit Sub

ErrorHandler:
          ReportError "SolverCBC", "LaunchCommandLine_CBC", True
          GoTo ExitSub
End Sub
