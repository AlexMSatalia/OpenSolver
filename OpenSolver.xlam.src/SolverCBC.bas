Attribute VB_Name = "SolverCBC"
Option Explicit

Public Const SolverTitle_CBC = "COIN-OR CBC (Linear solver)"
Public Const SolverDesc_CBC = "The COIN Branch and Cut solver (CBC) is the default solver for OpenSolver and is an open-source mixed-integer program (MIP) solver written in C++. CBC is an active open-source project led by John Forrest at www.coin-or.org."
Public Const SolverLink_CBC = "http://www.coin-or.org/Cbc/cbcuserguide.html"
Public Const SolverType_CBC = OpenSolver_SolverType.Linear

#If Mac Then
Public Const SolverName_CBC = "cbc"
#Else
Public Const SolverName_CBC = "cbc.exe"
#End If

Public Const SolverScript_CBC = "cbc" & ScriptExtension

Public Const SolutionFile_CBC = "modelsolution.txt"
Public Const CostRangesFile_CBC = "costranges.txt"
Public Const RHSRangesFile_CBC = "rhsranges.txt"

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

Sub CleanFiles_CBC(errorPrefix As String)
          ' Solution file
6051      DeleteFileAndVerify SolutionFilePath_CBC(), errorPrefix, "Unable to delete the CBC solver solution file: " & SolutionFilePath_CBC()
          ' Cost Range file
6052      DeleteFileAndVerify CostRangesFilePath_CBC(), errorPrefix, "Unable to delete the CBC solver sensitivity data file: " & CostRangesFilePath_CBC()
          ' RHS Range file
6053      DeleteFileAndVerify RHSRangesFilePath_CBC(), errorPrefix, "Unable to delete the CBC solver sensitivity data file: " & RHSRangesFilePath_CBC()
          ' Script file
6054      DeleteFileAndVerify ScriptFilePath_CBC(), errorPrefix, "Unable to delete the CBC solver script file: " & ScriptFilePath_CBC()
End Sub

Function About_CBC() As String
      ' Return string for "About" form
          Dim SolverPath As String, errorString As String
6055      If Not SolverAvailable_CBC(SolverPath, errorString) Then
6056          About_CBC = errorString
6057          Exit Function
6058      End If
          ' Assemble version info
6059      About_CBC = "CBC " & SolverBitness_CBC & "-bit" & _
                           " v" & SolverVersion_CBC & _
                           " at " & MakeSpacesNonBreaking(SolverPath)
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
6066          system ("chmod +x " & ConvertHfsPath(SolverPath))
#End If
          
6067      End If
End Function

Function SolverVersion_CBC() As String
      ' Get CBC version by running 'cbc -exit' at command line
          Dim SolverPath As String
6068      If Not SolverAvailable_CBC(SolverPath) Then
6069          SolverVersion_CBC = ""
6070          Exit Function
6071      End If
          
          ' Set up cbc to write version info to text file
          Dim logFile As String
6072      logFile = GetTempFilePath("cbcversion.txt")
6073      If FileOrDirExists(logFile) Then Kill logFile
          
          Dim RunPath As String, FileContents As String
6074      RunPath = ScriptFilePath_CBC()
6075      If FileOrDirExists(RunPath) Then Kill RunPath
6076      FileContents = QuotePath(ConvertHfsPath(SolverPath)) & " -exit" & " > " & QuotePath(ConvertHfsPath(logFile))
6077      CreateScriptFile RunPath, FileContents
          
          ' Run cbc
          Dim completed As Boolean
6078      completed = RunExternalCommand(ConvertHfsPath(RunPath), "", SW_HIDE, True) 'OSSolveSync
          
          ' Read version info back from output file
          Dim Line As String
6079      If FileOrDirExists(logFile) Then
6080          On Error GoTo ErrHandler
6081          Open logFile For Input As 1
6082          Line Input #1, Line
6083          Line Input #1, Line
6084          Close #1
6085          SolverVersion_CBC = right(Line, Len(Line) - 9)
6086          SolverVersion_CBC = left(SolverVersion_CBC, Len(SolverVersion_CBC) - 1)
6087      Else
6088          SolverVersion_CBC = ""
6089      End If
6090      Exit Function
          
ErrHandler:
6091      Close #1
6092      Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
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

Function GetExtraParameters_CBC(sheet As Worksheet, errorString As String) As String
          ' The user can define a set of parameters they want to pass to CBC; this gets them as a string
          ' Note: The named range MUST be on the current sheet
          Dim CBCParametersRange As Range, CBCExtraParametersString As String, i As Long
6103      errorString = ""
6104      If GetNamedRangeIfExistsOnSheet(sheet, "OpenSolver_CBCParameters", CBCParametersRange) Then
6105          If CBCParametersRange.Columns.Count <> 2 Then
6106              errorString = "The range OpenSolver_CBCParameters must be a two-column table."
6107              Exit Function
6108          End If
6109          For i = 1 To CBCParametersRange.Rows.Count
                  Dim ParamName As String, ParamValue As String
6110              ParamName = Trim(CBCParametersRange.Cells(i, 1))
6111              If ParamName <> "" Then
6112                  If left(ParamName, 1) <> "-" Then ParamName = "-" & ParamName
6113                  ParamValue = Trim(CBCParametersRange.Cells(i, 2))
6114                  CBCExtraParametersString = CBCExtraParametersString & " " & ParamName & " " & ConvertFromCurrentLocale(ParamValue)
6115              End If
6116          Next i
6117      End If
6118      GetExtraParameters_CBC = CBCExtraParametersString
End Function

Function CreateSolveScript_CBC(SolutionFilePathName As String, ExtraParametersString As String, SolveOptions As SolveOptionsType, s As COpenSolver) As String
          Dim CommandLineRunString As String, PrintingOptionString As String
          ' have to split up the command line as excel couldn't have a string longer than 255 characters??
6119      CommandLineRunString = " -directory " & QuotePath(ConvertHfsPath(left(GetTempFolder, Len(GetTempFolder) - 1))) _
                               & " -import " & QuotePath(ConvertHfsPath(s.ModelFilePathName)) _
                               & " -ratioGap " & str(SolveOptions.Tolerance) _
                               & " -seconds " & str(SolveOptions.maxTime) _
                               & ExtraParametersString _
                               & " -solve " _
                               & IIf(s.bGetDuals, " -printingOptions all ", "") _
                               & " -solution " & QuotePath(ConvertHfsPath(SolutionFilePathName))
          '-------------------sensitivity analysis-----------------------------------------------------------
          'extra command line option of -printingOptions rhs -solution rhsranges.txt gives the allowable increase for constraint rhs.
          '-this file has the increase as the third input and allowable decrease as the fifth input
          'extra command line option of -printingOptions objective - solution costranges.txt outputs the ranges on the costs to the costranges file
          '-this file has the increase as the fifth input and decrease as the third input
6120      PrintingOptionString = IIf(s.bGetDuals, " -printingOptions rhs  -solution " & RHSRangesFile_CBC & " -printingOptions objective -solution " & CostRangesFile_CBC, "")
                        
          Dim scriptFile As String, scriptFileContents As String
6121      scriptFile = ScriptFilePath_CBC()
6122      scriptFileContents = QuotePath(ConvertHfsPath(s.ExternalSolverPathName)) & CommandLineRunString & PrintingOptionString
6123      CreateScriptFile scriptFile, scriptFileContents
          
6124      CreateSolveScript_CBC = scriptFile
End Function

Function ReadModel_CBC(SolutionFilePathName As String, errorString As String, s As COpenSolver) As Boolean
          Dim LinearSolveStatusString As String
6125      ReadModel_CBC = False
6126      On Error GoTo ErrHandler
6127      Open SolutionFilePathName For Input As 1 ' supply path with filename
6128      Line Input #1, LinearSolveStatusString  ' Optimal - objective value              22
          ' Line Input #1, junk ' get rest of line
          Dim solutionExpected As Boolean
6129      solutionExpected = True
6130      If LinearSolveStatusString Like "Optimal*" Then
6131          s.SolveStatus = OpenSolverResult.Optimal
6132          s.SolveStatusString = "Optimal"
6133          s.LinearSolveStatus = LinearSolveResult.Optimal
              '
6134      ElseIf LinearSolveStatusString Like "Infeasible*" Then
6135          s.SolveStatus = OpenSolverResult.Infeasible
6136          s.SolveStatusString = "No Feasible Solution"
6137          s.LinearSolveStatus = LinearSolveResult.Infeasible
              '
6138      ElseIf LinearSolveStatusString Like "Integer infeasible*" Then
6139          s.SolveStatus = OpenSolverResult.Infeasible
6140          s.SolveStatusString = "No Feasible Integer Solution"
6141          s.LinearSolveStatus = LinearSolveResult.IntegerInfeasible
              '
6142      ElseIf LinearSolveStatusString Like "Unbounded*" Then
6143          s.SolveStatus = OpenSolverResult.Unbounded
6144          s.SolveStatusString = "No Solution Found (Unbounded)"
6145          s.LinearSolveStatus = LinearSolveResult.Unbounded
6146          solutionExpected = False
              '
6147      ElseIf LinearSolveStatusString Like "Stopped on time *" Then ' Stopped on iterations or time
6148          s.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
6149          s.SolveStatusString = "Stopped on Time Limit"
6150          If LinearSolveStatusString Like "*(no integer solution - continuous used)*" Then
6151              s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
6152          End If
6153          s.LinearSolveStatus = LinearSolveResult.SolveStopped
              '
6154      ElseIf LinearSolveStatusString Like "Stopped on iterations*" Then ' Stopped on iterations or time
6155          s.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
6156          s.SolveStatusString = "Stopped on Iteration Limit"
6157          If LinearSolveStatusString Like "*(no integer solution - continuous used)*" Then
6158              s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
6159          End If
6160          s.LinearSolveStatus = LinearSolveResult.SolveStopped
              '
6161      ElseIf LinearSolveStatusString Like "Stopped on difficulties*" Then ' Stopped on iterations or time
6162          s.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
6163          s.SolveStatusString = "Stopped on CBC difficulties"
6164          If LinearSolveStatusString Like "*(no integer solution - continuous used)*" Then
6165              s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
6166          End If
6167          s.LinearSolveStatus = LinearSolveResult.SolveStopped
              '
6168      ElseIf LinearSolveStatusString Like "Stopped on ctrl-c*" Then ' Stopped on iterations or time
6169          s.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
6170          s.SolveStatusString = "Stopped on Ctrl-C"
6171          If LinearSolveStatusString Like "*(no integer solution - continuous used)*" Then
6172              s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
6173          End If
6174          s.LinearSolveStatus = LinearSolveResult.SolveStopped
              '
6175      ElseIf LinearSolveStatusString Like "Status unknown*" Then
6176          errorString = "CBC solver did not solve the problem, suggesting there was an error in the CBC input parameters. The response was: " & vbCrLf _
               & LinearSolveStatusString _
               & vbCrLf & "The CBC command line can be found at:" _
               & vbCrLf & ScriptFilePath_CBC()
6177          GoTo ExitSub
6178      Else
6179          errorString = "The response from the CBC solver is not recognised. The response was: " & LinearSolveStatusString
6180          GoTo ExitSub
6181      End If
          
          ' Remove the double spaces from LinearSolveStatusString
6182      LinearSolveStatusString = Replace(LinearSolveStatusString, "    ", " ")
6183      LinearSolveStatusString = Replace(LinearSolveStatusString, "   ", " ")
6184      LinearSolveStatusString = Replace(LinearSolveStatusString, "  ", " ")
6185      LinearSolveStatusString = Replace(LinearSolveStatusString, "  ", " ")
6186      LinearSolveStatusString = Replace(LinearSolveStatusString, "  ", " ")
6187      LinearSolveStatusString = Replace(LinearSolveStatusString, "  ", " ")

6188      If solutionExpected Then
              ' We read in whatever solution CBC returned
6189          Application.StatusBar = "OpenSolver: Loading Solution... " & LinearSolveStatusString
              ' Zero the current decision variables
6190          s.AdjustableCells.Value2 = 0
              ' Faster code; put a zero into first adjustable cell, and copy it to all the adjustable cells
              ' AdjustableCells.Cells(0, 0).Value = 0
              ' AdjustableCells.Cells(0, 0).Copy
              ' AdjustableCells.PasteSpecial xlPasteValues
          
              ' Read in the Solution File
              ' This gives the non-zero? variable values
              ' Lines like:       0 AZ70                  15                      0
              ' ...being? : Index Name Value ReducedCost
              Dim Line As String, SplitLine() As String, index As Double, NameValue As String, value As Double, CBCConstraintIndex As Long
6191          If s.bGetDuals Then
                  Dim j As Long, row As Long, i As Long
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
6203                      SplitLine = Split(Line, " ")    ' 0 indexed; item 0 is the variable index
                          ' Skip over the blank items in the split (multiple delimiters give multiple items), getting the real items
6204                      i = 0
6205                      While SplitLine(i) = ""
6206                          i = i + 1
6207                      Wend
                          ' In the case of LpStatusInfeasible, we can get lines that start **. We strip the **
6208                      If SplitLine(i) = "**" Then i = i + 1
6209                      While SplitLine(i) = ""
6210                          i = i + 1
6211                      Wend
                          ' Get and check the index of the row
6212                      If Val(SplitLine(i)) <> CBCConstraintIndex Then
6213                          errorString = "While reading the CBC solution file, OpenSolver found an unexpected constraint row."
6214                          GoTo ExitSub
6215                      End If
6216                      i = i + 1
6217                      While SplitLine(i) = ""
6218                          i = i + 1
6219                      Wend
                          ' Get the constraint name; we don't use this
6220                      NameValue = SplitLine(i)
6221                      i = i + 1
6222                      While SplitLine(i) = ""
6223                          i = i + 1
6224                      Wend
6225                      s.FinalValueP(j) = SplitLine(i)
                          ' Skip the constraint LHS value - we don't need this
6226                      i = i + 1
6227                      While SplitLine(i) = ""
6228                          i = i + 1
6229                      Wend
                          ' Get the dual value
6230                      If s.ObjectiveSense = MaximiseObjective Then
6231                          value = -1 * Val(SplitLine(i))
                              'rConstraintList.Cells(row, 2).Value2 = Value
6232                      Else
6233                          value = Val(SplitLine(i))
                              'rConstraintList.Cells(row, 2).Value2 = Value
6234                      End If
6235                      s.ShadowPriceP(j) = value
6236                      If InStr(s.ShadowPriceP(j), "E-16") Then
6237                          s.ShadowPriceP(j) = "0"
6238                      End If
6239                      CBCConstraintIndex = CBCConstraintIndex + 1
6240                      j = j + 1
6241                  End If
6242              Next row
6243              ReadSensitivityData_CBC SolutionFilePathName, s
6244          End If
            
              ' Now we read in the decision variable values
6245          j = 1
6246          While Not EOF(1)
6247              Line Input #1, Line
6248              SplitLine = Split(Line, " ")    ' 0 indexed; item 0 is the variable index
                  ' Skip over the blank items in the split (multiple delimiters give multiple items), getting the real items
6249              i = 0
6250              While SplitLine(i) = ""
6251                  i = i + 1
6252              Wend
                  ' In the case of LpStatusInfeasible, we can get lines that start **. We strip the **
6253              If SplitLine(i) = "**" Then i = i + 1
6254              While SplitLine(i) = ""
6255                  i = i + 1
6256              Wend
                  ' Get the index of the variable
6257              index = Val(SplitLine(i))
6258              i = i + 1
6259              While SplitLine(i) = ""
6260                  i = i + 1
6261              Wend
                  ' Get the variable name, stripping any leading "_"
6262              NameValue = SplitLine(i)
6263              If left(NameValue, 1) = "_" Then NameValue = Mid(NameValue, 2) ' Strip any _ character added to make a valid name
6264              i = i + 1
6265              While SplitLine(i) = ""
6266                  i = i + 1
6267              Wend
6268              s.FinalVarValueP(j) = Val(SplitLine(i))

                  'Write to the sheet containing the decision variables (which may not be the active sheet)
                  'Value assigned to Value2 must be in US locale
6269              s.AdjustableCells.Worksheet.Range(NameValue).Value2 = Val(SplitLine(i)) 'ConvertFromCurrentLocale(s.FinalVarValueP(j))
                 
                  'ConvertFullLPFileVarNameToRange(name, AdjCellsSheetIndex).Value2 = Value
6270              If s.bGetDuals Then
6271                  i = i + 1
6272                  While SplitLine(i) = ""
6273                      i = i + 1
6274                  Wend
6275                  If s.ObjectiveSense = MaximiseObjective Then
6276                      value = -1 * Val(SplitLine(i))
6277                  Else
6278                      value = Val(SplitLine(i))
6279                  End If
6280                  s.ReducedCostsP(j) = str(value)
6281                  If InStr(s.ReducedCostsP(j), "E-16") Then
6282                      s.ReducedCostsP(j) = "0"
6283                  End If
6284                  s.VarCellP(j) = NameValue
6285              End If
6286              j = j + 1
6287          Wend

6288      End If
6289      Close #1
6290      ReadModel_CBC = True
ExitSub:
6291      s.LinearSolveStatusString = LinearSolveStatusString
          
6292      Exit Function
          
ErrHandler:
6293      Close #1
6294      Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Function

Sub ReadSensitivityData_CBC(SolutionFilePathName As String, s As COpenSolver)
          'Reads the two files with the limits on the bounds of shadow prices and reduced costs

          Dim RangeFilePathName As String, Stuff(5) As String, index2 As Long
          Dim Line As String, row As Long, j As Long, i As Long
          
          'Find the ranges on the constraints
6295      RangeFilePathName = left(SolutionFilePathName, InStrRev(SolutionFilePathName, Application.PathSeparator)) & RHSRangesFile_CBC
6296      On Error GoTo ErrHandler
6297      Open RangeFilePathName For Input As 2 ' supply path with filename
6298      Line Input #2, Line 'Dont want first line
6299      j = 1
6300      While Not EOF(2)
6301          Line Input #2, Line
6302          For i = 1 To 5
6303              index2 = InStr(Line, ",")
6304              Stuff(i) = left(Line, index2 - 1)
6305              If Stuff(i) = "1e-007" Then
6306                  Stuff(i) = "0"
6307              ElseIf InStr(Stuff(i), "E-16") Then
6308                  Stuff(i) = "0"
6309              End If
6310              Line = Mid(Line, index2 + 1)
6311          Next i
6312          s.IncreaseConP(j) = Stuff(3)
6313          s.DecreaseConP(j) = Stuff(5)
6314          j = j + 1
6315      Wend
6316      Close 2
          
6317      j = 1
          'Find the ranges on the variables
6318      RangeFilePathName = left(SolutionFilePathName, InStrRev(SolutionFilePathName, Application.PathSeparator)) & CostRangesFile_CBC
6319      Open RangeFilePathName For Input As 2 ' supply path with filename
6320      Line Input #2, Line 'Dont want first line
6321      row = s.NumRows + 2
6322      While Not EOF(2)
6323          Line Input #2, Line
6324          For i = 1 To 5
6325              index2 = InStr(Line, ",")
6326              Stuff(i) = left(Line, index2 - 1)
6327              If Stuff(i) = "1e-007" Then
6328                  Stuff(i) = "0"
6329              ElseIf InStr(Stuff(i), "E-16") Then
6330                  Stuff(i) = "0"
6331              End If
6332              Line = Mid(Line, index2 + 1)
6333          Next i
6334          If s.ObjectiveSense = MaximiseObjective Then
6335              s.IncreaseVarP(j) = Stuff(5)
6336              s.DecreaseVarP(j) = Stuff(3)
6337          Else
6338              s.IncreaseVarP(j) = Stuff(3)
6339              s.DecreaseVarP(j) = Stuff(5)
6340          End If
6341          j = j + 1
6342      Wend
6343      Close 2
6344      Exit Sub
          
ErrHandler:
6345      Close #2
6346      Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Sub

Sub LaunchCommandLine_CBC()
      ' Open the CBC solver with our last model loaded.
          ' If we have a worksheet open with a model, then we pass the solver options (max runtime etc) from this model to CBC. Otherwise, we don't pass any options.
6347      On Error GoTo errorHandler
          Dim errorPrefix  As String
6348      errorPrefix = ""
            
          Dim WorksheetAvailable As Boolean
6349      WorksheetAvailable = CheckWorksheetAvailable(SuppressDialogs:=True)
          
          Dim SolverPath As String, errorString As String
6350      If Not SolverAvailable_CBC(SolverPath, errorString) Then
6351          Err.Raise OpenSolver_CBCMissingError, Description:=errorString
6352      End If
          
          Dim ModelFilePathName As String
6353      ModelFilePathName = ModelFilePath("CBC")
          
          ' Get all the options that we pass to CBC when we solve the problem and pass them here as well
          ' Get the Solver Options, stored in named ranges with values such as "=0.12"
          ' Because these are NAMEs, they are always in English, not the local language, so get their value using Val
          Dim SolveOptions As SolveOptionsType, SolveOptionsString As String
6354      If WorksheetAvailable Then
6355          GetSolveOptions "'" & Replace(ActiveSheet.Name, "'", "''") & "'!", SolveOptions, errorString ' NB: We have to double any ' when we quote the sheet name
6356          If errorString = "" Then
6357             SolveOptionsString = " -ratioGap " & CStr(SolveOptions.Tolerance) & " -seconds " & CStr(SolveOptions.maxTime)
6358          End If
6359      End If
          
          Dim ExtraParametersString As String
6360      If WorksheetAvailable Then
6361          ExtraParametersString = GetExtraParameters_CBC(ActiveSheet, errorString)
6362          If errorString <> "" Then ExtraParametersString = ""
6363      End If
             
          Dim CBCRunString As String
6364      CBCRunString = " -directory " & QuotePath(ConvertHfsPath(left(GetTempFolder, Len(GetTempFolder) - 1))) _
                           & " -import " & QuotePath(ConvertHfsPath(ModelFilePathName)) _
                           & SolveOptionsString _
                           & ExtraParametersString _
                           & " -" ' Force CBC to accept commands from the command line
6365      RunExternalCommand QuotePath(ConvertHfsPath(SolverPath)) & CBCRunString, "", SW_SHOWNORMAL, False 'OSSolveSync

ExitSub:
6366      Exit Sub
errorHandler:
6367      MsgBox "OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source, , "OpenSolver Code Error"
6368      Resume ExitSub
End Sub
