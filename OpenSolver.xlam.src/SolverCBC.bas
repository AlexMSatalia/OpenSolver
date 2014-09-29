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
    ScriptFilePath_CBC = GetTempFilePath(SolverScript_CBC)
End Function

Function SolutionFilePath_CBC() As String
    SolutionFilePath_CBC = GetTempFilePath(SolutionFile_CBC)
End Function

Function CostRangesFilePath_CBC() As String
    CostRangesFilePath_CBC = GetTempFilePath(CostRangesFile_CBC)
End Function

Function RHSRangesFilePath_CBC() As String
    RHSRangesFilePath_CBC = GetTempFilePath(RHSRangesFile_CBC)
End Function

Sub CleanFiles_CBC(errorPrefix As String)
    ' Solution file
    DeleteFileAndVerify SolutionFilePath_CBC(), errorPrefix, "Unable to delete the CBC solver solution file: " & SolutionFilePath_CBC()
    ' Cost Range file
    DeleteFileAndVerify CostRangesFilePath_CBC(), errorPrefix, "Unable to delete the CBC solver sensitivity data file: " & CostRangesFilePath_CBC()
    ' RHS Range file
    DeleteFileAndVerify RHSRangesFilePath_CBC(), errorPrefix, "Unable to delete the CBC solver sensitivity data file: " & RHSRangesFilePath_CBC()
    ' Script file
    DeleteFileAndVerify ScriptFilePath_CBC(), errorPrefix, "Unable to delete the CBC solver script file: " & ScriptFilePath_CBC()
End Sub

Function About_CBC() As String
' Return string for "About" form
    Dim SolverPath As String, errorString As String
    If Not SolverAvailable_CBC(SolverPath, errorString) Then
        About_CBC = errorString
        Exit Function
    End If
    ' Assemble version info
    About_CBC = "CBC " & SolverBitness_CBC & "-bit" & _
                     " v" & SolverVersion_CBC & _
                     " at " & MakeSpacesNonBreaking(SolverPath)
End Function

Function SolverFilePath_CBC(errorString As String) As String
    SolverFilePath_CBC = SolverFilePath_Default("CBC", errorString)
End Function

Function SolverAvailable_CBC(Optional SolverPath As String, Optional errorString As String) As Boolean
' Returns true if CBC is available and sets SolverPath
    SolverPath = SolverFilePath_CBC(errorString)
    If SolverPath = "" Then
        SolverAvailable_CBC = False
    Else
        SolverAvailable_CBC = True

#If Mac Then
        ' Make sure cbc is executable on Mac
        system ("chmod +x " & ConvertHfsPath(SolverPath))
#End If
    
    End If
End Function

Function SolverVersion_CBC() As String
' Get CBC version by running 'cbc -exit' at command line
    Dim SolverPath As String
    If Not SolverAvailable_CBC(SolverPath) Then
        SolverVersion_CBC = ""
        Exit Function
    End If
    
    ' Set up cbc to write version info to text file
    Dim logFile As String
    logFile = GetTempFilePath("cbcversion.txt")
    If FileOrDirExists(logFile) Then Kill logFile
    
    Dim RunPath As String, FileContents As String
    RunPath = ScriptFilePath_CBC()
    If FileOrDirExists(RunPath) Then Kill RunPath
    FileContents = QuotePath(ConvertHfsPath(SolverPath)) & " -exit" & " > " & QuotePath(ConvertHfsPath(logFile))
    CreateScriptFile RunPath, FileContents
    
    ' Run cbc
    Dim completed As Boolean
    completed = OSSolveSync(ConvertHfsPath(RunPath), "", "", "", SW_HIDE, True)
    
    ' Read version info back from output file
    Dim Line As String
    If FileOrDirExists(logFile) Then
        On Error GoTo ErrHandler
        Open logFile For Input As 1
        Line Input #1, Line
        Line Input #1, Line
        Close #1
        SolverVersion_CBC = right(Line, Len(Line) - 9)
        SolverVersion_CBC = left(SolverVersion_CBC, Len(SolverVersion_CBC) - 1)
    Else
        SolverVersion_CBC = ""
    End If
    Exit Function
    
ErrHandler:
    Close #1
    Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Function

Function SolverBitness_CBC() As String
' Get Bitness of CBC solver
    Dim SolverPath As String
    If Not SolverAvailable_CBC(SolverPath) Then
        SolverBitness_CBC = ""
        Exit Function
    End If
    
    ' All Macs are 64-bit so we only provide 64-bit binaries
#If Mac Then
    SolverBitness_CBC = "64"
#Else
    If right(SolverPath, 10) = "64\cbc.exe" Then
        SolverBitness_CBC = "64"
    Else
        SolverBitness_CBC = "32"
    End If
#End If
End Function

Function GetExtraParameters_CBC(sheet As Worksheet, errorString As String) As String
    ' The user can define a set of parameters they want to pass to CBC; this gets them as a string
    ' Note: The named range MUST be on the current sheet
    Dim CBCParametersRange As Range, CBCExtraParametersString As String, i As Long
    errorString = ""
    If GetNamedRangeIfExistsOnSheet(sheet, "OpenSolver_CBCParameters", CBCParametersRange) Then
        If CBCParametersRange.Columns.Count <> 2 Then
            errorString = "The range OpenSolver_CBCParameters must be a two-column table."
            Exit Function
        End If
        For i = 1 To CBCParametersRange.Rows.Count
            Dim ParamName As String, ParamValue As String
            ParamName = Trim(CBCParametersRange.Cells(i, 1))
            If ParamName <> "" Then
                If left(ParamName, 1) <> "-" Then ParamName = "-" & ParamName
                ParamValue = Trim(CBCParametersRange.Cells(i, 2))
                CBCExtraParametersString = CBCExtraParametersString & " " & ParamName & " " & ConvertFromCurrentLocale(ParamValue)
            End If
        Next i
    End If
    GetExtraParameters_CBC = CBCExtraParametersString
End Function

Function CreateSolveScript_CBC(SolutionFilePathName As String, ExtraParametersString As String, SolveOptions As SolveOptionsType, s As COpenSolver) As String
    Dim CommandLineRunString As String, PrintingOptionString As String
    ' have to split up the command line as excel couldn't have a string longer than 255 characters??
    CommandLineRunString = " -directory " & QuotePath(ConvertHfsPath(left(GetTempFolder, Len(GetTempFolder) - 1))) _
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
    PrintingOptionString = IIf(s.bGetDuals, " -printingOptions rhs  -solution " & RHSRangesFile_CBC & " -printingOptions objective -solution " & CostRangesFile_CBC, "")
                  
    Dim scriptFile As String, scriptFileContents As String
    scriptFile = ScriptFilePath_CBC()
    scriptFileContents = QuotePath(ConvertHfsPath(s.ExternalSolverPathName)) & CommandLineRunString & PrintingOptionString
    CreateScriptFile scriptFile, scriptFileContents
    
    CreateSolveScript_CBC = scriptFile
End Function

Function ReadModel_CBC(SolutionFilePathName As String, errorString As String, s As COpenSolver) As Boolean
          Dim LinearSolveStatusString As String
20770     ReadModel_CBC = False
          On Error GoTo ErrHandler
20780     Open SolutionFilePathName For Input As 1 ' supply path with filename
20790     Line Input #1, LinearSolveStatusString  ' Optimal - objective value              22
          ' Line Input #1, junk ' get rest of line
          Dim solutionExpected As Boolean
20800     solutionExpected = True
20810     If LinearSolveStatusString Like "Optimal*" Then
20820         s.SolveStatus = OpenSolverResult.Optimal
20830         s.SolveStatusString = "Optimal"
20840         s.LinearSolveStatus = LinearSolveResult.Optimal
              '
20850     ElseIf LinearSolveStatusString Like "Infeasible*" Then
20860         s.SolveStatus = OpenSolverResult.Infeasible
20870         s.SolveStatusString = "No Feasible Solution"
20880         s.LinearSolveStatus = LinearSolveResult.Infeasible
              '
20890     ElseIf LinearSolveStatusString Like "Integer infeasible*" Then
20900         s.SolveStatus = OpenSolverResult.Infeasible
20910         s.SolveStatusString = "No Feasible Integer Solution"
20920         s.LinearSolveStatus = LinearSolveResult.IntegerInfeasible
              '
20930     ElseIf LinearSolveStatusString Like "Unbounded*" Then
20940         s.SolveStatus = OpenSolverResult.Unbounded
20950         s.SolveStatusString = "No Solution Found (Unbounded)"
20960         s.LinearSolveStatus = LinearSolveResult.Unbounded
20970         solutionExpected = False
              '
20980     ElseIf LinearSolveStatusString Like "Stopped on time *" Then ' Stopped on iterations or time
20990         s.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
21000         s.SolveStatusString = "Stopped on Time Limit"
21010         If LinearSolveStatusString Like "*(no integer solution - continuous used)*" Then
21020             s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
21030         End If
21040         s.LinearSolveStatus = LinearSolveResult.SolveStopped
              '
21050     ElseIf LinearSolveStatusString Like "Stopped on iterations*" Then ' Stopped on iterations or time
21060         s.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
21070         s.SolveStatusString = "Stopped on Iteration Limit"
21080         If LinearSolveStatusString Like "*(no integer solution - continuous used)*" Then
21090             s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
21100         End If
21110         s.LinearSolveStatus = LinearSolveResult.SolveStopped
              '
21120     ElseIf LinearSolveStatusString Like "Stopped on difficulties*" Then ' Stopped on iterations or time
21130         s.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
21140         s.SolveStatusString = "Stopped on CBC difficulties"
21150         If LinearSolveStatusString Like "*(no integer solution - continuous used)*" Then
21160             s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
21170         End If
21180         s.LinearSolveStatus = LinearSolveResult.SolveStopped
              '
21190     ElseIf LinearSolveStatusString Like "Stopped on ctrl-c*" Then ' Stopped on iterations or time
21200         s.SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
21210         s.SolveStatusString = "Stopped on Ctrl-C"
21220         If LinearSolveStatusString Like "*(no integer solution - continuous used)*" Then
21230             s.SolveStatusString = s.SolveStatusString & ": No integer solution found. Fractional solution returned."
21240         End If
21250         s.LinearSolveStatus = LinearSolveResult.SolveStopped
              '
21260     ElseIf LinearSolveStatusString Like "Status unknown*" Then
21270         errorString = "CBC solver did not solve the problem, suggesting there was an error in the CBC input parameters. The response was: " & vbCrLf _
               & LinearSolveStatusString _
               & vbCrLf & "The CBC command line can be found at:" _
               & vbCrLf & ScriptFilePath_CBC()
21280         GoTo ExitSub
21290     Else
21300         errorString = "The response from the CBC solver is not recognised. The response was: " & LinearSolveStatusString
21310         GoTo ExitSub
21320     End If
          
          ' Remove the double spaces from LinearSolveStatusString
21330     LinearSolveStatusString = Replace(LinearSolveStatusString, "    ", " ")
21340     LinearSolveStatusString = Replace(LinearSolveStatusString, "   ", " ")
21350     LinearSolveStatusString = Replace(LinearSolveStatusString, "  ", " ")
21360     LinearSolveStatusString = Replace(LinearSolveStatusString, "  ", " ")
21370     LinearSolveStatusString = Replace(LinearSolveStatusString, "  ", " ")
21380     LinearSolveStatusString = Replace(LinearSolveStatusString, "  ", " ")

21390     If solutionExpected Then
              ' We read in whatever solution CBC returned
21400         Application.StatusBar = "OpenSolver: Loading Solution... " & LinearSolveStatusString
              ' Zero the current decision variables
21410         s.AdjustableCells.Value2 = 0
              ' Faster code; put a zero into first adjustable cell, and copy it to all the adjustable cells
              ' AdjustableCells.Cells(0, 0).Value = 0
              ' AdjustableCells.Cells(0, 0).Copy
              ' AdjustableCells.PasteSpecial xlPasteValues
          
              ' Read in the Solution File
              ' This gives the non-zero? variable values
              ' Lines like:       0 AZ70                  15                      0
              ' ...being? : Index Name Value ReducedCost
              Dim Line As String, SplitLine() As String, index As Double, NameValue As String, value As Double, CBCConstraintIndex As Long
21420         If s.bGetDuals Then
                  Dim j As Long, row As Long, i As Long
                  CBCConstraintIndex = 0
                  
                  ' Throw away first constraint if it was from a seek objective model
21430             If s.ObjectiveSense = TargetObjective Then
                      Line Input #1, Line
                      CBCConstraintIndex = CBCConstraintIndex + 1
                  End If

21450             j = 1
21470             For row = 1 To s.NumRows
21480                 If s.GetSparseACount(row) = 0 Then
                          ' This constraint was not written to the model, as it had no coefficients. Just ignore it.
21490                     s.rConstraintList.Cells(row, 2).ClearContents
21500                 Else
21510                     Line Input #1, Line
21520                     SplitLine = Split(Line, " ")    ' 0 indexed; item 0 is the variable index
                          ' Skip over the blank items in the split (multiple delimiters give multiple items), getting the real items
21530                     i = 0
21540                     While SplitLine(i) = ""
21550                         i = i + 1
21560                     Wend
                          ' In the case of LpStatusInfeasible, we can get lines that start **. We strip the **
21570                     If SplitLine(i) = "**" Then i = i + 1
21580                     While SplitLine(i) = ""
21590                         i = i + 1
21600                     Wend
                          ' Get and check the index of the row
21610                     If Val(SplitLine(i)) <> CBCConstraintIndex Then
21620                         errorString = "While reading the CBC solution file, OpenSolver found an unexpected constraint row."
21630                         GoTo ExitSub
21640                     End If
21650                     i = i + 1
21660                     While SplitLine(i) = ""
21670                         i = i + 1
21680                     Wend
                          ' Get the constraint name; we don't use this
21690                     NameValue = SplitLine(i)
21700                     i = i + 1
21710                     While SplitLine(i) = ""
21720                         i = i + 1
21730                     Wend
21740                     s.FinalValueP(j) = SplitLine(i)
                          ' Skip the constraint LHS value - we don't need this
21750                     i = i + 1
21760                     While SplitLine(i) = ""
21770                         i = i + 1
21780                     Wend
                          ' Get the dual value
21790                     If s.ObjectiveSense = MaximiseObjective Then
21800                         value = -1 * Val(SplitLine(i))
                              'rConstraintList.Cells(row, 2).Value2 = Value
21810                     Else
21820                         value = Val(SplitLine(i))
                              'rConstraintList.Cells(row, 2).Value2 = Value
21830                     End If
21840                     s.ShadowPriceP(j) = value
21850                     If InStr(s.ShadowPriceP(j), "E-16") Then
21860                         s.ShadowPriceP(j) = "0"
21870                     End If
21880                     CBCConstraintIndex = CBCConstraintIndex + 1
21890                     j = j + 1
21900                 End If
21910             Next row
21920             ReadSensitivityData_CBC SolutionFilePathName, s
21930         End If
            
              ' Now we read in the decision variable values
21940         j = 1
21950         While Not EOF(1)
21960             Line Input #1, Line
21970             SplitLine = Split(Line, " ")    ' 0 indexed; item 0 is the variable index
                  ' Skip over the blank items in the split (multiple delimiters give multiple items), getting the real items
21980             i = 0
21990             While SplitLine(i) = ""
22000                 i = i + 1
22010             Wend
                  ' In the case of LpStatusInfeasible, we can get lines that start **. We strip the **
22020             If SplitLine(i) = "**" Then i = i + 1
22030             While SplitLine(i) = ""
22040                 i = i + 1
22050             Wend
                  ' Get the index of the variable
22060             index = Val(SplitLine(i))
22070             i = i + 1
22080             While SplitLine(i) = ""
22090                 i = i + 1
22100             Wend
                  ' Get the variable name, stripping any leading "_"
22110             NameValue = SplitLine(i)
22120             If left(NameValue, 1) = "_" Then NameValue = Mid(NameValue, 2) ' Strip any _ character added to make a valid name
22130             i = i + 1
22140             While SplitLine(i) = ""
22150                 i = i + 1
22160             Wend
22180             s.FinalVarValueP(j) = Val(SplitLine(i))
                  'Write to the sheet containing the decision variables (which may not be the active sheet)
                  'Value assigned to Value2 must be in US locale
22190             s.AdjustableCells.Worksheet.Range(NameValue).Value2 = ConvertFromCurrentLocale(s.FinalVarValueP(j))
                 
                  'ConvertFullLPFileVarNameToRange(name, AdjCellsSheetIndex).Value2 = Value
22200             If s.bGetDuals Then
22210                 i = i + 1
22220                 While SplitLine(i) = ""
22230                     i = i + 1
22240                 Wend
22250                 If s.ObjectiveSense = MaximiseObjective Then
22260                     value = -1 * Val(SplitLine(i))
22270                 Else
22280                     value = Val(SplitLine(i))
22290                 End If
22320                 s.ReducedCostsP(j) = str(value)
22330                 If InStr(s.ReducedCostsP(j), "E-16") Then
22340                     s.ReducedCostsP(j) = "0"
22350                 End If
22360                 s.VarCellP(j) = NameValue
22370             End If
22380             j = j + 1
22390         Wend

22400     End If
22410     Close #1
22420     ReadModel_CBC = True
ExitSub:
          s.LinearSolveStatusString = LinearSolveStatusString
          
          Exit Function
    
ErrHandler:
    Close #1
    Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Function

Sub ReadSensitivityData_CBC(SolutionFilePathName As String, s As COpenSolver)
          'Reads the two files with the limits on the bounds of shadow prices and reduced costs

          Dim RangeFilePathName As String, Stuff(5) As String, index2 As Long
          Dim Line As String, row As Long, j As Long, i As Long
          
          'Find the ranges on the constraints
          RangeFilePathName = left(SolutionFilePathName, InStrRev(SolutionFilePathName, PathDelimeter)) & RHSRangesFile_CBC
22460     On Error GoTo ErrHandler
          Open RangeFilePathName For Input As 2 ' supply path with filename
22470     Line Input #2, Line 'Dont want first line
22480     j = 1
22490     While Not EOF(2)
22500         Line Input #2, Line
22510         For i = 1 To 5
22520             index2 = InStr(Line, ",")
22530             Stuff(i) = left(Line, index2 - 1)
22540             If Stuff(i) = "1e-007" Then
22550                 Stuff(i) = "0"
22560             ElseIf InStr(Stuff(i), "E-16") Then
22570                 Stuff(i) = "0"
22580             End If
22590             Line = Mid(Line, index2 + 1)
22600         Next i
22610         s.IncreaseConP(j) = Stuff(3)
22620         s.DecreaseConP(j) = Stuff(5)
22630         j = j + 1
22640     Wend
22650     Close 2
          
22660     j = 1
          'Find the ranges on the variables
22670     RangeFilePathName = left(SolutionFilePathName, InStrRev(SolutionFilePathName, PathDelimeter)) & CostRangesFile_CBC
22680     Open RangeFilePathName For Input As 2 ' supply path with filename
22690     Line Input #2, Line 'Dont want first line
22700     row = s.NumRows + 2
22710     While Not EOF(2)
22740         Line Input #2, Line
22750         For i = 1 To 5
22760             index2 = InStr(Line, ",")
22770             Stuff(i) = left(Line, index2 - 1)
22780             If Stuff(i) = "1e-007" Then
22790                 Stuff(i) = "0"
22800             ElseIf InStr(Stuff(i), "E-16") Then
22810                 Stuff(i) = "0"
22820             End If
22830             Line = Mid(Line, index2 + 1)
22840         Next i
22850         If s.ObjectiveSense = MaximiseObjective Then
22860             s.IncreaseVarP(j) = Stuff(5)
22870             s.DecreaseVarP(j) = Stuff(3)
22880         Else
22890             s.IncreaseVarP(j) = Stuff(3)
22900             s.DecreaseVarP(j) = Stuff(5)
22910         End If
22920         j = j + 1
22930     Wend
22940     Close 2
          Exit Sub
    
ErrHandler:
    Close #2
    Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Sub

Sub LaunchCommandLine_CBC()
' Open the CBC solver with our last model loaded.
          ' If we have a worksheet open with a model, then we pass the solver options (max runtime etc) from this model to CBC. Otherwise, we don't pass any options.
27890     On Error GoTo errorHandler
          Dim errorPrefix  As String
27900     errorPrefix = ""
            
          Dim WorksheetAvailable As Boolean
27910     WorksheetAvailable = CheckWorksheetAvailable(SuppressDialogs:=True)
          
          Dim SolverPath As String, errorString As String
27930     If Not SolverAvailable_CBC(SolverPath, errorString) Then
              Err.Raise OpenSolver_CBCMissingError, Description:=errorString
          End If
          
          Dim ModelFilePathName As String
27950     ModelFilePathName = ModelFilePath("CBC")
          
          ' Get all the options that we pass to CBC when we solve the problem and pass them here as well
          ' Get the Solver Options, stored in named ranges with values such as "=0.12"
          ' Because these are NAMEs, they are always in English, not the local language, so get their value using Val
          Dim SolveOptions As SolveOptionsType, SolveOptionsString As String
27960     If WorksheetAvailable Then
27970         GetSolveOptions "'" & Replace(ActiveSheet.Name, "'", "''") & "'!", SolveOptions, errorString ' NB: We have to double any ' when we quote the sheet name
27980         If errorString = "" Then
27990            SolveOptionsString = " -ratioGap " & CStr(SolveOptions.Tolerance) & " -seconds " & CStr(SolveOptions.maxTime)
28000         End If
28010     End If
          
          Dim ExtraParametersString As String
28020     If WorksheetAvailable Then
28030         ExtraParametersString = GetExtraParameters_CBC(ActiveSheet, errorString)
28040         If errorString <> "" Then ExtraParametersString = ""
28050     End If
             
          Dim CBCRunString As String
28060     CBCRunString = " -directory " & QuotePath(ConvertHfsPath(left(GetTempFolder, Len(GetTempFolder) - 1))) _
                           & " -import " & QuotePath(ConvertHfsPath(ModelFilePathName)) _
                           & SolveOptionsString _
                           & ExtraParametersString _
                           & " -" ' Force CBC to accept commands from the command line
28070     OSSolveSync QuotePath(ConvertHfsPath(SolverPath)), CBCRunString, "", "", SW_SHOWNORMAL, False

ExitSub:
28080     Exit Sub
errorHandler:
28090     MsgBox "OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source, , "OpenSolver Code Error"
28100     Resume ExitSub
End Sub
