Attribute VB_Name = "OpenSolverUtils"
Option Explicit

' For sleep
#If Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Sub uSleep Lib "libc.dylib" Alias "usleep" (ByVal Microseconds As Long)
    #Else
        Private Declare Sub uSleep Lib "libc.dylib" Alias "usleep" (ByVal Microseconds As Long)
    #End If
#Else
    #If VBA7 Then
        Public Declare PtrSafe Sub mSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
    #Else
        Public Declare Sub mSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
    #End If
#End If  ' sleep

' For OS version
#If Win32 Then
    #If VBA7 Then
        Type OSVERSIONINFO
            dwOSVersionInfoSize As Long
            dwMajorVersion As Long
            dwMinorVersion As Long
            dwBuildNumber As Long
            dwPlatformId As Long
            szCSDVersion As String * 128
        End Type
        Private Declare PtrSafe Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
    #Else
        Type OSVERSIONINFO
            dwOSVersionInfoSize As Long
            dwMajorVersion As Long
            dwMinorVersion As Long
            dwBuildNumber As Long
            dwPlatformId As Long
            szCSDVersion As String * 128
        End Type
        Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
    #End If
#End If  ' OS version

'For OpenURL - Code Courtesy of Dev Ashish
#If Win32 Then
    #If VBA7 Then
        Private Declare PtrSafe Function apiShellExecute Lib "shell32.dll" _
            Alias "ShellExecuteA" _
            (ByVal hwnd As LongPtr, _
            ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) _
            As Long
    #Else
        Private Declare Function apiShellExecute Lib "shell32.dll" _
            Alias "ShellExecuteA" _
            (ByVal hwnd As Long, _
            ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) _
            As Long
    #End If
    
    Private Const WIN_NORMAL = 1         'Open Normal
    Private Const WIN_MAX = 2            'Open Maximized
    Private Const WIN_MIN = 3            'Open Minimized
    
    Private Const ERROR_SUCCESS = 32&
    Private Const ERROR_NO_ASSOC = 31&
    Private Const ERROR_OUT_OF_MEM = 0&
    Private Const ERROR_FILE_NOT_FOUND = 2&
    Private Const ERROR_PATH_NOT_FOUND = 3&
    Private Const ERROR_BAD_FORMAT = 11&
#End If ' OpenURL

#If Mac Then
    Public Const IsMac As Boolean = True
#Else
    Public Const IsMac As Boolean = False
#End If

' Complete the sleep interface, we just want a public mSleep function on both platforms
#If Mac Then
    Public Sub mSleep(Milliseconds As Long)
        uSleep Milliseconds * 1000
    End Sub
#End If  ' Mac

Function RemoveSheetNameFromString(s As String, sheet As Worksheet) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          ' Try with workbook name first
          Dim sheetName As String
          sheetName = "'[" & sheet.Parent.Name & "]" & Mid(EscapeSheetName(sheet, True), 2)
          If InStr(s, sheetName) Then
              RemoveSheetNameFromString = Replace(s, sheetName, vbNullString)
              GoTo ExitFunction
          End If
          
280       sheetName = EscapeSheetName(sheet)
281       If InStr(s, sheetName) Then
282           RemoveSheetNameFromString = Replace(s, sheetName, vbNullString)
283           GoTo ExitFunction
284       End If

290       RemoveSheetNameFromString = s

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverUtils", "RemoveSheetNameFromString") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function RemoveActiveSheetNameFromString(s As String) As String
          RemoveActiveSheetNameFromString = RemoveSheetNameFromString(s, ActiveSheet)
End Function

' Removes a "\n" character from the end of a string
Function StripTrailingNewline(Block As String) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          If Right(Block, Len(vbNewLine)) = vbNewLine Then
              Block = Left(Block, Len(Block) - Len(vbNewLine))
          End If
          StripTrailingNewline = Block

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverUtils", "StripTrailingNewline") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function StripWorksheetNameAndDollars(s As String, currentSheet As Worksheet) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          ' Remove the current worksheet name from a formula, along with any $
586       StripWorksheetNameAndDollars = RemoveSheetNameFromString(s, currentSheet)
588       StripWorksheetNameAndDollars = Replace(StripWorksheetNameAndDollars, "$", vbNullString)

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverUtils", "StripWorksheetNameAndDollars") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function EscapeSheetName(sheet As Worksheet, Optional ForceQuotes As Boolean = False) As String
1         EscapeSheetName = sheet.Name
          
          Dim SpecialChar As Variant, NeedsEscaping As Boolean
2         NeedsEscaping = False
3         For Each SpecialChar In Array("'", "!", "(", ")", "+", "-", " ")
4             If InStr(EscapeSheetName, SpecialChar) Then
5                 NeedsEscaping = True
6                 Exit For
7             End If
8         Next SpecialChar

9         If NeedsEscaping Then EscapeSheetName = Replace(EscapeSheetName, "'", "''")
10        If ForceQuotes Or NeedsEscaping Then EscapeSheetName = "'" & EscapeSheetName & "'"
          
11        EscapeSheetName = EscapeSheetName & "!"
End Function

Function ConvertFromCurrentLocale(ByVal s As String) As String
      ' Convert a formula or a range from the current locale into US locale
1               ConvertFromCurrentLocale = ConvertLocale(s, True)
End Function

Function ConvertToCurrentLocale(ByVal s As String) As String
      ' Convert a formula or a range from US locale into the current locale
1               ConvertToCurrentLocale = ConvertLocale(s, False)
End Function

Private Function ConvertLocale(ByVal s As String, ConvertToUS As Boolean) As String
      ' Convert strings between locales
      ' This will add a leading "=" if its not already there
      ' A blank string is returned if any errors occur
      ' This works by putting the expression into cell A1 on Sheet1 of the add-in!

          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' Static to track whether we can skip or not
          ' 0 = Haven't checked yet
          ' 1 = Can't skip
          ' 2 = Can skip
          Static SkipConvert As Long
          
3         If SkipConvert < 1 Then
4             SkipConvert = 1  ' Don't check again
              ' If we are in an english version of Excel and a known locale, we can skip the conversion
5             If Application.International(xlCountryCode) = 1 Then ' English language excel
                  Dim Country As Long
6                 Country = Application.International(xlCountrySetting)
7                 SkipConvert = 2
                  
8                 Select Case Country
                  Case 1  ' US
9                 Case 44 ' UK
10                Case 61 ' Australia
11                Case 64 ' NZ
12                Case Else: SkipConvert = 1
13                End Select
14            End If
15        End If
          
16        If SkipConvert = 2 Then
17            ConvertLocale = s
18            Exit Function
19        End If
              

          ' We turn off calculation & hide alerts as we don't want Excel popping up dialogs asking for references to other sheets
          Dim oldCalculation As Long
20        oldCalculation = Application.Calculation
          Dim oldDisplayAlerts As Boolean
21        oldDisplayAlerts = Application.DisplayAlerts

22        s = Trim(s)
          Dim equalsAdded As Boolean
23        If Left(s, 1) <> "=" Then
24            s = "=" & s
25            equalsAdded = True
26        End If
27        Application.Calculation = xlCalculationManual
28        Application.DisplayAlerts = False
          
29        If ConvertToUS Then
              ' Set FormulaLocal and get Formula
30            On Error GoTo DecimalFixer
31            ThisWorkbook.Sheets(1).Cells(1, 1).FormulaLocal = s
32            On Error GoTo ErrorHandler
33            s = ThisWorkbook.Sheets(1).Cells(1, 1).Formula
34        Else
              ' Set Formula and get FormulaLocal
35            On Error GoTo DecimalFixer
36            ThisWorkbook.Sheets(1).Cells(1, 1).Formula = s
37            On Error GoTo ErrorHandler
38            s = ThisWorkbook.Sheets(1).Cells(1, 1).FormulaLocal
39        End If
          
40        If equalsAdded Then
41            If Left(s, 1) = "=" Then s = Mid(s, 2)
42        End If
43        ConvertLocale = s

ExitFunction:
44        ThisWorkbook.Sheets(1).Cells(1, 1).Clear
45        Application.Calculation = oldCalculation
46        Application.DisplayAlerts = oldDisplayAlerts
47        If RaiseError Then RethrowError
48        Exit Function

DecimalFixer: 'Ensures decimal character used is correct.
49        If ConvertToUS Then
50            s = Replace(s, ".", Application.DecimalSeparator)
51        Else
52            s = Replace(s, Application.DecimalSeparator, ".")
53        End If
54        Resume

ErrorHandler:
55        If Not ReportError("OpenSolverUtils", "ConvertFromCurrentLocale") Then Resume
56        RaiseError = True
57        ConvertLocale = vbNullString
58        GoTo ExitFunction
End Function

Function GetSolverParametersDict(Solver As ISolver, sheet As Worksheet) As Dictionary
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim SolverParameters As Dictionary
3         Set SolverParameters = New Dictionary
          
          ' First we fill all info from the saved options. These can then be overridden by the parameters defined on the sheet
4         If PrecisionAvailable(Solver) Then SolverParameters.Add Key:=Solver.PrecisionName, Item:=GetPrecision(sheet)
5         If ToleranceAvailable(Solver) Then SolverParameters.Add Key:=Solver.ToleranceName, Item:=GetTolerance(sheet)
          
6         If TimeLimitAvailable(Solver) Then
              ' Trim TimeLimit to valid value - MAX_LONG seconds is still 68 years!
7             SolverParameters.Add Key:=Solver.TimeLimitName, Item:=Min(GetMaxTime(sheet), MAX_LONG)
8         End If
          
9         If IterationLimitAvailable(Solver) Then
              ' Trim IterationLimit to a valid integer
10            SolverParameters.Add Key:=Solver.IterationLimitName, Item:=Int(Min(GetMaxIterations(sheet), MAX_LONG))
11        End If
          
          ' The user can define a set of parameters they want to pass to the solver; this gets them as a dictionary. MUST be on the current sheet
          Dim SolverParametersRange As Range, i As Long
12        Set SolverParametersRange = GetSolverParameters(Solver.ShortName, sheet:=sheet)
13        If Not SolverParametersRange Is Nothing Then
14            ValidateSolverParameters SolverParametersRange
15            For i = 1 To SolverParametersRange.Rows.Count
                  Dim ParamName As String, ParamValue As String
16                ParamName = Trim(SolverParametersRange.Cells(i, 1))
17                If Len(ParamName) > 0 Then
18                    If SolverParameters.Exists(ParamName) Then SolverParameters.Remove ParamName
19                    ParamValue = SolverParametersRange.Cells(i, 2).value
20                    SolverParameters.Add Key:=ParamName, Item:=ParamValue
21                End If
22            Next i
23        End If

24        Set GetSolverParametersDict = SolverParameters

ExitFunction:
25        If RaiseError Then RethrowError
26        Exit Function

ErrorHandler:
27        If Not ReportError("OpenSolverUtils", "GetSolverParametersDict") Then Resume
28        RaiseError = True
29        GoTo ExitFunction
End Function

Function ParametersToKwargs(SolverParameters As Dictionary) As String
                Dim RaiseError As Boolean
1               RaiseError = False
2               On Error GoTo ErrorHandler

                Dim Key As Variant, result As String
3               For Each Key In SolverParameters.Keys
4                   result = result & Key & _
                             IIf(Len(SolverParameters.Item(Key)) > 0, "=" & StrExNoPlus(SolverParameters.Item(Key)), vbNullString) & " "
5               Next Key
6               ParametersToKwargs = Trim(result)

ExitFunction:
7               If RaiseError Then RethrowError
8               Exit Function

ErrorHandler:
9               If Not ReportError("OpenSolverUtils", "ParametersToKwargs") Then Resume
10              RaiseError = True
11              GoTo ExitFunction
End Function

Function ParametersToFlags(SolverParameters As Dictionary) As String
                Dim RaiseError As Boolean
1               RaiseError = False
2               On Error GoTo ErrorHandler

                Dim Key As Variant, result As String
3               For Each Key In SolverParameters.Keys
4                   result = result & IIf(Left(Key, 1) <> "-", "-", vbNullString) & Key & " " & StrExNoPlus(SolverParameters.Item(Key)) & " "
5               Next Key
6               ParametersToFlags = Trim(result)

ExitFunction:
7               If RaiseError Then RethrowError
8               Exit Function

ErrorHandler:
9               If Not ReportError("OpenSolverUtils", "ParametersToFlags") Then Resume
10              RaiseError = True
11              GoTo ExitFunction
End Function

Function ParametersToOptionsFileString(SolverParameters As Dictionary) As String
                Dim RaiseError As Boolean
1               RaiseError = False
2               On Error GoTo ErrorHandler
                
                Dim Key As Variant, result As String
3               For Each Key In SolverParameters.Keys
4                   result = result & Key & " " & StrExNoPlus(SolverParameters.Item(Key)) & vbNewLine
5               Next Key
                
6               ParametersToOptionsFileString = StripTrailingNewline(result)
                
ExitFunction:
7               If RaiseError Then RethrowError
8               Exit Function

ErrorHandler:
9               If Not ReportError("OpenSolverUtils", "ParametersToOptionsFileString") Then Resume
10              RaiseError = True
11              GoTo ExitFunction
End Function

Sub ParametersToOptionsFile(OptionsFilePath As String, SolverParameters As Dictionary)
                Dim RaiseError As Boolean
1               RaiseError = False
2               On Error GoTo ErrorHandler

3               DeleteFileAndVerify OptionsFilePath
                
                Dim FileNum As Integer
4               FileNum = FreeFile()
5               Open OptionsFilePath For Output As #FileNum
6               Print #FileNum, ParametersToOptionsFileString(SolverParameters)

ExitSub:
7               Close #FileNum
8               If RaiseError Then RethrowError
9               Exit Sub

ErrorHandler:
10              If Not ReportError("SolverFileNL", "OutputOptionsFile") Then Resume
11              RaiseError = True
12              GoTo ExitSub
End Sub

Function Max(ParamArray Vals() As Variant) As Variant
1         Max = Vals(LBound(Vals))
          
          Dim i As Long
2         For i = LBound(Vals) + 1 To UBound(Vals)
3             If Vals(i) > Max Then
4                 Max = Vals(i)
5             End If
6         Next i
End Function

Function Min(ParamArray Vals() As Variant) As Variant
1         Min = Vals(LBound(Vals))
          
          Dim i As Long
2         For i = LBound(Vals) + 1 To UBound(Vals)
3             If Vals(i) < Min Then
4                 Min = Vals(i)
5             End If
6         Next i
End Function

Function Create1x1Array(X As Variant) As Variant
          ' Create a 1x1 array containing the value x
          Dim v(1, 1) As Variant
1         v(1, 1) = X
2         Create1x1Array = v
End Function

Function StringArray(ParamArray Vals() As Variant) As String()
                ' Creates a string array from the input args
                Dim TempArray() As String
1               ReDim TempArray(LBound(Vals) To UBound(Vals))
                Dim i As Long
2               For i = LBound(Vals) To UBound(Vals)
3                   TempArray(i) = CStr(Vals(i))
4               Next i
5               StringArray = TempArray
End Function

Function ForceCalculate(prompt As String, Optional MinimiseUserInteraction As Boolean = False) As Boolean
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          #If Mac Then
              'In Excel 2011 the Application.CalculationState is not included:
              'http://sysmod.wordpress.com/2011/10/24/more-differences-mainly-vba/
              'Try calling 'Calculate' two times just to be safe? This will probably cause problems down the line, maybe Office 2014 will fix it?
3             Application.Calculate
4             Application.Calculate
5             ForceCalculate = True
          #Else
              'There appears to be a bug in Excel 2010 where the .Calculate does not always complete. We handle up to 3 such failures.
              ' We have seen this problem arise on large models.
6             Application.Calculate
7             If Application.CalculationState <> xlDone Then
8                 Application.Calculate
                  Dim i As Long
9                 For i = 1 To 10
10                    DoEvents
11                    mSleep 100
12                Next i
13            End If
14            If Application.CalculationState <> xlDone Then Application.Calculate
15            If Application.CalculationState <> xlDone Then
16                DoEvents
17                Application.CalculateFullRebuild
18                DoEvents
19            End If
          
              ' Check for circular references causing problems, which can happen if iterative calculation mode is enabled.
20            If Application.CalculationState <> xlDone Then
21                If Application.Iteration Then
22                    If MinimiseUserInteraction Then
23                        Application.Iteration = False
24                        Application.Calculate
25                    ElseIf MsgBox("Iterative calculation mode is enabled and may be interfering with the inital calculation. " & _
                                    "Would you like to try disabling iterative calculation mode to see if this fixes the problem?", _
                                    vbYesNo, _
                                    "OpenSolver: Iterative Calculation Mode Detected...") = vbYes Then
26                        Application.Iteration = False
27                        Application.Calculate
28                    End If
29                End If
30            End If
          
31            While Application.CalculationState <> xlDone
32                If MinimiseUserInteraction Then
33                    ForceCalculate = False
34                    GoTo ExitFunction
35                ElseIf MsgBox(prompt, _
                                vbCritical + vbRetryCancel + vbDefaultButton1, _
                                "OpenSolver: Calculation Error Occured...") = vbCancel Then
36                    ForceCalculate = False
37                    GoTo ExitFunction
38                Else 'Recalculate the workbook if the user wants to retry
39                    Application.Calculate
40                End If
41            Wend
42            ForceCalculate = True
          #End If

ExitFunction:
43        If RaiseError Then RethrowError
44        Exit Function

ErrorHandler:
45        If Not ReportError("OpenSolverUtils", "ForceCalculate") Then Resume
46        RaiseError = True
47        GoTo ExitFunction
End Function

Sub WriteToFile(intFileNum As Long, strData As String, Optional numSpaces As Long = 0, Optional AbortIfBlank As Boolean = False)
      ' Writes a string to the given file number, adds a newline, and number of spaces to front if specified
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If Len(strData) = 0 And AbortIfBlank Then GoTo ExitSub
4         Print #intFileNum, Space(numSpaces) & strData

ExitSub:
5         If RaiseError Then RethrowError
6         Exit Sub

ErrorHandler:
7         If Not ReportError("OpenSolverUtils", "WriteToFile") Then Resume
8         RaiseError = True
9         GoTo ExitSub
End Sub

Function MakeSpacesNonBreaking(Text As String) As String
      ' Replaces all spaces with NBSP char
1         MakeSpacesNonBreaking = Replace(Text, Chr(32), Chr(NBSP))
End Function

Function StripNonBreakingSpaces(Text As String) As String
      ' Replaces all spaces with NBSP char
1         StripNonBreakingSpaces = Replace(Text, Chr(NBSP), Chr(32))
End Function

Function Quote(Text As String) As String
1         Quote = """" & Text & """"
End Function

Function TrimBlankLines(s As String) As String
      ' Remove any blank lines at the beginning or end of s
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim Done As Boolean, NewLineSize As Integer
3         NewLineSize = Len(vbNewLine)
4         While Not Done
5             If Len(s) < NewLineSize Then
6                 Done = True
7             ElseIf Left(s, NewLineSize) = vbNewLine Then
8                s = Mid(s, NewLineSize + 1)
9             Else
10                Done = True
11            End If
12        Wend
13        Done = False
14        While Not Done
15            If Len(s) < NewLineSize Then
16                Done = True
17            ElseIf Right(s, NewLineSize) = vbNewLine Then
18               s = Left(s, Len(s) - NewLineSize)
19            Else
20                Done = True
21            End If
22        Wend
23        TrimBlankLines = s

ExitFunction:
24        If RaiseError Then RethrowError
25        Exit Function

ErrorHandler:
26        If Not ReportError("OpenSolverUtils", "TrimBlankLines") Then Resume
27        RaiseError = True
28        GoTo ExitFunction
End Function

Function IsZero(num As Double) As Boolean
      ' Returns true if a number is zero (within tolerance)
1         IsZero = IIf(Abs(num) < OpenSolver.EPSILON, True, False)
End Function

Function ZeroIfSmall(value As Double) As Double
1               ZeroIfSmall = IIf(IsZero(value), 0, value)
End Function

Function StrEx(d As Variant, Optional AddSign As Boolean = True) As String
      ' Convert a double to a string, always with a + or -. Also ensure we have "0.", not just "." for values between -1 and 1
              Dim s As String
1             On Error GoTo Abort
2             s = str(d)  ' check d is numeric and convert to string
3             s = Mid(s, 2)  ' remove the initial space (reserved by VB for the sign)
              ' ensure we have "0.", not just "."
4             StrEx = IIf(Left(s, 1) = ".", "0", vbNullString) & s
5             If AddSign Or d < 0 Then StrEx = IIf(d >= 0, "+", "-") & StrEx
6             Exit Function
Abort:
              ' d is not a number
7             StrEx = d
End Function

Function StrExNoPlus(d As Variant) As String
1         StrExNoPlus = StrEx(d, False)
End Function

Function IsAmericanNumber(s As String, Optional i As Long = 1) As Boolean
          ' Check this is a number like 3.45  or +1.23e-34
          ' This does NOT test for regional variations such as 12,34
          ' This code exists because
          '   val("12+3") gives 12 with no error
          '   Assigning a string to a double uses region-specific translation, so x="1,2" works in French
          '   IsNumeric("12,45") is true even on a US English system (and even worse...)
          '   IsNumeric("($1,23,,3.4,,,5,,E67$)")=True! See http://www.eggheadcafe.com/software/aspnet/31496070/another-vba-bug.aspx)

          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim MustBeInteger As Boolean, SeenDot As Boolean, SeenDigit As Boolean
3         MustBeInteger = i > 1   ' We call this a second time after seeing the "E", when only an int is allowed
4         IsAmericanNumber = False    ' Assume we fail
5         If Len(s) = 0 Then GoTo ExitFunction ' Not a number
6         If Mid(s, i, 1) = "+" Or Mid(s, i, 1) = "-" Then i = i + 1 ' Skip leading sign
7         For i = i To Len(s)
8             Select Case Asc(Mid(s, i, 1))
              Case Asc("E"), Asc("e")
9                 If MustBeInteger Or Not SeenDigit Then GoTo ExitFunction ' No exponent allowed (as must be a simple integer)
10                IsAmericanNumber = IsAmericanNumber(s, i + 1)   ' Process an int after the E
11                GoTo ExitFunction
12            Case Asc(".")
13                If SeenDot Then GoTo ExitFunction
14                SeenDot = True
15            Case Asc("0") To Asc("9")
16                SeenDigit = True
17            Case Else
18                GoTo ExitFunction   ' Not a valid char
19            End Select
20        Next i
          ' i As Long, AllowDot As Boolean
21        IsAmericanNumber = SeenDigit

ExitFunction:
22        If RaiseError Then RethrowError
23        Exit Function

ErrorHandler:
24        If Not ReportError("OpenSolverUtils", "IsAmericanNumber") Then Resume
25        RaiseError = True
26        GoTo ExitFunction
End Function

Function SplitWithoutRepeats(StringToSplit As String, Delimiter As String) As String()
      ' As Split() function, but treats consecutive delimiters as one
                Dim RaiseError As Boolean
1               RaiseError = False
2               On Error GoTo ErrorHandler

                Dim SplitValues() As String
3               SplitValues = Split(StringToSplit, Delimiter)
                ' Remove empty splits caused by consecutive delimiters
                Dim LastNonEmpty As Long, i As Long
4               LastNonEmpty = -1
5               For i = 0 To UBound(SplitValues)
6                   If Len(SplitValues(i)) > 0 Then
7                       LastNonEmpty = LastNonEmpty + 1
8                       SplitValues(LastNonEmpty) = SplitValues(i)
9                   End If
10              Next
11              ReDim Preserve SplitValues(0 To LastNonEmpty)
12              SplitWithoutRepeats = SplitValues

ExitFunction:
13              If RaiseError Then RethrowError
14              Exit Function

ErrorHandler:
15              If Not ReportError("OpenSolverUtils", "SplitWithoutRepeats") Then Resume
16              RaiseError = True
17              GoTo ExitFunction
End Function

Public Function TestKeyExists(ByRef col As Collection, Key As String) As Boolean
    On Error GoTo doesntExist:
          Dim Item As Variant
1         Set Item = col(Key)
2         TestKeyExists = True
3         Exit Function
          
doesntExist:
4         If Err.Number = 5 Then
5             TestKeyExists = False
6         Else
7             TestKeyExists = True
8         End If
          
End Function

Public Sub OpenURL(URL As String)
                Dim RaiseError As Boolean
1               RaiseError = False
2               On Error GoTo ErrorHandler

          #If Mac Then
3                   ExecAsync "open " & Quote(URL)
          #Else
                    ' We can't use ActiveWorkbook.FollowHyperlink as this seems to have some limit on
                    ' the length of the URL that we can pass
4                   fHandleFile URL, WIN_NORMAL
          #End If

ExitSub:
5               If RaiseError Then RethrowError
6               Exit Sub

ErrorHandler:
7               If Not ReportError("OpenSolverUtils", "OpenURL") Then Resume
8               RaiseError = True
9               GoTo ExitSub
End Sub

Public Function URLEncode(StringVal As String, Optional SpaceAsPlus As Boolean = False) As String
          Dim RaiseError As Boolean
1         RaiseError = False

          ' Starting in Excel 2013, this function is built in as WorksheetFunction.EncodeURL
          ' We can't include it without causing compilation errors on earlier versions, so we need our own
          
          ' From http://stackoverflow.com/a/218199
2         On Error GoTo ErrorHandler
3         Dim StringLen As Long: StringLen = Len(StringVal)
4         If StringLen > 0 Then
5             ReDim result(StringLen) As String
              Dim i As Long, CharCode As Integer
              Dim Char As String, Space As String

6             If SpaceAsPlus Then Space = "+" Else Space = "%20"

7             For i = 1 To StringLen
8                 Char = Mid$(StringVal, i, 1)
9                 CharCode = Asc(Char)
10                Select Case CharCode
                      Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
11                        result(i) = Char
12                    Case 32
13                        result(i) = Space
14                    Case 0 To 15
15                        result(i) = "%0" & Hex(CharCode)
16                    Case Else
17                        result(i) = "%" & Hex(CharCode)
18                End Select
19            Next i
20            URLEncode = Join(result, vbNullString)
21        End If

ExitFunction:
22        If RaiseError Then RethrowError
23        Exit Function

ErrorHandler:
24        If Not ReportError("OpenSolverUtils", "URLEncode") Then Resume
25        RaiseError = True
26        GoTo ExitFunction
End Function

#If Win32 Then
Private Sub fHandleFile(FilePath As String, WindowStyle As Long)
          ' Used to open a URL - Code Courtesy of Dev Ashish
          Dim lRet As Long
    #If VBA7 Then
              Dim hwnd As LongPtr
    #Else
              Dim hwnd As Long
    #End If
          'First try ShellExecute
1         lRet = apiShellExecute(hwnd, vbNullString, FilePath, vbNullString, vbNullString, WindowStyle)

2         If lRet <= ERROR_SUCCESS Then
3             Select Case lRet
                  Case ERROR_NO_ASSOC:
                      'Try the OpenWith dialog
                      Dim varTaskID As Variant
4                     varTaskID = Shell("rundll32.exe shell32.dll, OpenAs_RunDLL " & FilePath, WIN_NORMAL)
5                 Case ERROR_OUT_OF_MEM:
6                     RaiseGeneralError "Error: Out of Memory/Resources. Couldn't Execute!"
7                 Case ERROR_FILE_NOT_FOUND:
8                     RaiseGeneralError "Error: File not found.  Couldn't Execute!"
9                 Case ERROR_PATH_NOT_FOUND:
10                    RaiseGeneralError "Error: Path not found. Couldn't Execute!"
11                Case ERROR_BAD_FORMAT:
12                    RaiseGeneralError "Error:  Bad File Format. Couldn't Execute!"
13                Case Else:
14                    RaiseGeneralError "Unknown error when opening file"
15            End Select
16        End If
End Sub
#End If

Public Function SystemIs64Bit() As Boolean
          #If Mac Then
              ' Check output of uname -a
              Dim result As String
1             result = ExecCapture("uname -a")
2             SystemIs64Bit = (InStr(result, "x86_64") > 0)
          #Else
              ' Is true if the Windows system is a 64 bit one
              ' If Not Environ("ProgramFiles(x86)") = "" Then Is64Bit=True, or
              ' Is64bit = Len(Environ("ProgramW6432")) > 0; see:
              ' http://blog.johnmuellerbooks.com/2011/06/06/checking-the-vba-environment.aspx and
              ' http://www.mrexcel.com/forum/showthread.php?542727-Determining-If-OS-Is-32-Bit-Or-64-Bit-Using-VBA and
              ' http://stackoverflow.com/questions/6256140/how-to-detect-if-the-computer-is-x32-or-x64 and
              ' http://msdn.microsoft.com/en-us/library/ms684139%28v=vs.85%29.aspx
3             SystemIs64Bit = Len(Environ("ProgramFiles(x86)")) > 0
          #End If
End Function

Private Function VBAversion() As String
          #If VBA7 Then
1             VBAversion = "VBA7"
          #ElseIf VBA6 Then
2             VBAversion = "VBA6"
          #Else
3             VBAversion = "VBA"
          #End If
End Function

Private Function ExcelBitness() As String
          #If Win64 Then
1             ExcelBitness = "64"
          #Else
2             ExcelBitness = "32"
          #End If
End Function

Private Function ExcelLanguage() As String
          Dim Lang As Long
    #If Mac Then
              ' http://www.rondebruin.nl/mac/mac002.htm
1             Lang = Application.LocalizedLanguage
    #Else
2             Lang = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    #End If
3         ExcelLanguage = LanguageCodeToString(Lang)
End Function
    
Private Function LanguageCodeToString(Lang As Long)
          Dim Language As String
1         Select Case Lang
          Case 1033: Language = "English - US"
2         Case 1036: Language = "French"
3         Case 1031: Language = "German"
4         Case 1040: Language = "Italian"
5         Case 3082: Language = "Spanish - Spain (Modern Sort)"
6         Case 1034: Language = "Spanish - Spain (Traditional Sort)"
7         Case Else: Language = "Code " & Lang & "; see http://msdn.microsoft.com/en-US/goglobal/bb964664.aspx"
8         End Select
9         LanguageCodeToString = Language
End Function

Private Function OSFamily() As String
          #If Mac Then
1             OSFamily = "Mac"
          #Else
2             OSFamily = "Windows"
          #End If
End Function

Private Function OSVersion() As String
    #If Mac Then
1             OSVersion = Application.Clean(ExecCapture("sw_vers -productVersion"))
    #Else
              Dim info As OSVERSIONINFO
              Dim retvalue As Integer
2             info.dwOSVersionInfoSize = 148
3             info.szCSDVersion = Space$(128)
4             retvalue = GetVersionExA(info)
5             OSVersion = info.dwMajorVersion & "." & info.dwMinorVersion
    #End If
End Function

Private Function OSBitness() As String
1         OSBitness = IIf(SystemIs64Bit, "64", "32")
End Function

Private Function OSUsername() As String
    #If Mac Then
1             OSUsername = ExecCapture("whoami")
    #Else
2             OSUsername = Environ("USERNAME")
    #End If
End Function

Private Function OpenSolverDistribution() As String
          ' TODO replace with enum
1         OpenSolverDistribution = IIf(SolverIsPresent(CreateSolver("Bonmin")), "Advanced", "Linear")
End Function

Public Function EnvironmentString() As String
      ' Short encoding of key environment details
1         EnvironmentString = _
              OSFamily() & "/" & OSVersion() & "x" & OSBitness() & " " & _
              "Excel/" & Application.Version & "x" & ExcelBitness() & " " & _
              "OpenSolver/" & sOpenSolverVersion & "x" & OpenSolverDistribution()
End Function

Public Function EnvironmentSummary() As String
      ' Human-readable summary of key environment details
1         EnvironmentSummary = _
              "Version " & sOpenSolverVersion & " (" & sOpenSolverDate & ") " & _
              "running on " & OSBitness() & "-bit " & OSFamily() & " " & _
              OSVersion() & " with " & VBAversion() & " in " & ExcelBitness() & _
              "-bit Excel " & Application.Version
End Function

Public Function EnvironmentDetail() As String
      ' Full description of environment details
          Dim ProductCodeLine As String
    #If Win32 Then
1             ProductCodeLine = "Excel product code = " & Application.ProductCode & _
                                vbNewLine
    #End If
2         EnvironmentDetail = _
              "OpenSolver version " & sOpenSolverVersion & " (" & sOpenSolverDate & _
              "); Distribution=" & OpenSolverDistribution & vbNewLine & _
              "Location: " & _
              MakeSpacesNonBreaking(MakePathSafe(ThisWorkbook.FullName)) & _
              vbNewLine & vbNewLine & _
              "Excel " & Application.Version & "; build " & _
              Application.Build & "; " & ExcelBitness & "-bit; " & VBAversion & _
              vbNewLine & _
              ProductCodeLine & _
              "Excel language: " & ExcelLanguage & vbNewLine & _
              "OS: " & OSFamily & " " & OSVersion & "; " & OSBitness & "-bit" & _
              vbNewLine & _
              "Username: " & OSUsername
End Function

Public Function SolverSummary() As String
          Dim SolverShortName As Variant, Solver As ISolver
1         For Each SolverShortName In GetAvailableSolvers()
2             Set Solver = CreateSolver(CStr(SolverShortName))
3             If TypeOf Solver Is ISolverLocal Then
4                 SolverSummary = SolverSummary & AboutLocalSolver(Solver) & vbNewLine & vbNewLine
                  
                  ' If we are not correctly installed, we can break after the first such message
5                 If Not SolverDirIsPresent Then
6                     Exit Function
7                 End If
8             End If
9         Next SolverShortName
End Function

Sub UpdateStatusBar(Text As String, Optional Force As Boolean = False)
      ' Function for updating the status bar.
      ' Saves the last time the bar was updated and won't re-update until a specified amount of time has passed
      ' The bar can be forced to display the new text regardless of time with the Force argument.
      ' We only need to toggle ScreenUpdating on Mac
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

    #If Mac Then
              Dim ScreenStatus As Boolean
3             ScreenStatus = Application.ScreenUpdating
    #End If

          Static LastUpdate As Double
          Dim TimeDiff As Double
4         TimeDiff = (Now() - LastUpdate) * 86400  ' Time since last update in seconds

          ' Check if last update was long enough ago
5         If TimeDiff > 0.5 Or Force Then
6             LastUpdate = Now()
              
        #If Mac Then
7                 Application.ScreenUpdating = True
        #End If

8             Application.StatusBar = Text
9             DoEvents
10        End If

ExitSub:
    #If Mac Then
11            Application.ScreenUpdating = ScreenStatus
    #End If
12        If RaiseError Then RethrowError
13        Exit Sub

ErrorHandler:
14        If Not ReportError("OpenSolverUtils", "UpdateStatusBar") Then Resume
15        RaiseError = True
16        GoTo ExitSub
End Sub

Public Function MsgBoxEx(ByVal prompt As String, _
                Optional ByVal Options As VbMsgBoxStyle = 0&, _
                Optional ByVal Title As String = "OpenSolver", _
                Optional ByVal HelpFile As String, _
                Optional ByVal Context As Long, _
                Optional ByVal LinkTarget As String, _
                Optional ByVal LinkText As String, _
                Optional ByVal MoreDetailsButton As Boolean, _
                Optional ByVal ReportIssueButton As Boolean) _
        As VbMsgBoxResult

          ' Extends MsgBox with extra options:
          ' - First five args are the same as MsgBox, so any MsgBox calls can be swapped to MsgBoxEx
          ' - LinkTarget: a hyperlink will be included above the button if this is set
          ' - LinkText: the display text for the hyperlink. Defaults to the URL if not set
          ' - MoreDetailsButton: Shows a button that opens the error log
          ' - EmailReportButton: Shows a button that prepares an error report email
          
          Dim InteractiveStatus As Boolean
1         InteractiveStatus = Application.Interactive
          
2         If Len(LinkText) = 0 Then LinkText = LinkTarget
          
          Dim Button1 As String, Button2 As String, Button3 As String
          Dim Value1 As VbMsgBoxResult, Value2 As VbMsgBoxResult, Value3 As VbMsgBoxResult
          
          ' Get button types
3         Select Case Options Mod 8
          Case vbOKOnly
4             Button1 = "OK"
5             Value1 = vbOK
6         Case vbOKCancel
7             Button1 = "OK"
8             Value1 = vbOK
9             Button2 = "Cancel"
10            Value2 = vbCancel
11        Case vbAbortRetryIgnore
12            Button1 = "Abort"
13            Value1 = vbAbort
14            Button2 = "Retry"
15            Value2 = vbRetry
16            Button3 = "Ignore"
17            Value3 = vbIgnore
18        Case vbYesNoCancel
19            Button1 = "Yes"
20            Value1 = vbYes
21            Button2 = "No"
22            Value2 = vbNo
23            Button3 = "Cancel"
24            Value3 = vbCancel
25        Case vbYesNo
26            Button1 = "Yes"
27            Value1 = vbYes
28            Button2 = "No"
29            Value2 = vbNo
30        Case vbRetryCancel
31            Button1 = "Retry"
32            Value1 = vbRetry
33            Button2 = "Cancel"
34            Value2 = vbCancel
35        End Select
          
36        With New FMsgBoxEx
37            .cmdMoreDetails.Visible = MoreDetailsButton
38            .cmdReportIssue.Visible = ReportIssueButton
          
              ' Set up buttons
39            .cmdButton1.Caption = Button1
40            .cmdButton2.Caption = Button2
41            .cmdButton3.Caption = Button3
42            .cmdButton1.Tag = Value1
43            .cmdButton2.Tag = Value2
44            .cmdButton3.Tag = Value3
              
              ' Get default button
45            Select Case (Options / 256) Mod 4
              Case vbDefaultButton1 / 256
46                .cmdButton1.SetFocus
47            Case vbDefaultButton2 / 256
48                .cmdButton2.SetFocus
49            Case vbDefaultButton3 / 256
50                .cmdButton3.SetFocus
51            End Select
              ' Adjust default button if specified default is going to be hidden
52            If .ActiveControl.Tag = "0" Then .cmdButton1.SetFocus
          
              ' We need to unlock the textbox before writing to it on Mac
53            .txtMessage.Locked = False
54            .txtMessage.Text = prompt
55            .txtMessage.Locked = True
          
56            .lblLink.Caption = LinkText
57            .lblLink.ControlTipText = LinkTarget
          
58            .Caption = Title
              
59            .AutoLayout
              
60            Application.Interactive = True
61            .Show
62            Application.Interactive = InteractiveStatus
           
              ' If form was closed using [X], then it was also unloaded, so we set the default to vbCancel
63            MsgBoxEx = vbCancel
64            On Error Resume Next
65            MsgBoxEx = CLng(.Tag)
66            On Error GoTo 0
67        End With
End Function

Function ShowEscapeCancelMessage() As VbMsgBoxResult
1         ShowEscapeCancelMessage = MsgBox("You have pressed the Escape key. Do you wish to cancel?", _
                                           vbCritical + vbYesNo + vbDefaultButton1, _
                                           "OpenSolver - User Interrupt Occured...")
End Function

Function StringHasUnicode(TestString As String) As Boolean
      ' Quickly check for any characters that aren't ASCII
          Dim i As Long, CharCode As Long
1         For i = 1 To Len(TestString)
2             CharCode = AscW(Mid(TestString, i, 1))
3             If CharCode > 127 Or CharCode < 0 Then
4                 StringHasUnicode = True
5                 Exit Function
6             End If
7         Next i
8         StringHasUnicode = False
End Function

Function AddEquals(s As String) As String
1         AddEquals = IIf(Left(s, 1) <> "=", "=", vbNullString) & s
End Function
Function RemoveEquals(s As String) As String
1         RemoveEquals = IIf(Left(s, 1) <> "=", s, Mid(s, 2))
End Function


' Base64 encode/decode implementations for mac and windows
#If Mac Then
    Function EncodeBase64(ByVal str As String) As String
1             EncodeBase64 = ExecCapture("base64 <<< " & Quote(str))
    End Function
    
    Function DecodeBase64(ByVal str As String) As String
1             DecodeBase64 = ExecCapture("base64 --decode <<< " & str)
    End Function
#Else
    Function EncodeBase64(ByVal str As String) As String
        Dim arrData() As Byte
        Dim objXML As Object 'MSXML2.DOMDocument
        Dim objNode As Object 'MSXML2.IXMLDOMElement
        
        Dim RaiseError As Boolean
        RaiseError = False
        On Error GoTo ErrorHandler
    
        arrData = StrConv(str, vbFromUnicode)
        Set objXML = CreateObject("MSXML2.DOMDocument")
        Set objNode = objXML.createElement("b64")
        
        With objNode
            .DataType = "bin.base64"
            .nodeTypedValue = arrData
            EncodeBase64 = .Text
        End With
    
ExitFunction:
        Set objNode = Nothing
        Set objXML = Nothing
        If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
        Exit Function
    
ErrorHandler:
        If Not ReportError("OpenSolverUtils", "EncodeBase64") Then Resume
        RaiseError = True
        GoTo ExitFunction
        
    End Function
    
    ' Code by Tim Hastings
    Function DecodeBase64(ByVal strData As String) As String
              Dim RaiseError As Boolean
1             RaiseError = False
2             On Error GoTo ErrorHandler
    
              Dim objXML As Object 'MSXML2.DOMDocument
              Dim objNode As Object 'MSXML2.IXMLDOMElement
            
              ' Help from MSXML
3             Set objXML = CreateObject("MSXML2.DOMDocument")
4             Set objNode = objXML.createElement("b64")
5             objNode.DataType = "bin.base64"
6             objNode.Text = strData
7             DecodeBase64 = Stream_BinaryToString(objNode.nodeTypedValue)
            
    
ExitFunction:
8             Set objNode = Nothing
9             Set objXML = Nothing
10            If RaiseError Then RethrowError
11            Exit Function
    
ErrorHandler:
12            If Not ReportError("OpenSolverUtils", "DecodeBase64") Then Resume
13            RaiseError = True
14            GoTo ExitFunction
    End Function
    
    ' Code by Tim Hastings
    Function Stream_BinaryToString(Binary)
               Dim RaiseError As Boolean
1              RaiseError = False
2              On Error GoTo ErrorHandler
              
               Const adTypeText = 2
               Const adTypeBinary = 1
               
               'Create Stream object
               Dim BinaryStream 'As New Stream
3              Set BinaryStream = CreateObject("ADODB.Stream")
               
               'Specify stream type - we want To save binary data.
4              BinaryStream.Type = adTypeBinary
               
               'Open the stream And write binary data To the object
5              BinaryStream.Open
6              BinaryStream.Write Binary
               
               'Change stream type To text/string
7              BinaryStream.Position = 0
8              BinaryStream.Type = adTypeText
               
               'Specify charset For the output text (unicode) data.
9              BinaryStream.Charset = "us-ascii"
               
               'Open the stream And get text/string data from the object
10             Stream_BinaryToString = BinaryStream.ReadText
11             Set BinaryStream = Nothing
    
ExitFunction:
12             If RaiseError Then RethrowError
13             Exit Function
    
ErrorHandler:
14             If Not ReportError("OpenSolverUtils", "Stream_BinaryToString") Then Resume
15             RaiseError = True
16             GoTo ExitFunction
    End Function
#End If
