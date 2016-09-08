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
              RemoveSheetNameFromString = Replace(s, sheetName, "")
              GoTo ExitFunction
          End If
          
280       sheetName = EscapeSheetName(sheet)
281       If InStr(s, sheetName) Then
282           RemoveSheetNameFromString = Replace(s, sheetName, "")
283           GoTo ExitFunction
284       End If

290       RemoveSheetNameFromString = s

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
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
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
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
588       StripWorksheetNameAndDollars = Replace(StripWorksheetNameAndDollars, "$", "")

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverUtils", "StripWorksheetNameAndDollars") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function EscapeSheetName(sheet As Worksheet, Optional ForceQuotes As Boolean = False) As String
    EscapeSheetName = sheet.Name
    
    Dim SpecialChar As Variant, NeedsEscaping As Boolean
    NeedsEscaping = False
    For Each SpecialChar In Array("'", "!", "(", ")", "+", "-", " ")
        If InStr(EscapeSheetName, SpecialChar) Then
            NeedsEscaping = True
            Exit For
        End If
    Next SpecialChar

    If NeedsEscaping Then EscapeSheetName = Replace(EscapeSheetName, "'", "''")
    If ForceQuotes Or NeedsEscaping Then EscapeSheetName = "'" & EscapeSheetName & "'"
    
    EscapeSheetName = EscapeSheetName & "!"
End Function

Function ConvertFromCurrentLocale(ByVal s As String) As String
' Convert a formula or a range from the current locale into US locale
          ConvertFromCurrentLocale = ConvertLocale(s, True)
End Function

Function ConvertToCurrentLocale(ByVal s As String) As String
' Convert a formula or a range from US locale into the current locale
          ConvertToCurrentLocale = ConvertLocale(s, False)
End Function

Private Function ConvertLocale(ByVal s As String, ConvertToUS As Boolean) As String
' Convert strings between locales
' This will add a leading "=" if its not already there
' A blank string is returned if any errors occur
' This works by putting the expression into cell A1 on Sheet1 of the add-in!

          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          ' Static to track whether we can skip or not
          ' 0 = Haven't checked yet
          ' 1 = Can't skip
          ' 2 = Can skip
          Static SkipConvert As Long
          
          If SkipConvert < 1 Then
              SkipConvert = 1  ' Don't check again
              ' If we are in an english version of Excel and a known locale, we can skip the conversion
              If Application.International(xlCountryCode) = 1 Then ' English language excel
                  Dim Country As Long
                  Country = Application.International(xlCountrySetting)
                  SkipConvert = 2
                  
                  Select Case Country
                  Case 1  ' US
                  Case 44 ' UK
                  Case 61 ' Australia
                  Case 64 ' NZ
                  Case Else: SkipConvert = 1
                  End Select
              End If
          End If
          
          If SkipConvert = 2 Then
              ConvertLocale = s
              Exit Function
          End If
              

          ' We turn off calculation & hide alerts as we don't want Excel popping up dialogs asking for references to other sheets
          Dim oldCalculation As Long
291       oldCalculation = Application.Calculation
          Dim oldDisplayAlerts As Boolean
292       oldDisplayAlerts = Application.DisplayAlerts

294       s = Trim(s)
          Dim equalsAdded As Boolean
295       If Left(s, 1) <> "=" Then
296           s = "=" & s
297           equalsAdded = True
298       End If
299       Application.Calculation = xlCalculationManual
300       Application.DisplayAlerts = False
          
          If ConvertToUS Then
              ' Set FormulaLocal and get Formula
              On Error GoTo DecimalFixer
              ThisWorkbook.Sheets(1).Cells(1, 1).FormulaLocal = s
              On Error GoTo ErrorHandler
302           s = ThisWorkbook.Sheets(1).Cells(1, 1).Formula
          Else
              ' Set Formula and get FormulaLocal
              On Error GoTo DecimalFixer
              ThisWorkbook.Sheets(1).Cells(1, 1).Formula = s
              On Error GoTo ErrorHandler
              s = ThisWorkbook.Sheets(1).Cells(1, 1).FormulaLocal
          End If
          
303       If equalsAdded Then
304           If Left(s, 1) = "=" Then s = Mid(s, 2)
305       End If
306       ConvertLocale = s

ExitFunction:
          ThisWorkbook.Sheets(1).Cells(1, 1).Clear
          Application.Calculation = oldCalculation
          Application.DisplayAlerts = oldDisplayAlerts
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

DecimalFixer: 'Ensures decimal character used is correct.
          If ConvertToUS Then
              s = Replace(s, ".", Application.DecimalSeparator)
          Else
              s = Replace(s, Application.DecimalSeparator, ".")
          End If
          Resume

ErrorHandler:
          If Not ReportError("OpenSolverUtils", "ConvertFromCurrentLocale") Then Resume
          RaiseError = True
          ConvertLocale = ""
          GoTo ExitFunction
End Function

Function GetSolverParametersDict(Solver As ISolver, sheet As Worksheet) As Dictionary
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim SolverParameters As Dictionary
          Set SolverParameters = New Dictionary
          
          ' First we fill all info from the saved options. These can then be overridden by the parameters defined on the sheet
          If PrecisionAvailable(Solver) Then SolverParameters.Add Key:=Solver.PrecisionName, Item:=GetPrecision(sheet)
          If ToleranceAvailable(Solver) Then SolverParameters.Add Key:=Solver.ToleranceName, Item:=GetTolerance(sheet)
          
          If TimeLimitAvailable(Solver) Then
              ' Trim TimeLimit to valid value - MAX_LONG seconds is still 68 years!
              SolverParameters.Add Key:=Solver.TimeLimitName, Item:=Min(GetMaxTime(sheet), MAX_LONG)
          End If
          
          If IterationLimitAvailable(Solver) Then
              ' Trim IterationLimit to a valid integer
              SolverParameters.Add Key:=Solver.IterationLimitName, Item:=Int(Min(GetMaxIterations(sheet), MAX_LONG))
          End If
          
          ' The user can define a set of parameters they want to pass to the solver; this gets them as a dictionary. MUST be on the current sheet
          Dim SolverParametersRange As Range, i As Long
6104      Set SolverParametersRange = GetSolverParameters(Solver.ShortName, sheet:=sheet)
          If Not SolverParametersRange Is Nothing Then
6105          ValidateSolverParameters SolverParametersRange
6109          For i = 1 To SolverParametersRange.Rows.Count
                  Dim ParamName As String, ParamValue As String
6110              ParamName = Trim(SolverParametersRange.Cells(i, 1))
6111              If ParamName <> "" Then
                      If SolverParameters.Exists(ParamName) Then SolverParameters.Remove ParamName
6112                  ParamValue = SolverParametersRange.Cells(i, 2).value
6114                  SolverParameters.Add Key:=ParamName, Item:=ParamValue
6115              End If
6116          Next i
6117      End If

          Set GetSolverParametersDict = SolverParameters

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverUtils", "GetSolverParametersDict") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function ParametersToKwargs(SolverParameters As Dictionary) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim Key As Variant, result As String
          For Each Key In SolverParameters.Keys
              result = result & Key & _
                       IIf(SolverParameters.Item(Key) <> "", "=" & StrExNoPlus(SolverParameters.Item(Key)), "") & " "
          Next Key
          ParametersToKwargs = Trim(result)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverUtils", "ParametersToKwargs") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function ParametersToFlags(SolverParameters As Dictionary) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim Key As Variant, result As String
          For Each Key In SolverParameters.Keys
              result = result & IIf(Left(Key, 1) <> "-", "-", "") & Key & " " & StrExNoPlus(SolverParameters.Item(Key)) & " "
          Next Key
          ParametersToFlags = Trim(result)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverUtils", "ParametersToFlags") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function ParametersToOptionsFileString(SolverParameters As Dictionary) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim Key As Variant, result As String
          For Each Key In SolverParameters.Keys
              result = result & Key & " " & StrExNoPlus(SolverParameters.Item(Key)) & vbNewLine
          Next Key
          
          ParametersToOptionsFileString = StripTrailingNewline(result)
          
ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverUtils", "ParametersToOptionsFileString") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Sub ParametersToOptionsFile(OptionsFilePath As String, SolverParameters As Dictionary)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          DeleteFileAndVerify OptionsFilePath
          
          Dim FileNum As Integer
          FileNum = FreeFile()
          Open OptionsFilePath For Output As #FileNum
          Print #FileNum, ParametersToOptionsFileString(SolverParameters)

ExitSub:
          Close #FileNum
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("SolverFileNL", "OutputOptionsFile") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Function Max(ParamArray Vals() As Variant) As Variant
          Max = Vals(LBound(Vals))
          
          Dim i As Long
          For i = LBound(Vals) + 1 To UBound(Vals)
482           If Vals(i) > Max Then
483               Max = Vals(i)
486           End If
          Next i
End Function

Function Min(ParamArray Vals() As Variant) As Variant
          Min = Vals(LBound(Vals))
          
          Dim i As Long
          For i = LBound(Vals) + 1 To UBound(Vals)
482           If Vals(i) < Min Then
483               Min = Vals(i)
486           End If
          Next i
End Function

Function Create1x1Array(X As Variant) As Variant
          ' Create a 1x1 array containing the value x
          Dim v(1, 1) As Variant
492       v(1, 1) = X
493       Create1x1Array = v
End Function

Function StringArray(ParamArray Vals() As Variant) As String()
          ' Creates a string array from the input args
          Dim TempArray() As String
          ReDim TempArray(LBound(Vals) To UBound(Vals))
          Dim i As Long
          For i = LBound(Vals) To UBound(Vals)
              TempArray(i) = CStr(Vals(i))
          Next i
          StringArray = TempArray
End Function

Function ForceCalculate(prompt As String, Optional MinimiseUserInteraction As Boolean = False) As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          #If Mac Then
              'In Excel 2011 the Application.CalculationState is not included:
              'http://sysmod.wordpress.com/2011/10/24/more-differences-mainly-vba/
              'Try calling 'Calculate' two times just to be safe? This will probably cause problems down the line, maybe Office 2014 will fix it?
494           Application.Calculate
495           Application.Calculate
496           ForceCalculate = True
          #Else
              'There appears to be a bug in Excel 2010 where the .Calculate does not always complete. We handle up to 3 such failures.
              ' We have seen this problem arise on large models.
497           Application.Calculate
498           If Application.CalculationState <> xlDone Then
499               Application.Calculate
                  Dim i As Long
500               For i = 1 To 10
501                   DoEvents
502                   mSleep 100
503               Next i
504           End If
505           If Application.CalculationState <> xlDone Then Application.Calculate
506           If Application.CalculationState <> xlDone Then
507               DoEvents
508               Application.CalculateFullRebuild
509               DoEvents
510           End If
          
              ' Check for circular references causing problems, which can happen if iterative calculation mode is enabled.
511           If Application.CalculationState <> xlDone Then
512               If Application.Iteration Then
513                   If MinimiseUserInteraction Then
514                       Application.Iteration = False
515                       Application.Calculate
516                   ElseIf MsgBox("Iterative calculation mode is enabled and may be interfering with the inital calculation. " & _
                                    "Would you like to try disabling iterative calculation mode to see if this fixes the problem?", _
                                    vbYesNo, _
                                    "OpenSolver: Iterative Calculation Mode Detected...") = vbYes Then
517                       Application.Iteration = False
518                       Application.Calculate
519                   End If
520               End If
521           End If
          
522           While Application.CalculationState <> xlDone
523               If MinimiseUserInteraction Then
524                   ForceCalculate = False
525                   GoTo ExitFunction
526               ElseIf MsgBox(prompt, _
                                vbCritical + vbRetryCancel + vbDefaultButton1, _
                                "OpenSolver: Calculation Error Occured...") = vbCancel Then
527                   ForceCalculate = False
528                   GoTo ExitFunction
529               Else 'Recalculate the workbook if the user wants to retry
530                   Application.Calculate
531               End If
532           Wend
533           ForceCalculate = True
          #End If

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverUtils", "ForceCalculate") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Sub WriteToFile(intFileNum As Long, strData As String, Optional numSpaces As Long = 0, Optional AbortIfBlank As Boolean = False)
' Writes a string to the given file number, adds a newline, and number of spaces to front if specified
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          If Len(strData) = 0 And AbortIfBlank Then GoTo ExitSub
781       Print #intFileNum, Space(numSpaces) & strData

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("OpenSolverUtils", "WriteToFile") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Function MakeSpacesNonBreaking(Text As String) As String
' Replaces all spaces with NBSP char
784       MakeSpacesNonBreaking = Replace(Text, Chr(32), Chr(NBSP))
End Function

Function StripNonBreakingSpaces(Text As String) As String
' Replaces all spaces with NBSP char
784       StripNonBreakingSpaces = Replace(Text, Chr(NBSP), Chr(32))
End Function

Function Quote(Text As String) As String
745       Quote = """" & Text & """"
End Function

Function TrimBlankLines(s As String) As String
' Remove any blank lines at the beginning or end of s
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim Done As Boolean, NewLineSize As Integer
          NewLineSize = Len(vbNewLine)
611       While Not Done
612           If Len(s) < NewLineSize Then
613               Done = True
614           ElseIf Left(s, NewLineSize) = vbNewLine Then
615              s = Mid(s, NewLineSize + 1)
616           Else
617               Done = True
618           End If
619       Wend
620       Done = False
621       While Not Done
622           If Len(s) < NewLineSize Then
623               Done = True
624           ElseIf Right(s, NewLineSize) = vbNewLine Then
625              s = Left(s, Len(s) - NewLineSize)
626           Else
627               Done = True
628           End If
629       Wend
630       TrimBlankLines = s

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverUtils", "TrimBlankLines") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function IsZero(num As Double) As Boolean
' Returns true if a number is zero (within tolerance)
785       IsZero = IIf(Abs(num) < OpenSolver.EPSILON, True, False)
End Function

Function ZeroIfSmall(value As Double) As Double
          ZeroIfSmall = IIf(IsZero(value), 0, value)
End Function

Function StrEx(d As Variant, Optional AddSign As Boolean = True) As String
' Convert a double to a string, always with a + or -. Also ensure we have "0.", not just "." for values between -1 and 1
              Dim s As String
              On Error GoTo Abort
              s = str(d)  ' check d is numeric and convert to string
1912          s = Mid(s, 2)  ' remove the initial space (reserved by VB for the sign)
1913          ' ensure we have "0.", not just "."
1915          StrEx = IIf(Left(s, 1) = ".", "0", "") & s
              If AddSign Or d < 0 Then StrEx = IIf(d >= 0, "+", "-") & StrEx
              Exit Function
Abort:
              ' d is not a number
              StrEx = d
End Function

Function StrExNoPlus(d As Variant) As String
    StrExNoPlus = StrEx(d, False)
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
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim MustBeInteger As Boolean, SeenDot As Boolean, SeenDigit As Boolean
631       MustBeInteger = i > 1   ' We call this a second time after seeing the "E", when only an int is allowed
632       IsAmericanNumber = False    ' Assume we fail
633       If Len(s) = 0 Then GoTo ExitFunction ' Not a number
634       If Mid(s, i, 1) = "+" Or Mid(s, i, 1) = "-" Then i = i + 1 ' Skip leading sign
635       For i = i To Len(s)
636           Select Case Asc(Mid(s, i, 1))
              Case Asc("E"), Asc("e")
637               If MustBeInteger Or Not SeenDigit Then GoTo ExitFunction ' No exponent allowed (as must be a simple integer)
638               IsAmericanNumber = IsAmericanNumber(s, i + 1)   ' Process an int after the E
639               GoTo ExitFunction
640           Case Asc(".")
641               If SeenDot Then GoTo ExitFunction
642               SeenDot = True
643           Case Asc("0") To Asc("9")
644               SeenDigit = True
645           Case Else
646               GoTo ExitFunction   ' Not a valid char
647           End Select
648       Next i
          ' i As Long, AllowDot As Boolean
649       IsAmericanNumber = SeenDigit

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverUtils", "IsAmericanNumber") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function SplitWithoutRepeats(StringToSplit As String, Delimiter As String) As String()
' As Split() function, but treats consecutive delimiters as one
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim SplitValues() As String
          SplitValues = Split(StringToSplit, Delimiter)
          ' Remove empty splits caused by consecutive delimiters
          Dim LastNonEmpty As Long, i As Long
          LastNonEmpty = -1
          For i = 0 To UBound(SplitValues)
              If SplitValues(i) <> "" Then
                  LastNonEmpty = LastNonEmpty + 1
                  SplitValues(LastNonEmpty) = SplitValues(i)
              End If
          Next
          ReDim Preserve SplitValues(0 To LastNonEmpty)
          SplitWithoutRepeats = SplitValues

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverUtils", "SplitWithoutRepeats") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Public Function TestKeyExists(ByRef col As Collection, Key As String) As Boolean
          On Error GoTo doesntExist:
          Dim Item As Variant
2020      Set Item = col(Key)
2021      TestKeyExists = True
2022      Exit Function
          
doesntExist:
2023      If Err.Number = 5 Then
2024          TestKeyExists = False
2025      Else
2026          TestKeyExists = True
2027      End If
          
End Function

Public Sub OpenURL(URL As String)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          #If Mac Then
              ExecAsync "open " & Quote(URL)
          #Else
              ' We can't use ActiveWorkbook.FollowHyperlink as this seems to have some limit on
              ' the length of the URL that we can pass
              fHandleFile URL, WIN_NORMAL
          #End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("OpenSolverUtils", "OpenURL") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Public Function URLEncode(StringVal As String, Optional SpaceAsPlus As Boolean = False) As String
    Dim RaiseError As Boolean
    RaiseError = False

    ' Starting in Excel 2013, this function is built in as WorksheetFunction.EncodeURL
    ' We can't include it without causing compilation errors on earlier versions, so we need our own
    
    ' From http://stackoverflow.com/a/218199
    On Error GoTo ErrorHandler
    Dim StringLen As Long: StringLen = Len(StringVal)
    If StringLen > 0 Then
        ReDim result(StringLen) As String
        Dim i As Long, CharCode As Integer
        Dim Char As String, Space As String

        If SpaceAsPlus Then Space = "+" Else Space = "%20"

        For i = 1 To StringLen
            Char = Mid$(StringVal, i, 1)
            CharCode = Asc(Char)
            Select Case CharCode
                Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
                    result(i) = Char
                Case 32
                    result(i) = Space
                Case 0 To 15
                    result(i) = "%0" & Hex(CharCode)
                Case Else
                    result(i) = "%" & Hex(CharCode)
            End Select
        Next i
        URLEncode = Join(result, "")
    End If

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverUtils", "URLEncode") Then Resume
    RaiseError = True
    GoTo ExitFunction
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
    lRet = apiShellExecute(hwnd, vbNullString, FilePath, vbNullString, vbNullString, WindowStyle)

    If lRet <= ERROR_SUCCESS Then
        Select Case lRet
            Case ERROR_NO_ASSOC:
                'Try the OpenWith dialog
                Dim varTaskID As Variant
                varTaskID = Shell("rundll32.exe shell32.dll, OpenAs_RunDLL " & FilePath, WIN_NORMAL)
            Case ERROR_OUT_OF_MEM:
                Err.Raise OpenSolver_Error, Description:="Error: Out of Memory/Resources. Couldn't Execute!"
            Case ERROR_FILE_NOT_FOUND:
                Err.Raise OpenSolver_Error, Description:="Error: File not found.  Couldn't Execute!"
            Case ERROR_PATH_NOT_FOUND:
                Err.Raise OpenSolver_Error, Description:="Error: Path not found. Couldn't Execute!"
            Case ERROR_BAD_FORMAT:
                Err.Raise OpenSolver_Error, Description:="Error:  Bad File Format. Couldn't Execute!"
            Case Else:
                Err.Raise OpenSolver_Error, Description:="Unknown error when opening file"
        End Select
    End If
End Sub
#End If

Public Function SystemIs64Bit() As Boolean
          #If Mac Then
              ' Check output of uname -a
              Dim result As String
664           result = ExecCapture("uname -a")
665           SystemIs64Bit = (InStr(result, "x86_64") > 0)
          #Else
              ' Is true if the Windows system is a 64 bit one
              ' If Not Environ("ProgramFiles(x86)") = "" Then Is64Bit=True, or
              ' Is64bit = Len(Environ("ProgramW6432")) > 0; see:
              ' http://blog.johnmuellerbooks.com/2011/06/06/checking-the-vba-environment.aspx and
              ' http://www.mrexcel.com/forum/showthread.php?542727-Determining-If-OS-Is-32-Bit-Or-64-Bit-Using-VBA and
              ' http://stackoverflow.com/questions/6256140/how-to-detect-if-the-computer-is-x32-or-x64 and
              ' http://msdn.microsoft.com/en-us/library/ms684139%28v=vs.85%29.aspx
666           SystemIs64Bit = Environ("ProgramFiles(x86)") <> ""
          #End If
End Function

Private Function VBAversion() As String
          #If VBA7 Then
3517          VBAversion = "VBA7"
          #ElseIf VBA6 Then
3518          VBAversion = "VBA6"
          #Else
3516          VBAversion = "VBA"
          #End If
End Function

Private Function ExcelBitness() As String
          #If Win64 Then
3519          ExcelBitness = "64"
          #Else
3520          ExcelBitness = "32"
          #End If
End Function

Private Function ExcelLanguage() As String
    Dim Lang As Long
    #If Mac Then
        ' http://www.rondebruin.nl/mac/mac002.htm
        Lang = Application.LocalizedLanguage
    #Else
        Lang = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    #End If
    ExcelLanguage = LanguageCodeToString(Lang)
End Function
    
Private Function LanguageCodeToString(Lang As Long)
    Dim Language As String
    Select Case Lang
    Case 1033: Language = "English - US"
    Case 1036: Language = "French"
    Case 1031: Language = "German"
    Case 1040: Language = "Italian"
    Case 3082: Language = "Spanish - Spain (Modern Sort)"
    Case 1034: Language = "Spanish - Spain (Traditional Sort)"
    Case Else: Language = "Code " & Lang & "; see http://msdn.microsoft.com/en-US/goglobal/bb964664.aspx"
    End Select
    LanguageCodeToString = Language
End Function

Private Function OSFamily() As String
          #If Mac Then
3521          OSFamily = "Mac"
          #Else
3522          OSFamily = "Windows"
          #End If
End Function

Private Function OSVersion() As String
    #If Mac Then
        OSVersion = Application.Clean(ExecCapture("sw_vers -productVersion"))
    #Else
        Dim info As OSVERSIONINFO
        Dim retvalue As Integer
        info.dwOSVersionInfoSize = 148
        info.szCSDVersion = Space$(128)
        retvalue = GetVersionExA(info)
        OSVersion = info.dwMajorVersion & "." & info.dwMinorVersion
    #End If
End Function

Private Function OSBitness() As String
    OSBitness = IIf(SystemIs64Bit, "64", "32")
End Function

Private Function OSUsername() As String
    #If Mac Then
        OSUsername = ExecCapture("whoami")
    #Else
        OSUsername = Environ("USERNAME")
    #End If
End Function

Private Function OpenSolverDistribution() As String
    ' TODO replace with enum
    OpenSolverDistribution = IIf(SolverIsPresent(CreateSolver("Bonmin")), "Advanced", "Linear")
End Function

Public Function EnvironmentString() As String
' Short encoding of key environment details
    EnvironmentString = _
        OSFamily() & "/" & OSVersion() & "x" & OSBitness() & " " & _
        "Excel/" & Application.Version & "x" & ExcelBitness() & " " & _
        "OpenSolver/" & sOpenSolverVersion & "x" & OpenSolverDistribution()
End Function

Public Function EnvironmentSummary() As String
' Human-readable summary of key environment details
    EnvironmentSummary = _
        "Version " & sOpenSolverVersion & " (" & sOpenSolverDate & ") " & _
        "running on " & OSBitness() & "-bit " & OSFamily() & " " & _
        OSVersion() & " with " & VBAversion() & " in " & ExcelBitness() & _
        "-bit Excel " & Application.Version
End Function

Public Function EnvironmentDetail() As String
' Full description of environment details
    Dim ProductCodeLine As String
    #If Win32 Then
        ProductCodeLine = "Excel product code = " & Application.ProductCode & _
                          vbNewLine
    #End If
    EnvironmentDetail = _
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
    For Each SolverShortName In GetAvailableSolvers()
        Set Solver = CreateSolver(CStr(SolverShortName))
        If TypeOf Solver Is ISolverLocal Then
            SolverSummary = SolverSummary & AboutLocalSolver(Solver) & vbNewLine & vbNewLine
            
            ' If we are not correctly installed, we can break after the first such message
            If Not SolverDirIsPresent Then
                Exit Function
            End If
        End If
    Next SolverShortName
End Function

Sub UpdateStatusBar(Text As String, Optional Force As Boolean = False)
' Function for updating the status bar.
' Saves the last time the bar was updated and won't re-update until a specified amount of time has passed
' The bar can be forced to display the new text regardless of time with the Force argument.
' We only need to toggle ScreenUpdating on Mac
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    #If Mac Then
        Dim ScreenStatus As Boolean
        ScreenStatus = Application.ScreenUpdating
    #End If

    Static LastUpdate As Double
    Dim TimeDiff As Double
    TimeDiff = (Now() - LastUpdate) * 86400  ' Time since last update in seconds

    ' Check if last update was long enough ago
    If TimeDiff > 0.5 Or Force Then
        LastUpdate = Now()
        
        #If Mac Then
            Application.ScreenUpdating = True
        #End If

        Application.StatusBar = Text
        DoEvents
    End If

ExitSub:
    #If Mac Then
        Application.ScreenUpdating = ScreenStatus
    #End If
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Sub

ErrorHandler:
    If Not ReportError("OpenSolverUtils", "UpdateStatusBar") Then Resume
    RaiseError = True
    GoTo ExitSub
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
    InteractiveStatus = Application.Interactive
    
    If Len(LinkText) = 0 Then LinkText = LinkTarget
    
    Dim Button1 As String, Button2 As String, Button3 As String
    Dim Value1 As VbMsgBoxResult, Value2 As VbMsgBoxResult, Value3 As VbMsgBoxResult
    
    ' Get button types
    Select Case Options Mod 8
    Case vbOKOnly
        Button1 = "OK"
        Value1 = vbOK
    Case vbOKCancel
        Button1 = "OK"
        Value1 = vbOK
        Button2 = "Cancel"
        Value2 = vbCancel
    Case vbAbortRetryIgnore
        Button1 = "Abort"
        Value1 = vbAbort
        Button2 = "Retry"
        Value2 = vbRetry
        Button3 = "Ignore"
        Value3 = vbIgnore
    Case vbYesNoCancel
        Button1 = "Yes"
        Value1 = vbYes
        Button2 = "No"
        Value2 = vbNo
        Button3 = "Cancel"
        Value3 = vbCancel
    Case vbYesNo
        Button1 = "Yes"
        Value1 = vbYes
        Button2 = "No"
        Value2 = vbNo
    Case vbRetryCancel
        Button1 = "Retry"
        Value1 = vbRetry
        Button2 = "Cancel"
        Value2 = vbCancel
    End Select
    
    With New FMsgBoxEx
        .cmdMoreDetails.Visible = MoreDetailsButton
        .cmdReportIssue.Visible = ReportIssueButton
    
        ' Set up buttons
        .cmdButton1.Caption = Button1
        .cmdButton2.Caption = Button2
        .cmdButton3.Caption = Button3
        .cmdButton1.Tag = Value1
        .cmdButton2.Tag = Value2
        .cmdButton3.Tag = Value3
        
        ' Get default button
        Select Case (Options / 256) Mod 4
        Case vbDefaultButton1 / 256
            .cmdButton1.SetFocus
        Case vbDefaultButton2 / 256
            .cmdButton2.SetFocus
        Case vbDefaultButton3 / 256
            .cmdButton3.SetFocus
        End Select
        ' Adjust default button if specified default is going to be hidden
        If .ActiveControl.Tag = "0" Then .cmdButton1.SetFocus
    
        ' We need to unlock the textbox before writing to it on Mac
        .txtMessage.Locked = False
        .txtMessage.Text = prompt
        .txtMessage.Locked = True
    
        .lblLink.Caption = LinkText
        .lblLink.ControlTipText = LinkTarget
    
        .Caption = Title
        
        .AutoLayout
        
        Application.Interactive = True
        .Show
        Application.Interactive = InteractiveStatus
     
        ' If form was closed using [X], then it was also unloaded, so we set the default to vbCancel
        MsgBoxEx = vbCancel
        On Error Resume Next
        MsgBoxEx = CLng(.Tag)
        On Error GoTo 0
    End With
End Function

Function ShowEscapeCancelMessage() As VbMsgBoxResult
    ShowEscapeCancelMessage = MsgBox("You have pressed the Escape key. Do you wish to cancel?", _
                                     vbCritical + vbYesNo + vbDefaultButton1, _
                                     "OpenSolver - User Interrupt Occured...")
End Function

Function StringHasUnicode(TestString As String) As Boolean
' Quickly check for any characters that aren't ASCII
    Dim i As Long, CharCode As Long
    For i = 1 To Len(TestString)
        CharCode = AscW(Mid(TestString, i, 1))
        If CharCode > 127 Or CharCode < 0 Then
            StringHasUnicode = True
            Exit Function
        End If
    Next i
    StringHasUnicode = False
End Function

Function AddEquals(s As String) As String
    AddEquals = IIf(Left(s, 1) <> "=", "=", vbNullString) & s
End Function
Function RemoveEquals(s As String) As String
    RemoveEquals = IIf(Left(s, 1) <> "=", s, Mid(s, 2))
End Function

