Attribute VB_Name = "OpenSolverRangeUtils"
Option Explicit

Private SearchRangeNameCACHE As Collection

Sub SearchRangeName_DestroyCache()
' Destroy the name cache
' Andres Sommerhoff
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

697       If Not SearchRangeNameCACHE Is Nothing Then
698           While SearchRangeNameCACHE.Count > 0
699               SearchRangeNameCACHE.Remove 1
700           Wend
701       End If
702       Set SearchRangeNameCACHE = Nothing

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("OpenSolverRangeUtils", "SearchRangeName_DestroyCache") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Private Sub SearchRangeName_LoadCache(sheet As Worksheet)
' Save visible defined names in book in a cache to find them quickly
' Andres Sommerhoff
          Dim TestName As Name
          Dim rComp As Range
          Dim i As Long
          
          'Some checks in case the cache is an obsolete version
          Static LastNamesCount As Long
          Static LastSheetName As String
          Static LastFileName As String
          Dim CurrNamesCount As Long
          Dim CurrSheetName As String
          Dim CurrFileName As String

703       On Error Resume Next
704       CurrNamesCount = sheet.Parent.Names.Count
705       CurrSheetName = sheet.Name
706       CurrFileName = sheet.Parent.Name
          
707       If LastNamesCount <> CurrNamesCount _
             Or LastSheetName <> CurrSheetName _
             Or LastFileName <> CurrFileName Then
708                   SearchRangeName_DestroyCache  'Check confirm it is obsolote. Cache need to redone...
709       End If
              
710       If SearchRangeNameCACHE Is Nothing Then
711           Set SearchRangeNameCACHE = New Collection  'Start building a new Cache
712       Else
713           Exit Sub 'Cache is ok -> return back
714       End If

          'Here the Cache will be filled with visible range names only
715       For i = 1 To sheet.Parent.Names.Count
716           Set TestName = sheet.Parent.Names(i)
              
717           If TestName.Visible = True Then  'Iterate through the visible names only
                      ' Skip any references to external workbooks
718                   If Left$(TestName.RefersTo, 1) = "=" And InStr(TestName.RefersTo, "[") > 1 Then GoTo tryNext
719                   On Error GoTo tryerror
                      'Build the Cache with the range address as key (='sheet1'!$A$1:$B$3)
720                   Set rComp = TestName.RefersToRange
721                   SearchRangeNameCACHE.Add TestName, (rComp.Name)
722                   GoTo tryNext
tryerror:
723                   Resume tryNext
tryNext:
724           End If
725       Next i
          
726       LastNamesCount = CurrNamesCount
727       LastSheetName = CurrSheetName
728       LastFileName = CurrFileName
End Sub

Function SearchRangeInVisibleNames(r As Range) As Name
' Get a name from the cache if it exists
' Andres Sommerhoff
729       SearchRangeName_LoadCache r.Parent
730       On Error Resume Next
731       Set SearchRangeInVisibleNames = SearchRangeNameCACHE.Item((r.Name))
End Function

Function GetDisplayAddress(RefersTo As String, sheet As Worksheet, Optional showRangeName As Boolean = False) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim r As Range
          On Error Resume Next
          Set r = Range(RefersTo)
          If Err.Number <> 0 Then
              GetDisplayAddress = RemoveSheetNameFromString(RefersTo, ActiveSheet)
              GoTo ExitFunction
          End If
          On Error GoTo ErrorHandler
          
          Dim RefersToNames() As String, Offset As Long
          RefersToNames = Split(RefersTo, ",")
          Offset = LBound(RefersToNames) - 1  ' Add Offset to make array 1-indexed
          
          ' Make sure our range has the right number of areas
          If UBound(RefersToNames) - Offset <> r.Areas.Count Then
              Err.Raise OpenSolver_ModelError, Description:="The number of names does not match the number of areas in the range." & vbNewLine & _
                                                            "Names: " & RefersTo & vbNewLine & _
                                                            "Range: " & r.Address
          End If
          
          ' Include a sheet name if this range is not on the active sheet
          Dim Prefix As String
200       If r.Worksheet.Name <> sheet.Name Then
              Prefix = EscapeSheetName(r.Worksheet)
          End If
          
          Dim i As Long, AreaName As String, Rname As Name, R2 As Range
          For i = 1 To r.Areas.Count
              Set R2 = r.Areas(i)
              AreaName = Prefix & R2.Address
203           Set Rname = SearchRangeInVisibleNames(R2)
              If Not Rname Is Nothing Then
                  ' Check if the name was specified in the RefersTo
                  If Rname.Name <> RefersToNames(i + Offset) Then
                      If showRangeName Then
205                       AreaName = AreaName & " (" & StripWorksheetNameAndDollars(Rname.Name, R2.Worksheet) & ")"
                      End If
                  Else
                      AreaName = RefersToNames(i + Offset)
206               End If
207           End If
              GetDisplayAddress = GetDisplayAddress & "," & AreaName
          Next i
          
          ' Trim "," at beginning
          GetDisplayAddress = Mid(GetDisplayAddress, 2)
          ' Check it works!
          Set r = Range(GetDisplayAddress)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverRangeUtils", "GetDisplayAddress") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function GetRangeValues(r As Range) As Variant()
' This copies the values from a possible multi-area range into a variant
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim v() As Variant, i As Long
543       ReDim v(r.Areas.Count)
544       For i = 1 To r.Areas.Count
545           v(i) = r.Areas(i).Value2 ' Copy the entire area into the i'th entry of v
546       Next i
547       GetRangeValues = v

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverRangeUtils", "GetRangeValues") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Sub SetRangeValues(r As Range, v() As Variant)
' This copies the values from a variant into a possibly multi-area range; see GetRangeValues
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim i As Long
548       For i = 1 To r.Areas.Count
549           r.Areas(i).Value2 = v(i)
550       Next i

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("OpenSolverRangeUtils", "SetRangeValues") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Function GetOneCellInRange(r As Range, instance As Long) As Range
' Given an 'instance' between 1 and r.Count, return the instance'th cell in the range, where our count goes cross each row in turn (as does 'for each in range')
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim RowOffset As Long, ColOffset As Long
          Dim NumCols As Long
478       NumCols = r.Columns.Count
479       RowOffset = ((instance - 1) \ NumCols)
480       ColOffset = ((instance - 1) Mod NumCols)
481       Set GetOneCellInRange = r.Cells(1 + RowOffset, 1 + ColOffset)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverRangeUtils", "GetOneCellInRange") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function TestIntersect(ByRef R1 As Range, ByRef R2 As Range) As Boolean
          If R1 Is Nothing Or R2 Is Nothing Then
              TestIntersect = False
          Else
783           TestIntersect = Not (Intersect(R1, R2) Is Nothing)
          End If
End Function

Function CheckRangeContainsNoAmbiguousMergedCells(r As Range, BadCell As Range) As Boolean
' This checks that if the range contains any merged cells, those cells are the 'home' cell (top left) in the merged cell block
' and thus references to these cells are indeed to a unique cell
' If we have a cell that is not the top left of a merged cell, then this will be read as blank, and writing to this will effect other cells.
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

573       CheckRangeContainsNoAmbiguousMergedCells = True
574       If Not r.MergeCells Then
575           GoTo ExitFunction
576       End If
          Dim cell As Range
577       For Each cell In r
578           If cell.MergeCells Then
579               If cell.Address <> cell.MergeArea.Cells(1, 1).Address Then
580                   Set BadCell = cell
581                   CheckRangeContainsNoAmbiguousMergedCells = False
582                   GoTo ExitFunction
583               End If
584           End If
585       Next cell

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverRangeUtils", "CheckRangeContainsNoAmbiguousMergedCells") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function RemoveRangeOverlap(r As Range) As Range
' This creates a new range from r which does not contain any multiple repetitions of cells
' This works around the fact that Excel allows range like "A1:A2,A2:A3", which has a .count of 4 cells
' The Union function does NOT remove all overlaps; call this after the union to get a valid range
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

556       If r.Areas.Count = 1 Then
557           Set RemoveRangeOverlap = r
558           GoTo ExitFunction
559       End If
          Dim s As Range, i As Long
560       Set s = r.Areas(1)
561       For i = 2 To r.Areas.Count
562           If Intersect(s, r.Areas(i)) Is Nothing Then
                  ' Just take the standard union
563               Set s = Union(s, r.Areas(i))
564           Else
                  ' Merge these two ranges cell by cell; this seems to remove the overlap in my tests, but also see http://www.cpearson.com/excel/BetterUnion.aspx
                  ' Merge the smaller range into the larger
565               If s.Count < r.Areas(i).Count Then
566                   Set s = MergeRangesCellByCell(r.Areas(i), s)
567               Else
568                   Set s = MergeRangesCellByCell(s, r.Areas(i))
569               End If
570           End If
571       Next i
572       Set RemoveRangeOverlap = s

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverRangeUtils", "RemoveRangeOverlap") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function MergeRangesCellByCell(R1 As Range, R2 As Range) As Range
' This merges range r2 into r1 cell by cell.
' This shoulsd be fastest if range r2 is smaller than r1
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim result As Range, cell As Range
551       Set result = R1
552       For Each cell In R2
553           Set result = Union(result, cell)
554       Next cell
555       Set MergeRangesCellByCell = result

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverRangeUtils", "MergeRangesCellByCell") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function ProperUnion(R1 As Range, R2 As Range) As Range
' Return the union of r1 and r2, where r1 may be Nothing
' TODO: Handle the fact that Union will return a range with multiple copies of overlapping cells - does this matter?
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

534       If R1 Is Nothing Then
535           Set ProperUnion = R2
536       ElseIf R2 Is Nothing Then
537           Set ProperUnion = R1
540       ElseIf Not R1.Worksheet Is R2.Worksheet Then
              Set ProperUnion = Nothing
          Else
541           Set ProperUnion = Union(R1, R2)
542       End If

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverRangeUtils", "ProperUnion") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function ProperPrecedents(r As Range) As Range
' Gets the precedents of a range, returning nothing if there are no precedents or the range is nothing
    If r Is Nothing Then
        GoTo ExitFunction
    End If
    
    On Error GoTo ErrorHandler
    Set ProperPrecedents = r.Precedents
ExitFunction:
    Exit Function

ErrorHandler:
    If Err.Number = 1004 Then
        Set ProperPrecedents = Nothing
        Resume ExitFunction
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Function SetDifference(ByRef rng1 As Range, ByRef rng2 As Range) As Range
' Returns rng1 \ rng2 (set minus) i.e. all elements in rng1 that are not in rng2
' https://stackoverflow.com/a/17510237/4492726
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    Dim rngResult As Range
    Dim rngResultCopy As Range
    Dim rngIntersection As Range
    Dim rngArea1 As Range
    Dim rngArea2 As Range
    Dim lngTop As Long
    Dim lngLeft As Long
    Dim lngRight As Long
    Dim lngBottom As Long

    If rng1 Is Nothing Then
        Set rngResult = Nothing
    ElseIf rng2 Is Nothing Then
        Set rngResult = rng1
    ElseIf Not rng1.Worksheet Is rng2.Worksheet Then
        Set rngResult = rng1
    Else
        Set rngResult = rng1
        For Each rngArea2 In rng2.Areas
            If rngResult Is Nothing Then
                Exit For
            End If
            Set rngResultCopy = rngResult
            Set rngResult = Nothing
            For Each rngArea1 In rngResultCopy.Areas
                Set rngIntersection = Intersect(rngArea1, rngArea2)
                If rngIntersection Is Nothing Then
                    Set rngResult = ProperUnion(rngResult, rngArea1)
                Else
                    lngTop = rngIntersection.row - rngArea1.row
                    lngLeft = rngIntersection.Column - rngArea1.Column
                    lngRight = rngArea1.Column + rngArea1.Columns.Count - rngIntersection.Column - rngIntersection.Columns.Count
                    lngBottom = rngArea1.row + rngArea1.Rows.Count - rngIntersection.row - rngIntersection.Rows.Count
                    If lngTop > 0 Then
                        Set rngResult = ProperUnion(rngResult, rngArea1.Resize(lngTop, rngArea1.Columns.Count))
                    End If
                    If lngLeft > 0 Then
                        Set rngResult = ProperUnion(rngResult, rngArea1.Resize(rngArea1.Rows.Count - lngTop - lngBottom, lngLeft).Offset(lngTop, 0))
                    End If
                    If lngRight > 0 Then
                        Set rngResult = ProperUnion(rngResult, rngArea1.Resize(rngArea1.Rows.Count - lngTop - lngBottom, lngRight).Offset(lngTop, rngArea1.Columns.Count - lngRight))
                    End If
                    If lngBottom > 0 Then
                        Set rngResult = ProperUnion(rngResult, rngArea1.Resize(lngBottom, rngArea1.Columns.Count).Offset(rngArea1.Rows.Count - lngBottom, 0))
                    End If
                End If
            Next rngArea1
        Next rngArea2
    End If
    Set SetDifference = rngResult

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverRangeUtils", "SetDifference") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

Sub TestCellsForWriting(r As Range)
    ' We can't do r.Value2 = r.Value2 as this
    ' just sets the values from the first area in all areas
    Dim Area As Range
    For Each Area In r.Areas
        Area.Value2 = Area.Value2
    Next Area
End Sub
