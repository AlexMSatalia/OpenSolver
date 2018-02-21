Attribute VB_Name = "OpenSolverRangeUtils"
Option Explicit

Private SearchRangeNameCACHE As Collection

Sub SearchRangeName_DestroyCache()
      ' Destroy the name cache
      ' Andres Sommerhoff
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If Not SearchRangeNameCACHE Is Nothing Then
4             While SearchRangeNameCACHE.Count > 0
5                 SearchRangeNameCACHE.Remove 1
6             Wend
7         End If
8         Set SearchRangeNameCACHE = Nothing

ExitSub:
9         If RaiseError Then RethrowError
10        Exit Sub

ErrorHandler:
11        If Not ReportError("OpenSolverRangeUtils", "SearchRangeName_DestroyCache") Then Resume
12        RaiseError = True
13        GoTo ExitSub
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

1         On Error Resume Next
2         CurrNamesCount = sheet.Parent.Names.Count
3         CurrSheetName = sheet.Name
4         CurrFileName = sheet.Parent.Name
          
5         If LastNamesCount <> CurrNamesCount _
             Or LastSheetName <> CurrSheetName _
             Or LastFileName <> CurrFileName Then
6                     SearchRangeName_DestroyCache  'Check confirm it is obsolote. Cache need to redone...
7         End If
              
8         If SearchRangeNameCACHE Is Nothing Then
9             Set SearchRangeNameCACHE = New Collection  'Start building a new Cache
10        Else
11            Exit Sub 'Cache is ok -> return back
12        End If

          'Here the Cache will be filled with visible range names only
13        For i = 1 To sheet.Parent.Names.Count
14            Set TestName = sheet.Parent.Names(i)
              
15            If TestName.Visible = True Then  'Iterate through the visible names only
                      ' Skip any references to external workbooks
16                    If Left$(TestName.RefersTo, 1) = "=" And InStr(TestName.RefersTo, "[") > 1 Then GoTo tryNext
17                    On Error GoTo tryerror
                      'Build the Cache with the range address as key (='sheet1'!$A$1:$B$3)
18                    Set rComp = TestName.RefersToRange
19                    SearchRangeNameCACHE.Add TestName, (rComp.Name)
20                    GoTo tryNext
tryerror:
21                    Resume tryNext
tryNext:
22            End If
23        Next i
          
24        LastNamesCount = CurrNamesCount
25        LastSheetName = CurrSheetName
26        LastFileName = CurrFileName
End Sub

Function SearchRangeInVisibleNames(r As Range) As Name
      ' Get a name from the cache if it exists
      ' Andres Sommerhoff
1         SearchRangeName_LoadCache r.Parent
2         On Error Resume Next
3         Set SearchRangeInVisibleNames = SearchRangeNameCACHE.Item((r.Name))
End Function

Function GetCellName(cell As Range) As String
1         GetCellName = cell.Address(RowAbsolute:=False, ColumnAbsolute:=False)  ' Eg. A1
End Function

Function GetDisplayAddress(RefersTo As String, sheet As Worksheet, Optional showRangeName As Boolean = False) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim r As Range
3         On Error Resume Next
4         Set r = Range(RefersTo)
5         If Err.Number <> 0 Then
6             GetDisplayAddress = RemoveSheetNameFromString(RefersTo, ActiveSheet)
7             GoTo ExitFunction
8         End If
9         On Error GoTo ErrorHandler
          
          Dim RefersToNames() As String, Offset As Long
10        RefersToNames = Split(RefersTo, ",")
11        Offset = LBound(RefersToNames) - 1  ' Add Offset to make array 1-indexed
          
          ' Make sure our range has the right number of areas
12        If UBound(RefersToNames) - Offset <> r.Areas.Count Then
13            RaiseGeneralError "The number of names does not match the number of areas in the range." & vbNewLine & _
                                "Names: " & RefersTo & vbNewLine & _
                                "Range: " & r.Address
14        End If
          
          ' Include a sheet name if this range is not on the active sheet
          Dim Prefix As String
15        If r.Worksheet.Name <> sheet.Name Then
16            Prefix = EscapeSheetName(r.Worksheet)
17        End If
          
          Dim i As Long, AreaName As String, Rname As Name, R2 As Range
18        For i = 1 To r.Areas.Count
19            Set R2 = r.Areas(i)
20            AreaName = Prefix & R2.Address
21            Set Rname = SearchRangeInVisibleNames(R2)
22            If Not Rname Is Nothing Then
                  ' Check if the name was specified in the RefersTo
23                If Rname.Name <> RefersToNames(i + Offset) Then
24                    If showRangeName Then
25                        AreaName = AreaName & " (" & StripWorksheetNameAndDollars(Rname.Name, R2.Worksheet) & ")"
26                    End If
27                Else
28                    AreaName = RefersToNames(i + Offset)
29                End If
30            End If
31            GetDisplayAddress = GetDisplayAddress & "," & AreaName
32        Next i
          
          ' Trim "," at beginning
33        GetDisplayAddress = Mid(GetDisplayAddress, 2)
          ' Check it works!
34        Set r = Range(GetDisplayAddress)

ExitFunction:
35        If RaiseError Then RethrowError
36        Exit Function

ErrorHandler:
37        If Not ReportError("OpenSolverRangeUtils", "GetDisplayAddress") Then Resume
38        RaiseError = True
39        GoTo ExitFunction
End Function

Function GetRangeValues(r As Range) As Variant()
      ' This copies the values from a possible multi-area range into a variant
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim v() As Variant, i As Long
3         ReDim v(r.Areas.Count)
4         For i = 1 To r.Areas.Count
5             v(i) = r.Areas(i).Value2 ' Copy the entire area into the i'th entry of v
6         Next i
7         GetRangeValues = v

ExitFunction:
8         If RaiseError Then RethrowError
9         Exit Function

ErrorHandler:
10        If Not ReportError("OpenSolverRangeUtils", "GetRangeValues") Then Resume
11        RaiseError = True
12        GoTo ExitFunction
End Function

Sub SetRangeValues(r As Range, v() As Variant)
      ' This copies the values from a variant into a possibly multi-area range; see GetRangeValues
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim i As Long
3         For i = 1 To r.Areas.Count
4             r.Areas(i).Value2 = v(i)
5         Next i

ExitSub:
6         If RaiseError Then RethrowError
7         Exit Sub

ErrorHandler:
8         If Not ReportError("OpenSolverRangeUtils", "SetRangeValues") Then Resume
9         RaiseError = True
10        GoTo ExitSub
End Sub

Function GetOneCellInRange(r As Range, instance As Long) As Range
      ' Given an 'instance' between 1 and r.Count, return the instance'th cell in the range, where our count goes cross each row in turn (as does 'for each in range')
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim RowOffset As Long, ColOffset As Long
          Dim NumCols As Long
3         NumCols = r.Columns.Count
4         RowOffset = ((instance - 1) \ NumCols)
5         ColOffset = ((instance - 1) Mod NumCols)
6         Set GetOneCellInRange = r.Cells(1 + RowOffset, 1 + ColOffset)

ExitFunction:
7         If RaiseError Then RethrowError
8         Exit Function

ErrorHandler:
9         If Not ReportError("OpenSolverRangeUtils", "GetOneCellInRange") Then Resume
10        RaiseError = True
11        GoTo ExitFunction
End Function

Function TestIntersect(ByRef R1 As Range, ByRef R2 As Range) As Boolean
1         If R1 Is Nothing Or R2 Is Nothing Then
2             TestIntersect = False
3         Else
4             TestIntersect = Not (Intersect(R1, R2) Is Nothing)
5         End If
End Function

Function CheckRangeContainsNoAmbiguousMergedCells(r As Range, BadCell As Range) As Boolean
      ' This checks that if the range contains any merged cells, those cells are the 'home' cell (top left) in the merged cell block
      ' and thus references to these cells are indeed to a unique cell
      ' If we have a cell that is not the top left of a merged cell, then this will be read as blank, and writing to this will effect other cells.
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         CheckRangeContainsNoAmbiguousMergedCells = True
4         If Not r.MergeCells Then
5             GoTo ExitFunction
6         End If
          Dim cell As Range
7         For Each cell In r
8             If cell.MergeCells Then
9                 If cell.Address <> cell.MergeArea.Cells(1, 1).Address Then
10                    Set BadCell = cell
11                    CheckRangeContainsNoAmbiguousMergedCells = False
12                    GoTo ExitFunction
13                End If
14            End If
15        Next cell

ExitFunction:
16        If RaiseError Then RethrowError
17        Exit Function

ErrorHandler:
18        If Not ReportError("OpenSolverRangeUtils", "CheckRangeContainsNoAmbiguousMergedCells") Then Resume
19        RaiseError = True
20        GoTo ExitFunction
End Function

Function RemoveRangeOverlap(r As Range) As Range
      ' This creates a new range from r which does not contain any multiple repetitions of cells
      ' This works around the fact that Excel allows range like "A1:A2,A2:A3", which has a .count of 4 cells
      ' The Union function does NOT remove all overlaps; call this after the union to get a valid range
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If r.Areas.Count = 1 Then
4             Set RemoveRangeOverlap = r
5             GoTo ExitFunction
6         End If
          Dim s As Range, i As Long
7         Set s = r.Areas(1)
8         For i = 2 To r.Areas.Count
9             If Intersect(s, r.Areas(i)) Is Nothing Then
                  ' Just take the standard union
10                Set s = Union(s, r.Areas(i))
11            Else
                  ' Merge these two ranges cell by cell; this seems to remove the overlap in my tests, but also see http://www.cpearson.com/excel/BetterUnion.aspx
                  ' Merge the smaller range into the larger
12                If s.Count < r.Areas(i).Count Then
13                    Set s = MergeRangesCellByCell(r.Areas(i), s)
14                Else
15                    Set s = MergeRangesCellByCell(s, r.Areas(i))
16                End If
17            End If
18        Next i
19        Set RemoveRangeOverlap = s

ExitFunction:
20        If RaiseError Then RethrowError
21        Exit Function

ErrorHandler:
22        If Not ReportError("OpenSolverRangeUtils", "RemoveRangeOverlap") Then Resume
23        RaiseError = True
24        GoTo ExitFunction
End Function

Function MergeRangesCellByCell(R1 As Range, R2 As Range) As Range
      ' This merges range r2 into r1 cell by cell.
      ' This shoulsd be fastest if range r2 is smaller than r1
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim result As Range, cell As Range
3         Set result = R1
4         For Each cell In R2
5             Set result = Union(result, cell)
6         Next cell
7         Set MergeRangesCellByCell = result

ExitFunction:
8         If RaiseError Then RethrowError
9         Exit Function

ErrorHandler:
10        If Not ReportError("OpenSolverRangeUtils", "MergeRangesCellByCell") Then Resume
11        RaiseError = True
12        GoTo ExitFunction
End Function

Function ProperUnion(R1 As Range, R2 As Range) As Range
      ' Return the union of r1 and r2, where r1 may be Nothing
      ' TODO: Handle the fact that Union will return a range with multiple copies of overlapping cells - does this matter?
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If R1 Is Nothing Then
4             Set ProperUnion = R2
5         ElseIf R2 Is Nothing Then
6             Set ProperUnion = R1
7         ElseIf Not R1.Worksheet Is R2.Worksheet Then
8             Set ProperUnion = Nothing
9         Else
10            Set ProperUnion = Union(R1, R2)
11        End If

ExitFunction:
12        If RaiseError Then RethrowError
13        Exit Function

ErrorHandler:
14        If Not ReportError("OpenSolverRangeUtils", "ProperUnion") Then Resume
15        RaiseError = True
16        GoTo ExitFunction
End Function

Function ProperPrecedents(r As Range) As Range
      ' Gets the precedents of a range, returning nothing if there are no precedents or the range is nothing
1         If r Is Nothing Then
2             GoTo ExitFunction
3         End If
          
4         On Error GoTo ErrorHandler
5         Set ProperPrecedents = r.Precedents
ExitFunction:
6         Exit Function

ErrorHandler:
7         If Err.Number = 1004 Then
8             Set ProperPrecedents = Nothing
9             Resume ExitFunction
10        End If
11        RethrowError Err
End Function

Function ProperIntersect(R1 As Range, R2 As Range) As Range
1         If R1 Is Nothing Or R2 Is Nothing Then
2             Set ProperIntersect = Nothing
3         Else
4             Set ProperIntersect = Intersect(R1, R2)
5         End If
End Function

Function SetDifference(ByRef rng1 As Range, ByRef rng2 As Range) As Range
      ' Returns rng1 \ rng2 (set minus) i.e. all elements in rng1 that are not in rng2
      ' https://stackoverflow.com/a/17510237/4492726
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim rngResult As Range
          Dim rngResultCopy As Range
          Dim rngIntersection As Range
          Dim rngArea1 As Range
          Dim rngArea2 As Range
          Dim lngTop As Long
          Dim lngLeft As Long
          Dim lngRight As Long
          Dim lngBottom As Long

3         If rng1 Is Nothing Then
4             Set rngResult = Nothing
5         ElseIf rng2 Is Nothing Then
6             Set rngResult = rng1
7         ElseIf Not rng1.Worksheet Is rng2.Worksheet Then
8             Set rngResult = rng1
9         Else
10            Set rngResult = rng1
11            For Each rngArea2 In rng2.Areas
12                If rngResult Is Nothing Then
13                    Exit For
14                End If
15                Set rngResultCopy = rngResult
16                Set rngResult = Nothing
17                For Each rngArea1 In rngResultCopy.Areas
18                    Set rngIntersection = Intersect(rngArea1, rngArea2)
19                    If rngIntersection Is Nothing Then
20                        Set rngResult = ProperUnion(rngResult, rngArea1)
21                    Else
22                        lngTop = rngIntersection.row - rngArea1.row
23                        lngLeft = rngIntersection.Column - rngArea1.Column
24                        lngRight = rngArea1.Column + rngArea1.Columns.Count - rngIntersection.Column - rngIntersection.Columns.Count
25                        lngBottom = rngArea1.row + rngArea1.Rows.Count - rngIntersection.row - rngIntersection.Rows.Count
26                        If lngTop > 0 Then
27                            Set rngResult = ProperUnion(rngResult, rngArea1.Resize(lngTop, rngArea1.Columns.Count))
28                        End If
29                        If lngLeft > 0 Then
30                            Set rngResult = ProperUnion(rngResult, rngArea1.Resize(rngArea1.Rows.Count - lngTop - lngBottom, lngLeft).Offset(lngTop, 0))
31                        End If
32                        If lngRight > 0 Then
33                            Set rngResult = ProperUnion(rngResult, rngArea1.Resize(rngArea1.Rows.Count - lngTop - lngBottom, lngRight).Offset(lngTop, rngArea1.Columns.Count - lngRight))
34                        End If
35                        If lngBottom > 0 Then
36                            Set rngResult = ProperUnion(rngResult, rngArea1.Resize(lngBottom, rngArea1.Columns.Count).Offset(rngArea1.Rows.Count - lngBottom, 0))
37                        End If
38                    End If
39                Next rngArea1
40            Next rngArea2
41        End If
42        Set SetDifference = rngResult

ExitFunction:
43        If RaiseError Then RethrowError
44        Exit Function

ErrorHandler:
45        If Not ReportError("OpenSolverRangeUtils", "SetDifference") Then Resume
46        RaiseError = True
47        GoTo ExitFunction
End Function

Sub TestCellsForWriting(r As Range)
          ' We can't do r.Value2 = r.Value2 as this
          ' just sets the values from the first area in all areas
1         On Error GoTo ErrorHandler
          Dim Area As Range
2         For Each Area In r.Areas
3             Area.Value2 = Area.Value2
4         Next Area
5         Exit Sub
          
ErrorHandler:
6         RaiseUserError Err.Description
End Sub

Function RangeExists(r As String) As Boolean
    'https://stackoverflow.com/a/19179439
    Dim Test As Range
    On Error Resume Next
    Set Test = ActiveSheet.Range(r)
    RangeExists = Err.Number = 0
End Function
