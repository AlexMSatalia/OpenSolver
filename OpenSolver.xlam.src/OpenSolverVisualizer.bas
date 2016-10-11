Attribute VB_Name = "OpenSolverVisualizer"
Option Explicit

Private VisualizerHasShownError As Boolean  ' Track whether we've shown an error while trying to show the model

Private ShapeIndex As Long
Private HighlightColorIndex As Long   ' Used to rotate thru colours on different constraints
Private colHighlightOffsets As Collection

Function NextHighlightColor() As Long
1         HighlightColorIndex = HighlightColorIndex + 1
2         If HighlightColorIndex > 7 Then HighlightColorIndex = 1
3         Select Case HighlightColorIndex
              Case 1: NextHighlightColor = RGB(0, 0, 255) ' Blue
4             Case 2: NextHighlightColor = RGB(0, 128, 0)  ' Green
5             Case 3: NextHighlightColor = RGB(153, 0, 204) ' Purple
6             Case 4: NextHighlightColor = RGB(128, 0, 0) ' Brown
7             Case 5: NextHighlightColor = RGB(0, 204, 51) ' Light Green
8             Case 6: NextHighlightColor = RGB(255, 102, 0) ' Orange
9             Case 7: NextHighlightColor = RGB(204, 0, 153) ' Bright Purple
10        End Select
End Function

Function NextHighlightColor_SkipColor(colorToSkip As Long) As Long
          Dim newColour As Long
1         newColour = NextHighlightColor
2         If newColour = colorToSkip Then
3             newColour = NextHighlightColor
4         End If
5         NextHighlightColor_SkipColor = newColour
End Function

Function InitialiseHighlighting()
1         Set colHighlightOffsets = New Collection
2         ShapeIndex = 0
3         HighlightColorIndex = 0 ' Used to rotate thru colours on different constraints
End Function

Function SheetHasOpenSolverHighlighting(Optional w As Worksheet)
1         GetActiveSheetIfMissing w

          ' If we have a shape called OpenSolver1 then we are displaying highlighted data
2         If w.Shapes.Count = 0 Then GoTo NoHighlighting
          Dim s As Shape
3         SheetHasOpenSolverHighlighting = True
4         On Error Resume Next
5         Set s = w.Shapes("OpenSolver" & 1) ' This string is split up to avoid false positives on anti-virus scans
6         If Err.Number = 0 Then Exit Function ' Yes, we have highlighting
          ' Because the highlighting may be on another sheet, we also check all the shapes on this sheet
7         For Each s In w.Shapes
8             If s.Name Like "OpenSolver*" Then Exit Function
9         Next s
NoHighlighting:
10        SheetHasOpenSolverHighlighting = False
End Function

Function CreateLabelShape(w As Worksheet, Left As Long, Top As Long, Width As Long, Height As Long, Label As String, HighlightColor As Long) As Shape
' Create a label (as a msoShapeRectangle) and give it text. This is used for labelling obj function as min/max, and decision vars as binary or integer
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
    
          Dim s1 As Shape
3         Set s1 = w.Shapes.AddShape(msoShapeRectangle, Left, Top, Width, Height)
4         s1.Fill.Visible = True
5         s1.Fill.Solid
6         s1.Fill.ForeColor.RGB = RGB(255, 255, 255)
7         s1.Fill.Transparency = 0.2
8         s1.Line.Visible = False
9         s1.Shadow.Visible = msoFalse
10        With s1.TextFrame
11            .Characters.Text = Label
12            .Characters.Font.Size = 9
13            .Characters.Font.Color = HighlightColor
14            .HorizontalAlignment = xlHAlignLeft
15            .VerticalAlignment = xlVAlignTop
16            .MarginBottom = 0
17            .MarginLeft = 1
18            .MarginRight = 1
19            .MarginTop = 0
20            .AutoSize = True     ' Get width correct
21            .AutoSize = False
22        End With
23        ShapeIndex = ShapeIndex + 1
24        s1.Name = "OpenSolver" & ShapeIndex
25        s1.Height = Height      ' Force the specified height
26        Set CreateLabelShape = s1

ExitFunction:
27        If RaiseError Then RethrowError
28        Exit Function

ErrorHandler:
29        If Not ReportError("OpenSolverVisualizer", "CreateLabelShape") Then Resume
30        RaiseError = True
31        GoTo ExitFunction
End Function

Function AddLabelToRange(w As Worksheet, r As Range, voffset As Long, Height As Long, Label As String, HighlightColor As Long) As Shape
1         Set AddLabelToRange = CreateLabelShape(w, r.Left + 1, r.Top + voffset, r.Width, Height, Label, HighlightColor)
End Function

Function AddLabelToShape(w As Worksheet, s As Shape, voffset As Long, Height As Long, Label As String, HighlightColor As Long)
1         Set AddLabelToShape = CreateLabelShape(w, s.Left - 1, s.Top + voffset, s.Width, Height, Label, HighlightColor)
End Function

Function HighlightRange(r As Range, Label As String, HighlightColor As Long, Optional ShowFill As Boolean = False, Optional ShapeNamePrefix As String = "OpenSolver", Optional Bounds As Boolean = False) As ShapeRange
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          'Only show the model file that is on the active sheet to overcome hide bugs for shapes on different sheets
3         If r.Worksheet.Name <> ActiveSheet.Name Then GoTo ExitFunction
          
          Const HighlightingOffsetStep = 1
          ' We offset our highlighting so that successive highlights are still visible
          Dim Offset As Double, Key As String
4         Key = r.Address(RowAbsolute:=False, ColumnAbsolute:=False)
5         Offset = 0
6         On Error Resume Next
7         Offset = colHighlightOffsets(Key) ' eg A1
8         If Err.Number <> 0 Then
              ' Item does not exist in collection create it, with an offset of 2 ready to use next time
9             colHighlightOffsets.Add HighlightingOffsetStep, Key
10        Else
11            colHighlightOffsets.Remove Key
12            colHighlightOffsets.Add Offset + HighlightingOffsetStep, Key
13        End If
14        On Error GoTo ErrorHandler
          
          ' Handle merged cells
          Dim l As Single, t As Single, Right As Single, bottom As Single, R2 As Range, c As Range
15        If Not r.MergeCells Then
16            l = r.Left
17            t = r.Top
18            Right = l + r.Width
19            bottom = t + r.Height
20        Else
              ' This range contains merged cells. We use MergeArea to to find the area to highlight
              ' But, this only works for one cell, so we expand our size based on all cells
21            If r.Count = 1 Then
22                Set R2 = r.MergeArea
23                l = R2.Left
24                t = R2.Top
25                Right = l + R2.Width
26                bottom = t + R2.Height
27            Else
28                l = r.Left
29                t = r.Top
30                Right = l + r.Width
31                bottom = t + r.Height
32                For Each c In r
33                    If c.MergeCells Then
34                        Set R2 = c.MergeArea
35                        If R2.Left < l Then l = R2.Left
36                        If R2.Top < t Then t = R2.Top
37                        If R2.Left + R2.Width > Right Then Right = R2.Left + R2.Width
38                        If R2.Top + R2.Height > bottom Then bottom = R2.Top + R2.Height
39                    End If
40                Next c
41            End If
42        End If
          
          
          ' Use doubles here for more accuracy as we are summing terms, and so accummulating errors
          Dim left2 As Double, top2 As Double, right2 As Double, bottom2 As Double, Height As Double, Width As Double

          ' Draw enough shapes to cover the space; each shape has an Excel-set (undocumented?) maximum height (and width?)
          Dim isFirstShape As Boolean
43        isFirstShape = True
          Dim firstShapeIndex As Long
44        firstShapeIndex = ShapeIndex + 1

45        top2 = t + Offset
46        Do While bottom - top2 > 0.01  ' handle float rounding
              ' The height cannot exceed 169056.0; we allow some tolerance
47            Height = bottom - top2
48            If Height > 160000# Then
49                Height = 150000  ' This difference ensures we never end up with very small rectangle
50            End If
51            If isFirstShape And Height > 9500 Then
52                Height = 9000   ' The first shape we create has a height of 9000 to ensure we can rotate the text and have it show
                                  ' correctly; this works around an Excel 2007 bug
53            End If
54            isFirstShape = False
              
              ' Reset left2 for the inside loop
55            left2 = l + Offset

56            Do While Right - left2 > 0.01
57                Width = Right - left2
                  ' The height cannot exceed 169056.0; we allow some tolerance
58                If Width > 160000# Then
59                    Width = 150000   ' This difference ensures we never end up with very small rectangle
60                End If
        
                  Dim shapeName As String, s1 As Shape
61                ShapeIndex = ShapeIndex + 1
                  'If the constraint is not a bound then make a box for it
62                If Not Bounds Then
63                    shapeName = ShapeNamePrefix & ShapeIndex
64                    r.Worksheet.Shapes.AddShape(msoShapeRectangle, left2, top2, Width, Height).Name = shapeName
65                Else
                      'If the box is a bound we name it after the cell that it is in
66                    shapeName = ShapeNamePrefix & Key
67                    On Error Resume Next
                      Dim tmpName As String
68                    tmpName = r.Worksheet.Shapes(shapeName).Name
69                    On Error GoTo ErrorHandler
                      'If there hasn't been a bound on that cell then make a new cell
70                    If Len(tmpName) = 0 Then
71                        r.Worksheet.Shapes.AddShape(msoShapeRectangle, left2, top2, Width, Height).Name = shapeName
72                    Else
                          'If there has already been a bound then just add new text to it rather then making a new box
73                        Set s1 = r.Worksheet.Shapes(shapeName)
74                        s1.TextFrame.Characters.Text = s1.TextFrame.Characters.Text & "," & Label
75                        GoTo endLoop
76                    End If
77                End If
            
78                Set s1 = r.Worksheet.Shapes(shapeName)
              
                  Dim ShowOutline As Boolean
79                ShowOutline = Not ShowFill
               
80                If ShowOutline Then
81                    s1.Fill.Visible = False
82                    With s1.Line
83                        .Weight = 2
84                        .ForeColor.RGB = HighlightColor
85                    End With
86                Else
87                    s1.Line.Visible = False
88                    s1.Fill.Solid
89                    s1.Fill.Transparency = 0.6
90                    s1.Fill.ForeColor.RGB = HighlightColor
91                End If
92                s1.Shadow.Visible = msoFalse
        
93                With s1.TextFrame
94                    .Characters.Text = Label
95                    .Characters.Font.Color = HighlightColor
96                    .HorizontalAlignment = xlHAlignLeft ' xlHAlignCenter
                      ' "=", "<=", & ">=" will be centered
97                    If ((Height < 500) Or (Label = "=") Or (Label = ChrW(&H2265)) Or (Label = ChrW(&H2264))) Then
98                        .VerticalAlignment = xlVAlignCenter  ' Shape is small enought to have text fit on the screen when centered, so we center text
99                    Else
100                       .VerticalAlignment = xlVAlignTop   ' So we can see the name when scrolled to the top
101                   End If
102                   .MarginBottom = 0
103                   .MarginLeft = 2
104                   .MarginRight = 0
105                   .MarginTop = 2
106                   .Characters.Font.Bold = True
107               End With
endLoop:
108               left2 = left2 + Width
109           Loop
110           top2 = top2 + Height
111       Loop
          
          ' Create & return the shapeRange containing all the shapes we added
          ' Check we made a shape
112       If ShapeIndex >= firstShapeIndex Then
              Dim shapeNames(), i As Long
113           ReDim shapeNames(1 To ShapeIndex - firstShapeIndex + 1)
114           If Not Bounds Then
115               For i = firstShapeIndex To ShapeIndex
116                   shapeNames(i - firstShapeIndex + 1) = ShapeNamePrefix & i
117               Next i
118           Else
119               For i = firstShapeIndex To ShapeIndex
120                   shapeNames(i - firstShapeIndex + 1) = ShapeNamePrefix & Key
121               Next i
122           End If
123           Set HighlightRange = r.Worksheet.Shapes.Range(shapeNames)
124       End If

ExitFunction:
125       If RaiseError Then RethrowError
126       Exit Function

ErrorHandler:
127       If Not ReportError("OpenSolverVisualizer", "HighlightRange") Then Resume
128       RaiseError = True
129       GoTo ExitFunction
End Function

Function AddLabelledConnector(w As Worksheet, s1 As Shape, s2 As Shape, Label As String)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim t As Shapes, c As Shape
3         Set t = w.Shapes
4         Set c = t.AddConnector(msoConnectorStraight, 0, 0, 0, 0) ' msoConnectorCurve
5         With c.ConnectorFormat
6             .BeginConnect ConnectedShape:=s1, ConnectionSite:=1
7             .EndConnect ConnectedShape:=s2, ConnectionSite:=1
8         End With
9         c.RerouteConnections
10        c.Line.ForeColor = s1.Line.ForeColor
          ' The default styles can be changed, so we should set everything! We just do a few of the problem bits
11        c.Line.EndArrowheadStyle = msoArrowheadNone
12        c.Line.BeginArrowheadStyle = msoArrowheadNone
13        c.Line.DashStyle = msoLineSolid
14        c.Line.Weight = 0.75
15        c.Line.Style = msoLineSingle
16        c.Shadow.Visible = msoFalse
              
17        ShapeIndex = ShapeIndex + 1
18        c.Name = "OpenSolver" & ShapeIndex
          
          Dim s3 As Shape
19        Set s3 = t.AddShape(msoShapeRectangle, c.Left + c.Width / 2# - 30 / 2, c.Top + c.Height / 2# - 20 / 2, 30, 20)
                  
20        s3.Line.Visible = False
21        s3.Fill.Visible = False
22        s3.Shadow.Visible = msoFalse
23        If Len(Label) > 0 Then
24            With s3.TextFrame
25                .Characters.Text = Label
26                .MarginBottom = 0
27                .MarginLeft = 0
28                .MarginRight = 0
29                .MarginTop = 0
30                .HorizontalAlignment = xlHAlignCenter
31                .VerticalAlignment = xlVAlignCenter
32                .Characters.Font.Color = c.Line.ForeColor
33                .Characters.Font.Bold = True
34            End With
35        End If

36        ShapeIndex = ShapeIndex + 1
37        s3.Name = "OpenSolver" & ShapeIndex

ExitFunction:
38        If RaiseError Then RethrowError
39        Exit Function

ErrorHandler:
40        If Not ReportError("OpenSolverVisualizer", "AddLabelledConnector") Then Resume
41        RaiseError = True
42        GoTo ExitFunction

End Function

Sub HighlightConstraint(myDocument As Worksheet, LHSRange As Range, _
                        RHSRange As Range, ByVal RHSValue As String, ByVal sense As Long, _
                        ByVal Color As Long)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          ' Show a constraint of the form LHS <|=|> RHS.
          ' We always put the sign in the rightmost (or bottom-most) range so we read left-to-right or top-to-bottom.
          Dim Range1 As Range, Range2 As Range, Reversed As Boolean
          Dim s1 As ShapeRange, s2 As ShapeRange

          ' Get next color if none specified
3         If Color = 0 Then Color = NextHighlightColor
          
4         Reversed = False
5         If RHSRange Is Nothing And Len(RHSValue) > 0 Then
              ' We have a constant or formula in the constraint. Put into form RHS <|=|> Range1 (reversing the sense)
6             Set s1 = HighlightRange(LHSRange, RHSValue & SolverRelationAsUnicodeChar(4 - sense), Color, , , True)
7         ElseIf Not RHSRange Is Nothing Then
              ' If ranges overlaps on rows, then the top one becomes Range1
8             If ((RHSRange.Top >= LHSRange.Top And RHSRange.Top < LHSRange.Top + LHSRange.Height) _
              Or (LHSRange.Top >= RHSRange.Top And LHSRange.Top < RHSRange.Top + RHSRange.Height)) Then
                  ' Ranges over lap in rows. Range1 becomes the left most one
9                 If LHSRange.Left > RHSRange.Left Then
10                    Reversed = True
11                End If
12            ElseIf ((RHSRange.Left >= LHSRange.Left And RHSRange.Left < LHSRange.Left + LHSRange.Width) _
              Or (LHSRange.Left >= RHSRange.Left And LHSRange.Left < RHSRange.Left + RHSRange.Width)) Then
                  ' Ranges overlap in columns. Range1 becomes the top most one
13                If LHSRange.Top > RHSRange.Top Then
14                    Reversed = True
15                End If
16            Else
                  ' Ranges are in different rows with no overlap; top one becomes Range1
17                If LHSRange.Left >= RHSRange.Left + RHSRange.Width Then
18                    Reversed = True
19                End If
20            End If
              
21            If Reversed Then
22                Set Range1 = RHSRange
23                Set Range2 = LHSRange
24            Else
25                Set Range1 = LHSRange
26                Set Range2 = RHSRange
27            End If
          
28            Set s1 = HighlightRange(Range1, vbNullString, Color)
          
              ' Reverse the sense if the objects are shown in the reverse order
29            Set s2 = HighlightRange(Range2, SolverRelationAsUnicodeChar(IIf(Reversed, 4 - sense, sense)), Color)
              
30            If Range1.Worksheet.Name = Range2.Worksheet.Name And _
                 Range1.Worksheet.Name = ActiveSheet.Name And _
                 Not s1 Is Nothing And _
                 Not s2 Is Nothing Then
31                AddLabelledConnector Range1.Worksheet, s1(1), s2(1), vbNullString
32            End If
33        Else 'this was added if there is only a lhs that needs highlighting in linearity
34            Set s1 = HighlightRange(LHSRange, vbNullString, Color)
35        End If

ExitSub:
36        If RaiseError Then RethrowError
37        Exit Sub

ErrorHandler:
38        If Not ReportError("OpenSolverVisualizer", "HighlightConstraint") Then Resume
39        RaiseError = True
40        GoTo ExitSub
End Sub

Sub DeleteOpenSolverShapes(w As Worksheet)
          Dim s As Shape
1         For Each s In w.Shapes
2             If s.Name Like "OpenSolver*" Then
3                 s.Delete
4             End If
5         Next s
End Sub

Function HideSolverModel(Optional sheet As Worksheet) As Boolean
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         HideSolverModel = False
          
4         Application.EnableCancelKey = xlErrorHandler
          
          Dim ScreenStatus As Boolean
5         ScreenStatus = Application.ScreenUpdating
6         Application.ScreenUpdating = False
          
7         GetActiveSheetIfMissing sheet
8         DeleteOpenSolverShapes sheet

          ' Delete constraints on other sheets
9         Dim i As Long
10        For i = 1 To GetNumConstraints(sheet)
              Dim b As Boolean, rLHS As Range
11            b = False
12            Set rLHS = Nothing
              ' Set b to be true only if there is no error
13            On Error Resume Next
14            Set rLHS = GetConstraintLhs(i, sheet)
15            b = rLHS.Worksheet.Name <> sheet.Name
16            If b Then
17                DeleteOpenSolverShapes rLHS.Worksheet
18            End If
NextConstraint:
19        Next i
20        On Error GoTo ErrorHandler
          
21        HideSolverModel = True

ExitFunction:
22        Application.StatusBar = False
23        Application.ScreenUpdating = ScreenStatus
24        If RaiseError Then RethrowError
25        Exit Function

ErrorHandler:
26        If Not ReportError("OpenSolverVisualizer", "HideSolverModel") Then Resume
27        RaiseError = True
28        GoTo ExitFunction
          
End Function

Function ShowSolverModel(Optional sheet As Worksheet, Optional HandleError As Boolean = False) As Boolean
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         ShowSolverModel = False

4         Application.EnableCancelKey = xlErrorHandler
          
          Dim ScreenStatus As Boolean
5         ScreenStatus = Application.ScreenUpdating
6         Application.ScreenUpdating = False

7         GetActiveSheetIfMissing sheet
8         DeleteOpenSolverShapes sheet
9         InitialiseHighlighting
          
          ' Checks to see if a model exists internally
10        Dim AdjustableCells As Range
11        On Error Resume Next
12        Set AdjustableCells = GetDecisionVariablesNoOverlap(sheet)
          ' Don't try to highlight if we have no vars
13        If AdjustableCells Is Nothing Then
14            ShowSolverModel = False
15            GoTo ExitFunction
16        End If
          
          ' Highlight the decision variables
17        AddDecisionVariableHighlighting AdjustableCells
          
          Dim Errors As String
          
          ' Highlight the objective cell, if there is one
          Dim ObjRange As Range
18        Set ObjRange = GetObjectiveFunctionCell(sheet, Validate:=False)
19        If Not ObjRange Is Nothing Then
              Dim ObjType As ObjectiveSenseType, ObjectiveTargetValue As Double
20            ObjType = GetObjectiveSense(sheet)
21            If ObjType = TargetObjective Then ObjectiveTargetValue = GetObjectiveTargetValue(sheet)
22            AddObjectiveHighlighting ObjRange, ObjType, ObjectiveTargetValue
23        End If
          
          ' Count the correct number of constraints, and form the constraint
          Dim NumConstraints As Long
24        NumConstraints = GetNumConstraints(sheet)  ' Number of constraints entered in excel; can include ranges covering many constraints
          ' Note: Solver leaves around old constraints; the name <sheet>!solver_num gives the correct number of constraints (eg "=4")
          
25        UpdateStatusBar "OpenSolver: Displaying Problem... " & AdjustableCells.Count & " vars, " & NumConstraints & " Solver constraints", True
                  
          Dim NumVars As Long
26        NumVars = AdjustableCells.Count
          Dim BinaryCellsRange As Range
          Dim IntegerCellsRange As Range
          Dim NonAdjustableCellsRange As Range
                  
          ' Count the correct number of constraints, and form the constraint
          Dim constraint As Long
          Dim currentSheet As Worksheet
27        For constraint = 1 To NumConstraints
              
              Dim rLHS As Range
28            Set rLHS = Nothing
29            On Error Resume Next
30            Set rLHS = GetConstraintLhs(constraint, sheet)
31            If Err.Number <> 0 Then
32                Errors = Error & "Error: " & Err.Description & vbNewLine
33                GoTo NextConstraint
34            End If
35            On Error GoTo ErrorHandler

              Dim rel As Long

36            rel = GetConstraintRel(constraint, sheet)
                      
              Dim LHSCount As Double, Count As Double
37            LHSCount = rLHS.Count
38            Count = LHSCount
              Dim AllDecisionVariables As Boolean
39            AllDecisionVariables = False

40            If rel = RelationINT Or rel = RelationBIN Then
                  ' Track all variables that are integer or binary
41                If rel = RelationINT Then
42                    Set IntegerCellsRange = ProperUnion(IntegerCellsRange, rLHS)
43                Else
44                    Set BinaryCellsRange = ProperUnion(BinaryCellsRange, rLHS)
45                End If
                  ' Keep track of all non-adjustable cells that are int/bin
46                Set NonAdjustableCellsRange = ProperUnion(NonAdjustableCellsRange, SetDifference(rLHS, AdjustableCells))
47            Else
                  ' Constraint is a full equation with a RHS
                  Dim valRHS As Double, rRHS As Range, sRefersToRHS As String, RefersToFormula As Boolean
48                Set rRHS = Nothing
49                On Error Resume Next
50                Set rRHS = GetConstraintRhs(constraint, sRefersToRHS, valRHS, RefersToFormula, sheet)
51                If Err.Number <> 0 Then
52                    Errors = Error & "Error: " & Err.Description & vbNewLine
53                    GoTo NextConstraint
54                End If
55                On Error GoTo ErrorHandler
                  
56                If rLHS.Worksheet.Name <> sheet.Name Then
57                    Set currentSheet = rLHS.Worksheet
58                Else
59                    Set currentSheet = sheet
60                End If
                  
61                If rRHS Is Nothing Then
62                    sRefersToRHS = ConvertToCurrentLocale(StripWorksheetNameAndDollars(sRefersToRHS, currentSheet))
63                End If
64                HighlightConstraint currentSheet, rLHS, rRHS, sRefersToRHS, rel, 0  ' Show either a value or a formula from sRefersToRHS

65            End If
NextConstraint:
66        Next constraint

67        Set IntegerCellsRange = SetDifference(IntegerCellsRange, BinaryCellsRange)

          ' Mark integer and binary variables
          Dim selectedArea As Range
68        If NumVars > 200 Then
69            AddBinaryIntegerBlockLabels BinaryCellsRange, "binary"
70            AddBinaryIntegerBlockLabels IntegerCellsRange, "integer"
71        Else
72            AddBinaryIntegerIndividualLabels BinaryCellsRange, "b"
73            AddBinaryIntegerIndividualLabels IntegerCellsRange, "i"
74        End If

          ' Mark non-decision variables with int or bin constraints
75        If Not NonAdjustableCellsRange Is Nothing Then
76            For Each selectedArea In NonAdjustableCellsRange.Areas
77                HighlightRange selectedArea, vbNullString, RGB(255, 255, 0), True  ' Yellow highlight
78            Next selectedArea
79        End If
          
80        ShowSolverModel = True  ' success

ExitFunction:
81        Application.StatusBar = False ' Resume normal status bar behaviour
82        Application.ScreenUpdating = ScreenStatus
83        If RaiseError Then RethrowError
84        Exit Function

ErrorHandler:
85        If Not ReportError("OpenSolverVisualizer", "ShowSolverModel") Then Resume
          
          ' Only show an error once per Excel instance
86        If HandleError And Not VisualizerHasShownError Then
87            VisualizerHasShownError = True
88            MsgBox "There was an error while showing the model. Please let us know about this so that we can investigate the issue.", vbOKOnly, "OpenSolver Visualizer Error"
89            GoTo ExitFunction
90        End If
          
91        RaiseError = True
92        GoTo ExitFunction
End Function

Sub AddObjectiveHighlighting(ObjectiveRange As Range, ObjectiveType As ObjectiveSenseType, ObjectiveTargetValue As Double)
          ' Highlight the cell
          Dim CellHighlight As ShapeRange
1         Set CellHighlight = HighlightRange(ObjectiveRange, vbNullString, RGB(255, 0, 255)) ' Magenta highlight
          
          ' Add the label
          Dim CellLabel As String
2         CellLabel = "??? "
3         If ObjectiveType = MaximiseObjective Then CellLabel = "max "
4         If ObjectiveType = MinimiseObjective Then CellLabel = "min "
5         If ObjectiveType = TargetObjective Then CellLabel = "seek " & ObjectiveTargetValue
6         AddLabelToShape ActiveSheet, CellHighlight(1), -6, 10, CellLabel, RGB(0, 0, 0) ' Black text
End Sub


Sub AddDecisionVariableHighlighting(DecisionVariableRange As Range)
          Dim Area As Range
1         For Each Area In DecisionVariableRange.Areas
2             HighlightRange Area, vbNullString, RGB(255, 0, 255), True ' Magenta highlight
3         Next Area
          
End Sub

Sub AddBinaryIntegerIndividualLabels(CellsRange As Range, Label As String)
          Dim c As Range
1         If Not CellsRange Is Nothing Then
2             For Each c In CellsRange
3                 AddLabelToRange ActiveSheet, c, 1, 9, Label, RGB(0, 0, 0)
4             Next c
5         End If
End Sub

Sub AddBinaryIntegerBlockLabels(CellsRange As Range, Label As String)
          Dim selectedArea As Range, CellHighlight As ShapeRange
1         If Not CellsRange Is Nothing Then
2             For Each selectedArea In CellsRange.Areas
3                 Set CellHighlight = HighlightRange(selectedArea, vbNullString, RGB(255, 0, 255)) ' Magenta highlight
4                 AddLabelToShape ActiveSheet, CellHighlight(1), -6, 10, Label, RGB(0, 0, 0) ' Black text
5             Next selectedArea
6         End If
End Sub
