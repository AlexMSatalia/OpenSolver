Attribute VB_Name = "OpenSolverVisualizer"
Option Explicit

Private ShapeIndex As Long
Private HighlightColorIndex As Long   ' Used to rotate thru colours on different constraints
Private colHighlightOffsets As Collection

Function NextHighlightColor() As Long
3013      HighlightColorIndex = HighlightColorIndex + 1
3014      If HighlightColorIndex > 7 Then HighlightColorIndex = 1
3015      Select Case HighlightColorIndex
              Case 1: NextHighlightColor = RGB(0, 0, 255) ' Blue
3016          Case 2: NextHighlightColor = RGB(0, 128, 0)  ' Green
3017          Case 3: NextHighlightColor = RGB(153, 0, 204) ' Purple
3018          Case 4: NextHighlightColor = RGB(128, 0, 0) ' Brown
3019          Case 5: NextHighlightColor = RGB(0, 204, 51) ' Light Green
3020          Case 6: NextHighlightColor = RGB(255, 102, 0) ' Orange
3021          Case 7: NextHighlightColor = RGB(204, 0, 153) ' Bright Purple
3022      End Select
End Function

Function NextHighlightColor_SkipColor(colorToSkip As Long) As Long
          Dim newColour As Long
3023      newColour = NextHighlightColor
3024      If newColour = colorToSkip Then
3025          newColour = NextHighlightColor
3026      End If
3027      NextHighlightColor_SkipColor = newColour
End Function

Function InitialiseHighlighting()
3028      Set colHighlightOffsets = New Collection
3029      ShapeIndex = 0
3030      HighlightColorIndex = 0 ' Used to rotate thru colours on different constraints
End Function

Function SheetHasOpenSolverHighlighting(w As Worksheet)
          ' If we have a shape called OpenSolver1 then we are displaying highlighted data
3031      If w.Shapes.Count = 0 Then GoTo NoHighlighting
          Dim s As Shape
3032      SheetHasOpenSolverHighlighting = True
3033      On Error Resume Next
3034      Set s = w.Shapes("OpenSolver" & 1) ' This string is split up to avoid false positives on anti-virus scans
3035      If Err.Number = 0 Then Exit Function ' Yes, we have highlighting
          ' Because the highlighting may be on another sheet, we also check all the shapes on this sheet
3038      For Each s In w.Shapes
3039          If s.Name Like "OpenSolver*" Then Exit Function
3040      Next s
NoHighlighting:
3041      SheetHasOpenSolverHighlighting = False
End Function

Function CreateLabelShape(w As Worksheet, left As Long, top As Long, width As Long, height As Long, label As String, HighlightColor As Long) As Shape
' Create a label (as a msoShapeRectangle) and give it text. This is used for labelling obj function as min/max, and decision vars as binary or integer
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
    
          Dim s1 As Shape
3042      Set s1 = w.Shapes.AddShape(msoShapeRectangle, left, top, width, height)
3043      s1.Fill.Visible = True
3044      s1.Fill.Solid
3045      s1.Fill.ForeColor.RGB = RGB(255, 255, 255)
3046      s1.Fill.Transparency = 0.2
3047      s1.Line.Visible = False
3048      s1.Shadow.Visible = msoFalse
3049      With s1.TextFrame
3050          .Characters.Text = label
3051          .Characters.Font.Size = 9
3052          .Characters.Font.Color = HighlightColor
3053          .HorizontalAlignment = xlHAlignLeft
3054          .VerticalAlignment = xlVAlignTop
3055          .MarginBottom = 0
3056          .MarginLeft = 1
3057          .MarginRight = 1
3058          .MarginTop = 0
3059          .AutoSize = True     ' Get width correct
3060          .AutoSize = False
3061      End With
3062      ShapeIndex = ShapeIndex + 1
3063      s1.Name = "OpenSolver" & ShapeIndex
3064      s1.height = height      ' Force the specified height
3065      Set CreateLabelShape = s1

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverVisualizer", "CreateLabelShape") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function AddLabelToRange(w As Worksheet, r As Range, voffset As Long, height As Long, label As String, HighlightColor As Long) As Shape
3066      Set AddLabelToRange = CreateLabelShape(w, r.left + 1, r.top + voffset, r.width, height, label, HighlightColor)
End Function

Function AddLabelToShape(w As Worksheet, s As Shape, voffset As Long, height As Long, label As String, HighlightColor As Long)
3067      Set AddLabelToShape = CreateLabelShape(w, s.left - 1, s.top + voffset, s.width, height, label, HighlightColor)
End Function

Function HighlightRange(r As Range, label As String, HighlightColor As Long, Optional ShowFill As Boolean = False, Optional ShapeNamePrefix As String = "OpenSolver", Optional Bounds As Boolean = False) As ShapeRange
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          'Only show the model file that is on the active sheet to overcome hide bugs for shapes on different sheets
3068      If r.Worksheet.Name <> ActiveSheet.Name Then GoTo ExitFunction
          
          Const HighlightingOffsetStep = 1
          ' We offset our highlighting so that successive highlights are still visible
          Dim Offset As Double, Key As String
3069      Key = r.Address(RowAbsolute:=False, ColumnAbsolute:=False)
3070      Offset = 0
3071      On Error Resume Next
3072      Offset = colHighlightOffsets(Key) ' eg A1
3073      If Err.Number <> 0 Then
              ' Item does not exist in collection create it, with an offset of 2 ready to use next time
3074          colHighlightOffsets.Add HighlightingOffsetStep, Key
3075      Else
3076          colHighlightOffsets.Remove Key
3077          colHighlightOffsets.Add Offset + HighlightingOffsetStep, Key
3078      End If
3079      On Error GoTo ErrorHandler
          
          ' Handle merged cells
          Dim l As Single, t As Single, right As Single, bottom As Single, R2 As Range, c As Range
3080      If Not r.MergeCells Then
3081          l = r.left
3082          t = r.top
3083          right = l + r.width
3084          bottom = t + r.height
3085      Else
              ' This range contains merged cells. We use MergeArea to to find the area to highlight
              ' But, this only works for one cell, so we expand our size based on all cells
3086          If r.Count = 1 Then
3087              Set R2 = r.MergeArea
3088              l = R2.left
3089              t = R2.top
3090              right = l + R2.width
3091              bottom = t + R2.height
3092          Else
3093              l = r.left
3094              t = r.top
3095              right = l + r.width
3096              bottom = t + r.height
3097              For Each c In r
3098                  If c.MergeCells Then
3099                      Set R2 = c.MergeArea
3100                      If R2.left < l Then l = R2.left
3101                      If R2.top < t Then t = R2.top
3102                      If R2.left + R2.width > right Then right = R2.left + R2.width
3103                      If R2.top + R2.height > bottom Then bottom = R2.top + R2.height
3104                  End If
3105              Next c
3106          End If
3107      End If
          
          
          ' Use doubles here for more accuracy as we are summing terms, and so accummulating errors
          Dim left2 As Double, top2 As Double, right2 As Double, bottom2 As Double, height As Double, width As Double
3108      left2 = l + Offset
3109      top2 = t + Offset
          
          ' Draw enough shapes to cover the space; each shape has an Excel-set (undocumented?) maximum height (and width?)
          Dim isFirstShape As Boolean
3110      isFirstShape = True
          Dim firstShapeIndex As Long
3111      firstShapeIndex = ShapeIndex + 1
3112      Do
          ' The height cannot exceed 169056.0; we allow some tolerance
3113      height = bottom - top2
3114      If height > 160000# Then
3115          height = 150000  ' This difference ensures we never end up with very small rectangle
3116      End If
3117      If isFirstShape And height > 9500 Then
3118          height = 9000   ' The first shape we create has a height of 9000 to ensure we can rotate the text and have it show
                            ' correctly; this works around and Excel 2007 bug
3119      End If
3120      isFirstShape = False
        
          Dim shapeName As String, s1 As Shape
3121      ShapeIndex = ShapeIndex + 1
          'If the constraint is not a bound then make a box for it
3122      If Not Bounds Then
3123          shapeName = ShapeNamePrefix & ShapeIndex
3124          r.Worksheet.Shapes.AddShape(msoShapeRectangle, l + Offset, top2, right - l, height).Name = shapeName
3125      Else
              'If the box is a bound we name it after the cell that it is in
3126          shapeName = ShapeNamePrefix & Key
3127          On Error Resume Next
              Dim tmpName As String
3128          tmpName = r.Worksheet.Shapes(shapeName).Name
              On Error GoTo ErrorHandler
              'If there hasn't been a bound on that cell then make a new cell
3129          If tmpName = "" Then
3130              r.Worksheet.Shapes.AddShape(msoShapeRectangle, l + Offset, top2, right - l, height).Name = shapeName
3131          Else
                  'If there has already been a bound then just add new text to it rather then making a new box
3132              Set s1 = r.Worksheet.Shapes(shapeName)
3133              s1.TextFrame.Characters.Text = s1.TextFrame.Characters.Text & "," & label
3134              GoTo endLoop
3135          End If
3136      End If
        
3137      Set s1 = r.Worksheet.Shapes(shapeName)
              
          Dim ShowOutline As Boolean
3138      ShowOutline = Not ShowFill
        
3139      If ShowOutline Then
3140          s1.Fill.Visible = False
3141          With s1.Line
3142              .Weight = 2
3143              .ForeColor.RGB = HighlightColor
3144          End With
3145      Else
3146          s1.Line.Visible = False
3147          s1.Fill.Solid
3148          s1.Fill.Transparency = 0.6
3149          s1.Fill.ForeColor.RGB = HighlightColor
3150      End If
3151      s1.Shadow.Visible = msoFalse
        
3152      With s1.TextFrame
3153          .Characters.Text = label
3154          .Characters.Font.Color = HighlightColor
3155          .HorizontalAlignment = xlHAlignLeft ' xlHAlignCenter
              ' "=", "<=", & ">=" will be centered
3156          If ((height < 500) Or (label = "=") Or (label = ChrW(&H2265)) Or (label = ChrW(&H2264))) Then
3157              .VerticalAlignment = xlVAlignCenter  ' Shape is small enought to have text fit on the screen when centered, so we center text
3158          Else
3159              .VerticalAlignment = xlVAlignTop   ' So we can see the name when scrolled to the top
3160          End If
3161          .MarginBottom = 0
3162          .MarginLeft = 2
3163          .MarginRight = 0
3164          .MarginTop = 2
3165          .Characters.Font.Bold = True
3166      End With
endLoop:
3167      top2 = top2 + height
3168      Loop While bottom - top2 > 0.01  ' handle float rounding
          
          ' Create & return the shapeRange containing all the shapes we added
          Dim shapeNames(), i As Long
3169      ReDim shapeNames(ShapeIndex - firstShapeIndex + 1)
3170      If Not Bounds Then
3171          For i = firstShapeIndex To ShapeIndex
3172              shapeNames(i - firstShapeIndex + 1) = ShapeNamePrefix & i
3173          Next i
3174      Else
3175          For i = firstShapeIndex To ShapeIndex
3176              shapeNames(i - firstShapeIndex + 1) = ShapeNamePrefix & Key
3177          Next i
3178      End If
3179      Set HighlightRange = r.Worksheet.Shapes.Range(shapeNames)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverVisualizer", "HighlightRange") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function AddLabelledConnector(w As Worksheet, s1 As Shape, s2 As Shape, label As String)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim t As Shapes, c As Shape
3180      Set t = w.Shapes
3181      Set c = t.AddConnector(msoConnectorStraight, 0, 0, 0, 0) ' msoConnectorCurve
3182      With c.ConnectorFormat
3183          .BeginConnect ConnectedShape:=s1, ConnectionSite:=1
3184          .EndConnect ConnectedShape:=s2, ConnectionSite:=1
3185      End With
3186      c.RerouteConnections
3187      c.Line.ForeColor = s1.Line.ForeColor
          ' The default styles can be changed, so we should set everything! We just do a few of the problem bits
3188      c.Line.EndArrowheadStyle = msoArrowheadNone
3189      c.Line.BeginArrowheadStyle = msoArrowheadNone
3190      c.Line.DashStyle = msoLineSolid
3191      c.Line.Weight = 0.75
3192      c.Line.Style = msoLineSingle
3193      c.Shadow.Visible = msoFalse
              
3194      ShapeIndex = ShapeIndex + 1
3195      c.Name = "OpenSolver" & ShapeIndex
          
          Dim s3 As Shape
3196      Set s3 = t.AddShape(msoShapeRectangle, c.left + c.width / 2# - 30 / 2, c.top + c.height / 2# - 20 / 2, 30, 20)
                  
3197      s3.Line.Visible = False
3198      s3.Fill.Visible = False
3199      s3.Shadow.Visible = msoFalse
3200      If label <> "" Then
3201          With s3.TextFrame
3202              .Characters.Text = label
3203              .MarginBottom = 0
3204              .MarginLeft = 0
3205              .MarginRight = 0
3206              .MarginTop = 0
3207              .HorizontalAlignment = xlHAlignCenter
3208              .VerticalAlignment = xlVAlignCenter
3209              .Characters.Font.Color = c.Line.ForeColor
3210              .Characters.Font.Bold = True
3211          End With
3212      End If

3213      ShapeIndex = ShapeIndex + 1
3214      s3.Name = "OpenSolver" & ShapeIndex

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverVisualizer", "AddLabelledConnector") Then Resume
          RaiseError = True
          GoTo ExitFunction

End Function

Sub HighlightConstraint(myDocument As Worksheet, LHSRange As Range, _
                        RHSRange As Range, ByVal RHSValue As String, ByVal Sense As Long, _
                        ByVal Color As Long)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          ' Show a constraint of the form LHS <|=|> RHS.
          ' We always put the sign in the rightmost (or bottom-most) range so we read left-to-right or top-to-bottom.
          Dim Range1 As Range, Range2 As Range, Reversed As Boolean
          Dim s1 As ShapeRange, s2 As ShapeRange

          ' Get next color if none specified
3215      If Color = 0 Then Color = NextHighlightColor
          
3216      Reversed = False
3217      If RHSRange Is Nothing And RHSValue <> "" Then
              ' We have a constant or formula in the constraint. Put into form RHS <|=|> Range1 (reversing the sense)
3218          Set s1 = HighlightRange(LHSRange, RHSValue & SolverRelationAsUnicodeChar(4 - Sense), Color, , , True)
3219      ElseIf Not RHSRange Is Nothing Then
              ' If ranges overlaps on rows, then the top one becomes Range1
3220          If ((RHSRange.top >= LHSRange.top And RHSRange.top < LHSRange.top + LHSRange.height) _
              Or (LHSRange.top >= RHSRange.top And LHSRange.top < RHSRange.top + RHSRange.height)) Then
                  ' Ranges over lap in rows. Range1 becomes the left most one
3221              If LHSRange.left > RHSRange.left Then
3222                  Reversed = True
3223              End If
3224          ElseIf ((RHSRange.left >= LHSRange.left And RHSRange.left < LHSRange.left + LHSRange.width) _
              Or (LHSRange.left >= RHSRange.left And LHSRange.left < RHSRange.left + RHSRange.width)) Then
                  ' Ranges overlap in columns. Range1 becomes the top most one
3225              If LHSRange.top > RHSRange.top Then
3226                  Reversed = True
3227              End If
3228          Else
                  ' Ranges are in different rows with no overlap; top one becomes Range1
3229              If LHSRange.left >= RHSRange.left + RHSRange.width Then
3230                  Reversed = True
3231              End If
3232          End If
              
3233          If Reversed Then
3234              Set Range1 = RHSRange
3235              Set Range2 = LHSRange
3236          Else
3237              Set Range1 = LHSRange
3238              Set Range2 = RHSRange
3239          End If
          
3240          Set s1 = HighlightRange(Range1, "", Color)
          
              ' Reverse the sense if the objects are shown in the reverse order
3241          Set s2 = HighlightRange(Range2, SolverRelationAsUnicodeChar(IIf(Reversed, 4 - Sense, Sense)), Color)
              
3242          If Range1.Worksheet.Name = Range2.Worksheet.Name And Range1.Worksheet.Name = ActiveSheet.Name Then
3243              AddLabelledConnector Range1.Worksheet, s1(1), s2(1), ""
3244          End If
3245      Else 'this was added if there is only a lhs that needs highlighting in linearity
3246          Set s1 = HighlightRange(LHSRange, "", Color)
3247      End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("OpenSolverVisualizer", "HighlightConstraint") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Sub DeleteOpenSolverShapes(w As Worksheet)
          Dim s As Shape
3248      For Each s In w.Shapes
3249          If s.Name Like "OpenSolver*" Then
3250              s.Delete
3251          End If
3252      Next s
End Sub

Function HideSolverModel() As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

3253      HideSolverModel = False
          
3254      If Application.Workbooks.Count = 0 Then
3255          MsgBox "Error: No active workbook available", , "OpenSolver" & sOpenSolverVersion & " Error"
3256          GoTo ExitFunction
3257      End If
          
3258      Application.EnableCancelKey = xlErrorHandler
          
          Dim ScreenStatus As Boolean
          ScreenStatus = Application.ScreenUpdating
3260      Application.ScreenUpdating = False
          
          Dim sheet As Worksheet
3261      On Error Resume Next
          
3262      Set sheet = ActiveWorkbook.ActiveSheet
3263      If Err.Number <> 0 Then
3264          MsgBox "Error: Unable to access the active sheet", , "OpenSolver" & sOpenSolverVersion & " Error"
3265          GoTo ExitFunction
3266      End If
3267      On Error GoTo ErrorHandler
          
3268      DeleteOpenSolverShapes sheet
          Dim i As Long
          Dim NumOfConstraints As Long
          Dim sheetName As String
3269      sheetName = EscapeSheetName(ActiveWorkbook.ActiveSheet)
          
          ' Delete constraints on other sheets
          Dim b As Boolean, rLHS As Range
3273      For i = 1 To GetNumConstraints()
3274          b = False
              Set rLHS = Nothing
              ' Set b to be true only if there is no error
3275          On Error Resume Next
              Set rLHS = GetConstraintLhs(i)
3276          b = rLHS.Worksheet.Name <> ActiveWorkbook.ActiveSheet.Name
3277          If b Then
3278              DeleteOpenSolverShapes rLHS.Worksheet
3279          End If
NextConstraint:
3280      Next i
3281      On Error GoTo ErrorHandler
          
3282      HideSolverModel = True

ExitFunction:
3283      Application.StatusBar = False
3284      Application.ScreenUpdating = ScreenStatus
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverVisualizer", "HideSolverModel") Then Resume
          RaiseError = True
          GoTo ExitFunction
          
End Function

Function ShowSolverModel() As Boolean
          
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

3295      ShowSolverModel = False
          
3296      If Application.Workbooks.Count = 0 Then
3297          MsgBox "Error: No active workbook available", , "OpenSolver" & sOpenSolverVersion & " Error"
3298          GoTo ExitFunction
3299      End If

          ' Trap the Escape key
3300      Application.EnableCancelKey = xlErrorHandler
          
          Dim ScreenStatus As Boolean
          ScreenStatus = Application.ScreenUpdating
3302      Application.ScreenUpdating = False
          
          Dim sheetName As String, book As Workbook, sheet As Worksheet
          Dim NumConstraints  As Long

3303      On Error Resume Next
3304      sheetName = EscapeSheetName(ActiveWorkbook.ActiveSheet)
3305      If Err.Number <> 0 Then
3306          MsgBox "Error: Unable to access the active sheet", , "OpenSolver" & sOpenSolverVersion & " Error"
3307          GoTo ExitFunction
3308      End If
3309      On Error GoTo ErrorHandler
          
3310      Set book = ActiveWorkbook
          Set sheet = book.ActiveSheet
          
3311      DeleteOpenSolverShapes ActiveSheet
3312      InitialiseHighlighting
          
          ' Checks to see if a model exists internally
3316      Dim AdjustableCells As Range
          On Error Resume Next
          Set AdjustableCells = GetDecisionVariablesNoOverlap(book, sheet)
          ' Don't try to highlight if we have no vars
          If AdjustableCells Is Nothing Then
              ShowSolverModel = False
              GoTo ExitFunction
          End If
          
          ' Highlight the decision variables
3322      AddDecisionVariableHighlighting AdjustableCells
          
          Dim Errors As String
          
          ' Highlight the objective cell, if there is one
          Dim ObjRange As Range
3323      Set ObjRange = GetObjectiveFunctionCell(book, sheet)
          If Not ObjRange Is Nothing Then
              Dim ObjType As ObjectiveSenseType, ObjectiveTargetValue As Double
3325          ObjType = GetObjectiveSense(book, sheet)
3327          If ObjType = TargetObjective Then ObjectiveTargetValue = GetObjectiveTargetValue(book, sheet)
3328          AddObjectiveHighlighting ObjRange, ObjType, ObjectiveTargetValue
3330      End If
          
          ' Count the correct number of constraints, and form the constraint
3331      NumConstraints = GetNumConstraints()  ' Number of constraints entered in excel; can include ranges covering many constraints
          ' Note: Solver leaves around old constraints; the name <sheet>!solver_num gives the correct number of constraints (eg "=4")
          
3332      UpdateStatusBar "OpenSolver: Displaying Problem... " & AdjustableCells.Count & " vars, " & NumConstraints & " Solver constraints", True
                  
          Dim numVars As Long
3333      numVars = AdjustableCells.Count
          Dim BinaryCellsRange As Range
          Dim IntegerCellsRange As Range
          Dim NonAdjustableCellsRange As Range
                  
          ' Count the correct number of constraints, and form the constraint
          Dim constraint As Long
          Dim currentSheet As Worksheet
3334      For constraint = 1 To NumConstraints
              
              Dim rLHS As Range
              Set rLHS = Nothing
              On Error Resume Next
              Set rLHS = GetConstraintLhs(constraint, book, sheet)
              If Err.Number <> 0 Then
                  Errors = Error & "Error: " & Err.Description & vbNewLine
                  GoTo NextConstraint
              End If
              On Error GoTo ErrorHandler

              Dim rel As Long

3345          rel = GetConstraintRel(constraint, book, sheet)
                      
              Dim LHSCount As Double, Count As Double
3346          LHSCount = rLHS.Count
3347          Count = LHSCount
              Dim AllDecisionVariables As Boolean
3348          AllDecisionVariables = False

3349          If rel = RelationINT Or rel = RelationBIN Then
                  ' Track all variables that are integer or binary
3356              If rel = RelationINT Then
3357                  Set IntegerCellsRange = ProperUnion(IntegerCellsRange, rLHS)
3362              Else
3363                  Set BinaryCellsRange = ProperUnion(BinaryCellsRange, rLHS)
3368              End If
                  ' Keep track of all non-adjustable cells that are int/bin
                  Set NonAdjustableCellsRange = ProperUnion(NonAdjustableCellsRange, SetDifference(rLHS, AdjustableCells))
3373          Else
                  ' Constraint is a full equation with a RHS
                  Dim valRHS As Double, rRHS As Range, sRefersToRHS As String, RefersToFormula As Boolean
3374              Set rRHS = Nothing
                  On Error Resume Next
                  Set rRHS = GetConstraintRhs(constraint, sRefersToRHS, valRHS, RefersToFormula, book, sheet)
                  If Err.Number <> 0 Then
                      Errors = Error & "Error: " & Err.Description & vbNewLine
                      GoTo NextConstraint
                  End If
                  On Error GoTo ErrorHandler
                  
3384              If rLHS.Worksheet.Name <> ActiveWorkbook.ActiveSheet.Name Then
3385                  Set currentSheet = rLHS.Worksheet
3386              Else
3387                  Set currentSheet = ActiveSheet
3388              End If
                  
3389              If rRHS Is Nothing Then
3390                  sRefersToRHS = ConvertToCurrentLocale(StripWorksheetNameAndDollars(sRefersToRHS, currentSheet))
3393              End If
3394              HighlightConstraint currentSheet, rLHS, rRHS, sRefersToRHS, rel, 0  ' Show either a value or a formula from sRefersToRHS

3395          End If
NextConstraint:
3396      Next constraint

          Set IntegerCellsRange = SetDifference(IntegerCellsRange, BinaryCellsRange)

          ' Mark integer and binary variables
          Dim selectedArea As Range
3399      If numVars > 200 Then
3400          AddBinaryIntegerBlockLabels BinaryCellsRange, "binary"
3406          AddBinaryIntegerBlockLabels IntegerCellsRange, "integer"
3421      Else
3422          AddBinaryIntegerIndividualLabels BinaryCellsRange, "b"
              AddBinaryIntegerIndividualLabels IntegerCellsRange, "i"
3438      End If

          ' Mark non-decision variables with int or bin constraints
          If Not NonAdjustableCellsRange Is Nothing Then
              For Each selectedArea In NonAdjustableCellsRange.Areas
                  HighlightRange selectedArea, "", RGB(255, 255, 0), True  ' Yellow highlight
              Next selectedArea
          End If
          
3443      ShowSolverModel = True  ' success

ExitFunction:
3444      Application.StatusBar = False ' Resume normal status bar behaviour
3445      Application.ScreenUpdating = ScreenStatus
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverVisualizer", "ShowSolverModel") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Sub AddObjectiveHighlighting(ObjectiveRange As Range, ObjectiveType As ObjectiveSenseType, ObjectiveTargetValue As Double)
          ' Highlight the cell
          Dim CellHighlight As ShapeRange
3465      Set CellHighlight = HighlightRange(ObjectiveRange, "", RGB(255, 0, 255)) ' Magenta highlight
          
          ' Add the label
          Dim CellLabel As String
3466      CellLabel = "??? "
3467      If ObjectiveType = MaximiseObjective Then CellLabel = "max "
3468      If ObjectiveType = MinimiseObjective Then CellLabel = "min "
3469      If ObjectiveType = TargetObjective Then CellLabel = "seek " & ObjectiveTargetValue
3470      AddLabelToShape ActiveSheet, CellHighlight(1), -6, 10, CellLabel, RGB(0, 0, 0) ' Black text
End Sub


Sub AddDecisionVariableHighlighting(DecisionVariableRange As Range)
          Dim area As Range
3471      For Each area In DecisionVariableRange.Areas
3472          HighlightRange area, "", RGB(255, 0, 255), True ' Magenta highlight
3473      Next area
          
End Sub

Sub AddBinaryIntegerIndividualLabels(CellsRange As Range, label As String)
    Dim c As Range
    If Not CellsRange Is Nothing Then
        For Each c In CellsRange
            AddLabelToRange ActiveSheet, c, 1, 9, label, RGB(0, 0, 0)
        Next c
    End If
End Sub

Sub AddBinaryIntegerBlockLabels(CellsRange As Range, label As String)
    Dim selectedArea As Range, CellHighlight As ShapeRange
    If Not CellsRange Is Nothing Then
        For Each selectedArea In CellsRange.Areas
            Set CellHighlight = HighlightRange(selectedArea, "", RGB(255, 0, 255)) ' Magenta highlight
            AddLabelToShape ActiveSheet, CellHighlight(1), -6, 10, label, RGB(0, 0, 0) ' Black text
        Next selectedArea
    End If
End Sub
