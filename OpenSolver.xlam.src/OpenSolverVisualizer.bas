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
3034      Set s = w.Shapes("OpenSolver1")
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
End Function

Function AddLabelToRange(w As Worksheet, r As Range, voffset As Long, height As Long, label As String, HighlightColor As Long) As Shape
3066      Set AddLabelToRange = CreateLabelShape(w, r.left + 1, r.top + voffset, r.width, height, label, HighlightColor)
End Function

Function AddLabelToShape(w As Worksheet, s As Shape, voffset As Long, height As Long, label As String, HighlightColor As Long)
3067      Set AddLabelToShape = CreateLabelShape(w, s.left - 1, s.top + voffset, s.width, height, label, HighlightColor)
End Function

Function HighlightRange(r As Range, label As String, HighlightColor As Long, Optional ShowFill As Boolean = False, Optional ShapeNamePrefix As String = "OpenSolver", Optional Bounds As Boolean = False) As ShapeRange

          'Only show the model file that is on the active sheet to overcome hide bugs for shapes on different sheets
3068      If r.Worksheet.Name <> ActiveSheet.Name Then Exit Function
          
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
3079      On Error GoTo 0
          
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
End Function

Function AddLabelledConnector(w As Worksheet, s1 As Shape, s2 As Shape, label As String)
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
3195      c.Name = "OpenSolver " & ShapeIndex
          
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
3214      s3.Name = "OpenSolver " & ShapeIndex

End Function

Sub HighlightConstraint(myDocument As Worksheet, LHSRange As Range, _
                        ByVal RHSisRange As Boolean, RHSRange As Range, ByVal RHSValue As String, ByVal Sense As Long, _
                        ByVal Color As Long)
          ' Show a constraint of the form LHS <|=|> RHS.
          ' We always put the sign in the rightmost (or bottom-most) range so we read left-to-right or top-to-bottom.
          Dim Range1 As Range, Range2 As Range, Reversed As Boolean
          Dim s1 As ShapeRange, s2 As ShapeRange

          ' Get next color if none specified
3215      If Color = 0 Then Color = NextHighlightColor
          
3216      Reversed = False
3217      If Not RHSisRange And RHSValue <> "" Then
              ' We have a constant or formula in the constraint. Put into form RHS <|=|> Range1 (reversing the sense)
3218          Set s1 = HighlightRange(LHSRange, RHSValue & SolverRelationAsUnicodeChar(4 - Sense), Color, , , True)
3219      ElseIf RHSisRange Then
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

3253      HideSolverModel = False
          
3254      If Application.Workbooks.Count = 0 Then
3255          MsgBox "Error: No active workbook available", , "OpenSolver" & sOpenSolverVersion & " Error"
3256          Exit Function
3257      End If
          
          ' Trap the Escape key
3258      Application.EnableCancelKey = xlErrorHandler
3259      On Error GoTo errorHandler
          
3260      Application.ScreenUpdating = False
          
          Dim sheet As Worksheet
3261      On Error Resume Next
          
3262      Set sheet = ActiveWorkbook.ActiveSheet
3263      If Err.Number <> 0 Then
3264          MsgBox "Error: Unable to access the active sheet", , "OpenSolver" & sOpenSolverVersion & " Error"
3265          Exit Function
3266      End If
3267      On Error GoTo errorHandler
          
3268      DeleteOpenSolverShapes sheet
          Dim i As Long
          Dim NumOfConstraints As Long
          Dim sheetName As String
3269      sheetName = EscapeSheetName(ActiveWorkbook.ActiveSheet)
3270      On Error Resume Next ' There may not be a model on the sheet
3271      NumOfConstraints = Mid(Names(sheetName & "solver_num"), 2)
3272      On Error GoTo errorHandler
          
          ' Delete constraints on other sheets
          Dim b As Boolean
3273      For i = 1 To NumOfConstraints
              ' This code used to say On Error goto NextConstraint, but this failed because the error was never Resume'd
3274          b = False
              ' Set b to be true only if there is no error
3275          On Error Resume Next
3276          b = Range(sheetName & "solver_lhs" & i).Worksheet.Name <> ActiveWorkbook.ActiveSheet.Name
3277          If b Then
3278              DeleteOpenSolverShapes Range(sheetName & "solver_lhs" & i).Worksheet
3279          End If
NextConstraint:
3280      Next i
3281      On Error GoTo errorHandler
          
3282      HideSolverModel = True

ExitSub:
3283      Application.StatusBar = False
3284      Application.ScreenUpdating = True
          
3285      Exit Function
          
errorHandler:
3286      If Err.Number = 18 Then
3287          If MsgBox("You have pressed the Escape key. Do you wish to cancel?", _
                         vbCritical + vbYesNo + vbDefaultButton1, _
                         "OpenSolver: User Interrupt Occured...") = vbNo Then
3288              Resume
3289          Else
3290              Resume ExitSub
3291          End If
3292      End If
3293      MsgBox "OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source, , "OpenSolver" & sOpenSolverVersion & " Error"
3294      Resume ExitSub
          
End Function

Function ShowSolverModel() As Boolean
          
3295      ShowSolverModel = False
          
3296      If Application.Workbooks.Count = 0 Then
3297          MsgBox "Error: No active workbook available", , "OpenSolver" & sOpenSolverVersion & " Error"
3298          Exit Function
3299      End If

          ' Trap the Escape key
3300      Application.EnableCancelKey = xlErrorHandler
3301      On Error GoTo errorHandler
          
3302      Application.ScreenUpdating = False
          
          Dim i As Double, sheetName As String, book As Workbook, AdjustableCells As Range
          Dim NumConstraints  As Long

3303      On Error Resume Next
3304      sheetName = EscapeSheetName(ActiveWorkbook.ActiveSheet)
3305      If Err.Number <> 0 Then
3306          MsgBox "Error: Unable to access the active sheet", , "OpenSolver" & sOpenSolverVersion & " Error"
3307          Exit Function
3308      End If
3309      On Error GoTo errorHandler
          
3310      Set book = ActiveWorkbook
          
3311      DeleteOpenSolverShapes ActiveSheet
3312      InitialiseHighlighting
          
          ' We check to see if a model exists by getting the adjustable cells. We check for a name first, as this may contain =Sheet1!$C$2:$E$2,Sheet1!#REF!
          Dim n As Name
3313      On Error Resume Next
3314      Set n = Names(sheetName & "solver_adj")
3315      On Error Resume Next
3316      Set AdjustableCells = RemoveRangeOverlap(Range(sheetName & "solver_adj"))   ' Remove any overlap in the range defining the decision variables
3317      If Err.Number <> 0 Then
3318          MsgBox "Error: A model was found on the sheet " & ActiveWorkbook.ActiveSheet.Name & " but the decision variable cells (" & n & ") could not be interpreted. Please redefine the decision variable cells, and try again.", , "OpenSolver" & sOpenSolverVersion & " Error"
3319          GoTo ExitSub
3320      End If
3321      On Error GoTo errorHandler
          
          ' Highlight the decision variables
3322      AddDecisionVariableHighlighting AdjustableCells
          
          Dim Errors As String
          
          ' Highlight the objective cell, if there is one
          Dim ObjRange As Range
3323      If GetNamedRangeIfExists(book, sheetName & "solver_opt", ObjRange) Then
3324          Set ObjRange = Range(sheetName & "solver_opt")
              Dim ObjType As ObjectiveSenseType, temp As Long, ObjectiveTargetValue As Double
3325          ObjType = UnknownObjectiveSense
3326          If GetNamedIntegerIfExists(book, sheetName & "solver_typ", temp) Then ObjType = temp
3327          If ObjType = TargetObjective Then GetNamedNumericValueIfExists book, sheetName & "solver_val", ObjectiveTargetValue
3328          AddObjectiveHighlighting ObjRange, ObjType, ObjectiveTargetValue
3330      End If
          
          ' Count the correct number of constraints, and form the constraint
3331      NumConstraints = Mid(Names(sheetName & "solver_num"), 2)  ' Number of constraints entered in excel; can include ranges covering many constraints
          ' Note: Solver leaves around old constraints; the name <sheet>!solver_num gives the correct number of constraints (eg "=4")
          
3332      Application.StatusBar = "OpenSolver: Displaying Problem... " & AdjustableCells.Count & " vars, " & NumConstraints & " Solver constraints"
                  
          ' Process the decision variables as we need to compute their types (bin or int; a variable can be declared as both!)
          Dim numVars As Long
3333      numVars = AdjustableCells.Count
          Dim BinaryCellsRange As Range
          Dim IntegerCellsRange As Range
          Dim NonAdjustableCellsRange As Range
          ' Get names for all the variables so we can track their types
          Dim c As Range
                  
          ' Count the correct number of constraints, and form the constraint
          Dim constraint As Long
          Dim currentSheet As Worksheet
3334      For constraint = 1 To NumConstraints
              
              Dim isRangeLHS As Boolean, valLHS As Double, rLHS As Range, RangeRefersToError As Boolean, RefersToFormula As Boolean, sNameLHS As String, sRefersToLHS As String, isMissingLHS As Boolean
3335          sNameLHS = sheetName & "solver_lhs" & constraint
3336          GetNameAsValueOrRange book, sNameLHS, isMissingLHS, isRangeLHS, rLHS, RefersToFormula, RangeRefersToError, sRefersToLHS, valLHS
3337          If isMissingLHS Then
3338              Errors = Errors & "Error: The left hand side of constraint " & constraint & " is not defined (no 'solver_lhs" & constraint & "')." & vbCrLf
3339              GoTo NextConstraint
3340          End If
3341          If Not isRangeLHS Or RangeRefersToError Or RefersToFormula Then
3342              Errors = Errors & "Error: Range " & book.Names(sNameLHS).RefersTo & " is not a valid left hand side for a constraint." & vbCrLf
3343              GoTo NextConstraint
3344          End If

              Dim rel As Long

3345          rel = Mid(Names(sheetName & "solver_rel" & constraint), 2)
                      
              Dim LHSCount As Double, Count As Double
3346          LHSCount = rLHS.Count
3347          Count = LHSCount
              Dim AllDecisionVariables As Boolean
3348          AllDecisionVariables = False

3349          If rel = RelationINT Or rel = RelationBIN Then
              ' Make the LHS variables integer or binary
3356              If rel = RelationINT Then
3357                  Set IntegerCellsRange = ProperUnion(IntegerCellsRange, rLHS)
3362              Else
3363                  Set BinaryCellsRange = ProperUnion(BinaryCellsRange, rLHS)
3368              End If
                  ' Keep track of all non-adjustable cells that are int/bin
                  Set NonAdjustableCellsRange = ProperUnion(NonAdjustableCellsRange, SetDifference(rLHS, AdjustableCells))
3373          Else
                  ' Constraint is a full equation with a RHS
                  Dim isRangeRHS As Boolean, valRHS As Double, rRHS As Range, sNameRHS As String, sRefersToRHS As String, isMissingRHS As Boolean
3374              sNameRHS = sheetName & "solver_rhs" & constraint
3375              GetNameAsValueOrRange book, sNameRHS, isMissingRHS, isRangeRHS, rRHS, RefersToFormula, RangeRefersToError, sRefersToRHS, valRHS
3376              If isMissingRHS Then
3377                  Errors = Errors & "Error: The right hand side of constraint " & constraint & " is not defined (no 'solver_rhs" & constraint & "')." & vbCrLf
3378                  GoTo NextConstraint
3379              End If
3380              If RangeRefersToError Then
3381                  Errors = Errors & "Error: Range " & Mid(book.Names(sNameRHS), 2) & " is not a valid right hand side in a constraint." & vbCrLf
3382                  GoTo NextConstraint
3383              End If
                  
3384              If Range(sheetName & "solver_lhs" & constraint).Worksheet.Name <> ActiveWorkbook.ActiveSheet.Name Then
3385                  Set currentSheet = Range(sheetName & "solver_lhs" & constraint).Worksheet
3386              Else
3387                  Set currentSheet = ActiveSheet
3388              End If
                  
3389              If RefersToFormula Then
3390                  sRefersToRHS = ConvertToCurrentLocale(StripWorksheetNameAndDollars(sRefersToRHS, currentSheet))
3393              End If
3394              HighlightConstraint currentSheet, rLHS, isRangeRHS, rRHS, sRefersToRHS, rel, 0  ' Show either a value or a formula from sRefersToRHS

3395          End If
NextConstraint:
3396      Next constraint


          ' We now go thru and mark integer and binary variables
          Dim HighlightColor As Long
          Dim CellHighlight As ShapeRange
3397      HighlightColor = RGB(0, 0, 0)
3398      i = 0

          Dim selectedArea As Range
3399      If numVars > 200 Then
3400          If Not BinaryCellsRange Is Nothing Then
3401              For Each selectedArea In BinaryCellsRange.Areas
3402                  Set CellHighlight = HighlightRange(selectedArea, "", RGB(255, 0, 255)) ' Magenta highlight
3403                  AddLabelToShape ActiveSheet, CellHighlight(1), -6, 10, "Binary", RGB(0, 0, 0) ' Black text
3404              Next selectedArea
3405          End If
3406          If Not IntegerCellsRange Is Nothing Then
3407              If Not BinaryCellsRange Is Nothing Then
3408                  If Not BinaryCellsRange.Count = IntegerCellsRange.Count Then
3409                      For Each selectedArea In IntegerCellsRange.Areas
3410                          Set CellHighlight = HighlightRange(selectedArea, "", RGB(255, 0, 255)) ' Magenta highlight
3411                          AddLabelToShape ActiveSheet, CellHighlight(1), -6, 10, "integer", RGB(0, 0, 0) ' Black text
3412                      Next selectedArea
3413                  End If
3414              Else
3415                   For Each selectedArea In IntegerCellsRange.Areas
3416                      Set CellHighlight = HighlightRange(selectedArea, "", RGB(255, 0, 255)) ' Magenta highlight
3417                      AddLabelToShape ActiveSheet, CellHighlight(1), -6, 10, "integer", RGB(0, 0, 0) ' Black text
3418                  Next selectedArea
                      
3419               End If
3420          End If
3421      Else
3422          If Not BinaryCellsRange Is Nothing Then
3423              For Each c In BinaryCellsRange
3424                   AddLabelToRange ActiveSheet, c, 1, 9, "b", HighlightColor
3425              Next c
3426          End If
3427          If Not IntegerCellsRange Is Nothing Then
3428              For Each c In IntegerCellsRange
3429                  If Not BinaryCellsRange Is Nothing Then
3430                      If Intersect(c, BinaryCellsRange) Is Nothing Then
3431                          AddLabelToRange ActiveSheet, c, 1, 9, "i", HighlightColor
3432                      End If
3433                  Else
3434                       AddLabelToRange ActiveSheet, c, 1, 9, "i", HighlightColor
3435                  End If
3436              Next c
3437          End If
3438      End If

          If Not NonAdjustableCellsRange Is Nothing Then
              For Each selectedArea In NonAdjustableCellsRange.Areas
                  HighlightRange selectedArea, "", RGB(255, 255, 0), True  ' Yellow highlight
              Next selectedArea
          End If
          
3439      If Errors <> "" Then
3440          MsgBox Errors, , "OpenSolver Warning"
3441          GoTo ExitSub
3442      End If
          
3443      ShowSolverModel = True  ' success

ExitSub:
3444      Application.StatusBar = False ' Resume normal status bar behaviour
3445      Application.ScreenUpdating = True
          
3446      Exit Function
          
errorHandler:
3447      If Err.Number = 18 Then
3448          If MsgBox("You have pressed the Escape key. Do you wish to cancel?", _
                         vbCritical + vbYesNo + vbDefaultButton1, _
                         "OpenSolver: User Interrupt Occured...") = vbNo Then
3449              Resume 'continue on from where error occured
3450          Else
3451              Resume ExitSub
3452          End If
3453      End If
3454      MsgBox "OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source, , "OpenSolver Code Error"
          
3455      Resume ExitSub
3456      Resume
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

