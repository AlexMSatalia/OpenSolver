Attribute VB_Name = "OpenSolverVisualizer"
' OpenSolver
' Copyright Andrew Mason 2010
' http://www.OpenSolver.org
' This software is distributed under the terms of the GNU General Public License
''
' This file is part of OpenSolver.
'
' OpenSolver is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' OpenSolver is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with OpenSolver.  If not, see <http://www.gnu.org/licenses/>.
'

Option Explicit

Private ShapeIndex As Long
Private HighlightColorIndex As Long   ' Used to rotate thru colours on different constraints
Private colHighlightOffsets As Collection

Private ShowDataItemsInColour As Boolean ' For OpenSolverStudio

Function NextHighlightColor() As Long
30320     HighlightColorIndex = HighlightColorIndex + 1
30330     If HighlightColorIndex > 7 Then HighlightColorIndex = 1
30340     Select Case HighlightColorIndex
              Case 1: NextHighlightColor = RGB(0, 0, 255) ' Blue
30350         Case 2: NextHighlightColor = RGB(0, 128, 0)  ' Green
30360         Case 3: NextHighlightColor = RGB(153, 0, 204) ' Purple
30370         Case 4: NextHighlightColor = RGB(128, 0, 0) ' Brown
30380         Case 5: NextHighlightColor = RGB(0, 204, 51) ' Light Green
30390         Case 6: NextHighlightColor = RGB(255, 102, 0) ' Orange
30400         Case 7: NextHighlightColor = RGB(204, 0, 153) ' Bright Purple
30410     End Select
End Function

Function NextHighlightColor_SkipColor(colorToSkip As Long) As Long
          Dim newColour As Long
30420     newColour = NextHighlightColor
30430     If newColour = colorToSkip Then
30440         newColour = NextHighlightColor
30450     End If
30460     NextHighlightColor_SkipColor = newColour
End Function

Function InitialiseHighlighting()
30470     Set colHighlightOffsets = New Collection
30480     ShapeIndex = 0
30490     HighlightColorIndex = 0 ' Used to rotate thru colours on different constraints
End Function

'Function SheetHasOpenSolverModelHighlighting(w As Worksheet)
'    ' If we have a shape called OpenSolver1 then we are displaying a highlighted model
'    If w.Shapes.Count = 0 Then GoTo NoHighlighting
'    Dim s As Shape
'    SheetHasOpenSolverModelHighlighting = True
'    On Error Resume Next
'    Set s = w.Shapes("OpenSolver1")
'    If Err.Number = 0 Then Exit Function ' Yes, we have highlighting
'    ' Because the highlighting may be on another sheet, we also check a few shapes
'    For Each s In w.Shapes
'        If s.name Like "OpenSolver*" Then Exit Function
'    Next s
'NoHighlighting:
'    SheetHasOpenSolverModelHighlighting = False
'    Return
'End Function

'Function SheetHasOpenSolverStudioDataHighlighting(w As Worksheet)
'    ' If we have a shape called OpenSolverStudio1 then we are displaying highlighted data
'    If w.Shapes.Count = 0 Then GoTo NoHighlighting
'    Dim s As Shape
'    SheetHasOpenSolverStudioDataHighlighting = True
'    On Error Resume Next
'    Set s = w.Shapes("OpenSolverStudio1")
'    If Err.Number = 0 Then Exit Function ' Yes, we have highlighting
'    ' Because the highlighting may be on another sheet, we also check a few shapes
'    For Each s In w.Shapes
'        If s.name Like "OpenSolverStudio*" Then Exit Function
'    Next s
'NoHighlighting:
'    SheetHasOpenSolverStudioDataHighlighting = False
'    Return
'End Function

Function SheetHasOpenSolverHighlighting(w As Worksheet)
          ' If we have a shape called OpenSolverStudio1 then we are displaying highlighted data
30500     If w.Shapes.Count = 0 Then GoTo NoHighlighting
          Dim s As Shape
30510     SheetHasOpenSolverHighlighting = True
30520     On Error Resume Next
30530     Set s = w.Shapes("OpenSolver1")
30540     If Err.Number = 0 Then Exit Function ' Yes, we have highlighting
30550     Set s = w.Shapes("OpenSolverStudio1")
30560     If Err.Number = 0 Then Exit Function ' Yes, we have highlighting
          ' Because the highlighting may be on another sheet, we also check all the shapes on this sheet
30570     For Each s In w.Shapes
30580         If s.Name Like "OpenSolver*" Then Exit Function
30590     Next s
NoHighlighting:
30600     SheetHasOpenSolverHighlighting = False
End Function

Function CreateLabelShape(w As Worksheet, left As Long, top As Long, width As Long, height As Long, label As String, HighlightColor As Long) As Shape
          ' Create a label (as a msoShapeRectangle) and give it text. This is used for labelling obj function as min/max, and decision vars as binary or integer
          Dim s1 As Shape
30610     Set s1 = w.Shapes.AddShape(msoShapeRectangle, left, top, width, height)
30620     s1.Fill.Visible = True
          s1.Fill.Solid
30630     s1.Fill.ForeColor.RGB = RGB(255, 255, 255)
30640     s1.Fill.Transparency = 0.2
30650     s1.Line.Visible = False
          s1.Shadow.Visible = msoFalse
30660     With s1.TextFrame
30670         .Characters.Text = label
30680         .Characters.Font.Size = 9
30690         .Characters.Font.Color = HighlightColor
30700         .HorizontalAlignment = xlHAlignLeft
30710         .VerticalAlignment = xlVAlignTop
30720         .MarginBottom = 0
30730         .MarginLeft = 1
30740         .MarginRight = 1
30750         .MarginTop = 0
30760         .AutoSize = True     ' Get width correct
30770         .AutoSize = False
30780     End With
30790     ShapeIndex = ShapeIndex + 1
30800     s1.Name = "OpenSolver" & ShapeIndex
30810     s1.height = height      ' Force the specified height
30820     Set CreateLabelShape = s1
End Function

Function AddLabelToRange(w As Worksheet, r As Range, voffset As Long, height As Long, label As String, HighlightColor As Long) As Shape
30830     Set AddLabelToRange = CreateLabelShape(w, r.left + 1, r.top + voffset, r.width, height, label, HighlightColor)
End Function

Function AddLabelToShape(w As Worksheet, s As Shape, voffset As Long, height As Long, label As String, HighlightColor As Long)
30840     Set AddLabelToShape = CreateLabelShape(w, s.left - 1, s.top + voffset, s.width, height, label, HighlightColor)
          'Set s1 = w.Shapes.AddShape(msoShapeRectangle, s.left - 1, s.Top - height / 2, s.width, height)
End Function

Function HighlightRange(r As Range, label As String, HighlightColor As Long, Optional ShowFill As Boolean = False, Optional ShapeNamePrefix As String = "OpenSolver", Optional Bounds As Boolean = False) As ShapeRange

          'Only show the model file that is on the active sheet to overcome hide bugs for shapes on different sheets
30850     If r.Worksheet.Name <> ActiveSheet.Name Then Exit Function
          
          Const HighlightingOffsetStep = 1
          ' We offset our highlighting so that successive highlights are still visible
          Dim Offset As Double, key As String
30860     key = r.Address(RowAbsolute:=False, ColumnAbsolute:=False)
30870     Offset = 0
30880     On Error Resume Next
30890     Offset = colHighlightOffsets(key) ' eg A1
30900     If Err.Number <> 0 Then
              ' Item does not exist in collection create it, with an offset of 2 ready to use next time
30910         colHighlightOffsets.Add HighlightingOffsetStep, key
30920     Else
30930         colHighlightOffsets.Remove key
30940         colHighlightOffsets.Add Offset + HighlightingOffsetStep, key
30950     End If
30960     On Error GoTo 0
          
          ' Handle merged cells
          Dim l As Single, t As Single, right As Single, bottom As Single, R2 As Range, c As Range
30970     If Not r.MergeCells Then
30980         l = r.left
30990         t = r.top
31000         right = l + r.width
31010         bottom = t + r.height
31020     Else
              ' This range contains merged cells. We use MergeArea to to find the area to highlight
              ' But, this only works for one cell, so we expand our size based on all cells
31030         If r.Count = 1 Then
31040             Set R2 = r.MergeArea
31050             l = R2.left
31060             t = R2.top
31070             right = l + R2.width
31080             bottom = t + R2.height
31090         Else
31100             l = r.left
31110             t = r.top
31120             right = l + r.width
31130             bottom = t + r.height
31140             For Each c In r
31150                 If c.MergeCells Then
31160                     Set R2 = c.MergeArea
31170                     If R2.left < l Then l = R2.left
31180                     If R2.top < t Then t = R2.top
31190                     If R2.left + R2.width > right Then right = R2.left + R2.width
31200                     If R2.top + R2.height > bottom Then bottom = R2.top + R2.height
31210                 End If
31220             Next c
31230         End If
31240     End If
          
          
       ' Use doubles here for more accuracy as we are summing terms, and so accummulating errors
          Dim left2 As Double, top2 As Double, right2 As Double, bottom2 As Double, height As Double, width As Double
31250     left2 = l + Offset
31260     top2 = t + Offset
          
          ' Draw enough shapes to cover the space; each shape has an Excel-set (undocumented?) maximum height (and width?)
          Dim isFirstShape As Boolean
31270     isFirstShape = True
          Dim firstShapeIndex As Long
31280     firstShapeIndex = ShapeIndex + 1
31290     Do
        ' The height cannot exceed 169056.0; we allow some tolerance
31300   height = bottom - top2
31310   If height > 160000# Then
31320       height = 150000  ' This difference ensures we never end up with very small rectangle
31330   End If
31340   If isFirstShape And height > 9500 Then
31350       height = 9000   ' The first shape we create has a height of 9000 to ensure we can rotate the text and have it show
                            ' correctly; this works around and Excel 2007 bug
31360   End If
31370   isFirstShape = False
        
        Dim shapeName As String, s1 As Shape
31380   ShapeIndex = ShapeIndex + 1
        'If the constraint is not a bound then make a box for it
31390   If Not Bounds Then
31400       shapeName = ShapeNamePrefix & ShapeIndex
31410       r.Worksheet.Shapes.AddShape(msoShapeRectangle, l + Offset, top2, right - l, height).Name = shapeName
31420   Else
            'If the box is a bound we name it after the cell that it is in
31430       shapeName = ShapeNamePrefix & key
31440       On Error Resume Next
            Dim tmpName As String
31450       tmpName = r.Worksheet.Shapes(shapeName).Name
            'If there hasn't been a bound on that cell then make a new cell
31460       If tmpName = "" Then
31470           r.Worksheet.Shapes.AddShape(msoShapeRectangle, l + Offset, top2, right - l, height).Name = shapeName
31480       Else
                'If there has already been a bound then just add new text to it rather then making a new box
31490           Set s1 = r.Worksheet.Shapes(shapeName)
31500           s1.TextFrame.Characters.Text = s1.TextFrame.Characters.Text & "," & label
31510           GoTo endLoop
31520       End If
31530   End If
        
31540   Set s1 = r.Worksheet.Shapes(shapeName)
        'Set s1 = r.Worksheet.Shapes.AddShape(msoShapeRectangle, l + Offset, top2, right - l, height)
              
        Dim ShowOutline As Boolean
31550   ShowOutline = Not ShowFill
        
31560   If ShowOutline Then
31570       s1.Fill.Visible = False
31580       With s1.Line
31590           .Weight = 2
31600           .ForeColor.RGB = HighlightColor
31610       End With
31620   Else
31630       s1.Line.Visible = False
            s1.Fill.Solid
31640       s1.Fill.Transparency = 0.6
31650       s1.Fill.ForeColor.RGB = HighlightColor
31660   End If
        s1.Shadow.Visible = msoFalse
        
31670   With s1.TextFrame
31680       .Characters.Text = label
31690       .Characters.Font.Color = HighlightColor
31700       .HorizontalAlignment = xlHAlignLeft ' xlHAlignCenter
31710       If height < 500 Then
31720           .VerticalAlignment = xlVAlignCenter  ' Shape is small enought to have text fit on the screen when centered, so we center text
31730       Else
31740           .VerticalAlignment = xlVAlignTop   ' So we can see the name when scrolled to the top
31750       End If
31760       .MarginBottom = 0
31770       .MarginLeft = 2
31780       .MarginRight = 0
31790       .MarginTop = 2
31800       .Characters.Font.Bold = True
31810   End With
endLoop:
31820   top2 = top2 + height
31830     Loop While bottom - top2 > 0.01  ' handle float rounding
          
          ' Create & return the shapeRange containing all the shapes we added
          Dim shapeNames(), i As Long
31840     ReDim shapeNames(ShapeIndex - firstShapeIndex + 1)
31850     If Not Bounds Then
31860   For i = firstShapeIndex To ShapeIndex
31870       shapeNames(i - firstShapeIndex + 1) = ShapeNamePrefix & i
31880   Next i
31890     Else
31900   For i = firstShapeIndex To ShapeIndex
31910       shapeNames(i - firstShapeIndex + 1) = ShapeNamePrefix & key
31920   Next i
31930     End If
31940     Set HighlightRange = r.Worksheet.Shapes.Range(shapeNames)
          
          
      ' Dim s1 As Shape
      '21800     Set s1 = r.Worksheet.Shapes.AddShape(msoShapeRectangle, l + Offset, t + Offset, right - l, bottom - t)
      '
      '          Dim ShowOutline As Boolean
      '21810     ShowOutline = Not ShowFill
      '
      '21820     If ShowOutline Then
      '21830         s1.Fill.Visible = False
      '21840         With s1.Line
      '21850             .Weight = 2
      '21860             .ForeColor.RGB = HighlightColor
      '21870         End With
      '21880     Else
      '21890         s1.Line.Visible = False
      '21900         s1.Fill.Transparency = 0.6
      '21910         s1.Fill.ForeColor.RGB = HighlightColor
      '21920     End If
      '
      '21930     With s1.TextFrame
      '21940         .Characters.Text = label
      '21950         .Characters.Font.Color = HighlightColor
      '21960         .Characters.Font.Size = 11
      '21970         .HorizontalAlignment = xlHAlignLeft ' xlHAlignCenter
      '21980         .VerticalAlignment = xlVAlignCenter
      '21990         .MarginBottom = 0
      '22000         .MarginLeft = 2
      '22010         .MarginRight = 0
      '22020         .MarginTop = 0
      '22030         .Characters.Font.Bold = True
      '22040     End With
      '
      '22050     ShapeIndex = ShapeIndex + 1
      '22060     s1.name = ShapeNamePrefix & ShapeIndex
      '
      '22070     Set HighlightRange = s1
End Function

Function AddLabelledConnector(w As Worksheet, s1 As Shape, s2 As Shape, label As String)
          Dim t As Shapes, c As Shape
          
          'Set myDocument = Worksheets(1)
31950     Set t = w.Shapes
          'Set firstRect = t.AddShape(msoShapeRectangle, 100, 50, 200, 100)
          'Set secondRect = t.AddShape(msoShapeRectangle, 300, 300, 200, 100)
31960     Set c = t.AddConnector(msoConnectorStraight, 0, 0, 0, 0) ' msoConnectorCurve
31970     With c.ConnectorFormat
31980         .BeginConnect ConnectedShape:=s1, ConnectionSite:=1
31990         .EndConnect ConnectedShape:=s2, ConnectionSite:=1
32000     End With
32010     c.RerouteConnections
32020     c.Line.ForeColor = s1.Line.ForeColor
          ' The default styles can be changed, so we should set everything! We just do a few of the problem bits
32030     c.Line.EndArrowheadStyle = msoArrowheadNone
32040     c.Line.BeginArrowheadStyle = msoArrowheadNone
32050     c.Line.DashStyle = msoLineSolid
32060     c.Line.Weight = 0.75
32070     c.Line.Style = msoLineSingle
          c.Shadow.Visible = msoFalse
              
32080     ShapeIndex = ShapeIndex + 1
32090     c.Name = "OpenSolver " & ShapeIndex
          
          Dim s3 As Shape
32100     Set s3 = t.AddShape(msoShapeRectangle, c.left + c.width / 2# - 30 / 2, c.top + c.height / 2# - 20 / 2, 30, 20)
                  
32110     s3.Line.Visible = False
32120     s3.Fill.Visible = False
          s3.Shadow.Visible = msoFalse
32130     If label <> "" Then
32140         With s3.TextFrame
32150             .Characters.Text = label
32160             .MarginBottom = 0
32170             .MarginLeft = 0
32180             .MarginRight = 0
32190             .MarginTop = 0
32200             .HorizontalAlignment = xlHAlignCenter
32210             .VerticalAlignment = xlVAlignCenter
32220             .Characters.Font.Color = c.Line.ForeColor
32230             .Characters.Font.Bold = True
32240         End With
32250     End If

32260     ShapeIndex = ShapeIndex + 1
32270     s3.Name = "OpenSolver " & ShapeIndex

End Function

Sub HighlightConstraint(myDocument As Worksheet, LHSRange As Range, _
                        ByVal RHSisRange As Boolean, RHSRange As Range, ByVal RHSValue As String, ByVal Sense As Long, _
                        ByVal Color As Long)
          ' Show a constraint of the form LHS <|=|> RHS.
          ' We always put the sign in the rightmost (or bottom-most) range so we read left-to-right or top-to-bottom.
          Dim Range1 As Range, Range2 As Range, Reversed As Boolean
          Dim s1 As ShapeRange, s2 As ShapeRange

          ' Get next color if none specified
32280     If Color = 0 Then Color = NextHighlightColor
          
32290     Reversed = False
32300     If Not RHSisRange And RHSValue <> "" Then
              ' We have a constant or formula in the constraint. Put into form RHS <|=|> Range1 (reversing the sense)
32310         Set s1 = HighlightRange(LHSRange, RHSValue & SolverRelationAsUnicodeChar(4 - Sense), Color, , , True)
32320     ElseIf RHSisRange Then
              ' If ranges overlaps on rows, then the top one becomes Range1
32330         If ((RHSRange.top >= LHSRange.top And RHSRange.top < LHSRange.top + LHSRange.height) _
              Or (LHSRange.top >= RHSRange.top And LHSRange.top < RHSRange.top + RHSRange.height)) Then
                  ' Ranges over lap in rows. Range1 becomes the left most one
32340             If LHSRange.left > RHSRange.left Then
32350                 Reversed = True
32360             End If
32370         ElseIf ((RHSRange.left >= LHSRange.left And RHSRange.left < LHSRange.left + LHSRange.width) _
              Or (LHSRange.left >= RHSRange.left And LHSRange.left < RHSRange.left + RHSRange.width)) Then
                  ' Ranges overlap in columns. Range1 becomes the top most one
32380             If LHSRange.top > RHSRange.top Then
32390                 Reversed = True
32400             End If
32410         Else
                  ' Ranges are in different rows with no overlap; top one becomes Range1
32420             If LHSRange.left >= RHSRange.left + RHSRange.width Then
32430                 Reversed = True
32440             End If
32450         End If
              
32460         If Reversed Then
32470             Set Range1 = RHSRange
32480             Set Range2 = LHSRange
32490         Else
32500             Set Range1 = LHSRange
32510             Set Range2 = RHSRange
32520         End If
          
32530         Set s1 = HighlightRange(Range1, "", Color)
          
              ' Reverse the sense if the objects are shown in the reverse order
32540         Set s2 = HighlightRange(Range2, SolverRelationAsUnicodeChar(IIf(Reversed, 4 - Sense, Sense)), Color)
              
32550         If Range1.Worksheet.Name = Range2.Worksheet.Name And Range1.Worksheet.Name = ActiveSheet.Name Then
32560             AddLabelledConnector Range1.Worksheet, s1(1), s2(1), ""
32570         End If
32580     Else 'this was added if there is only a lhs that needs highlighting in linearity
32590         Set s1 = HighlightRange(LHSRange, "", Color)
              ' AddLabelledConnector myDocument, s1, s2, ""
32600     End If
End Sub

Sub DeleteOpenSolverShapes(w As Worksheet)
          Dim s As Shape
32610     For Each s In w.Shapes
32620         If s.Name Like "OpenSolver*" Then
32630             s.Delete
32640         End If
32650     Next s
End Sub

Function HideSolverModel() As Boolean

32660     HideSolverModel = False ' assume we fail
          
32670     If Application.Workbooks.Count = 0 Then
32680         MsgBox "Error: No active workbook available", , "OpenSolver" & sOpenSolverVersion & " Error"
32690         Exit Function
32700     End If
          
          ' We trap the Escape key which does an "on error"
          'xlDisabled = 0 'totally disables Esc / Ctrl-Break / Command-Period
          'xlInterrupt = 1 'go to debug
          'xlErrorHandler = 2 'go to error handler
          'Trappable error is #18
32710     Application.EnableCancelKey = xlErrorHandler
32720     On Error GoTo errorHandler
          
32730     Application.ScreenUpdating = False
          
          Dim sheet As Worksheet
32740     On Error Resume Next
          
32750     Set sheet = ActiveWorkbook.ActiveSheet
32760     If Err.Number <> 0 Then
32770         MsgBox "Error: Unable to access the active sheet", , "OpenSolver" & sOpenSolverVersion & " Error"
32780         Exit Function
32790     End If
32800     On Error GoTo errorHandler
          
32810     DeleteOpenSolverShapes sheet
          Dim i As Long
          Dim NumOfConstraints As Long
          Dim sheetName As String
32820     sheetName = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!" ' NB: We have to double any ' when we quote the sheet name
32830     On Error Resume Next ' There may not be a model on the sheet
32840     NumOfConstraints = Mid(Names(sheetName & "solver_num"), 2)
32850     On Error GoTo errorHandler
          
          ' Delete constraints on other sheets
          Dim b As Boolean
32860     For i = 1 To NumOfConstraints
              ' This code used to say On Error goto NextConstraint, but this failed because the error was never Resume'd
32870         b = False
              ' Set b to be true only if there is no error
32880         On Error Resume Next
32890         b = Range(sheetName & "solver_lhs" & i).Worksheet.Name <> ActiveWorkbook.ActiveSheet.Name
32900         If b Then
32910             DeleteOpenSolverShapes Range(sheetName & "solver_lhs" & i).Worksheet
32920         End If
NextConstraint:
32930     Next i
32940     On Error GoTo errorHandler
          
32950     HideSolverModel = True  ' Successful completion

ExitSub:
32960     Application.StatusBar = False ' Resume normal status bar behaviour
32970     Application.ScreenUpdating = True
          
32980     Exit Function
          
errorHandler:
32990     If Err.Number = 18 Then
33000         If MsgBox("You have pressed the Escape key. Do you wish to cancel?", _
                         vbCritical + vbYesNo + vbDefaultButton1, _
                         "OpenSolver: User Interrupt Occured...") = vbNo Then
33010             Resume 'continue on from where error occured
33020         Else
33030             Resume ExitSub
33040         End If
33050     End If
33060     MsgBox "OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source, , "OpenSolver" & sOpenSolverVersion & " Error"
33070     Resume ExitSub
          
End Function

Function ShowSolverModel() As Boolean
          
33080     ShowSolverModel = False ' Assume we fail
          
33090     If Application.Workbooks.Count = 0 Then
33100         MsgBox "Error: No active workbook available", , "OpenSolver" & sOpenSolverVersion & " Error"
33110         Exit Function
33120     End If

          ' We trap the Escape key which does an onerror
          'xlDisabled = 0 'totally disables Esc / Ctrl-Break / Command-Period
          'xlInterrupt = 1 'go to debug
          'xlErrorHandler = 2 'go to error handler
          'Trappable error is #18
33130     Application.EnableCancelKey = xlErrorHandler
33140     On Error GoTo errorHandler
          
33150     Application.ScreenUpdating = False
          
          Dim i As Double, sheetName As String, book As Workbook, AdjustableCells As Range
          Dim NumConstraints  As Long

33160     On Error Resume Next
33170     sheetName = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!" ' NB: We have to double any ' when we quote the sheet name
33180     If Err.Number <> 0 Then
33190         MsgBox "Error: Unable to access the active sheet", , "OpenSolver" & sOpenSolverVersion & " Error"
33200         Exit Function
33210     End If
33220     On Error GoTo errorHandler
          
33230     Set book = ActiveWorkbook
          
33240     DeleteOpenSolverShapes ActiveSheet
33250     InitialiseHighlighting
          
          ' We check to see if a model exists by getting the adjustable cells. We check for a name first, as this may contain =Sheet1!$C$2:$E$2,Sheet1!#REF!
          Dim n As Name
33260     On Error Resume Next
33270     Set n = Names(sheetName & "solver_adj")
      '23410     If Err.Number <> 0 Then
      ''23420         MsgBox "Error: No Solver model was found on sheet " & ActiveWorkbook.ActiveSheet.name, , "OpenSolver Error"
      ''23430         GoTo ExitSub
      ''23440     End If
33280     On Error Resume Next
33290     Set AdjustableCells = RemoveRangeOverlap(Range(sheetName & "solver_adj"))   ' Remove any overlap in the range defining the decision variables
33300     If Err.Number <> 0 Then
33310         MsgBox "Error: A model was found on the sheet " & ActiveWorkbook.ActiveSheet.Name & " but the decision variable cells (" & n & ") could not be interpreted. Please redefine the decision variable cells, and try again.", , "OpenSolver" & sOpenSolverVersion & " Error"
33320         GoTo ExitSub
33330     End If
33340     On Error GoTo errorHandler
          
          ' Highlight the decision variables
33350     AddDecisionVariableHighlighting AdjustableCells
          
          Dim Errors As String
          
          ' Highlight the objective cell, if there is one
          Dim ObjRange As Range
33360     If GetNamedRangeIfExists(book, sheetName & "solver_opt", ObjRange) Then
33370         Set ObjRange = Range(sheetName & "solver_opt")
              Dim ObjType As ObjectiveSenseType, temp As Long, ObjectiveTargetValue As Double
33380         ObjType = UnknownObjectiveSense
33390         If GetNamedIntegerIfExists(book, sheetName & "solver_typ", temp) Then ObjType = temp
33400         If ObjType = TargetObjective Then GetNamedNumericValueIfExists book, sheetName & "solver_val", ObjectiveTargetValue
33410         AddObjectiveHighlighting ObjRange, ObjType, ObjectiveTargetValue
33420     Else
      '              Dim ObjFunctErr As VbMsgBoxResult
      '              ObjFunctErr = MsgBox("Warning: No objective cell has been set. Do you still want to save the model?", vbYesNo, "OpenSolver")
      '              If ObjFunctErr = vbNo Then
      '                  Exit Function
      '              End If
      '23600         Errors = Errors & "Warning: No objective cell has been set." & vbCrLf
33430     End If
          
          ' Count the correct number of constraints, and form the constraint
33440     NumConstraints = Mid(Names(sheetName & "solver_num"), 2)  ' Number of constraints entered in excel; can include ranges covering many constraints
          ' Note: Solver leaves around old constraints; the name <sheet>!solver_num gives the correct number of constraints (eg "=4")
          
33450     Application.StatusBar = "OpenSolver: Displaying Problem... " & AdjustableCells.Count & " vars, " & NumConstraints & " Solver constraints"
                  
          ' Process the decision variables as we need to compute their types (bin or int; a variable can be declared as both!)
          'Dim NumVars As Long
          Dim numVars As Double
33460     numVars = AdjustableCells.Count
          'Dim VarNames() As String, VarTypes() As Range
          'Dim VarNamesCollection As New Collection
          'Dim VarTypesCollection As New Collection
          'Dim VarAreasCollection As New Collection
          'ReDim VarNames(NumVars), VarTypes(NumVars)
          Dim BinaryCellsRange As Range
          Dim IntegerCellsRange As Range
          ' Get names for all the variables so we can track their types
          Dim c As Range
          'i = 0
          'For Each c In AdjustableCells
              'i = i + 1
          '    VarNamesCollection.Add c, c.Address(RowAbsolute:=False, ColumnAbsolute:=False)
              'VarNames(i) = c.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) ' eg A1
          'Next c
                  
          ' Count the correct number of constraints, and form the constraint
          Dim constraint As Long
          Dim currentSheet As Worksheet
33470     For constraint = 1 To NumConstraints
              
              Dim isRangeLHS As Boolean, valLHS As Double, rLHS As Range, RangeRefersToError As Boolean, RefersToFormula As Boolean, sNameLHS As String, sRefersToLHS As String, isMissingLHS As Boolean
33480         sNameLHS = sheetName & "solver_lhs" & constraint
33490         GetNameAsValueOrRange book, sNameLHS, isMissingLHS, isRangeLHS, rLHS, RefersToFormula, RangeRefersToError, sRefersToLHS, valLHS
33500         If isMissingLHS Then
33510             Errors = Errors & "Error: The left hand side of constraint " & constraint & " is not defined (no 'solver_lhs" & constraint & "')." & vbCrLf
33520             GoTo NextConstraint
33530         End If
33540         If Not isRangeLHS Or RangeRefersToError Or RefersToFormula Then
33550             Errors = Errors & "Error: Range " & book.Names(sNameLHS).RefersTo & " is not a valid left hand side for a constraint." & vbCrLf
33560             GoTo NextConstraint
33570         End If

              Dim rel As Long

33580         rel = Mid(Names(sheetName & "solver_rel" & constraint), 2)
                      
              Dim LHSCount As Double, Count As Double
33590         LHSCount = rLHS.Count
33600         Count = LHSCount
              Dim AllDecisionVariables As Boolean
33610         AllDecisionVariables = False
      '        If rel = RelationInt Or rel = RelationBin Then
      '            ' Make the LHS variables integer or binary
      '            If isRangeLHS Then  ' It should always be a range
      '                For Each c In rLHS
      '                    Dim found As Boolean
      '                    found = False
      '                    For i = 1 To NumVars
      '                        If c.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) = VarNames(i) Then
      '                            found = True
      '                            'If Not TestKeyExists(VarTypesCollection, VarNames(i).Address) Then
      '                                'VarTypesCollection.Add VarNames(i).Formula, VarNames(i).Address
      '                            'End If
      '                            ' We allow for a variable to be specified as both int and binary, in which case it becomes binary
      '                            ' Note: RelationInt = 4, RelationBin = 5. Default of 0 means continuous
      '                            'VarTypes(i) = Max(VarTypes(i), rel)
      '
      '                            Exit For
      '                        End If
      '
      '                    Next i
33620         If rel = RelationINT Or rel = RelationBIN Then
              ' Make the LHS variables integer or binary
                  Dim intersection As Range
33630             Set intersection = Intersect(AdjustableCells, rLHS)
33640             If intersection Is Nothing Then
33650                 Errors = Errors & "Error: A cell specified as bin or int could not be found in the decision variable cells." & vbCrLf
33660                 GoTo NextConstraint
33670             End If
33680             If intersection.Count = rLHS.Count Then
33690                 If rel = RelationINT Then
33700                     If IntegerCellsRange Is Nothing Then
33710                         Set IntegerCellsRange = rLHS
33720                     Else
33730                         Set IntegerCellsRange = Union(IntegerCellsRange, rLHS)
33740                     End If
33750                 Else
33760                     If BinaryCellsRange Is Nothing Then
33770                         Set BinaryCellsRange = rLHS
33780                     Else
33790                         Set BinaryCellsRange = Union(BinaryCellsRange, rLHS)
33800                     End If
33810                 End If
33820             Else
33830                 Errors = Errors & "Error: A cell specified as bin or int could not be found in the decision variable cells." & vbCrLf
33840                 GoTo NextConstraint
33850             End If
33860         Else
                  ' Constraint is a full equation with a RHS
                  Dim isRangeRHS As Boolean, valRHS As Double, rRHS As Range, sNameRHS As String, sRefersToRHS As String, isMissingRHS As Boolean
33870             sNameRHS = sheetName & "solver_rhs" & constraint
33880             GetNameAsValueOrRange book, sNameRHS, isMissingRHS, isRangeRHS, rRHS, RefersToFormula, RangeRefersToError, sRefersToRHS, valRHS
33890             If isMissingRHS Then
33900                 Errors = Errors & "Error: The right hand side of constraint " & constraint & " is not defined (no 'solver_rhs" & constraint & "')." & vbCrLf
33910                 GoTo NextConstraint
33920             End If
33930             If RangeRefersToError Then
33940                 Errors = Errors & "Error: Range " & Mid(book.Names(sNameRHS), 2) & " is not a valid right hand side in a constraint." & vbCrLf
33950                 GoTo NextConstraint
33960             End If
                  
33970             If Range(sheetName & "solver_lhs" & constraint).Worksheet.Name <> ActiveWorkbook.ActiveSheet.Name Then
33980                 Set currentSheet = Range(sheetName & "solver_lhs" & constraint).Worksheet
33990             Else
34000                 Set currentSheet = ActiveSheet
34010             End If
                  
34020             If RefersToFormula Then
                      ' Shorten the formula (eg Test4!$M$11/4+Test4!$A$3) by removing the current sheet name and all $
34030                 sRefersToRHS = Replace(sRefersToRHS, currentSheet.Name & "!", "")   ' Remove names like Test4!
34040                 sRefersToRHS = Replace(sRefersToRHS, "'" & Replace(currentSheet.Name, "'", "''") & "'!", "") ' Remove names with spaces that are quoted, like 'Test 4'!, and 'Andrew''s'! (with escaped ')
34050                 sRefersToRHS = Replace(sRefersToRHS, "$", "")
34060             End If
34070             HighlightConstraint currentSheet, rLHS, isRangeRHS, rRHS, sRefersToRHS, rel, 0  ' Show either a value or a formula from sRefersToRHS

34080         End If
NextConstraint:
34090     Next constraint


          ' We now go thru and mark integer and binary variables
          Dim HighlightColor As Long
          Dim CellHighlight As ShapeRange
34100     HighlightColor = RGB(0, 0, 0)
34110     i = 0

          Dim selectedArea As Range
34120     If numVars > 200 Then
34130         If Not BinaryCellsRange Is Nothing Then
34140             For Each selectedArea In BinaryCellsRange.Areas
34150                 Set CellHighlight = HighlightRange(selectedArea, "", RGB(255, 0, 255)) ' Magenta highlight
34160                 AddLabelToShape ActiveSheet, CellHighlight(1), -6, 10, "Binary", RGB(0, 0, 0) ' Black text
34170             Next selectedArea
34180         End If
34190         If Not IntegerCellsRange Is Nothing Then
34200             If Not BinaryCellsRange Is Nothing Then
34210                 If Not BinaryCellsRange.Count = IntegerCellsRange.Count Then
34220                     For Each selectedArea In IntegerCellsRange.Areas
34230                         Set CellHighlight = HighlightRange(selectedArea, "", RGB(255, 0, 255)) ' Magenta highlight
34240                         AddLabelToShape ActiveSheet, CellHighlight(1), -6, 10, "integer", RGB(0, 0, 0) ' Black text
34250                     Next selectedArea
34260                 End If
34270             Else
34280                  For Each selectedArea In IntegerCellsRange.Areas
34290                     Set CellHighlight = HighlightRange(selectedArea, "", RGB(255, 0, 255)) ' Magenta highlight
34300                     AddLabelToShape ActiveSheet, CellHighlight(1), -6, 10, "integer", RGB(0, 0, 0) ' Black text
34310                 Next selectedArea
                      
34320              End If
34330         End If
34340     Else
34350         If Not BinaryCellsRange Is Nothing Then
34360             For Each c In BinaryCellsRange
34370                  AddLabelToRange ActiveSheet, c, 1, 9, "b", HighlightColor
34380             Next c
34390         End If
34400         If Not IntegerCellsRange Is Nothing Then
34410             For Each c In IntegerCellsRange
34420                 If Not BinaryCellsRange Is Nothing Then
34430                     If Intersect(c, BinaryCellsRange) Is Nothing Then
34440                         AddLabelToRange ActiveSheet, c, 1, 9, "i", HighlightColor
34450                     End If
34460                 Else
34470                      AddLabelToRange ActiveSheet, c, 1, 9, "i", HighlightColor
34480                 End If
34490             Next c
34500         End If
34510     End If
          
34520     If Errors <> "" Then
34530         MsgBox Errors, , "OpenSolver Warning"
34540         GoTo ExitSub
34550     End If
          
34560     ShowSolverModel = True  ' success

ExitSub:
34570     Application.StatusBar = False ' Resume normal status bar behaviour
34580     Application.ScreenUpdating = True
          
34590     Exit Function
          
errorHandler:
34600     If Err.Number = 18 Then
34610         If MsgBox("You have pressed the Escape key. Do you wish to cancel?", _
                         vbCritical + vbYesNo + vbDefaultButton1, _
                         "OpenSolver: User Interrupt Occured...") = vbNo Then
34620             Resume 'continue on from where error occured
34630         Else
34640             Resume ExitSub
34650         End If
34660     End If
34670     MsgBox "OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source, , "OpenSolver Code Error"
          
34680     Resume ExitSub
34690     Resume
End Function
'Iain dunning
Public Function TestKeyExists(ByRef col As Collection, key As String) As Boolean
          
          'MsgBox Key
    On Error GoTo doesntExist:
          Dim Item As Variant
          
34700     Set Item = col(key)
          
34710     TestKeyExists = True
34720     Exit Function
          
doesntExist:
34730     If Err.Number = 5 Then
34740         TestKeyExists = False
34750     Else
34760         TestKeyExists = True
34770     End If
          
End Function
'---------------------------------------------------------------------------------------------------------------
' Refactored code added 21st July
'---------------------------------------------------------------------------------------------------------------

Sub AddObjectiveHighlighting(ObjectiveRange As Range, ObjectiveType As ObjectiveSenseType, ObjectiveTargetValue As Double)
          ' AddObjectiveHighlighting
          ' Highlight the cell that is to be optimised, and add a label
          ' depending on the objective type
          ' Inputs:
          '   ObjectiveRange          Range that contains the cell to optimise
          '   ObjectiveType           Enum type that says what kind of objective we have
          ' Outputs:
          '   None
          
          ' Highlight the cell
          Dim CellHighlight As ShapeRange
34780     Set CellHighlight = HighlightRange(ObjectiveRange, "", RGB(255, 0, 255)) ' Magenta highlight
          
          ' Add the label
          Dim CellLabel As String
34790     CellLabel = "??? "
34800     If ObjectiveType = MaximiseObjective Then CellLabel = "max "
34810     If ObjectiveType = MinimiseObjective Then CellLabel = "min "
34820     If ObjectiveType = TargetObjective Then CellLabel = "seek " & ObjectiveTargetValue
34830     AddLabelToShape ActiveSheet, CellHighlight(1), -6, 10, CellLabel, RGB(0, 0, 0) ' Black text
End Sub


Sub AddDecisionVariableHighlighting(DecisionVariableRange As Range)
          ' AddDecisionVariableHighlighting
          ' Highlight decision variables. Note that this works even if decision
          ' variables are disjoint
          ' Inputs:
          '   DecisionVariableRange
          ' Outputs:
          '   None
          
          Dim area As Range
34840     For Each area In DecisionVariableRange.Areas
34850         HighlightRange area, "", RGB(255, 0, 255), True ' Magenta highlight
34860     Next area
          
End Sub

