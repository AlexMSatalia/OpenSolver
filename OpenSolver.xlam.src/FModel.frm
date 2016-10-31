VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FModel 
   Caption         =   "OpenSolver - Model"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   -4665
   ClientWidth     =   9840
   OleObjectBlob   =   "FModel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Mac Then
    Const FormWidthModel = 720
#Else
    Const FormWidthModel = 500
#End If

Private Constraints() As CConstraint
Private NumConstraints As Long

Private CurrentListIndex As Long  ' Store which constraint is currently showing
Private DisableConstraintListChange As Boolean  ' Disables the change event on the constraint list box
Private PreserveModel As Boolean  ' Used to persist model when re-showing form

Private RestoreHighlighting As Boolean
Private sheet As Worksheet

' Options that persist across form showings
Public ShowModelAfterSavingState As Boolean
Public ShowNamedRangesState As Boolean

' Resizing info
#If Mac Then
    Const MinHeight = 168
#Else
    Const MinHeight = 141
#End If

Private ResizeStartY As Double

' Function to map string rels to combobox index positions
' Assigning combobox by .value fails a lot on Mac
Function cboPosition(rel As String) As Integer
1         Select Case rel
          Case "="
2             cboPosition = 0
3         Case "<="
4             cboPosition = 1
5         Case ">="
6             cboPosition = 2
7         Case "int"
8             cboPosition = 3
9         Case "bin"
10            cboPosition = 4
11        Case "alldiff"
12            cboPosition = 5
13        End Select
End Function

Sub Disabler(TrueIfEnable As Boolean)
          ' Before calling Disabler, ensure that the focused control will not be disabled!
          ' This prevents tab-skipping into a RefEdit, and starting its event loop
          
1         lblDescHeader.Enabled = TrueIfEnable
2         lblDesc.Enabled = TrueIfEnable
3         cmdRunAutoModel.Enabled = TrueIfEnable

4         lblDiv1.Enabled = False

5         lblStep1.Enabled = TrueIfEnable
6         refObj.Enabled = TrueIfEnable
7         optMax.Enabled = TrueIfEnable
8         optMin.Enabled = TrueIfEnable
9         optTarget.Enabled = TrueIfEnable
10        txtObjTarget.Enabled = TrueIfEnable And GetFormObjectiveSense() = TargetObjective

11        lblDiv2.Enabled = False

12        lblStep2.Enabled = TrueIfEnable
13        refDecision.Enabled = TrueIfEnable

14        lblDiv3.Enabled = False

15        chkNonNeg.Enabled = TrueIfEnable
16        cmdCancelCon.Enabled = Not TrueIfEnable
17        cmdDelSelCon.Enabled = TrueIfEnable And lstConstraints.ListIndex > 0
18        chkNameRange.Enabled = TrueIfEnable

19        lblDiv4.Enabled = False

20        lblStep4.Enabled = TrueIfEnable

          ' Regardless of TrueIfEnable, disable the dual options if the solver doesn't support it
          Dim HasSensitivity As Boolean
21        HasSensitivity = SensitivityAnalysisAvailable(CreateSolver(GetChosenSolver(sheet)))

22        chkGetDuals.Enabled = TrueIfEnable And HasSensitivity
23        refDuals.Enabled = TrueIfEnable And HasSensitivity And chkGetDuals.value
24        chkGetDuals2.Enabled = TrueIfEnable And HasSensitivity
25        optUpdate.Enabled = TrueIfEnable And HasSensitivity And chkGetDuals2.value
26        optNew.Enabled = TrueIfEnable And HasSensitivity And chkGetDuals2.value

27        lblDiv5.Enabled = False
28        lblDiv6.Enabled = False

29        chkShowModel.Enabled = TrueIfEnable
30        cmdOptions.Enabled = TrueIfEnable
31        cmdBuild.Enabled = TrueIfEnable
32        cmdCancel.Enabled = TrueIfEnable
33        cmdReset.Enabled = TrueIfEnable
34        cmdChange.Enabled = TrueIfEnable
End Sub

Sub LoadModelFromSheet()
1         SetFormObjectiveSense GetObjectiveSense(sheet)
2         SetFormObjectiveTargetValue GetObjectiveTargetValue(sheet)
3         SetFormObjective GetObjectiveFunctionCellRefersTo(sheet)
4         SetFormDecisionVars GetDecisionVariablesRefersTo(sheet)
5         SetFormDuals GetDualsRefersTo(sheet)

6         NumConstraints = GetNumConstraints(sheet)
7         If NumConstraints > 0 Then
8             ReDim Constraints(1 To NumConstraints) As CConstraint
          
              Dim i As Long, LHSRefersTo As String, RHSRefersTo As String, Relation As RelationConsts
9             For i = 1 To NumConstraints
10                GetConstraintRefersTo i, LHSRefersTo, Relation, RHSRefersTo, sheet
11                Set Constraints(i) = New CConstraint
12                Constraints(i).Update LHSRefersTo, Relation, RHSRefersTo, sheet
13            Next i
14        End If

15        chkNonNeg.value = GetNonNegativity(sheet)
16        chkGetDuals.value = Len(GetFormDuals()) > 0
17        chkGetDuals2.value = GetDualsOnSheet(sheet)
18        chkGetDuals2_Click  ' Set enabled status of dual checkboxes

19        optUpdate.value = GetUpdateSensitivity(sheet)
20        optNew.value = Not optUpdate.value
End Sub

Private Sub PopulateConstraintListBox(Optional UpdateIndex As Long = -1)
          ' If UpdateIndex is specified, refresh that entry without redrawing the entire list

          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim DisplayString As String
          Dim oldLI As Long
          
          Dim showNamedRanges As Boolean
3         showNamedRanges = chkNameRange.value
4         If showNamedRanges Then SearchRangeName_DestroyCache
          
          ' Prevent change events from firing while we modify the list
5         DisableConstraintListChange = True
          
          
6         oldLI = lstConstraints.ListIndex
          
7         If UpdateIndex = -1 Then
8             lstConstraints.Clear
9             lstConstraints.AddItem "<Add new constraint>"
    
              Dim i As Long
10            For i = 1 To NumConstraints
11                lstConstraints.AddItem Constraints(i).ListDisplayString(sheet, showNamedRanges)
12            Next i
    
13        Else
14            If lstConstraints.ListCount > UpdateIndex Then lstConstraints.RemoveItem UpdateIndex
15            lstConstraints.AddItem Constraints(UpdateIndex).ListDisplayString(sheet, showNamedRanges), UpdateIndex
16        End If
          
          ' Restore selected index, forcing a valid selection if needed
17        lstConstraints.ListIndex = Max(Min(oldLI, lstConstraints.ListCount - 1), 0)

          ' Restore change event
18        DisableConstraintListChange = False
          
ExitSub:
19        If RaiseError Then RethrowError
20        Exit Sub

ErrorHandler:
21        If Not ReportError("CModel", "PopulateConstraintListBox") Then Resume
22        RaiseError = True
23        GoTo ExitSub
End Sub

Private Sub chkGetDuals_Click()
1         refDuals.Enabled = chkGetDuals.value
End Sub

Private Sub chkGetDuals2_Click()
1         optUpdate.Enabled = chkGetDuals2.value And chkGetDuals2.Enabled
2         optNew.Enabled = chkGetDuals2.value And chkGetDuals2.Enabled
End Sub

Private Sub chkNameRange_Click()
1         PopulateConstraintListBox
2         ShowNamedRangesState = chkNameRange.value
End Sub

Private Sub chkShowModel_Click()
1               ShowModelAfterSavingState = chkShowModel.value
End Sub

Private Sub cmdCancelCon_Click()
1         lstConstraints.SetFocus  ' Make sure refedits don't get focus
2         Disabler True
3         cmdAddCon.Enabled = False
4         RefreshConstraintPanel
End Sub

Private Sub cmdChange_Click()
          Dim frmSolverChange As FSolverChange
1         Set frmSolverChange = New FSolverChange
          
2         Me.Hide  '  Hide the model form so the refedit on the options form works, and to keep the focus clear
3         frmSolverChange.Show
4         Unload frmSolverChange
          
5         FormatCurrentSolver
6         PreserveModel = True
7         Me.Show
End Sub

Sub FormatCurrentSolver()
          Dim Solver As String
1         Solver = GetChosenSolver(sheet)
2         lblSolver.Caption = "Current Solver Engine: " & UCase(Left(Solver, 1)) & Mid(Solver, 2)
End Sub

Private Sub cmdOptions_Click()
          ' Save the current "Assume Non Negative" option so this is shown in the options dialog.
          ' The saved value gets updated on OK, which we then reflect in our Model dialog
1         SetNonNegativity chkNonNeg.value, sheet
          Dim frmOptions As FOptions
2         Set frmOptions = New FOptions
          
3         Me.Hide  ' Hide the model form so the refedit on the options form works, and to keep the focus clear
4         frmOptions.Show
          
5         Unload frmOptions
6         chkNonNeg.value = GetNonNegativity(sheet)

          ' Restore the original model form
7         PreserveModel = True
8         Me.Show
End Sub

Private Sub cmdReset_Click()
          'Check the user wants to reset the model
1         If MsgBox("This will reset the objective function, decision variables and constraints." _
                  & vbCrLf & " Are you sure you want to do this?", vbYesNo, "OpenSolver") = vbNo Then
2             Exit Sub
3         End If

          'Reset the objective function and the decision variables
4         SetFormObjective vbNullString
5         SetFormDecisionVars vbNullString

          ' Remove the constraints
6         NumConstraints = 0
7         PopulateConstraintListBox
End Sub

Private Sub optMax_Click()
1         txtObjTarget.Enabled = optTarget.value
End Sub

Private Sub optMin_Click()
1         txtObjTarget.Enabled = optTarget.value
End Sub

Private Sub optTarget_Click()
1         txtObjTarget.Enabled = optTarget.value
End Sub

Private Sub AlterConstraints(DoDisable As Boolean)
1               Disabler DoDisable
2               cmdAddCon.Enabled = Not DoDisable Or lstConstraints.ListIndex = 0
End Sub

Private Sub refConLHS_Change()
          ' Only fire the change event if the RefEdit is being edited directly
1         If ActiveControl.Name <> "refConLHS" Then Exit Sub
          
          ' refConLHS has focus, we don't need to worry about setting the focus
2         AlterConstraints Not HasConstraintChanged()
End Sub

Private Sub refConRHS_Change()
          ' Only fire the change event if the RefEdit is being edited directly
1         If ActiveControl.Name <> "refConRHS" Then Exit Sub
          
          ' refConRHS has focus, we don't need to worry about setting the focus
2         AlterConstraints Not HasConstraintChanged()
End Sub

Private Sub cboConRel_Change()
          ' Caller should ensure focus is set correctly
1         refConRHS.Enabled = RelationHasRHS(GetFormRelation())
2         AlterConstraints Not HasConstraintChanged()
End Sub

Private Function HasConstraintChanged() As Boolean
          Dim LHSChanged As Boolean, RelChanged As Boolean, RHSChanged As Boolean

          Dim Index As Long
1         Index = lstConstraints.ListIndex
          
          ' If there is a selected constraint, we check against the original values
          ' otherwise we compare to empty strings
          Dim OrigLHS As String, OrigRHSLocal As String
2         If Index >= 1 Then
3             With Constraints(Index)
4                 OrigLHS = .LHSRefersTo
5                 OrigRHSLocal = .RHSRefersToLocal
                  
6                 RelChanged = (GetFormRelation() <> .Relation)
7             End With
8         End If
          
9         LHSChanged = GetFormLHSRefersTo() <> OrigLHS
          ' Only check the RHS if the relation uses RHS
10        RHSChanged = RelationHasRHS(GetFormRelation()) And _
                       GetFormRHSRefersTo() <> OrigRHSLocal
          
11        HasConstraintChanged = LHSChanged Or RelChanged Or RHSChanged
End Function

Private Sub UserForm_Activate()
1         CenterForm
2         On Error GoTo ErrorHandler

3         cmdCancel.SetFocus

          ' Check if we have indicated to keep the model from the last time form was shown
4         If PreserveModel Then
5             PreserveModel = False
6         Else
7             UpdateStatusBar "Loading model...", True
8             Application.Cursor = xlWait
9             Application.ScreenUpdating = False

              ' Check we can even start
10            GetActiveSheetIfMissing sheet
             
11            SetAnyMissingDefaultSolverOptions sheet
          
12            If SheetHasOpenSolverHighlighting(sheet) Then
13                RestoreHighlighting = True
14                HideSolverModel sheet
15            End If
          
              ' Make sure sheet is up to date
16            Application.Calculate
              ' Remove the 'marching ants' showing if a range is copied.
              ' Otherwise, the ants stay visible, and visually conflict with
              ' our cell selection. The ants are also left behind on the
              ' screen. This works around an apparent bug (?) in Excel 2007.
17            Application.CutCopyMode = False
          
              ' Clear the form
18            SetFormObjectiveSense UnknownObjectiveSense
19            SetFormObjective vbNullString
20            SetFormDecisionVars vbNullString
21            SetFormLHSRefersTo vbNullString
22            SetFormRHSRefersTo vbNullString
23            lstConstraints.Clear
24            cboConRel.Clear
25            cboConRel.AddItem "="
26            cboConRel.AddItem "<="
27            cboConRel.AddItem ">="
28            cboConRel.AddItem "int"
29            cboConRel.AddItem "bin"
30            cboConRel.AddItem "alldiff"
31            SetFormRelation RelationEQ    ' We set an initial value just in case there is no model, and the user goes straight to AddNewConstraint

              'Find current solver
32            FormatCurrentSolver

              ' Load the model on the sheet into memory
33            LoadModelFromSheet
34        End If
          
35        PopulateConstraintListBox
36        RefreshConstraintPanel
              
37        chkNameRange.value = ShowNamedRangesState
38        chkShowModel.value = ShowModelAfterSavingState
        
ExitSub:
39        Application.StatusBar = False
40        Application.Cursor = xlDefault
41        Application.ScreenUpdating = True
42        Exit Sub

ErrorHandler:
43        If RestoreHighlighting Then ShowSolverModel sheet, HandleError:=True
44        Me.Hide
45        ReportError "FModel", "UserForm_Activate", True
46        GoTo ExitSub
End Sub


Private Sub cmdCancel_Click()
1         If RestoreHighlighting Then ShowSolverModel sheet, HandleError:=True
2         DoEvents
3         On Error Resume Next ' Just to be safe on our select
4         Application.CutCopyMode = False
5         ActiveCell.Select ' Just select one cell, choosing a cell that should be visible to avoid scrolling
6         Application.ScreenUpdating = True
7         Me.Hide
End Sub

' Make the [x] hide the form rather than unload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
          ' If CloseMode = vbFormControlMenu then we know the user
          ' clicked the [x] close button or Alt+F4 to close the form.
1         If CloseMode = vbFormControlMenu Then
2             cmdCancel_Click
3             Cancel = True
4         End If
End Sub

Private Sub cmdRunAutoModel_Click()
          ' Refedits on Mac don't work if more than one form is shown, so we need to hide it
          ' We also hide on windows to make sure that the forms don't hide each other
1         Me.Hide

          Dim AutoModel As CAutoModel
2         Set AutoModel = New CAutoModel
3         If AutoModel.BuildModel(sheet, MinimiseUserInteraction:=False, SaveAfterBuilding:=False) Then
              ' Copy results into FModel
4             SetFormObjective AutoModel.ObjectiveFunctionCellRefersTo
5             SetFormDecisionVars RangeToRefersTo(AutoModel.DecVarsRange)
6             SetFormObjectiveSense AutoModel.ObjSense
              
7             NumConstraints = AutoModel.Constraints.Count
8             ReDim Constraints(1 To NumConstraints) As CConstraint
              Dim c As CAutoModelConstraint, i As Long
9             i = 1
10            For Each c In AutoModel.Constraints
11                Set Constraints(i) = New CConstraint
12                With c
13                    Constraints(i).Update RangeToRefersTo(.LHS), .RelationType, RangeToRefersTo(.RHS), sheet
14                End With
15                i = i + 1
16            Next c
17        End If
          
18        lstConstraints.ListIndex = -1  ' Clear the selection
19        PreserveModel = True
20        Me.Show
End Sub

Private Sub cmdBuild_Click()
1         On Error GoTo ErrorHandler

2         DoEvents
3         On Error Resume Next
4         Application.CutCopyMode = False
5         ActiveCell.Select ' Just select one cell, choosing a cell that should be visible to avoid scrolling
6         Application.ScreenUpdating = True
7         On Error GoTo ErrorHandler

          Dim oldCalculationMode As Long
8         oldCalculationMode = Application.Calculation
9         Application.Calculation = xlCalculationManual
          
          
          '''''
          ' Check if the model is valid before saving anything!
          '''''
          
          ' Check objective
          Dim ObjectiveFunctionCellRefersTo As String
10        On Error GoTo BadObjRef
11        ObjectiveFunctionCellRefersTo = GetFormObjective()
12        ValidateObjectiveFunctionCellRefersTo ObjectiveFunctionCellRefersTo
13        On Error GoTo ErrorHandler

          ' Check objective sense
          Dim ObjectiveSense As ObjectiveSenseType
14        ObjectiveSense = GetFormObjectiveSense()
15        If ObjectiveSense = UnknownObjectiveSense Then
16            RaiseUserError "Please select an objective sense (minimise, maximise or target)."
17        End If

          ' Check objective target
18        If ObjectiveSense = TargetObjective Then
              Dim ObjectiveTarget As Double
19            On Error GoTo BadObjectiveTarget
20            ObjectiveTarget = GetFormObjectiveTargetValue()
21            On Error GoTo ErrorHandler
22        End If

          ' Check decision variables
          Dim DecisionVariablesRefersTo As String
23        On Error GoTo BadDecRef
24        DecisionVariablesRefersTo = GetFormDecisionVars()
25        ValidateDecisionVariablesRefersTo DecisionVariablesRefersTo
26        On Error GoTo ErrorHandler

          ' Check duals range
27        If chkGetDuals.value Then
              Dim DualsRefersTo As String
28            On Error GoTo BadDualsRef
29            DualsRefersTo = GetFormDuals()
30            ValidateDualsRefersTo DualsRefersTo
31            On Error GoTo ErrorHandler
32        End If

          ' Check all constraints are valid
          Dim c As Long
33        For c = 1 To NumConstraints
34            With Constraints(c)
35                On Error Resume Next
36                ValidateConstraintRefersTo .LHSRefersTo, .Relation, .RHSRefersTo, sheet
37                If Err.Number <> 0 Then
38                    MsgBox "There is an error in constraint " & c & ":" & vbNewLine & Err.Description
39                    GoTo ExitSub
40                End If
41            End With
42        Next c
43        On Error GoTo ErrorHandler
          
          ' Check int/bin constraints not set on non-decision variables before we start saving things
          ' TODO change to error ref #170
44        For c = 1 To NumConstraints
45            If Not RelationHasRHS(Constraints(c).Relation) Then
46                If Not SetDifference(Constraints(c).LHSRange, GetRefersToRange(DecisionVariablesRefersTo)) Is Nothing Then
47                    If MsgBox("This model has specified that a non-decision cell must take an integer/binary value. " & _
                                "This is a valid model, but not one that OpenSolver can solve. " & _
                                "Do you wish to continue with saving this model?", _
                                vbQuestion + vbYesNo, "OpenSolver - Warning") = vbYes Then
48                        Exit For
49                    Else
50                        GoTo ExitSub
51                    End If
52                End If
53            End If
54        Next c

          ''''''''
          ' Build is now confirmed, save everything to sheet
          ''''''''
55        SetDecisionVariablesRefersTo DecisionVariablesRefersTo, sheet
56        SetObjectiveFunctionCellRefersTo ObjectiveFunctionCellRefersTo, sheet
57        SetObjectiveSense ObjectiveSense, sheet
58        If ObjectiveSense = TargetObjective Then SetObjectiveTargetValue ObjectiveTarget, sheet
          
          ' Constraints
59        SetNumConstraints NumConstraints, sheet
60        For c = 1 To NumConstraints
61            With Constraints(c)
62                UpdateConstraintRefersTo c, .LHSRefersTo, .Relation, .RHSRefersTo, sheet
63            End With
64        Next c

          ' Options
65        SetNonNegativity chkNonNeg.value, sheet
          ' TODO: Only update these if doing sensitivity- needs a change to save chkGetDuals
          'If chkGetDuals.value Then
66            SetDualsRefersTo DualsRefersTo, sheet
67            SetDualsOnSheet chkGetDuals2.value, sheet
68            SetUpdateSensitivity optUpdate.value, sheet
          'End If

          ' Display on screen
69        If chkShowModel.value = True Then ShowSolverModel sheet, HandleError:=True
              
70        On Error GoTo CalculateFailed
71        Application.Calculate
72        On Error GoTo ErrorHandler
          
73        Me.Hide
74        GoTo ExitSub

CalculateFailed:
          ' Application.Calculate failed. Ignore error and try again
75        On Error GoTo ErrorHandler
76        Application.Calculate
77        Resume Next

BadObjRef:
          ' Couldn't turn the objective cell address into a range
78        MsgBox "Error: the cell address for the objective is invalid. " + _
                 "This must be a single cell. " & _
                 "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
79        refObj.SetFocus ' Set the focus back to the RefEdit
80        GoTo ExitSub
BadDecRef:
          ' Try with raw text if it might work that way!
          ' Converting the text to RefersTo can increase the length (with sheet names etc)
          ' and could cause the range text to become too long for Excel
81        If DecisionVariablesRefersTo <> Me.refDecision.Text Then
82            DecisionVariablesRefersTo = Me.refDecision.Text
83            Resume ' Try again on the validation
84        End If
          
          ' Couldn't turn the decision variable address into a range
85        MsgBox "Error: the cell range specified for the Variable Cells is invalid. " + _
                 "This must be a valid Excel range that does not exceed Excel's internal character count limits. " + _
                 "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
86        refDecision.SetFocus ' Set the focus back to the RefEdit
87        GoTo ExitSub
BadObjectiveTarget:
          ' Couldn't turn the objective target into a value
88        MsgBox "Error: the target value for the objective cell is invalid. " + _
                 "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
89        txtObjTarget.SetFocus ' Set the focus back to the target text box
90        GoTo ExitSub
BadDualsRef:
          ' Couldn't turn the dual cell into a range
91        MsgBox "Error: the cell for storing the shadow prices is invalid. " + _
                 "This must be a single cell. " & _
                 "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
92        refDuals.SetFocus ' Set the focus back to the refedit
93        GoTo ExitSub

ExitSub:
94        Application.Calculation = oldCalculationMode
95        Exit Sub

ErrorHandler:
96        ReportError "FModel", "cmdBuild_Click", True
97        GoTo ExitSub
End Sub

Private Sub cmdAddCon_Click()
1         On Error GoTo ErrorHandler

          Dim LHSRefersTo As String, RHSRefersTo As String, rel As RelationConsts
2         LHSRefersTo = GetFormLHSRefersTo()
3         RHSRefersTo = ConvertFromCurrentLocale(GetFormRHSRefersTo())
4         rel = GetFormRelation()

5         ValidateConstraintRefersTo LHSRefersTo, rel, RHSRefersTo, sheet
          
          ' Valid, now update
          
6         lstConstraints.SetFocus  ' Set focus back to constraint list to avoid disabling problems
7         AlterConstraints True

          Dim conIndex As Long
8         If cmdAddCon.Caption <> "Add constraint" Then
              ' Update constraint
              ' Note that we use the saved index value rather than lstConstraints.ListIndex in case
              ' this update was prompted by the user changing the index on an unsaved constraint.
              ' In this case, lstConstraints.ListIndex does NOT reflect the constraint being altered.
9             conIndex = CurrentListIndex
10        Else
              ' Add constraint
11            NumConstraints = NumConstraints + 1
12            ReDim Preserve Constraints(1 To NumConstraints) As CConstraint
13            conIndex = NumConstraints
14        End If
                   
15        Set Constraints(conIndex) = New CConstraint
16        Constraints(conIndex).Update LHSRefersTo, rel, RHSRefersTo, sheet
          
17        PopulateConstraintListBox conIndex
18        RefreshConstraintPanel
19        Exit Sub

ErrorHandler:
20        Application.DisplayAlerts = True
21        MsgBox Err.Description
22        Exit Sub

End Sub

Private Sub cmdDelSelCon_Click()
1         If lstConstraints.ListIndex = -1 Then
2             MsgBox "Please select a constraint!", vbOKOnly, "AutoModel"
3             Exit Sub
4         End If

          ' Set focus back to constraint box to avoid focus jumping around
5         lstConstraints.SetFocus

          Dim i As Long
6         For i = 1 To NumConstraints - 1
7             If i >= lstConstraints.ListIndex Then
8                 Set Constraints(i) = Constraints(i + 1)
9             End If
10        Next i
11        NumConstraints = NumConstraints - 1
12        If NumConstraints > 0 Then ReDim Preserve Constraints(1 To NumConstraints) As CConstraint
          
          ' Update form
13        PopulateConstraintListBox
14        lstConstraints_Change
End Sub

Private Sub lstConstraints_Change()
1         If DisableConstraintListChange Then Exit Sub

2         If cmdCancelCon.Enabled Then
              Dim SaveChanges As Boolean
3             If cmdAddCon.Caption = "Update constraint" Then
4                 SaveChanges = (MsgBox("You have made changes to the current constraint." & vbNewLine & _
                                        "Do you want to save these changes?", vbYesNo) = vbYes)
5             Else
6                 SaveChanges = (MsgBox("You have entered a constraint." & vbNewLine & _
                                        "Do you want to save this as a new constraint?", vbYesNo) = vbYes)
7             End If
8             If SaveChanges Then
9                 cmdAddCon_Click
10            End If
11            AlterConstraints True
12        End If
          
13        RefreshConstraintPanel
          
End Sub

Private Sub RefreshConstraintPanel()
          ' Save the index of the constraint we will display
1         CurrentListIndex = lstConstraints.ListIndex

2         If lstConstraints.ListIndex = -1 Then
3             Exit Sub
4         End If
5         If lstConstraints.ListIndex = 0 Then
              'Add constraint
6             SetFormLHSRefersTo vbNullString
7             SetFormRHSRefersTo vbNullString
8             cboConRel_Change ' Set the RHS to be active based on the last constraint
9             cmdAddCon.Enabled = True
10            cmdAddCon.Caption = "Add constraint"
11            cmdDelSelCon.Enabled = False
12            Application.CutCopyMode = False
13            ActiveCell.Select ' Just select one cell, choosing a cell that should be visible to avoid scrolling
14            Application.ScreenUpdating = True
15        Else
              ' Update constraint
16            cmdAddCon.Enabled = False
17            cmdAddCon.Caption = "Update constraint"
18            cmdDelSelCon.Enabled = True

              ' Set text and marching ants; the select/copy operations can throw errors if on different sheets
19            On Error Resume Next
20            ActiveCell.Select   ' We may fail in the next steps, so we cancel any old highlighting
21            Application.CutCopyMode = False
              Dim copyRange As Range
              
22            With Constraints(lstConstraints.ListIndex)
23                SetFormLHSRefersTo .LHSRefersTo
24                SetFormRelation .Relation
25                SetFormRHSRefersTo .RHSRefersTo

26                Set copyRange = ProperUnion(.LHSRange, .RHSRange)
27            End With
28
29            cboConRel_Change

              ' Will fail if LHS and RHS are different shape
              ' Silently fail, nothing that can be done about it
30            copyRange.Select
31            copyRange.Copy
32            Application.ScreenUpdating = True
33        End If
End Sub

Private Sub lstConstraints_Exit(ByVal Cancel As MSForms.ReturnBoolean)
          ' When the focus leaves this list, we want to remove any highlighting shown by selected cells
1         ActiveCell.Select
End Sub

Private Sub UserForm_Initialize()
1         AutoLayout
2         CenterForm
End Sub

Private Sub AutoLayout()
1         AutoFormat Me.Controls
          
2         Me.Width = FormWidthModel
          
3         With cmdRunAutoModel
4             .Caption = "AutoModel"
5             .Width = FormButtonWidth * 1.3
6             .Top = FormMargin
7             .Left = LeftOfForm(Me.Width, .Width)
8         End With
          
9         With lblDescHeader
10            .Left = FormMargin
11            .Top = cmdRunAutoModel.Top + FormSpacing
12            .Caption = "What is AutoModel?"
13        End With
          
14        With lblDesc
15            .Left = lblDescHeader.Left
16            .Top = Below(cmdRunAutoModel)
17            .Caption = "AutoModel is a feature of OpenSolver that tries to automatically determine " & _
                         "the problem you are trying to optimise by observing the structure of the " & _
                         "spreadsheet. It will turn its best guess into a Solver model, which you can " & _
                         "then edit in this window."
18            AutoHeight lblDesc, Me.Width - 2 * FormMargin
19        End With
          
20        With lblDiv1
21            .Left = lblDescHeader.Left
22            .Top = Below(lblDesc)
23            .Width = lblDesc.Width
24            .Height = FormDivHeight
25            .BackColor = FormDivBackColor
26        End With
          
27        With lblStep1
28            .Caption = "Objective Cell:"
29            .Left = lblDescHeader.Left
30            .Top = Below(lblDiv1)
31            AutoHeight lblStep1, Me.Width, True
32        End With
          
33        With txtObjTarget
34            .Width = cmdRunAutoModel.Width
35            .Top = lblStep1.Top
36            .Left = LeftOfForm(Me.Width, .Width)
37        End With
          
38        With optTarget
39            .Caption = "target value:"
40            AutoHeight optTarget, Me.Width, True
41            .Top = lblStep1.Top
42            .Left = LeftOf(txtObjTarget, .Width)
43        End With
          
44        With optMin
45            .Caption = "minimise"
46            AutoHeight optMin, Me.Width, True
47            .Top = lblStep1.Top
48            .Left = LeftOf(optTarget, .Width)
49        End With
          
50        With optMax
51            .Caption = "maximise"
52            AutoHeight optMax, Me.Width, True
53            .Top = lblStep1.Top
54            .Left = LeftOf(optMin, .Width)
55        End With
          
56        With refObj
57            .Left = RightOf(lblStep1)
58            .Top = lblStep1.Top
59            .Width = LeftOf(optMax, .Left)
60        End With
          
61        With lblDiv2
62            .Left = lblDescHeader.Left
63            .Top = Below(optMax)
64            .Width = lblDesc.Width
65            .Height = FormDivHeight
66            .BackColor = FormDivBackColor
67        End With
          
68        With lblStep2
69            .Caption = "Variable Cells:"
70            .Left = lblDescHeader.Left
71            .Top = Below(lblDiv2)
72            AutoHeight lblStep2, Me.Width, True
73        End With
          
74        With refDecision
75            .Height = 2 * refObj.Height
76            .Left = refObj.Left
77            .Top = lblStep2.Top
78            .Width = LeftOfForm(Me.Width, .Left)
79        End With
          
80        With lblDiv3
81            .Left = lblDescHeader.Left
82            .Top = Below(refDecision)
83            .Width = lblDesc.Width
84            .Height = FormDivHeight
85            .BackColor = FormDivBackColor
86        End With
          
87        With lblStep3
88            .Caption = "Constraints:"
89            .Left = lblDescHeader.Left
90            .Top = Below(lblDiv3)
91            .Width = lblDesc.Width
92        End With
          
93        lblConstraintGroup.Top = Below(lblStep3, False)
          
94        With cboConRel
95            .Width = cmdRunAutoModel.Width / 2
96            .Height = refObj.Height
97            .Top = lblConstraintGroup.Top + FormSpacing
98            .Left = LeftOfForm(Me.Width, .Width) - FormSpacing
99        End With
          
100       With refConLHS
101           .Width = cboConRel.Width * 3
102           .Height = refObj.Height
103           .Top = cboConRel.Top
104           .Left = LeftOf(cboConRel, .Width)
105       End With
          
106       With refConRHS
107           .Width = refConLHS.Width
108           .Left = refConLHS.Left
109           .Height = refObj.Height
110           .Top = Below(refConLHS)
111       End With
          
112       With cmdAddCon
113           .Caption = "Add constraint"
114           .Left = refConLHS.Left
115           .Top = Below(refConRHS)
116           .Width = cboConRel.Width * 2
117       End With
          
118       With cmdCancelCon
119           .Caption = "Cancel"
120           .Left = RightOf(cmdAddCon)
121           .Top = cmdAddCon.Top
122           .Width = cmdAddCon.Width
123       End With
          
124       With lblConstraintGroup
125           .Left = refConLHS.Left - FormSpacing
126           .Width = FormSpacing * 3 + cmdAddCon.Width + cmdCancelCon.Width
127           .Height = FormSpacing * 4 + refConLHS.Height + refConRHS.Height + cmdAddCon.Height
128       End With

          ' Hack: hide lblConstraintGroup on Excel builds over 73xx (ref #256)
          #If Win32 Then
              If Application.Build > 7300 Then
                  lblConstraintGroup.Visible = False
              End If
          #End If
          
129       With cmdDelSelCon
130           .Caption = "Delete selected constraint"
131           .Left = lblConstraintGroup.Left
132           .Top = Below(lblConstraintGroup)
133           .Width = lblConstraintGroup.Width
134       End With
          
135       With chkNonNeg
136           .Caption = "Make unconstrained variable cells non-negative"
137           .Left = lblConstraintGroup.Left
138           .Top = Below(cmdDelSelCon)
139           .Width = lblConstraintGroup.Width
140       End With
          
141       With chkNameRange
142           .Left = lblConstraintGroup.Left
143           .Width = lstConstraints.Width
144           .Caption = "Show named ranges in constraint list"
145           .Top = Below(chkNonNeg, False)
146       End With
          
147       With lstConstraints
148           .Left = lblDescHeader.Left
149           .Top = lblConstraintGroup.Top
150           .Height = MinHeight
151           .Width = LeftOf(lblConstraintGroup, .Left)
152       End With
          
153       With lblDiv4
154           .Left = lblDescHeader.Left
155           .Width = lblDesc.Width
156           .Height = FormDivHeight
157           .BackColor = FormDivBackColor
158       End With
          
159       With lblStep4
160           .Caption = "Sensitivity Analysis"
161           .Left = lblDescHeader.Left
162           AutoHeight lblStep4, Me.Width, True
163       End With
          
164       With chkGetDuals
165           .Caption = "List sensitivity analysis on the same sheet with top left cell:"
166           .Left = RightOf(lblStep4)
167           AutoHeight chkGetDuals, Me.Width, True
168       End With
          
169       With refDuals
170           .Left = RightOf(chkGetDuals)
171           .Width = LeftOfForm(Me.Width, .Left)
172           .Height = refObj.Height
173       End With
          
174       With chkGetDuals2
175           .Caption = "Output sensitivity analysis:"
176           .Left = chkGetDuals.Left
177           AutoHeight chkGetDuals2, Me.Width, True
178       End With
          
179       With optUpdate
180           .Caption = "updating any previous output sheet"
181           .Left = RightOf(chkGetDuals2)
182           AutoHeight optUpdate, Me.Width, True
183       End With
          
184       With optNew
185           .Caption = "on a new sheet"
186           .Left = RightOf(optUpdate)
187           AutoHeight optNew, Me.Width, True
188       End With
          
189       With lblDiv5
190           .Left = lblDescHeader.Left
191           .Width = lblDesc.Width
192           .Height = FormDivHeight
193           .BackColor = FormDivBackColor
194       End With
          
195       With lblStep5
196           .Left = lblDescHeader.Left
197           .Caption = "Solver Engine:"
198       End With
          
199       With cmdChange
200           .Width = cmdRunAutoModel.Width
201           .Left = LeftOfForm(Me.Width, .Width)
202           .Caption = "Solver Engine..."
203       End With
          
204       With lblSolver
205           .Width = LeftOf(cmdChange, .Left)
206       End With
          
207       With lblDiv6
208           .Left = lblDescHeader.Left
209           .Width = lblDesc.Width
210           .Height = FormDivHeight
211           .BackColor = FormDivBackColor
212       End With
              
213       With chkShowModel
214           .Left = lblDescHeader.Left
215           AutoHeight chkShowModel, Me.Width, True
216       End With
          
217       With cmdCancel
218           .Width = cmdRunAutoModel.Width
219           .Caption = "Cancel"
220           .Left = LeftOfForm(Me.Width, .Width)
221           .Cancel = True
222       End With
          
223       With cmdBuild
224           .Width = cmdRunAutoModel.Width
225           .Caption = "Save Model"
226           .Left = LeftOf(cmdCancel, .Width)
227       End With
          
228       With cmdOptions
229           .Width = cmdRunAutoModel.Width
230           .Caption = "Options..."
231           .Left = LeftOf(cmdBuild, .Width)
232       End With
          
233       With cmdReset
234           .Width = cmdRunAutoModel.Width
235           .Caption = "Clear Model"
236           .Left = LeftOf(cmdOptions, .Width)
237       End With
          
          ' Add resizer
238       With lblResizer
        #If Mac Then
                  ' Mac labels don't fire MouseMove events correctly
239               .Visible = False
        #End If
240           .Caption = "o"
241           With .Font
242               .Name = "Marlett"
243               .Charset = 2
244               .Size = 10
245           End With
246           .AutoSize = True
247           .Left = Me.Width - .Width
248           .MousePointer = fmMousePointerSizeNWSE
249           .BackStyle = fmBackStyleTransparent
250       End With
          
          ' Set the vertical positions of the lower half of the form
251       UpdateLayout
          
252       Me.Width = Me.Width + FormWindowMargin
          
253       Me.BackColor = FormBackColor
254       Me.Caption = "OpenSolver - Model"
End Sub

Private Sub UpdateLayout(Optional ChangeY As Single = 0)
      ' Do the layout of the lower half of the form, changing the height of the list box by ChangeY
          Dim NewHeight As Double
1         NewHeight = Max(lstConstraints.Height + ChangeY, MinHeight)
          
2         lstConstraints.Height = NewHeight
              
          ' Cascade the updated height
3         lblDiv4.Top = Below(lstConstraints)
4         lblStep4.Top = Below(lblDiv4)
5         chkGetDuals.Top = lblStep4.Top
6         refDuals.Top = lblStep4.Top
7         chkGetDuals2.Top = Below(chkGetDuals, False)
8         optUpdate.Top = chkGetDuals2.Top
9         optNew.Top = chkGetDuals2.Top
10        lblDiv5.Top = Below(optNew, False)
11        lblStep5.Top = Below(lblDiv5)
12        cmdChange.Top = lblStep5.Top
13        lblSolver.Top = lblStep5.Top + FormButtonHeight - FormTextHeight
14        lblDiv6.Top = Below(cmdChange)
15        chkShowModel.Top = Below(lblDiv6)
16        cmdCancel.Top = chkShowModel.Top
17        cmdBuild.Top = chkShowModel.Top
18        cmdOptions.Top = chkShowModel.Top
19        cmdReset.Top = chkShowModel.Top
20        Me.Height = FormHeight(cmdCancel)
21        lblResizer.Top = Me.InsideHeight - lblResizer.Height
End Sub

Private Sub lblResizer_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
1         If Button = 1 Then
2             ResizeStartY = Y
3         End If
End Sub

Private Sub lblResizer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
1         If Button = 1 Then
        #If Mac Then
                  ' Mac reports delta already
2                 UpdateLayout Y
        #Else
3                 UpdateLayout (Y - ResizeStartY)
        #End If
4         End If
End Sub

Private Sub CenterForm()
1         Me.Top = CenterFormTop(Me.Height)
2         Me.Left = CenterFormLeft(Me.Width)
End Sub

Private Function GetFormObjectiveSense() As ObjectiveSenseType
1         If optMax.value = True Then GetFormObjectiveSense = MaximiseObjective
2         If optMin.value = True Then GetFormObjectiveSense = MinimiseObjective
3         If optTarget.value = True Then GetFormObjectiveSense = TargetObjective
End Function
Private Sub SetFormObjectiveSense(ObjectiveSense As ObjectiveSenseType)
1         Select Case ObjectiveSense
          Case MaximiseObjective: optMax.value = True
2         Case MinimiseObjective: optMin.value = True
3         Case TargetObjective:   optTarget.value = True
4         Case UnknownObjectiveSense
5             optMax.value = False
6             optMin.value = False
7             optTarget.value = False
8         End Select
End Sub

Private Function GetFormObjectiveTargetValue() As Double
1         GetFormObjectiveTargetValue = CDbl(txtObjTarget.Text)
End Function
Private Sub SetFormObjectiveTargetValue(ObjectiveTargetValue As Double)
1         txtObjTarget.Text = CStr(ObjectiveTargetValue)
End Sub

Private Function GetFormObjective() As String
1         GetFormObjective = RefEditToRefersTo(refObj.Text)
End Function
Private Sub SetFormObjective(ObjectiveRefersTo As String)
1         refObj.Text = GetDisplayAddress(ObjectiveRefersTo, sheet, False)
End Sub

Private Function GetFormLHSRefersTo() As String
1         GetFormLHSRefersTo = RefEditToRefersTo(refConLHS.Text)
End Function
Private Sub SetFormLHSRefersTo(newLHSRefersTo As String)
1         refConLHS.Text = GetDisplayAddress(newLHSRefersTo, sheet, False)
End Sub

Private Function GetFormRHSRefersTo() As String
1         If RelationHasRHS(GetFormRelation()) Then
              ' We don't convert from current locale here to avoid causing errors in the RefEdit events
2             GetFormRHSRefersTo = RefEditToRefersTo(refConRHS.Text)
3         Else
4             GetFormRHSRefersTo = vbNullString
5         End If
End Function
Private Sub SetFormRHSRefersTo(newRHSRefersTo As String)
1         refConRHS.Text = ConvertToCurrentLocale(GetDisplayAddress(newRHSRefersTo, sheet, False))
End Sub

Private Function GetFormRelation() As RelationConsts
1         GetFormRelation = RelationStringToEnum(cboConRel.Text)
End Function
Private Sub SetFormRelation(newRelation As RelationConsts)
1         cboConRel.ListIndex = cboPosition(RelationEnumToString(newRelation))
End Sub

' We allow multiple area ranges for the variables, which requires ConvertLocale as delimiter can vary
Private Function GetFormDecisionVars() As String
1         GetFormDecisionVars = ConvertFromCurrentLocale(RefEditToRefersTo(refDecision.Text))
End Function
Private Sub SetFormDecisionVars(newDecisionVars As String)
1         refDecision.Text = ConvertToCurrentLocale(GetDisplayAddress(newDecisionVars, sheet, False))
End Sub

Private Function GetFormDuals() As String
1         GetFormDuals = RefEditToRefersTo(refDuals.Text)
End Function
Private Sub SetFormDuals(newDuals As String)
1         refDuals.Text = GetDisplayAddress(newDuals, sheet, False)
End Sub
