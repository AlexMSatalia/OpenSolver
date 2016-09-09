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
4148      Select Case rel
          Case "="
4149          cboPosition = 0
4150      Case "<="
4151          cboPosition = 1
4152      Case ">="
4153          cboPosition = 2
4154      Case "int"
4155          cboPosition = 3
4156      Case "bin"
4157          cboPosition = 4
4158      Case "alldiff"
4159          cboPosition = 5
4160      End Select
End Function

Sub Disabler(TrueIfEnable As Boolean)
          ' Before calling Disabler, ensure that the focused control will not be disabled!
          ' This prevents tab-skipping into a RefEdit, and starting its event loop
          
4161      lblDescHeader.Enabled = TrueIfEnable
4162      lblDesc.Enabled = TrueIfEnable
4163      cmdRunAutoModel.Enabled = TrueIfEnable

4164      lblDiv1.Enabled = False

4165      lblStep1.Enabled = TrueIfEnable
4166      refObj.Enabled = TrueIfEnable
4167      optMax.Enabled = TrueIfEnable
4168      optMin.Enabled = TrueIfEnable
4169      optTarget.Enabled = TrueIfEnable
4170      txtObjTarget.Enabled = TrueIfEnable And GetFormObjectiveSense() = TargetObjective

4171      lblDiv2.Enabled = False

4172      lblStep2.Enabled = TrueIfEnable
4173      refDecision.Enabled = TrueIfEnable

4174      lblDiv3.Enabled = False

4175      chkNonNeg.Enabled = TrueIfEnable
4176      cmdCancelCon.Enabled = Not TrueIfEnable
4177      cmdDelSelCon.Enabled = TrueIfEnable And lstConstraints.ListIndex > 0
          chkNameRange.Enabled = TrueIfEnable

4178      lblDiv4.Enabled = False

4179      lblStep4.Enabled = TrueIfEnable

          ' Regardless of TrueIfEnable, disable the dual options if the solver doesn't support it
          Dim HasSensitivity As Boolean
          HasSensitivity = SensitivityAnalysisAvailable(CreateSolver(GetChosenSolver(sheet)))

4180      chkGetDuals.Enabled = TrueIfEnable And HasSensitivity
4184      refDuals.Enabled = TrueIfEnable And HasSensitivity And chkGetDuals.value
4181      chkGetDuals2.Enabled = TrueIfEnable And HasSensitivity
          optUpdate.Enabled = TrueIfEnable And HasSensitivity And chkGetDuals2.value
          optNew.Enabled = TrueIfEnable And HasSensitivity And chkGetDuals2.value

4185      lblDiv5.Enabled = False
4186      lblDiv6.Enabled = False

4187      chkShowModel.Enabled = TrueIfEnable
4188      cmdOptions.Enabled = TrueIfEnable
4189      cmdBuild.Enabled = TrueIfEnable
4190      cmdCancel.Enabled = TrueIfEnable
          cmdReset.Enabled = TrueIfEnable
          cmdChange.Enabled = TrueIfEnable
End Sub

Sub LoadModelFromSheet()
          SetFormObjectiveSense GetObjectiveSense(sheet)
          SetFormObjectiveTargetValue GetObjectiveTargetValue(sheet)
          SetFormObjective GetObjectiveFunctionCellRefersTo(sheet)
4215      SetFormDecisionVars GetDecisionVariablesRefersTo(sheet)
4218      SetFormDuals GetDualsRefersTo(sheet)

          NumConstraints = GetNumConstraints(sheet)
          If NumConstraints > 0 Then
              ReDim Constraints(1 To NumConstraints) As CConstraint
          
              Dim i As Long, LHSRefersTo As String, RHSRefersTo As String, Relation As RelationConsts
              For i = 1 To NumConstraints
                  GetConstraintRefersTo i, LHSRefersTo, Relation, RHSRefersTo, sheet
                  Set Constraints(i) = New CConstraint
                  Constraints(i).Update LHSRefersTo, Relation, RHSRefersTo, sheet
              Next i
          End If

4213      chkNonNeg.value = GetNonNegativity(sheet)
4216      chkGetDuals.value = Len(GetFormDuals()) > 0
          chkGetDuals2.value = GetDualsOnSheet(sheet)
4237      chkGetDuals2_Click  ' Set enabled status of dual checkboxes

          optUpdate.value = GetUpdateSensitivity(sheet)
4242      optNew.value = Not optUpdate.value
End Sub

Private Sub PopulateConstraintListBox(Optional UpdateIndex As Long = -1)
          ' If UpdateIndex is specified, refresh that entry without redrawing the entire list

          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim DisplayString As String
          Dim oldLI As Long
          
          Dim showNamedRanges As Boolean
          showNamedRanges = chkNameRange.value
3984      If showNamedRanges Then SearchRangeName_DestroyCache
          
          ' Prevent change events from firing while we modify the list
          DisableConstraintListChange = True
          
          
3985      oldLI = lstConstraints.ListIndex
          
          If UpdateIndex = -1 Then
3986          lstConstraints.Clear
3987          lstConstraints.AddItem "<Add new constraint>"
    
              Dim i As Long
3988          For i = 1 To NumConstraints
4001              lstConstraints.AddItem Constraints(i).ListDisplayString(sheet, showNamedRanges)
4002          Next i
    
          Else
              If lstConstraints.ListCount > UpdateIndex Then lstConstraints.RemoveItem UpdateIndex
              lstConstraints.AddItem Constraints(UpdateIndex).ListDisplayString(sheet, showNamedRanges), UpdateIndex
          End If
          
          ' Restore selected index, forcing a valid selection if needed
4006      lstConstraints.ListIndex = Max(Min(oldLI, lstConstraints.ListCount - 1), 0)

          ' Restore change event
          DisableConstraintListChange = False
          
ExitSub:
          If RaiseError Then RethrowError
          Exit Sub

ErrorHandler:
          If Not ReportError("CModel", "PopulateConstraintListBox") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Private Sub chkGetDuals_Click()
4250      refDuals.Enabled = chkGetDuals.value
End Sub

Private Sub chkGetDuals2_Click()
4253      optUpdate.Enabled = chkGetDuals2.value And chkGetDuals2.Enabled
4254      optNew.Enabled = chkGetDuals2.value And chkGetDuals2.Enabled
End Sub

Private Sub chkNameRange_Click()
4256      PopulateConstraintListBox
          ShowNamedRangesState = chkNameRange.value
End Sub

Private Sub chkShowModel_Click()
          ShowModelAfterSavingState = chkShowModel.value
End Sub

Private Sub cmdCancelCon_Click()
          lstConstraints.SetFocus  ' Make sure refedits don't get focus
4259      Disabler True
4260      cmdAddCon.Enabled = False
4262      RefreshConstraintPanel
End Sub

Private Sub cmdChange_Click()
    Dim frmSolverChange As FSolverChange
    Set frmSolverChange = New FSolverChange
    
    Me.Hide  '  Hide the model form so the refedit on the options form works, and to keep the focus clear
    frmSolverChange.Show
    Unload frmSolverChange
    
    FormatCurrentSolver
    PreserveModel = True
    Me.Show
End Sub

Sub FormatCurrentSolver()
    Dim Solver As String
    Solver = GetChosenSolver(sheet)
    lblSolver.Caption = "Current Solver Engine: " & UCase(Left(Solver, 1)) & Mid(Solver, 2)
End Sub

Private Sub cmdOptions_Click()
          ' Save the current "Assume Non Negative" option so this is shown in the options dialog.
          ' The saved value gets updated on OK, which we then reflect in our Model dialog
4267      SetNonNegativity chkNonNeg.value, sheet
          Dim frmOptions As FOptions
          Set frmOptions = New FOptions
          
          Me.Hide  ' Hide the model form so the refedit on the options form works, and to keep the focus clear
4268      frmOptions.Show
          
          Unload frmOptions
4270      chkNonNeg.value = GetNonNegativity(sheet)

          ' Restore the original model form
          PreserveModel = True
          Me.Show
End Sub

Private Sub cmdReset_Click()
          'Check the user wants to reset the model
4274      If MsgBox("This will reset the objective function, decision variables and constraints." _
                  & vbCrLf & " Are you sure you want to do this?", vbYesNo, "OpenSolver") = vbNo Then
4275          Exit Sub
4276      End If

          'Reset the objective function and the decision variables
4277      SetFormObjective vbNullString
4278      SetFormDecisionVars vbNullString

          ' Remove the constraints
4280      NumConstraints = 0
4283      PopulateConstraintListBox
End Sub

Private Sub optMax_Click()
4285      txtObjTarget.Enabled = optTarget.value
End Sub

Private Sub optMin_Click()
4287      txtObjTarget.Enabled = optTarget.value
End Sub

Private Sub optTarget_Click()
4291      txtObjTarget.Enabled = optTarget.value
End Sub

Private Sub AlterConstraints(DoDisable As Boolean)
          Disabler DoDisable
          cmdAddCon.Enabled = Not DoDisable Or lstConstraints.ListIndex = 0
End Sub

Private Sub refConLHS_Change()
    ' Only fire the change event if the RefEdit is being edited directly
    If ActiveControl.Name <> "refConLHS" Then Exit Sub
    
    ' refConLHS has focus, we don't need to worry about setting the focus
    AlterConstraints Not HasConstraintChanged()
End Sub

Private Sub refConRHS_Change()
    ' Only fire the change event if the RefEdit is being edited directly
    If ActiveControl.Name <> "refConRHS" Then Exit Sub
    
    ' refConRHS has focus, we don't need to worry about setting the focus
    AlterConstraints Not HasConstraintChanged()
End Sub

Private Sub cboConRel_Change()
    ' Caller should ensure focus is set correctly
    refConRHS.Enabled = RelationHasRHS(GetFormRelation())
    AlterConstraints Not HasConstraintChanged()
End Sub

Private Function HasConstraintChanged() As Boolean
    Dim LHSChanged As Boolean, RelChanged As Boolean, RHSChanged As Boolean

    Dim Index As Long
    Index = lstConstraints.ListIndex
    
    ' If there is a selected constraint, we check against the original values
    ' otherwise we compare to empty strings
    Dim OrigLHS As String, OrigRHSLocal As String
    If Index >= 1 Then
        With Constraints(Index)
            OrigLHS = .LHSRefersTo
            OrigRHSLocal = .RHSRefersToLocal
            
            RelChanged = (GetFormRelation() <> .Relation)
        End With
    End If
    
    LHSChanged = GetFormLHSRefersTo() <> OrigLHS
    ' Only check the RHS if the relation uses RHS
    RHSChanged = RelationHasRHS(GetFormRelation()) And _
                 GetFormRHSRefersTo() <> OrigRHSLocal
    
    HasConstraintChanged = LHSChanged Or RelChanged Or RHSChanged
End Function

Private Sub UserForm_Activate()
          CenterForm
          On Error GoTo ErrorHandler

          cmdCancel.SetFocus

          ' Check if we have indicated to keep the model from the last time form was shown
4350      If PreserveModel Then
              PreserveModel = False
          Else
              UpdateStatusBar "Loading model...", True
              Application.Cursor = xlWait
              Application.ScreenUpdating = False

              ' Check we can even start
              GetActiveSheetIfMissing sheet
             
4349          SetAnyMissingDefaultSolverOptions sheet
          
4351          If SheetHasOpenSolverHighlighting(sheet) Then
                  RestoreHighlighting = True
                  HideSolverModel sheet
              End If
          
              ' Make sure sheet is up to date
4352          Application.Calculate
              ' Remove the 'marching ants' showing if a range is copied.
              ' Otherwise, the ants stay visible, and visually conflict with
              ' our cell selection. The ants are also left behind on the
              ' screen. This works around an apparent bug (?) in Excel 2007.
4353          Application.CutCopyMode = False
          
              ' Clear the form
4354          SetFormObjectiveSense UnknownObjectiveSense
4356          SetFormObjective vbNullString
4357          SetFormDecisionVars vbNullString
4358          SetFormLHSRefersTo vbNullString
              SetFormRHSRefersTo vbNullString
4360          lstConstraints.Clear
4361          cboConRel.Clear
4362          cboConRel.AddItem "="
4363          cboConRel.AddItem "<="
4364          cboConRel.AddItem ">="
4365          cboConRel.AddItem "int"
4366          cboConRel.AddItem "bin"
4367          cboConRel.AddItem "alldiff"
4368          SetFormRelation RelationEQ    ' We set an initial value just in case there is no model, and the user goes straight to AddNewConstraint

              'Find current solver
              FormatCurrentSolver

              ' Load the model on the sheet into memory
4379          LoadModelFromSheet
          End If
          
          PopulateConstraintListBox
          RefreshConstraintPanel
              
          chkNameRange.value = ShowNamedRangesState
          chkShowModel.value = ShowModelAfterSavingState
        
ExitSub:
          Application.StatusBar = False
          Application.Cursor = xlDefault
          Application.ScreenUpdating = True
          Exit Sub

ErrorHandler:
          If RestoreHighlighting Then ShowSolverModel sheet, HandleError:=True
          Me.Hide
          ReportError "FModel", "UserForm_Activate", True
          GoTo ExitSub
End Sub


Private Sub cmdCancel_Click()
          If RestoreHighlighting Then ShowSolverModel sheet, HandleError:=True
4386      DoEvents
4387      On Error Resume Next ' Just to be safe on our select
4388      Application.CutCopyMode = False
4389      ActiveCell.Select ' Just select one cell, choosing a cell that should be visible to avoid scrolling
4390      Application.ScreenUpdating = True
          Me.Hide
End Sub

' Make the [x] hide the form rather than unload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' If CloseMode = vbFormControlMenu then we know the user
    ' clicked the [x] close button or Alt+F4 to close the form.
    If CloseMode = vbFormControlMenu Then
        cmdCancel_Click
        Cancel = True
    End If
End Sub

Private Sub cmdRunAutoModel_Click()
    ' Refedits on Mac don't work if more than one form is shown, so we need to hide it
    ' We also hide on windows to make sure that the forms don't hide each other
    Me.Hide

    Dim AutoModel As CAutoModel
    Set AutoModel = New CAutoModel
    If AutoModel.BuildModel(sheet, MinimiseUserInteraction:=False, SaveAfterBuilding:=False) Then
        ' Copy results into FModel
        SetFormObjective AutoModel.ObjectiveFunctionCellRefersTo
        SetFormDecisionVars RangeToRefersTo(AutoModel.DecVarsRange)
        SetFormObjectiveSense AutoModel.ObjSense
        
        NumConstraints = AutoModel.Constraints.Count
        ReDim Constraints(1 To NumConstraints) As CConstraint
        Dim c As CAutoModelConstraint, i As Long
        i = 1
        For Each c In AutoModel.Constraints
            Set Constraints(i) = New CConstraint
            With c
                Constraints(i).Update RangeToRefersTo(.LHS), .RelationType, RangeToRefersTo(.RHS), sheet
            End With
            i = i + 1
        Next c
    End If
    
    lstConstraints.ListIndex = -1  ' Clear the selection
    PreserveModel = True
    Me.Show
End Sub

Private Sub cmdBuild_Click()
          On Error GoTo ErrorHandler

4408      DoEvents
4409      On Error Resume Next
4410      Application.CutCopyMode = False
4411      ActiveCell.Select ' Just select one cell, choosing a cell that should be visible to avoid scrolling
4412      Application.ScreenUpdating = True
4413      On Error GoTo ErrorHandler

          Dim oldCalculationMode As Long
          oldCalculationMode = Application.Calculation
          Application.Calculation = xlCalculationManual
          
          
          '''''
          ' Check if the model is valid before saving anything!
          '''''
          
          ' Check objective
          Dim ObjectiveFunctionCellRefersTo As String
          On Error GoTo BadObjRef
          ObjectiveFunctionCellRefersTo = GetFormObjective()
          ValidateObjectiveFunctionCellRefersTo ObjectiveFunctionCellRefersTo
          On Error GoTo ErrorHandler

          ' Check objective sense
          Dim ObjectiveSense As ObjectiveSenseType
4421      ObjectiveSense = GetFormObjectiveSense()
4429      If ObjectiveSense = UnknownObjectiveSense Then
4430          RaiseUserError "Please select an objective sense (minimise, maximise or target)."
4432      End If

          ' Check objective target
4423      If ObjectiveSense = TargetObjective Then
              Dim ObjectiveTarget As Double
4425          On Error GoTo BadObjectiveTarget
4426          ObjectiveTarget = GetFormObjectiveTargetValue()
4427          On Error GoTo ErrorHandler
4428      End If

          ' Check decision variables
          Dim DecisionVariablesRefersTo As String
          On Error GoTo BadDecRef
          DecisionVariablesRefersTo = GetFormDecisionVars()
          ValidateDecisionVariablesRefersTo DecisionVariablesRefersTo
          On Error GoTo ErrorHandler

          ' Check duals range
4441      If chkGetDuals.value Then
              Dim DualsRefersTo As String
              On Error GoTo BadDualsRef
              DualsRefersTo = GetFormDuals()
              ValidateDualsRefersTo DualsRefersTo
              On Error GoTo ErrorHandler
4445      End If

          ' Check all constraints are valid
          Dim c As Long
          For c = 1 To NumConstraints
              With Constraints(c)
                  On Error Resume Next
                  ValidateConstraintRefersTo .LHSRefersTo, .Relation, .RHSRefersTo, sheet
                  If Err.Number <> 0 Then
                      MsgBox "There is an error in constraint " & c & ":" & vbNewLine & Err.Description
                      GoTo ExitSub
                  End If
              End With
          Next c
          On Error GoTo ErrorHandler
          
          ' Check int/bin constraints not set on non-decision variables before we start saving things
          ' TODO change to error ref #170
3924      For c = 1 To NumConstraints
3925          If Not RelationHasRHS(Constraints(c).Relation) Then
3927              If Not SetDifference(Constraints(c).LHSRange, GetRefersToRange(DecisionVariablesRefersTo)) Is Nothing Then
                      If MsgBox("This model has specified that a non-decision cell must take an integer/binary value. " & _
                                "This is a valid model, but not one that OpenSolver can solve. " & _
                                "Do you wish to continue with saving this model?", _
                                vbQuestion + vbYesNo, "OpenSolver - Warning") = vbYes Then
                          Exit For
                      Else
                          GoTo ExitSub
                      End If
                  End If
3934          End If
3935      Next c

          ''''''''
          ' Build is now confirmed, save everything to sheet
          ''''''''
3906      SetDecisionVariablesRefersTo DecisionVariablesRefersTo, sheet
3911      SetObjectiveFunctionCellRefersTo ObjectiveFunctionCellRefersTo, sheet
3916      SetObjectiveSense ObjectiveSense, sheet
3918      If ObjectiveSense = TargetObjective Then SetObjectiveTargetValue ObjectiveTarget, sheet
          
          ' Constraints
3936      SetNumConstraints NumConstraints, sheet
3937      For c = 1 To NumConstraints
              With Constraints(c)
                  UpdateConstraintRefersTo c, .LHSRefersTo, .Relation, .RHSRefersTo, sheet
              End With
3968      Next c

          ' Options
          SetNonNegativity chkNonNeg.value, sheet
          ' TODO: Only update these if doing sensitivity- needs a change to save chkGetDuals
          'If chkGetDuals.value Then
              SetDualsRefersTo DualsRefersTo, sheet
4252          SetDualsOnSheet chkGetDuals2.value, sheet
              SetUpdateSensitivity optUpdate.value, sheet
          'End If

          ' Display on screen
4449      If chkShowModel.value = True Then ShowSolverModel sheet, HandleError:=True
              
4450      On Error GoTo CalculateFailed
4451      Application.Calculate
4452      On Error GoTo ErrorHandler
          
          Me.Hide
4453      GoTo ExitSub

CalculateFailed:
          ' Application.Calculate failed. Ignore error and try again
4454      On Error GoTo ErrorHandler
4455      Application.Calculate
4456      Resume Next

BadObjRef:
          ' Couldn't turn the objective cell address into a range
4457      MsgBox "Error: the cell address for the objective is invalid. " + _
                 "This must be a single cell. " & _
                 "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
4458      refObj.SetFocus ' Set the focus back to the RefEdit
          GoTo ExitSub
BadDecRef:
          ' Try with raw text if it might work that way!
          ' Converting the text to RefersTo can increase the length (with sheet names etc)
          ' and could cause the range text to become too long for Excel
          If DecisionVariablesRefersTo <> Me.refDecision.Text Then
              DecisionVariablesRefersTo = Me.refDecision.Text
              Resume ' Try again on the validation
          End If
          
          ' Couldn't turn the decision variable address into a range
4461      MsgBox "Error: the cell range specified for the Variable Cells is invalid. " + _
                 "This must be a valid Excel range that does not exceed Excel's internal character count limits. " + _
                 "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
4462      refDecision.SetFocus ' Set the focus back to the RefEdit
          GoTo ExitSub
BadObjectiveTarget:
          ' Couldn't turn the objective target into a value
4465      MsgBox "Error: the target value for the objective cell is invalid. " + _
                 "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
4466      txtObjTarget.SetFocus ' Set the focus back to the target text box
          GoTo ExitSub
BadDualsRef:
          ' Couldn't turn the dual cell into a range
4469      MsgBox "Error: the cell for storing the shadow prices is invalid. " + _
                 "This must be a single cell. " & _
                 "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
4470      refDuals.SetFocus ' Set the focus back to the refedit
          GoTo ExitSub

ExitSub:
          Application.Calculation = oldCalculationMode
          Exit Sub

ErrorHandler:
          ReportError "FModel", "cmdBuild_Click", True
          GoTo ExitSub
End Sub

Private Sub cmdAddCon_Click()
4496      On Error GoTo ErrorHandler

          Dim LHSRefersTo As String, RHSRefersTo As String, rel As RelationConsts
          LHSRefersTo = GetFormLHSRefersTo()
4513      RHSRefersTo = ConvertFromCurrentLocale(GetFormRHSRefersTo())
          rel = GetFormRelation()

          ValidateConstraintRefersTo LHSRefersTo, rel, RHSRefersTo, sheet
          
          ' Valid, now update
          
          lstConstraints.SetFocus  ' Set focus back to constraint list to avoid disabling problems
4558      AlterConstraints True

          Dim conIndex As Long
4561      If cmdAddCon.Caption <> "Add constraint" Then
              ' Update constraint
              ' Note that we use the saved index value rather than lstConstraints.ListIndex in case
              ' this update was prompted by the user changing the index on an unsaved constraint.
              ' In this case, lstConstraints.ListIndex does NOT reflect the constraint being altered.
4562          conIndex = CurrentListIndex
4585      Else
              ' Add constraint
              NumConstraints = NumConstraints + 1
              ReDim Preserve Constraints(1 To NumConstraints) As CConstraint
              conIndex = NumConstraints
          End If
                   
          Set Constraints(conIndex) = New CConstraint
          Constraints(conIndex).Update LHSRefersTo, rel, RHSRefersTo, sheet
          
4608      PopulateConstraintListBox conIndex
          RefreshConstraintPanel
4609      Exit Sub

ErrorHandler:
4620      Application.DisplayAlerts = True
4622      MsgBox Err.Description
4624      Exit Sub

End Sub

Private Sub cmdDelSelCon_Click()
4626      If lstConstraints.ListIndex = -1 Then
4627          MsgBox "Please select a constraint!", vbOKOnly, "AutoModel"
4628          Exit Sub
4629      End If

          ' Set focus back to constraint box to avoid focus jumping around
          lstConstraints.SetFocus

          Dim i As Long
          For i = 1 To NumConstraints - 1
              If i >= lstConstraints.ListIndex Then
                  Set Constraints(i) = Constraints(i + 1)
              End If
          Next i
          NumConstraints = NumConstraints - 1
          If NumConstraints > 0 Then ReDim Preserve Constraints(1 To NumConstraints) As CConstraint
          
          ' Update form
4631      PopulateConstraintListBox
          lstConstraints_Change
End Sub

Private Sub lstConstraints_Change()
          If DisableConstraintListChange Then Exit Sub

4633      If cmdCancelCon.Enabled Then
              Dim SaveChanges As Boolean
4634          If cmdAddCon.Caption = "Update constraint" Then
4635              SaveChanges = (MsgBox("You have made changes to the current constraint." & vbNewLine & _
                                        "Do you want to save these changes?", vbYesNo) = vbYes)
4640          Else
4641              SaveChanges = (MsgBox("You have entered a constraint." & vbNewLine & _
                                        "Do you want to save this as a new constraint?", vbYesNo) = vbYes)
4646          End If
              If SaveChanges Then
4643              cmdAddCon_Click
              End If
4647          AlterConstraints True
4651      End If
          
          RefreshConstraintPanel
          
End Sub

Private Sub RefreshConstraintPanel()
          ' Save the index of the constraint we will display
          CurrentListIndex = lstConstraints.ListIndex

4653      If lstConstraints.ListIndex = -1 Then
4654          Exit Sub
4655      End If
4656      If lstConstraints.ListIndex = 0 Then
              'Add constraint
4659          SetFormLHSRefersTo vbNullString
              SetFormRHSRefersTo vbNullString
4663          cboConRel_Change ' Set the RHS to be active based on the last constraint
4664          cmdAddCon.Enabled = True
4665          cmdAddCon.Caption = "Add constraint"
4666          cmdDelSelCon.Enabled = False
4668          Application.CutCopyMode = False
4669          ActiveCell.Select ' Just select one cell, choosing a cell that should be visible to avoid scrolling
4670          Application.ScreenUpdating = True
4672      Else
              ' Update constraint
4677          cmdAddCon.Enabled = False
4678          cmdAddCon.Caption = "Update constraint"
4679          cmdDelSelCon.Enabled = True

              ' Set text and marching ants; the select/copy operations can throw errors if on different sheets
4681          On Error Resume Next
4682          ActiveCell.Select   ' We may fail in the next steps, so we cancel any old highlighting
4683          Application.CutCopyMode = False
              Dim copyRange As Range
              
              With Constraints(lstConstraints.ListIndex)
                  SetFormLHSRefersTo .LHSRefersTo
4687              SetFormRelation .Relation
4688              SetFormRHSRefersTo .RHSRefersTo

                  Set copyRange = ProperUnion(.LHSRange, .RHSRange)
              End With
4686
4700          cboConRel_Change

              ' Will fail if LHS and RHS are different shape
              ' Silently fail, nothing that can be done about it
4701          copyRange.Select
4702          copyRange.Copy
4703          Application.ScreenUpdating = True
4705      End If
End Sub

Private Sub lstConstraints_Exit(ByVal Cancel As MSForms.ReturnBoolean)
          ' When the focus leaves this list, we want to remove any highlighting shown by selected cells
4706      ActiveCell.Select
End Sub

Private Sub UserForm_Initialize()
    AutoLayout
    CenterForm
End Sub

Private Sub AutoLayout()
    AutoFormat Me.Controls
    
    Me.Width = FormWidthModel
    
    With cmdRunAutoModel
        .Caption = "AutoModel"
        .Width = FormButtonWidth * 1.3
        .Top = FormMargin
        .Left = LeftOfForm(Me.Width, .Width)
    End With
    
    With lblDescHeader
        .Left = FormMargin
        .Top = cmdRunAutoModel.Top + FormSpacing
        .Caption = "What is AutoModel?"
    End With
    
    With lblDesc
        .Left = lblDescHeader.Left
        .Top = Below(cmdRunAutoModel)
        .Caption = "AutoModel is a feature of OpenSolver that tries to automatically determine " & _
                   "the problem you are trying to optimise by observing the structure of the " & _
                   "spreadsheet. It will turn its best guess into a Solver model, which you can " & _
                   "then edit in this window."
        AutoHeight lblDesc, Me.Width - 2 * FormMargin
    End With
    
    With lblDiv1
        .Left = lblDescHeader.Left
        .Top = Below(lblDesc)
        .Width = lblDesc.Width
        .Height = FormDivHeight
        .BackColor = FormDivBackColor
    End With
    
    With lblStep1
        .Caption = "Objective Cell:"
        .Left = lblDescHeader.Left
        .Top = Below(lblDiv1)
        AutoHeight lblStep1, Me.Width, True
    End With
    
    With txtObjTarget
        .Width = cmdRunAutoModel.Width
        .Top = lblStep1.Top
        .Left = LeftOfForm(Me.Width, .Width)
    End With
    
    With optTarget
        .Caption = "target value:"
        AutoHeight optTarget, Me.Width, True
        .Top = lblStep1.Top
        .Left = LeftOf(txtObjTarget, .Width)
    End With
    
    With optMin
        .Caption = "minimise"
        AutoHeight optMin, Me.Width, True
        .Top = lblStep1.Top
        .Left = LeftOf(optTarget, .Width)
    End With
    
    With optMax
        .Caption = "maximise"
        AutoHeight optMax, Me.Width, True
        .Top = lblStep1.Top
        .Left = LeftOf(optMin, .Width)
    End With
    
    With refObj
        .Left = RightOf(lblStep1)
        .Top = lblStep1.Top
        .Width = LeftOf(optMax, .Left)
    End With
    
    With lblDiv2
        .Left = lblDescHeader.Left
        .Top = Below(optMax)
        .Width = lblDesc.Width
        .Height = FormDivHeight
        .BackColor = FormDivBackColor
    End With
    
    With lblStep2
        .Caption = "Variable Cells:"
        .Left = lblDescHeader.Left
        .Top = Below(lblDiv2)
        AutoHeight lblStep2, Me.Width, True
    End With
    
    With refDecision
        .Height = 2 * refObj.Height
        .Left = refObj.Left
        .Top = lblStep2.Top
        .Width = LeftOfForm(Me.Width, .Left)
    End With
    
    With lblDiv3
        .Left = lblDescHeader.Left
        .Top = Below(refDecision)
        .Width = lblDesc.Width
        .Height = FormDivHeight
        .BackColor = FormDivBackColor
    End With
    
    With lblStep3
        .Caption = "Constraints:"
        .Left = lblDescHeader.Left
        .Top = Below(lblDiv3)
        .Width = lblDesc.Width
    End With
    
    lblConstraintGroup.Top = Below(lblStep3, False)
    
    With cboConRel
        .Width = cmdRunAutoModel.Width / 2
        .Height = refObj.Height
        .Top = lblConstraintGroup.Top + FormSpacing
        .Left = LeftOfForm(Me.Width, .Width) - FormSpacing
    End With
    
    With refConLHS
        .Width = cboConRel.Width * 3
        .Height = refObj.Height
        .Top = cboConRel.Top
        .Left = LeftOf(cboConRel, .Width)
    End With
    
    With refConRHS
        .Width = refConLHS.Width
        .Left = refConLHS.Left
        .Height = refObj.Height
        .Top = Below(refConLHS)
    End With
    
    With cmdAddCon
        .Caption = "Add constraint"
        .Left = refConLHS.Left
        .Top = Below(refConRHS)
        .Width = cboConRel.Width * 2
    End With
    
    With cmdCancelCon
        .Caption = "Cancel"
        .Left = RightOf(cmdAddCon)
        .Top = cmdAddCon.Top
        .Width = cmdAddCon.Width
    End With
    
    With lblConstraintGroup
        .Left = refConLHS.Left - FormSpacing
        .Width = FormSpacing * 3 + cmdAddCon.Width + cmdCancelCon.Width
        .Height = FormSpacing * 4 + refConLHS.Height + refConRHS.Height + cmdAddCon.Height
    End With
    
    With cmdDelSelCon
        .Caption = "Delete selected constraint"
        .Left = lblConstraintGroup.Left
        .Top = Below(lblConstraintGroup)
        .Width = lblConstraintGroup.Width
    End With
    
    With chkNonNeg
        .Caption = "Make unconstrained variable cells non-negative"
        .Left = lblConstraintGroup.Left
        .Top = Below(cmdDelSelCon)
        .Width = lblConstraintGroup.Width
    End With
    
    With chkNameRange
        .Left = lblConstraintGroup.Left
        .Width = lstConstraints.Width
        .Caption = "Show named ranges in constraint list"
        .Top = Below(chkNonNeg, False)
    End With
    
    With lstConstraints
        .Left = lblDescHeader.Left
        .Top = lblConstraintGroup.Top
        .Height = MinHeight
        .Width = LeftOf(lblConstraintGroup, .Left)
    End With
    
    With lblDiv4
        .Left = lblDescHeader.Left
        .Width = lblDesc.Width
        .Height = FormDivHeight
        .BackColor = FormDivBackColor
    End With
    
    With lblStep4
        .Caption = "Sensitivity Analysis"
        .Left = lblDescHeader.Left
        AutoHeight lblStep4, Me.Width, True
    End With
    
    With chkGetDuals
        .Caption = "List sensitivity analysis on the same sheet with top left cell:"
        .Left = RightOf(lblStep4)
        AutoHeight chkGetDuals, Me.Width, True
    End With
    
    With refDuals
        .Left = RightOf(chkGetDuals)
        .Width = LeftOfForm(Me.Width, .Left)
        .Height = refObj.Height
    End With
    
    With chkGetDuals2
        .Caption = "Output sensitivity analysis:"
        .Left = chkGetDuals.Left
        AutoHeight chkGetDuals2, Me.Width, True
    End With
    
    With optUpdate
        .Caption = "updating any previous output sheet"
        .Left = RightOf(chkGetDuals2)
        AutoHeight optUpdate, Me.Width, True
    End With
    
    With optNew
        .Caption = "on a new sheet"
        .Left = RightOf(optUpdate)
        AutoHeight optNew, Me.Width, True
    End With
    
    With lblDiv5
        .Left = lblDescHeader.Left
        .Width = lblDesc.Width
        .Height = FormDivHeight
        .BackColor = FormDivBackColor
    End With
    
    With lblStep5
        .Left = lblDescHeader.Left
        .Caption = "Solver Engine:"
    End With
    
    With cmdChange
        .Width = cmdRunAutoModel.Width
        .Left = LeftOfForm(Me.Width, .Width)
        .Caption = "Solver Engine..."
    End With
    
    With lblSolver
        .Width = LeftOf(cmdChange, .Left)
    End With
    
    With lblDiv6
        .Left = lblDescHeader.Left
        .Width = lblDesc.Width
        .Height = FormDivHeight
        .BackColor = FormDivBackColor
    End With
        
    With chkShowModel
        .Left = lblDescHeader.Left
        AutoHeight chkShowModel, Me.Width, True
    End With
    
    With cmdCancel
        .Width = cmdRunAutoModel.Width
        .Caption = "Cancel"
        .Left = LeftOfForm(Me.Width, .Width)
        .Cancel = True
    End With
    
    With cmdBuild
        .Width = cmdRunAutoModel.Width
        .Caption = "Save Model"
        .Left = LeftOf(cmdCancel, .Width)
    End With
    
    With cmdOptions
        .Width = cmdRunAutoModel.Width
        .Caption = "Options..."
        .Left = LeftOf(cmdBuild, .Width)
    End With
    
    With cmdReset
        .Width = cmdRunAutoModel.Width
        .Caption = "Clear Model"
        .Left = LeftOf(cmdOptions, .Width)
    End With
    
    ' Add resizer
    With lblResizer
        #If Mac Then
            ' Mac labels don't fire MouseMove events correctly
            .Visible = False
        #End If
        .Caption = "o"
        With .Font
            .Name = "Marlett"
            .Charset = 2
            .Size = 10
        End With
        .AutoSize = True
        .Left = Me.Width - .Width
        .MousePointer = fmMousePointerSizeNWSE
        .BackStyle = fmBackStyleTransparent
    End With
    
    ' Set the vertical positions of the lower half of the form
    UpdateLayout
    
    Me.Width = Me.Width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - Model"
End Sub

Private Sub UpdateLayout(Optional ChangeY As Single = 0)
' Do the layout of the lower half of the form, changing the height of the list box by ChangeY
    Dim NewHeight As Double
    NewHeight = Max(lstConstraints.Height + ChangeY, MinHeight)
    
    lstConstraints.Height = NewHeight
        
    ' Cascade the updated height
    lblDiv4.Top = Below(lstConstraints)
    lblStep4.Top = Below(lblDiv4)
    chkGetDuals.Top = lblStep4.Top
    refDuals.Top = lblStep4.Top
    chkGetDuals2.Top = Below(chkGetDuals, False)
    optUpdate.Top = chkGetDuals2.Top
    optNew.Top = chkGetDuals2.Top
    lblDiv5.Top = Below(optNew, False)
    lblStep5.Top = Below(lblDiv5)
    cmdChange.Top = lblStep5.Top
    lblSolver.Top = lblStep5.Top + FormButtonHeight - FormTextHeight
    lblDiv6.Top = Below(cmdChange)
    chkShowModel.Top = Below(lblDiv6)
    cmdCancel.Top = chkShowModel.Top
    cmdBuild.Top = chkShowModel.Top
    cmdOptions.Top = chkShowModel.Top
    cmdReset.Top = chkShowModel.Top
    Me.Height = FormHeight(cmdCancel)
    lblResizer.Top = Me.InsideHeight - lblResizer.Height
End Sub

Private Sub lblResizer_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        ResizeStartY = Y
    End If
End Sub

Private Sub lblResizer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        #If Mac Then
            ' Mac reports delta already
            UpdateLayout Y
        #Else
            UpdateLayout (Y - ResizeStartY)
        #End If
    End If
End Sub

Private Sub CenterForm()
    Me.Top = CenterFormTop(Me.Height)
    Me.Left = CenterFormLeft(Me.Width)
End Sub

Private Function GetFormObjectiveSense() As ObjectiveSenseType
    If optMax.value = True Then GetFormObjectiveSense = MaximiseObjective
    If optMin.value = True Then GetFormObjectiveSense = MinimiseObjective
    If optTarget.value = True Then GetFormObjectiveSense = TargetObjective
End Function
Private Sub SetFormObjectiveSense(ObjectiveSense As ObjectiveSenseType)
    Select Case ObjectiveSense
    Case MaximiseObjective: optMax.value = True
    Case MinimiseObjective: optMin.value = True
    Case TargetObjective:   optTarget.value = True
    Case UnknownObjectiveSense
        optMax.value = False
        optMin.value = False
        optTarget.value = False
    End Select
End Sub

Private Function GetFormObjectiveTargetValue() As Double
    GetFormObjectiveTargetValue = CDbl(txtObjTarget.Text)
End Function
Private Sub SetFormObjectiveTargetValue(ObjectiveTargetValue As Double)
    txtObjTarget.Text = CStr(ObjectiveTargetValue)
End Sub

Private Function GetFormObjective() As String
    GetFormObjective = RefEditToRefersTo(refObj.Text)
End Function
Private Sub SetFormObjective(ObjectiveRefersTo As String)
    refObj.Text = GetDisplayAddress(ObjectiveRefersTo, sheet, False)
End Sub

Private Function GetFormLHSRefersTo() As String
    GetFormLHSRefersTo = RefEditToRefersTo(refConLHS.Text)
End Function
Private Sub SetFormLHSRefersTo(newLHSRefersTo As String)
    refConLHS.Text = GetDisplayAddress(newLHSRefersTo, sheet, False)
End Sub

Private Function GetFormRHSRefersTo() As String
    If RelationHasRHS(GetFormRelation()) Then
        ' We don't convert from current locale here to avoid causing errors in the RefEdit events
        GetFormRHSRefersTo = RefEditToRefersTo(refConRHS.Text)
    Else
        GetFormRHSRefersTo = vbNullString
    End If
End Function
Private Sub SetFormRHSRefersTo(newRHSRefersTo As String)
    refConRHS.Text = ConvertToCurrentLocale(GetDisplayAddress(newRHSRefersTo, sheet, False))
End Sub

Private Function GetFormRelation() As RelationConsts
    GetFormRelation = RelationStringToEnum(cboConRel.Text)
End Function
Private Sub SetFormRelation(newRelation As RelationConsts)
    cboConRel.ListIndex = cboPosition(RelationEnumToString(newRelation))
End Sub

' We allow multiple area ranges for the variables, which requires ConvertLocale as delimiter can vary
Private Function GetFormDecisionVars() As String
    GetFormDecisionVars = ConvertFromCurrentLocale(RefEditToRefersTo(refDecision.Text))
End Function
Private Sub SetFormDecisionVars(newDecisionVars As String)
    refDecision.Text = ConvertToCurrentLocale(GetDisplayAddress(newDecisionVars, sheet, False))
End Sub

Private Function GetFormDuals() As String
    GetFormDuals = RefEditToRefersTo(refDuals.Text)
End Function
Private Sub SetFormDuals(newDuals As String)
    refDuals.Text = GetDisplayAddress(newDuals, sheet, False)
End Sub
