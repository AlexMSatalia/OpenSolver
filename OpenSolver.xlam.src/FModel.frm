VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FModel 
   Caption         =   "OpenSolver - Model"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9840
   OleObjectBlob   =   "FModel.frx":0000
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
Const MinHeight = 140

Private model As CModel

Private ConChangedMode As Boolean
Private DontRepop As Boolean
Private IsLoadingModel As Boolean
Private PreserveModel As Boolean  ' Used to persist model when re-showing form

Private RestoreHighlighting As Boolean
Private sheet As Worksheet

Private IsResizing As Boolean
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

4161      lblDescHeader.Enabled = TrueIfEnable
4162      lblDesc.Enabled = TrueIfEnable
4163      cmdRunAutoModel.Enabled = TrueIfEnable

4164      lblDiv1.Enabled = False

4165      lblStep1.Enabled = TrueIfEnable
4166      refObj.Enabled = TrueIfEnable
4167      optMax.Enabled = TrueIfEnable
4168      optMin.Enabled = TrueIfEnable
4169      optTarget.Enabled = TrueIfEnable
4170      txtObjTarget.Enabled = TrueIfEnable And optTarget.value

4171      lblDiv2.Enabled = False

4172      lblStep2.Enabled = TrueIfEnable
4173      refDecision.Enabled = TrueIfEnable

4174      lblDiv3.Enabled = False

4175      chkNonNeg.Enabled = TrueIfEnable
4176      cmdCancelCon.Enabled = Not TrueIfEnable
4177      cmdDelSelCon.Enabled = TrueIfEnable
          chkNameRange.Enabled = TrueIfEnable

4178      lblDiv4.Enabled = False

4179      lblStep4.Enabled = TrueIfEnable

4180      chkGetDuals.Enabled = TrueIfEnable
4181      chkGetDuals2.Enabled = TrueIfEnable
4182      chkGetDuals2_Click

4184      refDuals.Enabled = TrueIfEnable And chkGetDuals.value And chkGetDuals.Enabled

4185      lblDiv5.Enabled = False
4186      lblDiv6.Enabled = False

4187      chkShowModel.Enabled = TrueIfEnable
4188      cmdOptions.Enabled = TrueIfEnable
4189      cmdBuild.Enabled = TrueIfEnable
4190      cmdCancel.Enabled = TrueIfEnable
          cmdReset.Enabled = TrueIfEnable
          cmdChange.Enabled = TrueIfEnable
          
          Dim Solver As ISolver
          Set Solver = CreateSolver(GetChosenSolver())

4203      If Not SensitivityAnalysisAvailable(Solver) Then
              ' Disable dual options
4204          chkGetDuals2.Enabled = False
4205          chkGetDuals.Enabled = False
4206          chkGetDuals2_Click
4208      End If

          '============================================================================================
          'NOTE: Beware that RefEdits cannot be enabled last in this sub as they seem to grab the focus
          '       and create weird errors
          '============================================================================================
End Sub

Sub UpdateFormFromMemory()
4209      If model.ObjectiveSense = MaximiseObjective Then optMax.value = True
4210      If model.ObjectiveSense = MinimiseObjective Then optMin.value = True
4211      If model.ObjectiveSense = TargetObjective Then optTarget.value = True
4212      txtObjTarget.Text = CStr(model.ObjectiveTarget)   ' Always show the target (which may just be 0)

4213      chkNonNeg.value = model.NonNegativityAssumption
4216      chkGetDuals.value = Not model.Duals Is Nothing

4214      refObj.Text = GetDisplayAddress(model.ObjectiveFunctionCell, False)
4215      refDecision.Text = ConvertToCurrentLocale(GetDisplayAddress(model.DecisionVariables))
4218      refDuals.Text = GetDisplayAddress(model.Duals, False)

4222      model.PopulateConstraintListBox lstConstraints, chkNameRange.value
4223      lstConstraints_Change

          chkGetDuals2.value = GetDualsOnSheet()
4237      chkGetDuals2_Click

          optUpdate.value = GetUpdateSensitivity()
4242      optNew.value = Not optUpdate.value
End Sub

Private Sub chkGetDuals_Click()
4250      refDuals.Enabled = chkGetDuals.value
End Sub

Private Sub chkGetDuals2_Click()
4253      optUpdate.Enabled = chkGetDuals2.value And chkGetDuals2.Enabled
4254      optNew.Enabled = chkGetDuals2.value And chkGetDuals2.Enabled
End Sub

Private Sub chkNameRange_Click()
4256      model.PopulateConstraintListBox lstConstraints, chkNameRange.value
4257      lstConstraints_Change
End Sub

Private Sub cmdCancelCon_Click()
4259      Disabler True
4260      cmdAddCon.Enabled = False
4261      ConChangedMode = False
4262      lstConstraints_Change
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
    Solver = GetChosenSolver()
    lblSolver.Caption = "Current Solver Engine: " & UCase(Left(Solver, 1)) & Mid(Solver, 2)
End Sub

Private Sub cmdOptions_Click()
          ' Save the current "Assume Non Negative" option so this is shown in the options dialog.
          ' The saved value gets updated on OK, which we then reflect in our Model dialog
4267      SetNonNegativity chkNonNeg.value
          Dim frmOptions As FOptions
          Set frmOptions = New FOptions
          
          Me.Hide  ' Hide the model form so the refedit on the options form works, and to keep the focus clear
4268      frmOptions.Show
          
          Unload frmOptions
4270      chkNonNeg.value = GetNonNegativity

          
          ' Restore the original model form
          model.NonNegativityAssumption = chkNonNeg.value
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
4277      refObj.Text = ""
4278      refDecision.Text = ""

          ' Remove the constraints
4280      Set model.Constraints = New Collection
4283      model.PopulateConstraintListBox lstConstraints, chkNameRange.value

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
          cmdAddCon.Enabled = Not DoDisable
          ConChangedMode = Not DoDisable
End Sub

Private Sub refConLHS_Change()
    If IsLoadingModel Then Exit Sub
    AlterConstraints Not HasConstraintChanged()
End Sub

Private Sub refConRHS_Change()
    If IsLoadingModel Then Exit Sub
    AlterConstraints Not HasConstraintChanged()
End Sub

Private Sub cboConRel_Change()
    refConRHS.Enabled = RelationHasRHS(RelationStringToEnum(cboConRel.Text))
    AlterConstraints Not HasConstraintChanged()

End Sub

Private Function HasConstraintChanged() As Boolean
    Dim LHSChanged As Boolean, RelChanged As Boolean, RHSChanged As Boolean

    Dim OrigLHS As String, OrigRHS As String
    ' If there is a selected constraint, we check against the original values
    ' otherwise we compare to empty strings
    If lstConstraints.ListIndex >= 1 And Not model.Constraints Is Nothing Then
        Dim con As CConstraint
        Set con = model.Constraints(lstConstraints.ListIndex)
        
        ' Check relation
        If cboConRel.Text <> RelationEnumToString(con.RelationType) Then
            RelChanged = True
        End If
        
        ' Set original values
        OrigLHS = con.LHS.Address
        If con.RHS Is Nothing Then
            OrigRHS = con.RHSstring
        Else
            OrigRHS = con.RHS.Address
        End If
    End If
    LHSChanged = RemoveActiveSheetNameFromString(refConLHS.Text) <> OrigLHS
    ' Only check the RHS if the relation uses RHS
    RHSChanged = RelationHasRHS(RelationStringToEnum(cboConRel.Text)) And RemoveActiveSheetNameFromString(refConRHS.Text) <> OrigRHS
    
    HasConstraintChanged = LHSChanged Or RelChanged Or RHSChanged
End Function

Private Sub UserForm_Activate()
          CenterForm
          On Error GoTo ErrorHandler

          ' Check we can even start
          GetActiveSheetIfMissing sheet

          UpdateStatusBar "Loading model...", True
          Application.Cursor = xlWait
          Application.ScreenUpdating = False

          cmdCancel.SetFocus

4349      SetAnyMissingDefaultSolverOptions sheet

          ' Check if we have indicated to keep the model from the last time form was shown
4350      If PreserveModel Then
              PreserveModel = False
          Else
              Set model = New CModel
              model.LoadFromSheet
          End If

          IsLoadingModel = True

4351      If SheetHasOpenSolverHighlighting(sheet) Then
              RestoreHighlighting = True
              HideSolverModel sheet
          End If
          
          ' Make sure sheet is up to date
4352      Application.Calculate
          ' Remove the 'marching ants' showing if a range is copied.
          ' Otherwise, the ants stay visible, and visually conflict with
          ' our cell selection. The ants are also left behind on the
          ' screen. This works around an apparent bug (?) in Excel 2007.
4353      Application.CutCopyMode = False
          
          ' Clear the form
4354      optMax.value = False
4355      optMin.value = False
4356      refObj.Text = ""
4357      refDecision.Text = ""
4358      refConLHS.Text = ""
4359      refConRHS.Text = ""
4360      lstConstraints.Clear
4361      cboConRel.Clear
4362      cboConRel.AddItem "="
4363      cboConRel.AddItem "<="
4364      cboConRel.AddItem ">="
4365      cboConRel.AddItem "int"
4366      cboConRel.AddItem "bin"
4367      cboConRel.AddItem "alldiff"
4368      cboConRel.ListIndex = cboPosition("=")    ' We set an initial value just in case there is no model, and the user goes straight to AddNewConstraint

          'Find current solver
          FormatCurrentSolver

          ' Load the model on the sheet into memory
4374      AlterConstraints True
          IsLoadingModel = False
4378      DoEvents
4379      UpdateFormFromMemory
        
ExitSub:
          Application.StatusBar = False
          Application.Cursor = xlDefault
          Application.ScreenUpdating = True
          
          ' Take focus away from refEdits
4381      DoEvents
4382      Repaint
4383      DoEvents
          Exit Sub

ErrorHandler:
          If RestoreHighlighting Then ShowSolverModel sheet
          Me.Hide
          ReportError "FModel", "UserForm_Activate", True
          GoTo ExitSub
End Sub


Private Sub cmdCancel_Click()
          If RestoreHighlighting Then ShowSolverModel sheet
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

    Dim NewModel As CModel
    Set NewModel = New CModel
    If RunAutoModel(sheet, False, NewModel) Then Set model = NewModel
    
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

          ' Pull possibly update objective info into model
4414      On Error GoTo BadObjRef
4415      If Trim(refObj.Text) = "" Then
4416          Set model.ObjectiveFunctionCell = Nothing
4417      Else
4418          Set model.ObjectiveFunctionCell = sheet.Range(refObj.Text)
4419      End If
4420      On Error GoTo ErrorHandler

          ' Get the objective sense
4421      If optMax.value = True Then model.ObjectiveSense = MaximiseObjective
4422      If optMin.value = True Then model.ObjectiveSense = MinimiseObjective
4423      If optTarget.value = True Then
4424          model.ObjectiveSense = TargetObjective
4425          On Error GoTo BadObjectiveTarget
4426          model.ObjectiveTarget = CDbl(txtObjTarget.Text)
4427          On Error GoTo ErrorHandler
4428      End If
4429      If model.ObjectiveSense = UnknownObjectiveSense Then
4430          Err.Raise OpenSolver_ModelError, Description:="Please select an objective sense (minimise, maximise or target)."
4432      End If

          ' We allow multiple area ranges here, which requires ConvertFromCurrentLocale as delimiter can vary
4433      On Error GoTo BadDecRef
4434      If Trim(refDecision.Text) = "" Then
4435          Set model.DecisionVariables = Nothing
4436      Else
4437          Set model.DecisionVariables = sheet.Range(ConvertFromCurrentLocale(refDecision.Text))
4438      End If
4439      On Error GoTo ErrorHandler

4440      On Error GoTo BadDualsRef
4441      If chkGetDuals.value = False Or Trim(refDuals.Text) = "" Then
4442          Set model.Duals = Nothing
4443      Else
4444          Set model.Duals = sheet.Range(refDuals.Text)
4445      End If
4446      On Error GoTo ErrorHandler

4447      model.NonNegativityAssumption = chkNonNeg.value
          
          ' BuildModel fails if build is aborted by the user
4448      If Not model.BuildModel(sheet) Then GoTo ExitSub

          ' We know the save is confirmed now, so we can update values that aren't stored in the model
4252      SetDualsOnSheet chkGetDuals2.value, sheet
          SetUpdateSensitivity optUpdate.value, sheet

          ' Display on screen
4449      If chkShowModel.value = True Then OpenSolverVisualizer.ShowSolverModel sheet
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
                 "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
4458      refObj.SetFocus ' Set the focus back to the RefEdit
          GoTo ExitSub
          '----------------------------------------------------------------
BadDecRef:
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
                 "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
4470      refDuals.SetFocus ' Set the focus back to the target text box
          GoTo ExitSub

ExitSub:
          Application.Calculation = oldCalculationMode
4474      DoEvents ' Try to stop RefEdit bugs
          Exit Sub

ErrorHandler:
          ReportError "FModel", "cmdBuild_Click", True
          GoTo ExitSub
End Sub

Private Sub cmdAddCon_Click()
4496      On Error GoTo ErrorHandler

          Dim rngLHS As Range
          
          ' LEFT HAND SIDE
          Dim LHSisRange As Boolean, LHSisFormula As Boolean, LHSIsValueWithEqual As Boolean, LHSIsValueWithoutEqual As Boolean
4497      TestStringForConstraint refConLHS.Text, LHSisRange, LHSisFormula, LHSIsValueWithEqual, LHSIsValueWithoutEqual

4498      If LHSisRange = False Then
4499          MsgBox "Left-hand-side of constraint must be a range."
4500          Exit Sub
4501      End If
4506      Set rngLHS = Range(Trim(refConLHS.Text))

          ' RELATION
4508      If Len(cboConRel.Text) = 0 Then ' Should not happen
4509          MsgBox "Please select a relation such as = or <="
4510          Exit Sub
4511      End If

          Dim rel As RelationConsts
          rel = RelationStringToEnum(cboConRel.Text)
          
          ' RIGHT HAND SIDE
          
          Dim internalRHS As String, rngRHS As Range
4513      If RelationHasRHS(rel) Then
4514          If Trim(refConRHS.Text) = "" Then
4515              MsgBox "Please enter a right-hand-side!"
4516              Exit Sub
4517          End If

              Dim RHSisRange As Boolean, RHSIsFormula As Boolean, RHSIsValueWithEqual As Boolean, RHSIsValueWithoutEqual As Boolean
4518          TestStringForConstraint refConRHS.Text, RHSisRange, RHSIsFormula, RHSIsValueWithEqual, RHSIsValueWithoutEqual

4519          If Not RHSisRange And Not RHSIsFormula _
              And Not RHSIsValueWithEqual And Not RHSIsValueWithoutEqual Then
4520              MsgBox "The right-hand-side of a constraint can be either:" + vbNewLine + _
                         "A single-cell range (e.g. =A4)" + vbNewLine + _
                         "A multi-cell range of the same size as the LHS (e.g. =A4:B5)" + vbNewLine + _
                         "A single constant value (e.g. =2)" + vbNewLine + _
                         "A formula returning a single value (eg =sin(A4)"
4521              Exit Sub
4522          End If

4523          If RHSisRange Then
4524              Set rngRHS = Range(Trim(refConRHS.Text))
4531          Else
4532              internalRHS = refConRHS.Text
              End If
4557      End If

          ValidateConstraint rngLHS, rel, rngRHS, internalRHS

4558      AlterConstraints True

          Dim curCon As CConstraint
4561      If cmdAddCon.Caption <> "Add constraint" Then
              ' Update constraint
4562          Set curCon = model.Constraints(lstConstraints.ListIndex)
4585      Else
              ' Add constraint
              Set curCon = New CConstraint
4607          model.Constraints.Add curCon
          End If
          
4586      curCon.Init rngLHS, rel, Nothing, rngRHS, internalRHS

4608      If Not DontRepop Then model.PopulateConstraintListBox lstConstraints, chkNameRange.value
4609      Exit Sub

ErrorHandler:
4620      Application.DisplayAlerts = True
4621      OpenSolverSheet.Range("A1").FormulaLocal = "" ' This must be blank to ensure no risk of dialogs being shown trying to locate a sheet
4622      MsgBox Err.Description
4623      DoEvents ' Try to stop RefEdit bugs
4624      Exit Sub

End Sub

Private Sub cmdDelSelCon_Click()
4626      If lstConstraints.ListIndex = -1 Then
4627          MsgBox "Please select a constraint!", vbOKOnly, "AutoModel"
4628          Exit Sub
4629      End If

          ' Remove it
4630      model.Constraints.Remove lstConstraints.ListIndex

          ' Update form
4631      model.PopulateConstraintListBox lstConstraints, chkNameRange.value
End Sub

Private Sub lstConstraints_Change()
4633      If ConChangedMode = True And Not IsLoadingModel Then
              Dim SaveChanges As Boolean
4634          If cmdAddCon.Caption = "Update constraint" Then
4635              SaveChanges = (MsgBox("You have made changes to the current constraint." & vbNewLine & _
                                        "Do you want to save these changes?", vbYesNo) = vbYes)
4640          Else
4641              SaveChanges = (MsgBox("You have entered a constraint." & vbNewLine & _
                                        "Do you want to save this as a new constraint?", vbYesNo) = vbYes)
4646          End If
              If SaveChanges Then
                  DontRepop = True
4643              cmdAddCon_Click
4644              DontRepop = False
              End If
4647          AlterConstraints True
4650          model.PopulateConstraintListBox lstConstraints, chkNameRange.value
4651      End If

4653      If lstConstraints.ListIndex = -1 Then
4654          Exit Sub
4655      End If
4656      If lstConstraints.ListIndex = 0 Then
              'Add constraint
4657          refConLHS.Enabled = True
4658          DoEvents
4659          refConLHS.Text = ""
4660          DoEvents
              refConRHS.Text = ""
4662          DoEvents
4663          cboConRel_Change ' Set the RHS to be active based on the last constraint
4664          cmdAddCon.Enabled = True
4665          cmdAddCon.Caption = "Add constraint"
4666          cmdDelSelCon.Enabled = False
4667          DoEvents
4668          Application.CutCopyMode = False
4669          ActiveCell.Select ' Just select one cell, choosing a cell that should be visible to avoid scrolling
4670          Application.ScreenUpdating = True
4671          Exit Sub
4672      Else
              ' Update constraint
4673          refConLHS.Enabled = True
4674          DoEvents
4675          refConRHS.Enabled = True
4676          DoEvents
4677          cmdAddCon.Enabled = False
4678          cmdAddCon.Caption = "Update constraint"
4679          cmdDelSelCon.Enabled = True
4680          DoEvents

              ' Set text and marching ants; the select/copy operations can throw errors if on different sheets
4681          On Error Resume Next
4682          ActiveCell.Select   ' We may fail in the next steps, se we cancel any old highlighting
4683          Application.CutCopyMode = False
              Dim copyRange As Range, curCon As CConstraint
              Set curCon = model.Constraints(lstConstraints.ListIndex)
4684          With curCon
4685              refConLHS.Text = GetDisplayAddress(.LHS, False)
4686              Set copyRange = .LHS
4687              cboConRel.ListIndex = cboPosition(RelationEnumToString(.RelationType))
4688              refConRHS.Text = ""
4689              If Not .RHS Is Nothing Then
4690                  refConRHS.Text = GetDisplayAddress(.RHS, False)
4691                  Set copyRange = ProperUnion(copyRange, .RHS)
4692              ElseIf .RHS Is Nothing And .RHSstring <> "" Then
4693                  Dim newRHS As String
                      newRHS = ConvertToCurrentLocale(.RHSstring)
                      If Mid(newRHS, 1) = "=" Then newRHS = Mid(newRHS, 2)
                      refConRHS.Text = RemoveActiveSheetNameFromString(newRHS)
4698              End If
4699          End With
4700          cboConRel_Change

              ' Will fail if LHS and RHS are different shape
              ' Silently fail, nothing that can be done about it
              ' ALSO fails at the Union step if on different shapes
4701          copyRange.Select
4702          copyRange.Copy
4703          Application.ScreenUpdating = True

4704          Exit Sub
4705      End If
End Sub

Private Sub lstConstraints_Exit(ByVal Cancel As MSForms.ReturnBoolean)
          ' When the focus leaves this list, we want to remove any highlighting shown by selected cells
4706      ActiveCell.Select
End Sub

Sub TestStringForConstraint(ByVal TheString As String, _
                            RefersToRange As Boolean, _
                            RefersToFormula As Boolean, _
                            RefersToValueWithEqual As Boolean, _
                            RefersToValueWithoutEqual As Boolean)
4707      TheString = Trim(TheString)
4708      If Len(TheString) = 0 Then Exit Sub

          ' Test for RANGE
4709      On Error Resume Next
          Dim r As Range
4710      Set r = Range(TheString)
4711      RefersToRange = (Err.Number = 0)

          ' Test for ...
4712      If Not RefersToRange Then
              ' Not a range, might be constant?
4713          If Mid(TheString, 1, 1) <> "=" Then
                  ' Not sure what this is, but assume its OK - if an equal is added
4714              RefersToValueWithoutEqual = True
4715          Else
                  ' Test for a numeric constant, in US format
4716              If IsAmericanNumber(Mid(TheString, 2)) Then
4717                  RefersToValueWithEqual = True
4718              Else
                      'FORMULA
4719                  RefersToFormula = True
4720              End If
4721          End If
4722      End If

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
    
    With lstConstraints
        .Left = lblDescHeader.Left
        .Top = lblConstraintGroup.Top
        .Height = MinHeight
        .Width = LeftOf(lblConstraintGroup, .Left)
    End With
    
    With chkNameRange
        .Left = lblDescHeader.Left
        .Width = lstConstraints.Width
        .Caption = "Show named ranges"
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
    IsResizing = False
    
    ' Set the vertical positions of the lower half of the form
    UpdateLayout
    
    Me.Width = Me.Width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - Model"
End Sub

Private Sub UpdateLayout(Optional ChangeY As Single = 0)
' Do the layout of the lower half of the form, changing the height of the list box by ChangeY
    Dim NewHeight As Double
    NewHeight = lstConstraints.Height + ChangeY
    If NewHeight < MinHeight Then NewHeight = MinHeight
    
    lstConstraints.Height = NewHeight
        
    ' Cascade the updated height
    chkNameRange.Top = Below(lstConstraints)
    lblDiv4.Top = Below(chkNameRange)
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
