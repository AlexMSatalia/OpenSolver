VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmModel 
   Caption         =   "OpenSolver - Model"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9840
   OleObjectBlob   =   "frmModel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmModel"
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

Private ListItem As Long
Private ConChangedMode As Boolean
Private DontRepop As Boolean

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
4182      optUpdate.Enabled = chkGetDuals2.value And TrueIfEnable
4183      optNew.Enabled = chkGetDuals2.value And TrueIfEnable

4184      refDuals.Enabled = TrueIfEnable And chkGetDuals.value And chkGetDuals.Enabled

4185      lblDiv5.Enabled = False
4186      lblDiv6.Enabled = False

4187      chkShowModel.Enabled = TrueIfEnable
4188      cmdOptions.Enabled = TrueIfEnable
4189      cmdBuild.Enabled = TrueIfEnable
4190      cmdCancel.Enabled = TrueIfEnable
          cmdReset.Enabled = TrueIfEnable
          cmdChange.Enabled = TrueIfEnable

          Dim Solver As String
4199      If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!OpenSolver_ChosenSolver", Solver) Then
4200          Solver = "CBC"
4201          Call SetNameOnSheet("OpenSolver_ChosenSolver", "=" & Solver)
4202      End If

4203      If Not SolverHasSensitivityAnalysis(Solver) Then
              ' Disable dual options
4204          chkGetDuals2.Enabled = False
4205          chkGetDuals.Enabled = False
4206          optUpdate.Enabled = False
4207          optNew.Enabled = False
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

4214      If Not model.ObjectiveFunctionCell Is Nothing Then refObj.Text = GetDisplayAddress(model.ObjectiveFunctionCell, False)

4215      If Not model.DecisionVariables Is Nothing Then refDecision.Text = GetDisplayAddressInCurrentLocale(model.DecisionVariables)

4216      chkGetDuals.value = Not model.Duals Is Nothing
4217      If model.Duals Is Nothing Then
4218          refDuals.Text = ""
4219      Else
4220          refDuals.Text = GetDisplayAddress(model.Duals, False)
4221      End If

4222      model.PopulateConstraintListBox lstConstraints
4223      lstConstraints_Change

          Dim sheetName As String, value As String, ResetDualsNewSheet As Boolean
4224      sheetName = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!"
4225      If GetNameValueIfExists(ActiveWorkbook, sheetName & "OpenSolver_DualsNewSheet", value) Then
4226          chkGetDuals2.value = value
              ' If checkbox is null, then the stored value was not 'True' or 'False'. We should reset to false
4227          If IsNull(chkGetDuals2.value) Then
4228              ResetDualsNewSheet = True
4229          End If
4230      Else
4231          ResetDualsNewSheet = True
4232      End If

4233      If ResetDualsNewSheet Then
4234          Call SetNameOnSheet("OpenSolver_DualsNewSheet", "=FALSE")
4235          chkGetDuals2.value = False
4236      End If

4237      optUpdate.Enabled = chkGetDuals2.value
4238      optNew.Enabled = chkGetDuals2.value
4239      If GetNameValueIfExists(ActiveWorkbook, sheetName & "OpenSolver_UpdateSensitivity", value) Then
4240          If value = "TRUE" Then
4241            optUpdate.value = value
4242          Else
4243            optNew.value = True
4244          End If
4245      Else
4246          Call SetNameOnSheet("OpenSolver_UpdateSensitivity", "=TRUE")
4247          optUpdate.value = True
4248      End If
End Sub

Private Sub chkGetDuals_Click()
4250      refDuals.Enabled = chkGetDuals.value
End Sub

Private Sub chkGetDuals2_Click()
4252      Call SetNameOnSheet("OpenSolver_DualsNewSheet", "=" & chkGetDuals2.value)
4253      optUpdate.Enabled = chkGetDuals2.value
4254      optNew.Enabled = chkGetDuals2.value
End Sub

Private Sub chkNameRange_Click()
4256      model.PopulateConstraintListBox lstConstraints
4257      lstConstraints_Change
End Sub

Private Sub cmdCancelCon_Click()
4259      Disabler True
4260      cmdAddCon.Enabled = False
4261      ConChangedMode = False
4262      lstConstraints_Change
End Sub


Private Sub cmdChange_Click()
#If Mac Then
4264      frmSolverChange.Show
#Else
4265      frmSolverChange.Show vbModal
#End If
End Sub

Private Sub cmdOptions_Click()
          ' Save the current "Assume Non Negative" option so this is shown in the options dialog.
          ' The saved value gets updated on OK, which we then reflect in our Model dialog
          Dim s As String
4267      SetSolverNameOnSheet "neg", IIf(chkNonNeg.value, "=1", "=2")

#If Mac Then
4268      frmOptions.Show
#Else
4269      frmOptions.Show vbModal
#End If
4270      If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_neg", s) Then    ' This should always be true
4271          chkNonNeg.value = s = "1"
4272      End If
End Sub

'--------------------------------------------------------------------------------------
'Reset Button
'Deletes the objective function, decision variables and all the constraints in the model
'---------------------------------------------------------------------------------------

Private Sub cmdReset_Click()
          Dim NumConstraints As Single, i As Long

          'Check the user wants to reset the model
4274      If MsgBox("This will reset the objective function, decision variables and constraints." _
                  & vbCrLf & " Are you sure you want to do this?", vbYesNo, "OpenSolver") = vbNo Then
4275          Exit Sub
4276      End If

          'Reset the objective function and the decision variables
4277      refObj.Text = ""
4278      refDecision.Text = ""

          'Find the number of constraints in model
4279      NumConstraints = model.Constraints.Count

          ' Remove the constraints
4280      For i = 1 To NumConstraints
4281          model.Constraints.Remove 1
4282      Next i

          ' Update constraints form
4283      model.PopulateConstraintListBox lstConstraints

End Sub

Private Sub optMax_Click()
4285      txtObjTarget.Enabled = optTarget.value
End Sub

Private Sub optMin_Click()
4287      txtObjTarget.Enabled = optTarget.value
End Sub

Private Sub optNew_Click()
4289      Call SetNameOnSheet("OpenSolver_UpdateSensitivity", "=" & optUpdate.value)
End Sub

Private Sub optTarget_Click()
4291      txtObjTarget.Enabled = optTarget.value
End Sub

Private Sub optUpdate_Click()
4293      Call SetNameOnSheet("OpenSolver_UpdateSensitivity", "=" & optUpdate.value)
End Sub

Private Sub AlterConstraints(DoDisable As Boolean)
          Disabler DoDisable
          cmdAddCon.Enabled = Not DoDisable
          ConChangedMode = Not DoDisable
End Sub

Private Sub refConLHS_Change()
          ' Compare to expected value
          Dim DoDisable As Boolean
4295      If ListItem >= 1 And Not model.Constraints Is Nothing Then
4297          If refConLHS.Text <> model.Constraints(ListItem).LHS.Address Then
4298              DoDisable = False
4301          Else
4302              DoDisable = True
4305          End If
4306      ElseIf ListItem = 0 Then
4307          If refConLHS.Text <> "" Then
4308              DoDisable = False
4311          Else
                  DoDisable = True
4315          End If
4316      End If
          AlterConstraints DoDisable
End Sub

Private Sub refConRHS_Change()
          ' Compare to expected value
          Dim DoDisable As Boolean
4318      If ListItem >= 1 And Not model.Constraints Is Nothing Then
              Dim origRHS As String
4319          If model.Constraints(ListItem).RHS Is Nothing Then
4320              origRHS = model.Constraints(ListItem).RHSstring
4321          Else
4322              origRHS = model.Constraints(ListItem).RHS.Address
4323          End If
4324          If refConRHS.Text <> origRHS Then
4325              DoDisable = False
4328          Else
4329              DoDisable = True
4332          End If
4333      ElseIf ListItem = 0 Then
4334          If refConLHS.Text <> "" Then
4335              DoDisable = False
4338          Else
4339              DoDisable = True
4342          End If
4343      End If
          AlterConstraints DoDisable
End Sub

Private Sub UserForm_Activate()
          ' Check we can even start
4345      If Not CheckWorksheetAvailable Then
4346          Unload Me
4347          Exit Sub
4348      End If

4349      SetAnyMissingDefaultExcel2007SolverOptions
4350      Set model = New CModel

4351      If SheetHasOpenSolverHighlighting(ActiveSheet) Then HideSolverModel
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
          Dim Solver As String
4369      If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!OpenSolver_ChosenSolver", Solver) Then
4370          lblSolver.Caption = "Current Solver Engine: " & UCase(left(Solver, 1)) & Mid(Solver, 2)
4371      Else
              lblSolver.Caption = "Current Solver Engine: CBC"
4372      End If
          ' Load the model on the sheet into memory
4373      ListItem = -1
4374      ConChangedMode = False
4375      DontRepop = False
4376      Disabler True
4377      model.LoadFromSheet
4378      DoEvents
4379      UpdateFormFromMemory
4380      DoEvents
          ' Take focus away from refEdits
4381      DoEvents
4382      Repaint
4383      DoEvents
End Sub


Private Sub cmdCancel_Click()
4386      DoEvents
4387      On Error Resume Next ' Just to be safe on our select
4388      Application.CutCopyMode = False
4389      ActiveCell.Select ' Just select one cell, choosing a cell that should be visible to avoid scrolling
4390      Application.ScreenUpdating = True
          Me.Hide
End Sub

Private Sub cmdRunAutoModel_Click()
          ' Try and guess the objective
          Dim status As String
4392      status = model.FindObjective(ActiveSheet)

          ' Get it in memory
4393      Load frmAutoModel
          ' Pass it the model reference
4394      Set frmAutoModel.model = model
4395      frmAutoModel.GuessObjStatus = status

4396      Select Case status
              Case "NoSense", "SenseNoCell"
#If Mac Then
                  'MacAutoModel.Show
                  ' Mac can't use a RefEdit properly when multiple forms are open since it doesn't support modeless forms properly
                  ' Trying to fix this using DoEvents and hiding forms seems to be hard. Random RefEdits take focus in places, so diasabling for now
4397              MsgBox ("Couldn't find objective cell, and couldn't finish as a result.")
#Else
4398              frmAutoModel.Show vbModal
#End If
4399          Case Else ' Found objective
4400              model.FindVarsAndCons IsFirstTime:=True
4401      End Select

          ' Force the automatically created model to be AssumeNonNegative
4402      model.NonNegativityAssumption = True

4403      UpdateFormFromMemory
4404      DoEvents
4405      Application.StatusBar = False
End Sub

Private Sub cmdBuild_Click()
4408      DoEvents
4409      On Error Resume Next
4410      Application.CutCopyMode = False
4411      ActiveCell.Select ' Just select one cell, choosing a cell that should be visible to avoid scrolling
4412      Application.ScreenUpdating = True
4413      On Error GoTo 0

          Dim oldCalculationMode As Long
          oldCalculationMode = Application.Calculation
          Application.Calculation = xlCalculationManual

          ' Pull possibly update objective info into model
4414      On Error GoTo BadObjRef
4415      If Trim(refObj.Text) = "" Then
4416          Set model.ObjectiveFunctionCell = Nothing
4417      Else
4418          Set model.ObjectiveFunctionCell = Range(refObj.Text)
4419      End If
4420      On Error GoTo errorHandler

          ' Get the objective sense
4421      If optMax.value = True Then model.ObjectiveSense = MaximiseObjective
4422      If optMin.value = True Then model.ObjectiveSense = MinimiseObjective
4423      If optTarget.value = True Then
4424          model.ObjectiveSense = TargetObjective
4425          On Error GoTo BadObjectiveTarget
4426          model.ObjectiveTarget = CDbl(txtObjTarget.Text)
4427          On Error GoTo errorHandler
4428      End If
4429      If model.ObjectiveSense = UnknownObjectiveSense Then
4430          MsgBox "Error: Please select an objective sense (minimise, maximise or target).", vbExclamation + vbOKOnly, "OpenSolver"
4431          Exit Sub
4432      End If

          '----------------------------------------------------------------
          ' Pull possibly updated decision variable info into model
          ' We allow multiple area ranges here, which requires ConvertFromCurrentLocale as delimiter can vary
4433      On Error GoTo BadDecRef
4434      If Trim(refDecision.Text) = "" Then
4435          Set model.DecisionVariables = Nothing
4436      Else
4437          Set model.DecisionVariables = Range(ConvertFromCurrentLocale(refDecision.Text))
4438      End If
4439      On Error GoTo errorHandler

          '----------------------------------------------------------------
          ' Pull possibly updated dual storage cells
4440      On Error GoTo BadDualsRef

4441      If chkGetDuals.value = False Or Trim(refDuals.Text) = "" Then
4442          Set model.Duals = Nothing
4443      Else
4444          Set model.Duals = Range(refDuals.Text)
4445      End If
4446      On Error GoTo errorHandler

          '----------------------------------------------------------------
          ' Do it
4447      model.NonNegativityAssumption = chkNonNeg.value

4448      model.BuildModel


          '----------------------------------------------------------------
          ' Display on screen
4449      If chkShowModel.value = True Then OpenSolverVisualizer.ShowSolverModel
4450      On Error GoTo CalculateFailed
4451      Application.Calculate
4452      On Error GoTo errorHandler
          '----------------------------------------------------------------
          ' Finish
          Application.Calculation = oldCalculationMode
          Me.Hide
4453      Exit Sub

          '----------------------------------------------------------------
CalculateFailed:
          ' Application.Calculate failed. Ignore error and try again
4454      On Error GoTo errorHandler
4455      Application.Calculate
4456      Resume Next

          '----------------------------------------------------------------
BadObjRef:
          ' Couldn't turn the objective cell address into a range
4457      MsgBox "Error: the cell address for the objective is invalid. " + _
                  "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
4458      refObj.SetFocus ' Set the focus back to the RefEdit
          GoTo cleanUp
          '----------------------------------------------------------------
BadDecRef:
          ' Couldn't turn the decision variable address into a range
4461      MsgBox "Error: the cell range specified for the Variable Cells is invalid. " + _
                 "This must be a valid Excel range that does not exceed Excel's internal character count limits. " + _
                  "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
4462      refDecision.SetFocus ' Set the focus back to the RefEdit
          GoTo cleanUp
BadObjectiveTarget:
          ' Couldn't turn the objective target into a value
4465      MsgBox "Error: the target value for the objective cell is invalid. " + _
                  "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
4466      txtObjTarget.SetFocus ' Set the focus back to the target text box
          GoTo cleanUp
BadDualsRef:
          ' Couldn't turn the dual cell into a range
4469      MsgBox "Error: the cell for storing the shadow prices is invalid. " + _
                  "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
4470      refDuals.SetFocus ' Set the focus back to the target text box
          GoTo cleanUp
errorHandler:
4473      MsgBox "While constructing the model, OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source & " (frmModel::cmdBuild_Click)", , "OpenSolver Code Error"
          GoTo cleanUp
cleanUp:
          Application.Calculation = oldCalculationMode
4474      DoEvents ' Try to stop RefEdit bugs
4475      Exit Sub
End Sub

Private Sub cboConRel_Change()
4477      If cboConRel.Text = "=" Then refConRHS.Enabled = True
4478      If cboConRel.Text = "<=" Then refConRHS.Enabled = True
4479      If cboConRel.Text = ">=" Then refConRHS.Enabled = True
4480      If cboConRel.Text = "int" Or cboConRel.Text = "bin" Or cboConRel.Text = "alldiff" Then
4481          refConRHS.Enabled = False
4482      End If

4483      If ListItem >= 1 And Not model.Constraints Is Nothing Then
4485          AlterConstraints (cboConRel.Text = model.Constraints(ListItem).ConstraintType)
4494      End If
End Sub

Private Sub cmdAddCon_Click()
4496      On Error GoTo errorHandler

          Dim rngLHS As Range, rngRHS As Range
          Dim IsRestrict As Boolean

          '================================================================
          ' Validation
          ' Solver enforces the following requirements.
          ' The LHS mut be a range with one or more cells (one area!)
          ' The RHS can be either:
          '   A single-cell range (=A4)
          '   A multi-cell range of the same size as the LHS (=A4:B5)
          '   A single constant value (eg =2)
          '   A formula returning a single value (eg =sin(A4))
          '----------------------------------------------------------------
          '
          'TODO: This needs tidying up to handle locales prooperly. Need to check that it converts values/formaulae to the current local when it shows them, and
          '      back again when it saves them
          '      We currently do this by putting the RHS into a cell and using .formula and .formulalocal to make this conversion. This won't always be needed
          '      This code also distinguishes a leading '='; this should not be needed
          '
          ' LEFT HAND SIDE
          Dim LHSisRange As Boolean, LHSisFormula As Boolean, LHSIsValueWithEqual As Boolean, LHSIsValueWithoutEqual As Boolean
4497      TestStringForConstraint refConLHS.Text, LHSisRange, LHSisFormula, LHSIsValueWithEqual, LHSIsValueWithoutEqual

4498      If LHSisRange = False Then
              ' The string in the LHS refedit does not describe a range
4499          MsgBox "Left-hand-side of constraint must be a range."
4500          Exit Sub
4501      End If
4502      If Range(Trim(refConLHS.Text)).Areas.Count > 1 Then
              ' The LHS is multiple areas - not allowed
4503          MsgBox "Left-hand-side of constraint must have only one area."
4504          Exit Sub
4505      End If
4506      Set rngLHS = Range(Trim(refConLHS.Text))

          '----------------------------------------------------------------
          ' RIGHT HAND SIDE
          Dim RHSisRange As Boolean, RHSisFormula As Boolean, RHSIsValueWithEqual As Boolean, RHSIsValueWithoutEqual As Boolean
          Dim strRel As String
4507      strRel = cboConRel.Text
4508      If strRel = "" Then ' Should not happen
4509          MsgBox "Please select a relation such as = or <="
4510          Exit Sub
4511      End If
4512      IsRestrict = Not ((strRel = "=") Or (strRel = "<=") Or (strRel = ">="))
4513      If Not IsRestrict Then
4514          If Trim(refConRHS.Text) = "" Then
4515              MsgBox "Please enter a right-hand-side!"
4516              Exit Sub
4517          End If

4518          TestStringForConstraint refConRHS.Text, RHSisRange, RHSisFormula, RHSIsValueWithEqual, RHSIsValueWithoutEqual

4519          If Not RHSisRange And Not RHSisFormula _
              And Not RHSIsValueWithEqual And Not RHSIsValueWithoutEqual Then
4520              MsgBox "The right-hand-side of a constraint can be either:" + vbNewLine + _
                          "A single-cell range (e.g. =A4)" + vbNewLine + _
                          "A multi-cell range of the same size as the LHS (e.g. =A4:B5)" + vbNewLine + _
                          "A single constant value (e.g. =2)" + vbNewLine + _
                          "A formula returning a single value (eg =sin(A4)"
4521              Exit Sub
4522          End If

4523          If RHSisRange Then
                  ' If it is single cell, thats OK
                  ' If it is multi cell, it must match cell count for LHS
4524              Set rngRHS = Range(Trim(refConRHS.Text))
4525              If rngRHS.Count > 1 Then
4526                  If rngRHS.Count <> rngLHS.Count Then
                          ' Mismatch!
4527                      MsgBox "Right-hand-side of constraint has more than one cell, and does not match the number of cells on the left-hand-side."
4528                      Exit Sub
4529                  End If
4530              End If
4531          End If

              ' If not a range then evaluate to see if it's legit
              ' Evaluate is not locale-friendly
              ' So we put it in a cell on the internal sheet, then get it back
              ' We need to prefix the formula with an "=" otherwise formula such as 'sheet name'!A1 get entered as a string constant (becaused of the leading ')
              Dim internalRHS As String
4532          internalRHS = Trim(refConRHS.Text)

              ' Turn off dialog display; we do not want try to open a workbook with a name of the worksheet! This happens if the formula comes from a worksheet
              ' whose name contains a space
4533          Application.DisplayAlerts = False
4534          On Error GoTo ErrorHandler_CannotInterpretRHS
4535          OpenSolverSheet.Range("A1").FormulaLocal = IIf(left(internalRHS, 1) = "=", "", "=") & refConRHS.Text
4536          internalRHS = OpenSolverSheet.Range("A1").Formula
4537          OpenSolverSheet.Range("A1").Clear ' This must be blank to ensure no risk of dialogs being shown trying to locate a sheet
4538          Application.DisplayAlerts = True

4539          If Not RHSisRange Then
                  ' Can we evaluate this function or constant?
                  Dim varReturn As Variant
4540              varReturn = ActiveSheet.Evaluate(internalRHS) ' Must be worksheet.evaluate to get references to names local to the sheet
4541              If VBA.VarType(varReturn) = vbError Then
4542                  MsgBox "The formula or value for the RHS is not valid. Please check and try again."
4543                  refConRHS.SetFocus
4544                  DoEvents
4545                  Exit Sub
4546              End If
4547          End If

              ' If it isn't a range, lets convert any cell references to absolute
              ' Will fail if refConRHS has a non-English locale number
4548          If left(internalRHS, 1) <> "=" Then
4549              varReturn = Application.ConvertFormula("=" + internalRHS, FromReferenceStyle:=xlA1, ToReferenceStyle:=xlA1, ToAbsolute:=xlAbsolute)
4550          Else
4551              varReturn = Application.ConvertFormula(internalRHS, FromReferenceStyle:=xlA1, ToReferenceStyle:=xlA1, ToAbsolute:=xlAbsolute)
4552          End If
4553          If (VBA.VarType(varReturn) = vbError) Then
                  ' Its valid, but couldn't convert to standard form, probably because not A1... just leave it
4554          Else
                  ' Always comes back with a = at the start
                  ' Unfortunately, return value will have wrong locale...
                  ' But not much can be done with that?
4555              refConRHS.Text = Mid(varReturn, 2, Len(varReturn))
4556          End If


4557      End If

4558      AlterConstraints True

          '================================================================
          ' Update constraint?
4561      If cmdAddCon.Caption <> "Add constraint" Then

              'With model.Constraints(f.lstConstraints.ListIndex)
4562          With model.Constraints(ListItem)
4563              Set .LHS = rngLHS
4564              Set .Relation = Nothing
4565              .ConstraintType = strRel
4566              If IsRestrict Then
4567                  Set .RHS = Nothing
4568                  .RHSstring = ""
4569              Else
4570                  If RHSisRange Then
4571                      Set .RHS = rngRHS
4572                      .RHSstring = ""
4573                  Else
4574                      Set .RHS = Nothing
4575                      If left(.RHSstring, 1) <> "=" Then
4576                          .RHSstring = "=" + refConRHS.Text
4577                      Else
4578                          .RHSstring = refConRHS.Text
4579                      End If
4580                  End If
4581              End If
4582          End With

4583          If Not DontRepop Then model.PopulateConstraintListBox lstConstraints
4584          Exit Sub
4585      Else
          '================================================================
          ' Add constraint
              Dim NewConstraint As New CConstraint
4586          With NewConstraint
4587              Set .LHS = rngLHS
4588              Set .Relation = Nothing
4589              .ConstraintType = strRel
4590              If IsRestrict Then
4591                  Set .RHS = Nothing
4592                  .RHSstring = ""
4593              Else
4594                  If RHSisRange Then
4595                      Set .RHS = rngRHS
4596                      .RHSstring = ""
4597                  Else
4598                      Set .RHS = Nothing
4599                      If left(.RHSstring, 1) <> "=" Then
4600                          .RHSstring = "=" + refConRHS.Text ' This has been converted above into a US-locale formula, value or reference
4601                      Else
4602                          .RHSstring = refConRHS.Text ' This has been converted above into a US-locale formula, value or reference
4603                      End If
4604                  End If
4605              End If
4606          End With

4607          model.Constraints.Add NewConstraint ', NewConstraint.GetKey
4608          If Not DontRepop Then model.PopulateConstraintListBox lstConstraints
4609          Exit Sub
4610      End If

4611      Application.DisplayAlerts = True
4612      OpenSolverSheet.Range("A1").Clear ' This must be blank to ensure no risk of dialogs being shown trying to locate a sheet
4613      Exit Sub

ErrorHandler_CannotInterpretRHS:
4614      Application.DisplayAlerts = True
4615      OpenSolverSheet.Range("A1").Clear ' This must be blank to ensure no risk of dialogs being shown trying to locate a sheet
          ' Couldn't turn the RHS into a formula
4616      MsgBox "The formula or value for the RHS is not valid. Please check and try again."
4617      refConRHS.SetFocus
4618      DoEvents ' Try to stop RefEdit bugs
4619      Exit Sub
errorHandler:
4620      Application.DisplayAlerts = True
4621      OpenSolverSheet.Range("A1").FormulaLocal = "" ' This must be blank to ensure no risk of dialogs being shown trying to locate a sheet
4622      MsgBox "While constructing the model, OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source & " (frmModel::cmdBuild_Click)", , "OpenSolver Code Error"
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
4631      model.PopulateConstraintListBox lstConstraints
End Sub

Private Sub lstConstraints_Change()
4633      If ConChangedMode = True Then
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
4650          model.PopulateConstraintListBox lstConstraints
4651      End If

4652      ListItem = lstConstraints.ListIndex

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
              Dim copyRange As Range
4684          With model.Constraints(lstConstraints.ListIndex)
4685              refConLHS.Text = GetDisplayAddress(.LHS, False)
4686              Set copyRange = .LHS
4687              cboConRel.ListIndex = cboPosition(.ConstraintType)
4688              refConRHS.Text = ""
4689              If Not .RHS Is Nothing Then
4690                  refConRHS.Text = GetDisplayAddress(.RHS, False)
4691                  Set copyRange = ProperUnion(copyRange, .RHS)
4692              ElseIf .RHS Is Nothing And .RHSstring <> "" Then
4693                  If Mid(.RHSstring, 1, 1) = "=" Then
4694                      refConRHS.Text = RemoveActiveSheetNameFromString(Mid(.RHSstring, 2, Len(.RHSstring)))
4695                  Else
4696                      refConRHS.Text = RemoveActiveSheetNameFromString(.RHSstring)
4697                  End If
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
End Sub

Private Sub AutoLayout()
    AutoFormat Me.Controls
    
    Me.width = FormWidthModel
    
    With cmdRunAutoModel
        .Caption = "AutoModel"
        .width = FormButtonWidth * 1.3
        .top = FormMargin
        .left = Me.width - FormMargin - .width
    End With
    
    With lblDescHeader
        .left = FormMargin
        .top = cmdRunAutoModel.top + FormSpacing
        .Caption = "What is AutoModel?"
    End With
    
    With lblDesc
        .left = lblDescHeader.left
        .top = cmdRunAutoModel.top + cmdRunAutoModel.height + FormSpacing
        .Caption = "AutoModel is a feature of OpenSolver that tries to automatically determine " & _
                   "the problem you are trying to optimise by observing the structure of the " & _
                   "spreadsheet. It will turn its best guess into a Solver model, which you can " & _
                   "then edit in this window."
        .width = Me.width - 2 * FormMargin
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
        .width = Me.width - 2 * FormMargin
    End With
    
    With lblDiv1
        .left = lblDescHeader.left
        .top = lblDesc.top + lblDesc.height + FormSpacing
        .width = lblDesc.width
        .height = FormDivHeight
        .BackColor = FormDivBackColor
    End With
    
    With lblStep1
        .Caption = "Objective Cell:"
        .left = lblDescHeader.left
        .top = lblDiv1.top + lblDiv1.height + FormSpacing
        .width = Me.width
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
    End With
    
    With txtObjTarget
        .width = cmdRunAutoModel.width
        .top = lblStep1.top
        .left = Me.width - FormMargin - .width
    End With
    
    With optTarget
        .Caption = "target value:"
        .width = Me.width
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
        .top = lblStep1.top
        .left = txtObjTarget.left - FormSpacing - .width
    End With
    
    With optMin
        .Caption = "minimise"
        .width = Me.width
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
        .top = lblStep1.top
        .left = optTarget.left - FormSpacing - .width
    End With
    
    With optMax
        .Caption = "maximise"
        .width = Me.width
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
        .top = lblStep1.top
        .left = optMin.left - FormSpacing - .width
    End With
    
    With refObj
        .left = lblStep1.left + lblStep1.width + FormSpacing
        .top = lblStep1.top
        .width = optMax.left - FormSpacing - .left
        .height = FormTextBoxHeight
    End With
    
    With lblDiv2
        .left = lblDescHeader.left
        .top = optMax.top + optMax.height + FormSpacing
        .width = lblDesc.width
        .height = FormDivHeight
        .BackColor = FormDivBackColor
    End With
    
    With lblStep2
        .Caption = "Variable Cells:"
        .left = lblDescHeader.left
        .top = lblDiv2.top + lblDiv2.height + FormSpacing
        .width = Me.width
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
    End With
    
    With refDecision
        .height = 2 * FormTextHeight
        .left = refObj.left
        .top = lblStep2.top
        .width = Me.width - FormMargin - .left
    End With
    
    With lblDiv3
        .left = lblDescHeader.left
        .top = refDecision.top + refDecision.height + FormSpacing
        .width = lblDesc.width
        .height = FormDivHeight
        .BackColor = FormDivBackColor
    End With
    
    With lblStep3
        .Caption = "Constraints:"
        .left = lblDescHeader.left
        .top = lblDiv3.top + lblDiv3.height + FormSpacing
        .width = lblDesc.width
    End With
    
    lblConstraintGroup.top = lblStep3.top + lblStep3.height
    
    With cboConRel
        .width = cmdRunAutoModel.width / 2
        .height = refObj.height
        .top = lblConstraintGroup.top + FormSpacing
        .left = Me.width - FormMargin - FormSpacing - .width
    End With
    
    With refConLHS
        .width = cboConRel.width * 3
        .height = refObj.height
        .top = cboConRel.top
        .left = cboConRel.left - FormSpacing - .width
    End With
    
    With refConRHS
        .width = refConLHS.width
        .left = refConLHS.left
        .height = refObj.height
        .top = refConLHS.top + refConLHS.height + FormSpacing
    End With
    
    With cmdAddCon
        .Caption = "Add constraint"
        .left = refConLHS.left
        .top = refConRHS.top + refConLHS.height + FormSpacing
        .width = cboConRel.width * 2
    End With
    
    With cmdCancelCon
        .Caption = "Cancel"
        .left = cmdAddCon.left + cmdAddCon.width + FormSpacing
        .top = cmdAddCon.top
        .width = cmdAddCon.width
    End With
    
    With lblConstraintGroup
        .left = refConLHS.left - FormSpacing
        .width = FormSpacing * 3 + cmdAddCon.width + cmdCancelCon.width
        .height = FormSpacing * 4 + refConLHS.height + refConRHS.height + cmdAddCon.height
    End With
    
    With cmdDelSelCon
        .Caption = "Delete selected constraint"
        .left = lblConstraintGroup.left
        .top = lblConstraintGroup.top + lblConstraintGroup.height + FormSpacing
        .width = lblConstraintGroup.width
    End With
    
    With chkNonNeg
        .Caption = "Make unconstrainted variable cells non-negative"
        .left = lblConstraintGroup.left
        .top = cmdDelSelCon.top + cmdDelSelCon.height + FormSpacing
        .width = lblConstraintGroup.width
    End With
    
    With lstConstraints
        .left = lblDescHeader.left
        .top = lblConstraintGroup.top
        .height = MinHeight
        .width = lblConstraintGroup.left - .left - FormSpacing
    End With
    
    With chkNameRange
        .left = lblDescHeader.left
        .width = lstConstraints.width
        .Caption = "Show named ranges"
    End With
    
    With lblDiv4
        .left = lblDescHeader.left
        .width = lblDesc.width
        .height = FormDivHeight
        .BackColor = FormDivBackColor
    End With
    
    With lblStep4
        .Caption = "Sensitivity Analysis"
        .left = lblDescHeader.left
        .width = Me.width
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
    End With
    
    With chkGetDuals
        .Caption = "List sensitivity analysis on the same sheet with top left cell:"
        .left = lblStep4.left + lblStep4.width + FormSpacing
        .width = Me.width
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
    End With
    
    With refDuals
        .left = chkGetDuals.left + chkGetDuals.width + FormSpacing
        .width = Me.width - FormMargin - .left
        .height = refObj.height
    End With
    
    With chkGetDuals2
        .Caption = "Output sensitivity analysis:"
        .left = chkGetDuals.left
        .width = Me.width
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
    End With
    
    With optUpdate
        .Caption = "updating any previous output sheet"
        .left = chkGetDuals2.left + chkGetDuals2.width + FormSpacing
        .width = Me.width
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
    End With
    
    With optNew
        .Caption = "on a new sheet"
        .left = optUpdate.left + optUpdate.width + FormSpacing
        .width = Me.width
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
    End With
    
    With lblDiv5
        .left = lblDescHeader.left
        .width = lblDesc.width
        .height = FormDivHeight
        .BackColor = FormDivBackColor
    End With
    
    With lblStep5
        .left = lblDescHeader.left
        .Caption = "Solver Engine:"
    End With
    
    With cmdChange
        .width = cmdRunAutoModel.width
        .left = Me.width - FormMargin - .width
        .Caption = "Solver Engine..."
    End With
    
    With lblSolver
        .width = cmdChange.left - FormSpacing - .left
    End With
    
    With lblDiv6
        .left = lblDescHeader.left
        .width = lblDesc.width
        .height = FormDivHeight
        .BackColor = FormDivBackColor
    End With
        
    With chkShowModel
        .left = lblDescHeader.left
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
    End With
    
    With cmdCancel
        .width = cmdRunAutoModel.width
        .Caption = "Cancel"
        .left = Me.width - FormMargin - .width
    End With
    
    With cmdBuild
        .width = cmdRunAutoModel.width
        .Caption = "Save Model"
        .left = cmdCancel.left - FormSpacing - .width
    End With
    
    With cmdOptions
        .width = cmdRunAutoModel.width
        .Caption = "Options..."
        .left = cmdBuild.left - FormSpacing - .width
    End With
    
    With cmdReset
        .width = cmdRunAutoModel.width
        .Caption = "Clear Model"
        .left = cmdOptions.left - FormSpacing - .width
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
        .left = Me.width - .width
        .MousePointer = fmMousePointerSizeNWSE
        .BackStyle = fmBackStyleTransparent
    End With
    IsResizing = False
    
    ' Set the vertical positions of the lower half of the form
    UpdateLayout
    
    Me.width = Me.width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - Model"
End Sub

Private Sub UpdateLayout(Optional ChangeY As Double = 0)
' Do the layout of the lower half of the form, changing the height of the list box by ChangeY
    Dim NewHeight As Double
    NewHeight = lstConstraints.height + ChangeY
    If NewHeight < MinHeight Then NewHeight = MinHeight
    
    lstConstraints.height = NewHeight
        
    ' Cascade the updated height
    chkNameRange.top = lstConstraints.top + lstConstraints.height + FormSpacing
    lblDiv4.top = chkNameRange.top + chkNameRange.height + FormSpacing
    lblStep4.top = lblDiv4.top + lblDiv4.height + FormSpacing
    chkGetDuals.top = lblStep4.top
    refDuals.top = lblStep4.top
    chkGetDuals2.top = chkGetDuals.top + chkGetDuals.height
    optUpdate.top = chkGetDuals2.top
    optNew.top = chkGetDuals2.top
    lblDiv5.top = optNew.top + optNew.height
    lblStep5.top = lblDiv5.top + lblDiv5.height + FormSpacing
    cmdChange.top = lblStep5.top
    lblSolver.top = lblStep5.top + FormButtonHeight - FormTextHeight
    lblDiv6.top = cmdChange.top + cmdChange.height + FormSpacing
    chkShowModel.top = lblDiv6.top + lblDiv6.height + FormSpacing
    cmdCancel.top = chkShowModel.top
    cmdBuild.top = chkShowModel.top
    cmdOptions.top = chkShowModel.top
    cmdReset.top = chkShowModel.top
    Me.height = cmdCancel.top + cmdCancel.height + FormMargin + FormTitleHeight
    lblResizer.top = Me.InsideHeight - lblResizer.height
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
