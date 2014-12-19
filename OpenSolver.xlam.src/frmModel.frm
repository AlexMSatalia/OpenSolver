VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmModel 
   Caption         =   "OpenSolver - Model"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9855
   OleObjectBlob   =   "frmModel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------
' OpenSolver
' http://www.opensolver.org
' This software is distributed under the terms of the GNU General Public License
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
'--------------------------------------------------------------------
' FILE DESCRIPTION
' frmModel
' Userform around the functionality in CModel
' Allows user to build and edit models.
'
' Created by:       IRD
'--------------------------------------------------------------------

Option Explicit

' The (formely) only global needed: the handle to the Model instance
Private model As CModel

' Constraint editing mode
Private ListItem As Long
Private ConChangedMode As Boolean
Private DontRepop As Boolean

Private m_clsResizer As CResizer
Public MinHeight As Long

Private OpenedBefore As Boolean
Private ContractedBefore As Boolean

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

Sub Disabler(TrueIfEnable As Boolean, f As UserForm)

4161      f.lblDescHeader.Enabled = TrueIfEnable
4162      f.lblDesc.Enabled = TrueIfEnable
4163      f.cmdRunAutoModel.Enabled = TrueIfEnable
          
4164      f.frameDiv1.Enabled = False
          
4165      f.lblStep1.Enabled = TrueIfEnable
4166      f.refObj.Enabled = TrueIfEnable
4167      f.optMax.Enabled = TrueIfEnable
4168      f.optMin.Enabled = TrueIfEnable
4169      f.optTarget.Enabled = TrueIfEnable
4170      f.txtObjTarget.Enabled = TrueIfEnable And f.optTarget.value
          
4171      f.frameDiv2.Enabled = False
          
4172      f.lblStep2.Enabled = TrueIfEnable
4173      f.refDecision.Enabled = TrueIfEnable
          
4174      f.frameDiv3.Enabled = False
          
4175      f.chkNonNeg.Enabled = TrueIfEnable
4176      f.cmdCancelCon.Enabled = Not TrueIfEnable
4177      f.cmdDelSelCon.Enabled = TrueIfEnable
          f.chkNameRange.Enabled = TrueIfEnable
          
4178      f.frameDiv4.Enabled = False
          
4179      f.lblDuals.Enabled = TrueIfEnable

4180      f.chkGetDuals.Enabled = TrueIfEnable
4181      f.chkGetDuals2.Enabled = TrueIfEnable
4182      f.optUpdate.Enabled = f.chkGetDuals2.value
4183      f.optNew.Enabled = f.chkGetDuals2.value
          
4184      f.refDuals.Enabled = TrueIfEnable And f.chkGetDuals.value And f.chkGetDuals.Enabled
          
4185      f.frameDiv5.Enabled = False
4186      f.frameDiv6.Enabled = False
          
4187      f.chkShowModel.Enabled = TrueIfEnable
4188      f.cmdOptions.Enabled = TrueIfEnable
4189      f.cmdBuild.Enabled = TrueIfEnable
4190      f.cmdCancel.Enabled = TrueIfEnable
          f.cmdReset.Enabled = TrueIfEnable
          f.cmdChange.Enabled = TrueIfEnable
#If Mac Then
4191      MacOptions.chkPerformLinearityCheck.Enabled = True
4192      MacOptions.txtTol.Enabled = True
4193      MacOptions.txtMaxIter.Enabled = True
4194      MacOptions.txtPre.Enabled = True
#Else
4195      frmOptions.chkPerformLinearityCheck.Enabled = True
4196      frmOptions.txtTol.Enabled = True
4197      frmOptions.txtMaxIter.Enabled = True
4198      frmOptions.txtPre.Enabled = True
#End If
        
          Dim Solver As String
4199      If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!OpenSolver_ChosenSolver", Solver) Then
4200          Solver = "CBC"
4201          Call SetNameOnSheet("OpenSolver_ChosenSolver", "=" & Solver)
4202      End If
            
          
4203      If Not SolverHasSensitivityAnalysis(Solver) Then
              ' Disable dual options
4204          f.chkGetDuals2.Enabled = False
4205          f.chkGetDuals.Enabled = False
4206          f.optUpdate.Enabled = False
4207          f.optNew.Enabled = False
4208      End If
          
          '============================================================================================
          'NOTE: Beware that RefEdits cannot be enabled last in this sub as they seem to grab the focus
          '       and create weird errors
          '============================================================================================
End Sub

'--------------------------------------------------------------------
' UpdateFormFromMemory
' Populates the UserForm using the internal model.
'
' Written by:       IRD
'--------------------------------------------------------------------
Sub UpdateFormFromMemory(f As UserForm)
4209      If model.ObjectiveSense = MaximiseObjective Then f.optMax.value = True
4210      If model.ObjectiveSense = MinimiseObjective Then f.optMin.value = True
4211      If model.ObjectiveSense = TargetObjective Then f.optTarget.value = True   ' AJM 20110907
4212      f.txtObjTarget.Text = CStr(model.ObjectiveTarget)   ' AJM 20110907 Always show the target (which may just be 0)
          
4213      f.chkNonNeg.value = model.NonNegativityAssumption
         
4214      If Not model.ObjectiveFunctionCell Is Nothing Then f.refObj.Text = GetDisplayAddress(model.ObjectiveFunctionCell, False)
          
4215      If Not model.DecisionVariables Is Nothing Then f.refDecision.Text = GetDisplayAddressInCurrentLocale(model.DecisionVariables)
          
4216      f.chkGetDuals.value = Not model.Duals Is Nothing
4217      If model.Duals Is Nothing Then
4218          f.refDuals.Text = ""
4219      Else
4220          f.refDuals.Text = GetDisplayAddress(model.Duals, False)
4221      End If
                    
4222      model.PopulateConstraintListBox f.lstConstraints
4223      ModelLstConstraintsChange f

      '          On Error GoTo nameUndefined
      '          f.chkGetDuals2.Value = Mid(Names("'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!OpenSolver_DualsNewSheet").Value, 2)
          Dim sheetName As String, value As String, ResetDualsNewSheet As Boolean
4224      sheetName = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!"
4225      If GetNameValueIfExists(ActiveWorkbook, sheetName & "OpenSolver_DualsNewSheet", value) Then
4226          f.chkGetDuals2.value = value
              ' If checkbox is null, then the stored value was not 'True' or 'False'. We should reset to false
4227          If IsNull(f.chkGetDuals2.value) Then
4228              ResetDualsNewSheet = True
4229          End If
4230      Else
4231          ResetDualsNewSheet = True
4232      End If
          
4233      If ResetDualsNewSheet Then
4234          Call SetNameOnSheet("OpenSolver_DualsNewSheet", "=FALSE")
4235          f.chkGetDuals2.value = False
4236      End If
          
4237      f.optUpdate.Enabled = f.chkGetDuals2.value
4238      f.optNew.Enabled = f.chkGetDuals2.value
4239      If GetNameValueIfExists(ActiveWorkbook, sheetName & "OpenSolver_UpdateSensitivity", value) Then
4240          If value = "TRUE" Then
4241            f.optUpdate.value = value
4242          Else
4243            f.optNew.value = True
4244          End If
4245      Else
4246          Call SetNameOnSheet("OpenSolver_UpdateSensitivity", "=TRUE")
4247          f.optUpdate.value = True
4248      End If
      '          Exit Sub
          
      'nameUndefined:
      '          Call SetNameOnSheet("OpenSolver_DualsNewSheet", "=FALSE")
      '          chkGetDuals2.Value = False 'Mid(Names("'" & Replace(ActiveWorkbook.ActiveSheet.name, "'", "''") & "'!OpenSolver_DualsNewSheet").Value, 2)
          
End Sub

Private Sub chkGetDuals_Click()
4249            frmModel.UpdateGetDuals Me
End Sub

Public Sub UpdateGetDuals(f As UserForm)
4250      f.refDuals.Enabled = f.chkGetDuals.value
End Sub

Private Sub chkGetDuals2_Click()
4251            frmModel.UpdateGetDuals2 Me
End Sub

Public Sub UpdateGetDuals2(f As UserForm)
4252      Call SetNameOnSheet("OpenSolver_DualsNewSheet", "=" & f.chkGetDuals2.value)
4253      f.optUpdate.Enabled = f.chkGetDuals2.value
4254      f.optNew.Enabled = f.chkGetDuals2.value
End Sub

Private Sub chkNameRange_Click()
4255            frmModel.UpdateNameRange Me
End Sub

Public Sub UpdateNameRange(f As UserForm)
          'Call UpdateFormFromMemory
4256      model.PopulateConstraintListBox f.lstConstraints
4257      ModelLstConstraintsChange f
End Sub

Private Sub cmdCancelCon_Click()
4258            frmModel.ModelCancel Me
End Sub

Public Sub ModelCancel(f As UserForm)
4259      frmModel.Disabler True, f
4260      f.cmdAddCon.Enabled = False
4261      ConChangedMode = False
4262      ModelLstConstraintsChange f
End Sub


Private Sub cmdChange_Click()
4263            frmModel.ModelSolverClick Me
End Sub

Public Sub ModelSolverClick(f As UserForm)
#If Mac Then
4264      MacSolverChange.Show
#Else
4265      frmSolverChange.Show vbModal
#End If
End Sub

Private Sub cmdOptions_Click()
4266            frmModel.ModelOptionsClick Me
End Sub

Public Sub ModelOptionsClick(f As UserForm)
          ' Save the current "Assume Non Negative" option so this is shown in the options dialog.
          ' The saved value gets updated on OK, which we then reflect in our Model dialog
          Dim s As String
4267      SetSolverNameOnSheet "neg", IIf(f.chkNonNeg.value, "=1", "=2")
              
#If Mac Then
4268      MacOptions.Show
#Else
4269      frmOptions.Show vbModal
#End If
4270      If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_neg", s) Then    ' This should always be true
4271          f.chkNonNeg.value = s = "1"
4272      End If
End Sub

'--------------------------------------------------------------------------------------
'Reset Button
'Deletes the objective function, decision variables and all the constraints in the model
'---------------------------------------------------------------------------------------

Private Sub cmdReset_Click()
4273            frmModel.ModelReset Me
End Sub

Public Sub ModelReset(f As UserForm)
          Dim NumConstraints As Single, i As Long
                  
          'Check the user wants to reset the model
4274      If MsgBox("This will reset the objective function, decision variables and constraints." _
                  & vbCrLf & " Are you sure you want to do this?", vbYesNo, "OpenSolver") = vbNo Then
4275          Exit Sub
4276      End If

          'Reset the objective function and the decision variables
4277      f.refObj.Text = ""
4278      f.refDecision.Text = ""
              
          'Find the number of constraints in model
4279      NumConstraints = model.Constraints.Count
          
          ' Remove the constraints
4280      For i = 1 To NumConstraints
4281          model.Constraints.Remove 1
4282      Next i

          ' Update constraints form
4283      model.PopulateConstraintListBox f.lstConstraints

End Sub

Private Sub optMax_Click()
4284            frmModel.ModelMaxClick Me
End Sub

Public Sub ModelMaxClick(f As UserForm)
4285      f.txtObjTarget.Enabled = f.optTarget.value
End Sub

Private Sub optMin_Click()
4286            frmModel.ModelMinClick Me
End Sub

Public Sub ModelMinClick(f As UserForm)
4287      txtObjTarget.Enabled = f.optTarget.value
End Sub

Private Sub optNew_Click()
4288            frmModel.ModelNewClick Me
End Sub

Public Sub ModelNewClick(f As UserForm)
4289      Call SetNameOnSheet("OpenSolver_UpdateSensitivity", "=" & f.optUpdate.value)
End Sub

Private Sub optTarget_Click()
4290            frmModel.ModelTargetClick Me
End Sub

Public Sub ModelTargetClick(f As UserForm)
4291      f.txtObjTarget.Enabled = f.optTarget.value
End Sub

Private Sub optUpdate_Click()
4292            frmModel.ModelUpdateClick Me
End Sub

Public Sub ModelUpdateClick(f As UserForm)
4293      Call SetNameOnSheet("OpenSolver_UpdateSensitivity", "=" & f.optUpdate.value)
End Sub

Private Sub refConLHS_Change()
4294            frmModel.ModelChangeLHS Me
End Sub

Public Sub ModelChangeLHS(f As UserForm)
          ' Compare to expected value
4295      If ListItem >= 1 And Not model.Constraints Is Nothing Then
              Dim origLHS As String
              
4296          origLHS = model.Constraints(ListItem).LHS.Address
4297          If f.refConLHS.Text <> origLHS Then
4298              Disabler False, f
4299              f.cmdAddCon.Enabled = True
4300              ConChangedMode = True
4301          Else
4302              Disabler True, f
4303              f.cmdAddCon.Enabled = False
4304              ConChangedMode = False
4305          End If
4306      ElseIf ListItem = 0 Then
4307          If f.refConLHS.Text <> "" Then
4308              Disabler False, f
4309              f.cmdAddCon.Enabled = True
4310              ConChangedMode = True
4311          Else
4312              Disabler True, f
4313              f.cmdAddCon.Enabled = False
4314              ConChangedMode = False
4315          End If
4316      End If
End Sub

Private Sub refConRHS_Change()
4317            frmModel.ModelChangeRHS Me
End Sub

Public Sub ModelChangeRHS(f As UserForm)
          ' Compare to expected value
4318      If ListItem >= 1 And Not model.Constraints Is Nothing Then
              Dim origRHS As String
4319          If model.Constraints(ListItem).RHS Is Nothing Then
4320              origRHS = model.Constraints(ListItem).RHSstring
4321          Else
4322              origRHS = model.Constraints(ListItem).RHS.Address
4323          End If
4324          If f.refConRHS.Text <> origRHS Then
4325              Disabler False, f
4326              f.cmdAddCon.Enabled = True
4327              ConChangedMode = True
4328          Else
4329              Disabler True, f
4330              f.cmdAddCon.Enabled = False
4331              ConChangedMode = False
4332          End If
4333      ElseIf ListItem = 0 Then
4334          If f.refConLHS.Text <> "" Then
4335              Disabler False, f
4336              f.cmdAddCon.Enabled = True
4337              ConChangedMode = True
4338          Else
4339              Disabler True, f
4340              f.cmdAddCon.Enabled = False
4341              ConChangedMode = False
4342          End If
4343      End If
End Sub

'--------------------------------------------------------------------
' UserForm_Activate [event]
' Called when the form is shown.
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub UserForm_Activate()
4344            frmModel.ModelActivate Me
End Sub

Public Sub ModelActivate(f As UserForm)
          ' Check we can even start
4345      If Not CheckWorksheetAvailable Then
4346          Unload Me
4347          Exit Sub
4348      End If
          ' Set any default solver options if none have been set yet
4349      SetAnyMissingDefaultExcel2007SolverOptions
          ' Create a new model object
4350      Set model = New CModel
          ' Hides the current model, if its showing
4351      If SheetHasOpenSolverHighlighting(ActiveSheet) Then HideSolverModel
          ' Make sure sheet is up to date
4352      Application.Calculate
          ' Remove the 'marching ants' showing if a range is copied.
          ' Otherwise, the ants stay visible, and visually conflict with
          ' our cell selection. The ants are also left behind on the
          ' screen. This works around an apparent bug (?) in Excel 2007.
4353      Application.CutCopyMode = False
          ' Clear the form
4354      f.optMax.value = False
4355      f.optMin.value = False
4356      f.refObj.Text = ""
4357      f.refDecision.Text = ""
4358      f.refConLHS.Text = ""
4359      f.refConRHS.Text = ""
4360      f.lstConstraints.Clear
4361      f.cboConRel.Clear
4362      f.cboConRel.AddItem "="
4363      f.cboConRel.AddItem "<="
4364      f.cboConRel.AddItem ">="
4365      f.cboConRel.AddItem "int"
4366      f.cboConRel.AddItem "bin"
4367      f.cboConRel.AddItem "alldiff"
4368      f.cboConRel.ListIndex = cboPosition("=")    ' We set an initial value just in case there is no model, and the user goes straight to AddNewConstraint
          
          'Find current solver
          Dim SolverName As String
4369      If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!OpenSolver_ChosenSolver", SolverName) Then
4370          f.lblSolver.Caption = "Current Solver Engine: " & UCase(left(SolverName, 1)) & Mid(SolverName, 2)
4371      Else: f.lblSolver.Caption = "Current Solver Engine: CBC"
4372      End If
          ' Load the model on the sheet into memory
4373      ListItem = -1
4374      ConChangedMode = False
4375      DontRepop = False
4376      Disabler True, f
4377      model.LoadFromSheet
4378      DoEvents
4379      UpdateFormFromMemory f

            
          MinHeight = 434.25
          If Not OpenedBefore Then
              Set m_clsResizer = New CResizer
              m_clsResizer.Add Me
              f.cmdReset.left = f.cmdReset.left - m_clsResizer.width
              f.cmdOptions.left = f.cmdOptions.left - m_clsResizer.width
              f.cmdBuild.left = f.cmdBuild.left - m_clsResizer.width
              f.cmdCancel.left = f.cmdCancel.left - m_clsResizer.width
          End If
          OpenedBefore = True

4380      DoEvents
          ' Take focus away from refEdits
4381      DoEvents
          'cmdCancel.SetFocus
          'DoEvents
4382      f.Repaint
          'cmdCancel.SetFocus
4383      DoEvents
End Sub


Private Sub cmdCancel_Click()
4384            frmModel.ModelCancelClick Me
4385            Me.Hide
End Sub

Public Sub ModelCancelClick(f As UserForm)
4386      DoEvents
4387      On Error Resume Next ' Just to be safe on our select
4388      Application.CutCopyMode = False
4389      ActiveCell.Select ' Just select one cell, choosing a cell that should be visible to avoid scrolling
4390      Application.ScreenUpdating = True
End Sub


'--------------------------------------------------------------------
' cmdBuild_Click [event]
' Turn the model into a Solver model
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub cmdRunAutoModel_Click()
4391            frmModel.ModelRunAutoModel Me
End Sub

Public Sub ModelRunAutoModel(f As UserForm)
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

4403      UpdateFormFromMemory f
4404      DoEvents
4405      Application.StatusBar = False
End Sub


'--------------------------------------------------------------------
' cmdBuild_Click [event]
' Turn the model into a Solver model
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub cmdBuild_Click()
4406           frmModel.ModelBuild Me
4407           Me.Hide
End Sub

Public Sub ModelBuild(f As UserForm)
4408      DoEvents
4409      On Error Resume Next
4410      Application.CutCopyMode = False
4411      ActiveCell.Select ' Just select one cell, choosing a cell that should be visible to avoid scrolling
4412      Application.ScreenUpdating = True
4413      On Error GoTo 0
          
      '          If (model.DecisionVariables Is Nothing) Or (model.ObjectiveFunctionCell Is Nothing) _
      '                    Or model.Constraints.Count = 0 Then
      '              Me.Hide
      '              Exit Sub
      '          End If
          '----------------------------------------------------------------
          ' Pull possibly update objective info into model
4414      On Error GoTo BadObjRef
      'On Error Resume Next

4415      If Trim(f.refObj.Text) = "" Then
4416          Set model.ObjectiveFunctionCell = Nothing
4417      Else
4418          Set model.ObjectiveFunctionCell = Range(f.refObj.Text)
4419      End If
4420      On Error GoTo errorHandler
          
          ' Get the objective sense
4421      If f.optMax.value = True Then model.ObjectiveSense = MaximiseObjective
4422      If f.optMin.value = True Then model.ObjectiveSense = MinimiseObjective
4423      If f.optTarget.value = True Then
4424          model.ObjectiveSense = TargetObjective
4425          On Error GoTo BadObjectiveTarget
4426          model.ObjectiveTarget = CDbl(f.txtObjTarget.Text)
4427          On Error GoTo errorHandler
4428      End If
4429      If model.ObjectiveSense = UnknownObjectiveSense Then
4430          MsgBox "Error: Please select an objective sense (minimise, maximise or target).", vbExclamation + vbOKOnly, "OpenSolver"
4431          Exit Sub
4432      End If
          
          '----------------------------------------------------------------
          ' Pull possibly updated decision variable info into model ConvertFromCurrentLocale
          ' We allow multiple=area ranges here, which requires
4433      On Error GoTo BadDecRef
      'On Error Resume Next
4434      If Trim(f.refDecision.Text) = "" Then
4435          Set model.DecisionVariables = Nothing
4436      Else
4437          Set model.DecisionVariables = Range(ConvertFromCurrentLocale(f.refDecision.Text))
4438      End If
4439      On Error GoTo errorHandler

          '----------------------------------------------------------------
          ' Pull possibly updated dual storage cells
4440      On Error GoTo BadDualsRef

4441      If f.chkGetDuals.value = False Or Trim(f.refDuals.Text) = "" Then
4442          Set model.Duals = Nothing
4443      Else
4444          Set model.Duals = Range(f.refDuals.Text)
4445      End If
4446      On Error GoTo errorHandler
          
          '----------------------------------------------------------------
          ' Do it
4447      model.NonNegativityAssumption = f.chkNonNeg.value
          
4448      model.BuildModel
          
          
          '----------------------------------------------------------------
          ' Display on screen
4449      If f.chkShowModel.value = True Then OpenSolverVisualizer.ShowSolverModel
4450      On Error GoTo CalculateFailed
4451      Application.Calculate
4452      On Error GoTo errorHandler
          '----------------------------------------------------------------
          ' Finish
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
4458      f.refObj.SetFocus ' Set the focus back to the RefEdit
4459      DoEvents ' Try to stop RefEdit bugs
4460      Exit Sub
          '----------------------------------------------------------------
BadDecRef:
          ' Couldn't turn the decision variable address into a range
4461      MsgBox "Error: the cell address for the decision variables is invalid. " + _
                  "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
4462      f.refDecision.SetFocus ' Set the focus back to the RefEdit
4463      DoEvents ' Try to stop RefEdit bugs
4464      Exit Sub
BadObjectiveTarget:
          ' Couldn't turn the objective target into a value
4465      MsgBox "Error: the target value for the objective cell is invalid. " + _
                  "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
4466      f.txtObjTarget.SetFocus ' Set the focus back to the target text box
4467      DoEvents ' Try to stop RefEdit bugs
4468      Exit Sub
BadDualsRef:
          ' Couldn't turn the dual cell into a range
4469      MsgBox "Error: the cell for storing the shadow prices is invalid. " + _
                  "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
4470      f.refDuals.SetFocus ' Set the focus back to the target text box
4471      DoEvents ' Try to stop RefEdit bugs
4472      Exit Sub
errorHandler:
4473      MsgBox "While constructing the model, OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source & " (frmModel::cmdBuild_Click)", , "OpenSolver Code Error"
4474      DoEvents ' Try to stop RefEdit bugs
4475      Exit Sub
End Sub


'--------------------------------------------------------------------
' cboConRel_Change [event]
' Called whenever the constraint type changes
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub cboConRel_Change()
4476            frmModel.ModelChangeConRel Me
End Sub

Public Sub ModelChangeConRel(f As UserForm)
4477      If f.cboConRel.Text = "=" Then f.refConRHS.Enabled = True
4478      If f.cboConRel.Text = "<=" Then f.refConRHS.Enabled = True
4479      If f.cboConRel.Text = ">=" Then f.refConRHS.Enabled = True
4480      If f.cboConRel.Text = "int" Or f.cboConRel.Text = "bin" Or f.cboConRel.Text = "alldiff" Then
4481          f.refConRHS.Enabled = False
              'f.refConRHS.Text = ""
4482      End If
          
4483      If ListItem >= 1 And Not model.Constraints Is Nothing Then
              Dim origREL As String
              
4484          origREL = model.Constraints(ListItem).ConstraintType
4485          If f.cboConRel.Text <> origREL Then
4486              Disabler False, f
4487              f.cmdAddCon.Enabled = True
4488              ConChangedMode = True
4489          Else
4490              Disabler True, f
4491              f.cmdAddCon.Enabled = False
4492              ConChangedMode = False
4493          End If
4494      End If
End Sub


'--------------------------------------------------------------------
' cmdAddCon_Click [event]
' Add a constraint, assuming it validates
' OR
' Update an existing constraint, again assuming validation
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub cmdAddCon_Click()
4495            frmModel.ModelAddConstraint Me
End Sub
Public Sub ModelAddConstraint(f As UserForm)
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
4497      TestStringForConstraint f.refConLHS.Text, LHSisRange, LHSisFormula, LHSIsValueWithEqual, LHSIsValueWithoutEqual
          
4498      If LHSisRange = False Then
              ' The string in the LHS refedit does not describe a range
4499          MsgBox "Left-hand-side of constraint must be a range."
4500          Exit Sub
4501      End If
4502      If Range(Trim(f.refConLHS.Text)).Areas.Count > 1 Then
              ' The LHS is multiple areas - not allowed
4503          MsgBox "Left-hand-side of constraint must have only one area."
4504          Exit Sub
4505      End If
4506      Set rngLHS = Range(Trim(f.refConLHS.Text))
          
          '----------------------------------------------------------------
          ' RIGHT HAND SIDE
          Dim RHSisRange As Boolean, RHSisFormula As Boolean, RHSIsValueWithEqual As Boolean, RHSIsValueWithoutEqual As Boolean
          Dim strRel As String
4507      strRel = f.cboConRel.Text
4508      If strRel = "" Then ' Should not happen as of 20/9/2011 (AJM)
4509          MsgBox "Please select a relation such as = or <="
4510          Exit Sub
4511      End If
4512      IsRestrict = Not ((strRel = "=") Or (strRel = "<=") Or (strRel = ">="))
4513      If Not IsRestrict Then
4514          If Trim(f.refConRHS.Text) = "" Then
4515              MsgBox "Please enter a right-hand-side!"
4516              Exit Sub
4517          End If
              
4518          TestStringForConstraint f.refConRHS.Text, RHSisRange, RHSisFormula, RHSIsValueWithEqual, RHSIsValueWithoutEqual
              
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
4524              Set rngRHS = Range(Trim(f.refConRHS.Text))
4525              If rngRHS.Count > 1 Then
4526                  If rngRHS.Count <> rngLHS.Count Then
                          ' Mismatch!
4527                      MsgBox "Right-hand-side of constraint has more than one cell, and does not match the number of cells on the left-hand-side."
4528                      Exit Sub
4529                  End If
4530              End If
4531          End If
              
              ' If not a range then evaluate to see if its legit
              ' Evaluate is not locale-friendly
              ' So we put it in a cell on the internal sheet, then get it back
              ' AJM 20/9/2011: We need to prefix the formula with an "=" otherwise formula such as 'sheet name'!A1 get entered as a string constant (becaused of the leading ')
              Dim internalRHS As String
4532          internalRHS = Trim(f.refConRHS.Text)
              
              ' Turn off dialog display; we do not want try to open a workbook with a name of the worksheet! This happens if the formula comes from a worksheet
              ' whose name contains a space
4533          Application.DisplayAlerts = False
4534          On Error GoTo ErrorHandler_CannotInterpretRHS
4535          OpenSolverSheet.Range("A1").FormulaLocal = IIf(left(internalRHS, 1) = "=", "", "=") & f.refConRHS.Text
4536          internalRHS = OpenSolverSheet.Range("A1").Formula
4537          OpenSolverSheet.Range("A1").Clear ' This must be blank to ensure no risk of dialogs being shown trying to locate a sheet
4538          Application.DisplayAlerts = True
              
4539          If Not RHSisRange Then
                  ' Can we evaluate this function or constant?
                  Dim varReturn As Variant
4540              varReturn = ActiveSheet.Evaluate(internalRHS) ' Must be worksheet.evaluate to get references to names local to the sheet
4541              If VBA.VarType(varReturn) = vbError Then
4542                  MsgBox "The formula or value for the RHS is not valid. Please check and try again."
4543                  f.refConRHS.SetFocus
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
4555              f.refConRHS.Text = Mid(varReturn, 2, Len(varReturn))
4556          End If
              
          
4557      End If
          
4558      Disabler True, f
4559      f.cmdAddCon.Enabled = False
4560      ConChangedMode = False
              
          '================================================================
          ' Update constraint?
4561      If f.cmdAddCon.Caption <> "Add constraint" Then
          
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
4576                          .RHSstring = "=" + f.refConRHS.Text
4577                      Else
4578                          .RHSstring = f.refConRHS.Text
4579                      End If
4580                  End If
4581              End If
4582          End With

4583          If Not DontRepop Then model.PopulateConstraintListBox f.lstConstraints
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
4600                          .RHSstring = "=" + f.refConRHS.Text ' This has been converted above into a US-locale formula, value or reference
4601                      Else
4602                          .RHSstring = f.refConRHS.Text ' This has been converted above into a US-locale formula, value or reference
4603                      End If
4604                  End If
4605              End If
4606          End With
              
4607          model.Constraints.Add NewConstraint ', NewConstraint.GetKey
4608          If Not DontRepop Then model.PopulateConstraintListBox f.lstConstraints
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
4617      f.refConRHS.SetFocus
4618      DoEvents ' Try to stop RefEdit bugs
4619      Exit Sub
errorHandler:
4620      Application.DisplayAlerts = True
4621      OpenSolverSheet.Range("A1").FormulaLocal = "" ' This must be blank to ensure no risk of dialogs being shown trying to locate a sheet
4622      MsgBox "While constructing the model, OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source & " (frmModel::cmdBuild_Click)", , "OpenSolver Code Error"
4623      DoEvents ' Try to stop RefEdit bugs
4624      Exit Sub

End Sub


'--------------------------------------------------------------------
' cmdDelSelCon_Click [event]
' Delete selected constraint
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub cmdDelSelCon_Click()
4625           frmModel.ModelDeleteConstraint Me
End Sub
Public Sub ModelDeleteConstraint(f As UserForm)

4626      If f.lstConstraints.ListIndex = -1 Then
4627          MsgBox "Please select a constraint!", vbOKOnly, "AutoModel"
4628          Exit Sub
4629      End If
            
          ' Remove it
4630      model.Constraints.Remove f.lstConstraints.ListIndex
          
          ' Update form
4631      model.PopulateConstraintListBox f.lstConstraints
End Sub


'--------------------------------------------------------------------
' lstConstraints_Change [event]
' Selection in constraints box changes
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub lstConstraints_Change()
4632      frmModel.ModelLstConstraintsChange Me
End Sub

Public Sub ModelLstConstraintsChange(f As UserForm)
          
4633      If ConChangedMode = True Then
4634          If f.cmdAddCon.Caption = "Update constraint" Then
4635              If MsgBox("You have made changes to the current constraint." _
                      + vbNewLine + "Do you want to save these changes?", vbYesNo) = vbYes Then
          
                      'Debug.Print "Doing cmdAddCon_Click"
4636                  DontRepop = True
4637                  f.cmdAddCon_Click
4638                  DontRepop = False
                      ''Debug.Print "Done."
4639              End If
4640          Else
4641              If MsgBox("You have entered a constraint." _
                      + vbNewLine + "Do you want to save this as a new constraint?", vbYesNo) = vbYes Then
          
                      ''Debug.Print "Doing cmdAddCon_Click"
4642                  DontRepop = True
4643                  f.cmdAddCon_Click
4644                  DontRepop = False
                      ''Debug.Print "Done."
4645              End If
4646          End If
          
4647          Disabler True, f
4648          f.cmdAddCon.Enabled = False
4649          ConChangedMode = False
              'ListItem = f.lstConstraints.ListIndex
4650          model.PopulateConstraintListBox f.lstConstraints
              'f.lstConstraints.ListIndex = ListItem
4651      End If
          
4652      ListItem = f.lstConstraints.ListIndex
          
4653      If f.lstConstraints.ListIndex = -1 Then
4654          Exit Sub
4655      End If
4656      If f.lstConstraints.ListIndex = 0 Then
              'Add constraint
4657          f.refConLHS.Enabled = True
4658          DoEvents
4659          f.refConLHS.Text = ""
4660          DoEvents
              'refConRHS.Enabled = True
              'DoEvents
4661          f.refConRHS.Text = ""
4662          DoEvents
4663          ModelChangeConRel f ' AJM: Force the RHS to be active only if the current relation is =, < or >, which is set based on the last constraint
4664          f.cmdAddCon.Enabled = True
4665          f.cmdAddCon.Caption = "Add constraint"
4666          f.cmdDelSelCon.Enabled = False
4667          DoEvents
4668          Application.CutCopyMode = False
4669          ActiveCell.Select ' Just select one cell, choosing a cell that should be visible to avoid scrolling
4670          Application.ScreenUpdating = True
              'refConLHS.select
4671          Exit Sub
4672      Else
              ' Update constraint
4673          f.refConLHS.Enabled = True
4674          DoEvents
4675          f.refConRHS.Enabled = True
4676          DoEvents
4677          f.cmdAddCon.Enabled = False
4678          f.cmdAddCon.Caption = "Update constraint"
4679          f.cmdDelSelCon.Enabled = True
4680          DoEvents

              ' Set text and marching ants; the select/copy operations can throw errors if on different sheets
4681          On Error Resume Next
4682          ActiveCell.Select   ' We may fail in the next steps, se we cancel any old highlighting
4683          Application.CutCopyMode = False
              Dim copyRange As Range
4684          With model.Constraints(f.lstConstraints.ListIndex)
4685              f.refConLHS.Text = GetDisplayAddress(.LHS, False)
4686              Set copyRange = .LHS
4687              f.cboConRel.ListIndex = cboPosition(.ConstraintType)
4688              f.refConRHS.Text = ""
4689              If Not .RHS Is Nothing Then
4690                  f.refConRHS.Text = GetDisplayAddress(.RHS, False)
4691                  Set copyRange = ProperUnion(copyRange, .RHS)
4692              ElseIf .RHS Is Nothing And .RHSstring <> "" Then
4693                  If Mid(.RHSstring, 1, 1) = "=" Then
4694                      f.refConRHS.Text = RemoveActiveSheetNameFromString(Mid(.RHSstring, 2, Len(.RHSstring)))
4695                  Else
4696                      f.refConRHS.Text = RemoveActiveSheetNameFromString(.RHSstring)
4697                  End If
4698              End If
4699          End With
4700          ModelChangeConRel f
              
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


'--------------------------------------------------------------------
' TestStringForConstraint
' Based on GetNameAsValueOrRange
' Returns information on the kind of value string is, and whether
' it is suitable for a constraint
'
' Written by:       IRD
'--------------------------------------------------------------------
Sub TestStringForConstraint(ByVal TheString As String, _
                            RefersToRange As Boolean, _
                            RefersToFormula As Boolean, _
                            RefersToValueWithEqual As Boolean, _
                            RefersToValueWithoutEqual As Boolean)
4707      TheString = Trim(TheString) ' AJM: Remove any leading/trailing spaces
          
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

Public Sub MoveItems(ChangeY As Single)
        chkNameRange.top = chkNameRange.top + ChangeY
        frameDiv4.top = frameDiv4.top + ChangeY
        lblDuals.top = lblDuals.top + ChangeY
        chkGetDuals.top = chkGetDuals.top + ChangeY
        chkGetDuals2.top = chkGetDuals2.top + ChangeY
        optUpdate.top = optUpdate.top + ChangeY
        refDuals.top = refDuals.top + ChangeY
        optNew.top = optNew.top + ChangeY
        frameDiv6.top = frameDiv6.top + ChangeY
        Label5.top = Label5.top + ChangeY
        lblSolver.top = lblSolver.top + ChangeY
        cmdChange.top = cmdChange.top + ChangeY
        Frame3.top = Frame3.top + ChangeY
        chkShowModel.top = chkShowModel.top + ChangeY
        cmdReset.top = cmdReset.top + ChangeY
        cmdOptions.top = cmdOptions.top + ChangeY
        cmdBuild.top = cmdBuild.top + ChangeY
        cmdCancel.top = cmdCancel.top + ChangeY
        
        If Me.height + ChangeY >= MinHeight Then
            lstConstraints.height = Me.height - 294
            lblDesc.Caption = "AutoModel is a feature of OpenSolver that tries to automatically " _
                            & "determine the problem you are trying to optimise by observing the " _
                            & "structure of the spreadsheet. It will turn its best guess into a " _
                            & "Solver model, which you can then edit in this window."
            lblDesc.height = 24
            
            If ContractedBefore Then
                frameDiv1.top = 57
                lblStep1.top = 64
                refObj.top = 64
                optMax.top = 64
                Label2.top = 64
                optMin.top = 64
                Label3.top = 64
                optTarget.top = 64
                Label4.top = 64
                txtObjTarget.top = 64
                frameDiv2.top = 85.05
                lblStep2.top = 94
                refDecision.top = 94
                frameDiv3.top = 136
                lblStep3.top = 142
                lstConstraints.top = 160
                Label1.top = 159.95
                refConLHS.top = 166
                cboConRel.top = 166
                refConRHS.top = 189.95
                cmdAddCon.top = 213.95
                cmdCancelCon.top = 213.95
                cmdDelSelCon.top = 244
                chkNonNeg.top = 268
                ContractedBefore = False
            End If
        Else
            ContractedBefore = True
            lblDesc.Caption = ""
            lblDesc.height = 0
            frameDiv1.top = frameDiv1.top + ChangeY
            lblStep1.top = lblStep1.top + ChangeY
            refObj.top = refObj.top + ChangeY
            optMax.top = optMax.top + ChangeY
            Label2.top = Label2.top + ChangeY
            optMin.top = optMin.top + ChangeY
            Label3.top = Label3.top + ChangeY
            optTarget.top = optTarget.top + ChangeY
            Label4.top = Label4.top + ChangeY
            txtObjTarget.top = txtObjTarget.top + ChangeY
            frameDiv2.top = frameDiv2.top + ChangeY
            lblStep2.top = lblStep2.top + ChangeY
            refDecision.top = refDecision.top + ChangeY
            frameDiv3.top = frameDiv3.top + ChangeY
            lblStep3.top = lblStep3.top + ChangeY
            lstConstraints.top = lstConstraints.top + ChangeY
            Label1.top = Label1.top + ChangeY
            refConLHS.top = refConLHS.top + ChangeY
            cboConRel.top = cboConRel.top + ChangeY
            refConRHS.top = refConRHS.top + ChangeY
            cmdAddCon.top = cmdAddCon.top + ChangeY
            cmdCancelCon.top = cmdCancelCon.top + ChangeY
            cmdDelSelCon.top = cmdDelSelCon.top + ChangeY
            chkNonNeg.top = chkNonNeg.top + ChangeY
        End If
End Sub


