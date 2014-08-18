VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmModel 
   Caption         =   "OpenSolver - Model"
   ClientHeight    =   8265
   ClientLeft      =   42
   ClientTop       =   378
   ClientWidth     =   9856
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

' Function to map string rels to combobox index positions
' Assigning combobox by .value fails a lot on Mac
Function cboPosition(rel As String) As Integer
    Select Case rel
    Case "="
        cboPosition = 0
    Case "<="
        cboPosition = 1
    Case ">="
        cboPosition = 2
    Case "int"
        cboPosition = 3
    Case "bin"
        cboPosition = 4
    Case "alldiff"
        cboPosition = 5
    End Select
End Function

Sub Disabler(TrueIfEnable As Boolean, f As UserForm)

42390     f.lblDescHeader.Enabled = TrueIfEnable
42400     f.lblDesc.Enabled = TrueIfEnable
42410     f.cmdRunAutoModel.Enabled = TrueIfEnable
          
42420     f.frameDiv1.Enabled = False
          
42430     f.lblStep1.Enabled = TrueIfEnable
42440     f.refObj.Enabled = TrueIfEnable
42450     f.optMax.Enabled = TrueIfEnable
42460     f.optMin.Enabled = TrueIfEnable
42470     f.optTarget.Enabled = TrueIfEnable
42480     f.txtObjTarget.Enabled = TrueIfEnable And f.optTarget.value
          
42490     f.frameDiv2.Enabled = False
          
42500     f.lblStep2.Enabled = TrueIfEnable
42510     f.refDecision.Enabled = TrueIfEnable
          
42520     f.frameDiv3.Enabled = False
          
42530     f.chkNonNeg.Enabled = TrueIfEnable
42540     f.cmdCancelCon.Enabled = Not TrueIfEnable
42550     f.cmdDelSelCon.Enabled = TrueIfEnable
          
42560     f.frameDiv4.Enabled = False
          
42570     f.lblDuals.Enabled = TrueIfEnable

42580     f.chkGetDuals.Enabled = TrueIfEnable
42590     f.chkGetDuals2.Enabled = TrueIfEnable
42600     f.optUpdate.Enabled = f.chkGetDuals2.value
42610     f.optNew.Enabled = f.chkGetDuals2.value
          
42620     f.refDuals.Enabled = TrueIfEnable And f.chkGetDuals.value And f.chkGetDuals.Enabled
          
42630     f.frameDiv5.Enabled = False
42640     f.frameDiv6.Enabled = False
          
42660     f.chkShowModel.Enabled = TrueIfEnable
42670     f.cmdOptions.Enabled = TrueIfEnable
42680     f.cmdBuild.Enabled = TrueIfEnable
42690     f.cmdCancel.Enabled = TrueIfEnable
#If Mac Then
          MacOptions.chkLinear.Enabled = True
          MacOptions.chkPerformLinearityCheck.Enabled = True
          MacOptions.txtTol.Enabled = True
          MacOptions.txtMaxIter.Enabled = True
          MacOptions.txtPre.Enabled = True
#Else
42700     frmOptions.chkLinear.Enabled = True
42710     frmOptions.chkPerformLinearityCheck.Enabled = True
42720     frmOptions.txtTol.Enabled = True
42730     frmOptions.txtMaxIter.Enabled = True
42740     frmOptions.txtPre.Enabled = True
#End If
        
          Dim Solver As String
42750     If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!OpenSolver_ChosenSolver", Solver) Then
              Solver = "CBC"
              Call SetNameOnSheet("OpenSolver_ChosenSolver", "=" & Solver)
42770     End If
            
          
          If Not SolverHasSensitivityAnalysis(Solver) Then
42780         ' Disable dual options
              f.chkGetDuals2.Enabled = False
42790         f.chkGetDuals.Enabled = False
42800         f.optUpdate.Enabled = False
42810         f.optNew.Enabled = False
          End If
          
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
42890     If model.ObjectiveSense = MaximiseObjective Then f.optMax.value = True
42900     If model.ObjectiveSense = MinimiseObjective Then f.optMin.value = True
42910     If model.ObjectiveSense = TargetObjective Then f.optTarget.value = True   ' AJM 20110907
42920     f.txtObjTarget.Text = CStr(model.ObjectiveTarget)   ' AJM 20110907 Always show the target (which may just be 0)
          
42930     f.chkNonNeg.value = model.NonNegativityAssumption
         
42940     If Not model.ObjectiveFunctionCell Is Nothing Then f.refObj.Text = GetDisplayAddress(model.ObjectiveFunctionCell, False)
          
42950     If Not model.DecisionVariables Is Nothing Then f.refDecision.Text = GetDisplayAddressInCurrentLocale(model.DecisionVariables)
          
42960     f.chkGetDuals.value = Not model.Duals Is Nothing
42970     If model.Duals Is Nothing Then
42980         f.refDuals.Text = ""
42990     Else
43000         f.refDuals.Text = GetDisplayAddress(model.Duals, False)
43010     End If
                    
43020     model.PopulateConstraintListBox f.lstConstraints
43030     ModelLstConstraintsChange f

      '          On Error GoTo nameUndefined
      '          f.chkGetDuals2.Value = Mid(Names("'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!OpenSolver_DualsNewSheet").Value, 2)
          Dim sheetName As String, value As String, ResetDualsNewSheet As Boolean
43040     sheetName = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!"
43050     If GetNameValueIfExists(ActiveWorkbook, sheetName & "OpenSolver_DualsNewSheet", value) Then
43060         f.chkGetDuals2.value = value
              ' If checkbox is null, then the stored value was not 'True' or 'False'. We should reset to false
              If IsNull(f.chkGetDuals2.value) Then
                  ResetDualsNewSheet = True
              End If
43070     Else
43080         ResetDualsNewSheet = True
43100     End If
          
          If ResetDualsNewSheet Then
              Call SetNameOnSheet("OpenSolver_DualsNewSheet", "=FALSE")
              f.chkGetDuals2.value = False
          End If
          
43110     f.optUpdate.Enabled = f.chkGetDuals2.value
43120     f.optNew.Enabled = f.chkGetDuals2.value
43130     If GetNameValueIfExists(ActiveWorkbook, sheetName & "OpenSolver_UpdateSensitivity", value) Then
43140         If value = "TRUE" Then
43150           f.optUpdate.value = value
43160         Else
43170           f.optNew.value = True
43180         End If
43190     Else
43200         Call SetNameOnSheet("OpenSolver_UpdateSensitivity", "=TRUE")
43210         f.optUpdate.value = True
43220     End If
      '          Exit Sub
          
      'nameUndefined:
      '          Call SetNameOnSheet("OpenSolver_DualsNewSheet", "=FALSE")
      '          chkGetDuals2.Value = False 'Mid(Names("'" & Replace(ActiveWorkbook.ActiveSheet.name, "'", "''") & "'!OpenSolver_DualsNewSheet").Value, 2)
          
End Sub

Private Sub chkGetDuals_Click()
          frmModel.UpdateGetDuals Me
End Sub

Public Sub UpdateGetDuals(f As UserForm)
43230     f.refDuals.Enabled = f.chkGetDuals.value
End Sub

Private Sub chkGetDuals2_Click()
          frmModel.UpdateGetDuals2 Me
End Sub

Public Sub UpdateGetDuals2(f As UserForm)
43240     Call SetNameOnSheet("OpenSolver_DualsNewSheet", "=" & f.chkGetDuals2.value)
43250     f.optUpdate.Enabled = f.chkGetDuals2.value
43260     f.optNew.Enabled = f.chkGetDuals2.value
End Sub

Private Sub chkNameRange_Click()
          frmModel.UpdateNameRange Me
End Sub

Public Sub UpdateNameRange(f As UserForm)
          'Call UpdateFormFromMemory
          model.PopulateConstraintListBox f.lstConstraints
43280     ModelLstConstraintsChange f
End Sub

Private Sub cmdCancelCon_Click()
          frmModel.ModelCancel Me
End Sub

Public Sub ModelCancel(f As UserForm)
43290     frmModel.Disabler True, f
43300     f.cmdAddCon.Enabled = False
43310     ConChangedMode = False
43320     ModelLstConstraintsChange f
End Sub


Private Sub cmdChange_Click()
          frmModel.ModelSolverClick Me
End Sub

Public Sub ModelSolverClick(f As UserForm)
#If Mac Then
          MacSolverChange.Show
#Else
43330     frmSolverChange.Show vbModal
#End If
End Sub

Private Sub cmdOptions_Click()
          frmModel.ModelOptionsClick Me
End Sub

Public Sub ModelOptionsClick(f As UserForm)
          ' Save the current "Assume Non Negative" option so this is shown in the options dialog.
          ' The saved value gets updated on OK, which we then reflect in our Model dialog
          Dim s As String
43340     SetSolverNameOnSheet "neg", IIf(f.chkNonNeg.value, "=1", "=2")
              
#If Mac Then
          MacOptions.Show
#Else
43350     frmOptions.Show vbModal
#End If
43360     If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_neg", s) Then    ' This should always be true
43370         f.chkNonNeg.value = s = "1"
43380     End If
End Sub

'--------------------------------------------------------------------------------------
'Reset Button
'Deletes the objective function, decision variables and all the constraints in the model
'---------------------------------------------------------------------------------------

Private Sub cmdReset_Click()
          frmModel.ModelReset Me
End Sub

Public Sub ModelReset(f As UserForm)
          Dim NumConstraints As Single, i As Long
                  
          'Check the user wants to reset the model
43390     If MsgBox("This will reset the objective function, decision variables and constraints." _
                  & vbCrLf & " Are you sure you want to do this?", vbYesNo, "OpenSolver") = vbNo Then
43400         Exit Sub
43410     End If

          'Reset the objective function and the decision variables
43420     f.refObj.Text = ""
43430     f.refDecision.Text = ""
              
          'Find the number of constraints in model
43440     NumConstraints = model.Constraints.Count
          
          ' Remove the constraints
43450     For i = 1 To NumConstraints
43460         model.Constraints.Remove 1
43470     Next i

          ' Update constraints form
43480     model.PopulateConstraintListBox f.lstConstraints

End Sub

Private Sub optMax_Click()
          frmModel.ModelMaxClick Me
End Sub

Public Sub ModelMaxClick(f As UserForm)
43490     f.txtObjTarget.Enabled = f.optTarget.value
End Sub

Private Sub optMin_Click()
          frmModel.ModelMinClick Me
End Sub

Public Sub ModelMinClick(f As UserForm)
43500     txtObjTarget.Enabled = f.optTarget.value
End Sub

Private Sub optNew_Click()
          frmModel.ModelNewClick Me
End Sub

Public Sub ModelNewClick(f As UserForm)
43510     Call SetNameOnSheet("OpenSolver_UpdateSensitivity", "=" & f.optUpdate.value)
End Sub

Private Sub optTarget_Click()
          frmModel.ModelTargetClick Me
End Sub

Public Sub ModelTargetClick(f As UserForm)
43520     f.txtObjTarget.Enabled = f.optTarget.value
End Sub

Private Sub optUpdate_Click()
          frmModel.ModelUpdateClick Me
End Sub

Public Sub ModelUpdateClick(f As UserForm)
43530     Call SetNameOnSheet("OpenSolver_UpdateSensitivity", "=" & f.optUpdate.value)
End Sub

Private Sub refConLHS_Change()
          frmModel.ModelChangeLHS Me
End Sub

Public Sub ModelChangeLHS(f As UserForm)
          ' Compare to expected value
43540     If ListItem >= 1 And Not model.Constraints Is Nothing Then
              Dim origLHS As String
              
43550         origLHS = model.Constraints(ListItem).LHS.Address
43560         If f.refConLHS.Text <> origLHS Then
43570             Disabler False, f
43580             f.cmdAddCon.Enabled = True
43590             ConChangedMode = True
43600         Else
43610             Disabler True, f
43620             f.cmdAddCon.Enabled = False
43630             ConChangedMode = False
43640         End If
43650     ElseIf ListItem = 0 Then
43660         If f.refConLHS.Text <> "" Then
43670             Disabler False, f
43680             f.cmdAddCon.Enabled = True
43690             ConChangedMode = True
43700         Else
43710             Disabler True, f
43720             f.cmdAddCon.Enabled = False
43730             ConChangedMode = False
43740         End If
43750     End If
End Sub

Private Sub refConRHS_Change()
          frmModel.ModelChangeRHS Me
End Sub

Public Sub ModelChangeRHS(f As UserForm)
          ' Compare to expected value
43760     If ListItem >= 1 And Not model.Constraints Is Nothing Then
              Dim origRHS As String
43770         If model.Constraints(ListItem).RHS Is Nothing Then
43780             origRHS = model.Constraints(ListItem).RHSstring
43790         Else
43800             origRHS = model.Constraints(ListItem).RHS.Address
43810         End If
43820         If f.refConRHS.Text <> origRHS Then
43830             Disabler False, f
43840             f.cmdAddCon.Enabled = True
43850             ConChangedMode = True
43860         Else
43870             Disabler True, f
43880             f.cmdAddCon.Enabled = False
43890             ConChangedMode = False
43900         End If
43910     ElseIf ListItem = 0 Then
43920         If f.refConLHS.Text <> "" Then
43930             Disabler False, f
43940             f.cmdAddCon.Enabled = True
43950             ConChangedMode = True
43960         Else
43970             Disabler True, f
43980             f.cmdAddCon.Enabled = False
43990             ConChangedMode = False
44000         End If
44010     End If
End Sub

'--------------------------------------------------------------------
' UserForm_Activate [event]
' Called when the form is shown.
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub UserForm_Activate()
          frmModel.ModelActivate Me
End Sub

Public Sub ModelActivate(f As UserForm)
          ' Check we can even start
44020     If Not CheckWorksheetAvailable Then
44030         Unload Me
44040         Exit Sub
44050     End If
          ' Set any default solver options if none have been set yet
44060     SetAnyMissingDefaultExcel2007SolverOptions
          ' Create a new model object
44070     Set model = New CModel
          ' Hides the current model, if its showing
44080     If SheetHasOpenSolverHighlighting(ActiveSheet) Then HideSolverModel
          ' Make sure sheet is up to date
44090     Application.Calculate
          ' Remove the 'marching ants' showing if a range is copied.
          ' Otherwise, the ants stay visible, and visually conflict with
          ' our cell selection. The ants are also left behind on the
          ' screen. This works around an apparent bug (?) in Excel 2007.
44100     Application.CutCopyMode = False
          ' Clear the form
44110     f.optMax.value = False
44120     f.optMin.value = False
44130     f.refObj.Text = ""
44140     f.refDecision.Text = ""
44150     f.refConLHS.Text = ""
44160     f.refConRHS.Text = ""
44170     f.lstConstraints.Clear
44180     f.cboConRel.Clear
44190     f.cboConRel.AddItem "="
44200     f.cboConRel.AddItem "<="
44210     f.cboConRel.AddItem ">="
44220     f.cboConRel.AddItem "int"
44230     f.cboConRel.AddItem "bin"
44240     f.cboConRel.AddItem "alldiff"
44250     f.cboConRel.ListIndex = cboPosition("=")    ' We set an initial value just in case there is no model, and the user goes straight to AddNewConstraint
          
          'Find current solver
          Dim solverName As String
44260     If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!OpenSolver_ChosenSolver", solverName) Then
44270         f.lblSolver.Caption = "Current Solver Engine: " & UCase(left(solverName, 1)) & Mid(solverName, 2)
44280     Else: f.lblSolver.Caption = "Current Solver Engine: CBC"
44290     End If
          ' Load the model on the sheet into memory
44300     ListItem = -1
44310     ConChangedMode = False
44320     DontRepop = False
44330     Disabler True, f
44340     model.LoadFromSheet
44350     DoEvents
44360     UpdateFormFromMemory f
44370     DoEvents
          ' Take focus away from refEdits
44380     DoEvents
          'cmdCancel.SetFocus
          'DoEvents
44390     f.Repaint
          'cmdCancel.SetFocus
44400     DoEvents
End Sub


Private Sub cmdCancel_Click()
          frmModel.ModelCancelClick Me
          Me.Hide
End Sub

Public Sub ModelCancelClick(f As UserForm)
44410     DoEvents
44420     On Error Resume Next ' Just to be safe on our select
44430     Application.CutCopyMode = False
44440     ActiveCell.Select ' Just select one cell, choosing a cell that should be visible to avoid scrolling
44450     Application.ScreenUpdating = True
End Sub


'--------------------------------------------------------------------
' cmdBuild_Click [event]
' Turn the model into a Solver model
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub cmdRunAutoModel_Click()
          frmModel.ModelRunAutoModel Me
End Sub

Public Sub ModelRunAutoModel(f As UserForm)
          ' Try and guess the objective
          Dim status As String
44470     status = model.FindObjective(ActiveSheet)

          ' Get it in memory
44480     Load frmAutoModel
          ' Pass it the model reference
44490     Set frmAutoModel.model = model
44500     frmAutoModel.GuessObjStatus = status
          
44510     Select Case status
              Case "NoSense", "SenseNoCell"
#If Mac Then
                  'MacAutoModel.Show
                  ' Mac can't use a RefEdit properly when multiple forms are open since it doesn't support modeless forms properly
                  MsgBox ("Couldn't find objective cell, and couldn't finish as a result.")
#Else
44520             frmAutoModel.Show vbModal
#End If
44550         Case Else ' Found objective
44560             model.FindVarsAndCons IsFirstTime:=True
44570     End Select

          ' Force the automatically created model to be a linear one, and turn on AssumeNonNegative
44580     model.NonNegativityAssumption = True
44590     SetSolverNameOnSheet "lin", "=1"
          Dim s As String
44600     If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_eng", s) Then SetSolverNameOnSheet "eng", "=2"
                      ' Set this for Solver 2010 models, but only if this name is already defined

44610     UpdateFormFromMemory f
44620     DoEvents
44630     Application.StatusBar = False
End Sub


'--------------------------------------------------------------------
' cmdBuild_Click [event]
' Turn the model into a Solver model
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub cmdBuild_Click()
         frmModel.ModelBuild Me
         Me.Hide
End Sub

Public Sub ModelBuild(f As UserForm)
44640     DoEvents
44650     On Error Resume Next
44660     Application.CutCopyMode = False
44670     ActiveCell.Select ' Just select one cell, choosing a cell that should be visible to avoid scrolling
44680     Application.ScreenUpdating = True
44690     On Error GoTo 0
          
      '          If (model.DecisionVariables Is Nothing) Or (model.ObjectiveFunctionCell Is Nothing) _
      '                    Or model.Constraints.Count = 0 Then
      '              Me.Hide
      '              Exit Sub
      '          End If
          '----------------------------------------------------------------
          ' Pull possibly update objective info into model
44700     On Error GoTo BadObjRef
      'On Error Resume Next

44710     If Trim(f.refObj.Text) = "" Then
44720         Set model.ObjectiveFunctionCell = Nothing
44730     Else
44740         Set model.ObjectiveFunctionCell = Range(f.refObj.Text)
44750     End If
44760     On Error GoTo errorHandler
          
          ' Get the objective sense
44770     If f.optMax.value = True Then model.ObjectiveSense = MaximiseObjective
44780     If f.optMin.value = True Then model.ObjectiveSense = MinimiseObjective
44790     If f.optTarget.value = True Then
44800         model.ObjectiveSense = TargetObjective
44810         On Error GoTo BadObjectiveTarget
44820         model.ObjectiveTarget = CDbl(f.txtObjTarget.Text)
44830         On Error GoTo errorHandler
44840     End If
44850     If model.ObjectiveSense = UnknownObjectiveSense Then
44860         MsgBox "Error: Please select an objective sense (minimise, maximise or target).", vbExclamation + vbOKOnly, "OpenSolver"
44870         Exit Sub
44880     End If
          
          '----------------------------------------------------------------
          ' Pull possibly updated decision variable info into model ConvertFromCurrentLocale
          ' We allow multiple=area ranges here, which requires
44890     On Error GoTo BadDecRef
      'On Error Resume Next
44900     If Trim(f.refDecision.Text) = "" Then
44910         Set model.DecisionVariables = Nothing
44920     Else
44930         Set model.DecisionVariables = Range(ConvertFromCurrentLocale(f.refDecision.Text))
44940     End If
44950     On Error GoTo errorHandler

          '----------------------------------------------------------------
          ' Pull possibly updated dual storage cells
44960     On Error GoTo BadDualsRef

44970     If f.chkGetDuals.value = False Or Trim(f.refDuals.Text) = "" Then
44980         Set model.Duals = Nothing
44990     Else
45000         Set model.Duals = Range(f.refDuals.Text)
45010     End If
45020     On Error GoTo errorHandler
          
          '----------------------------------------------------------------
          ' Do it
45030     model.NonNegativityAssumption = f.chkNonNeg.value
          
45040     model.BuildModel
          
          
          '----------------------------------------------------------------
          ' Display on screen
45050     If f.chkShowModel.value = True Then OpenSolverVisualizer.ShowSolverModel
          On Error GoTo CalculateFailed
45060     Application.Calculate
          On Error GoTo errorHandler
          '----------------------------------------------------------------
          ' Finish
45080     Exit Sub

          '----------------------------------------------------------------
CalculateFailed:
          ' Application.Calculate failed. Ignore error and try again
          On Error GoTo errorHandler
          Application.Calculate
          Resume Next
          
          '----------------------------------------------------------------
BadObjRef:
          ' Couldn't turn the objective cell address into a range
45090     MsgBox "Error: the cell address for the objective is invalid. " + _
                  "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
45100     f.refObj.SetFocus ' Set the focus back to the RefEdit
45110     DoEvents ' Try to stop RefEdit bugs
45120     Exit Sub
          '----------------------------------------------------------------
BadDecRef:
          ' Couldn't turn the decision variable address into a range
45130     MsgBox "Error: the cell address for the decision variables is invalid. " + _
                  "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
45140     f.refDecision.SetFocus ' Set the focus back to the RefEdit
45150     DoEvents ' Try to stop RefEdit bugs
45160     Exit Sub
BadObjectiveTarget:
          ' Couldn't turn the objective target into a value
45170     MsgBox "Error: the target value for the objective cell is invalid. " + _
                  "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
45180     f.txtObjTarget.SetFocus ' Set the focus back to the target text box
45190     DoEvents ' Try to stop RefEdit bugs
45200     Exit Sub
BadDualsRef:
          ' Couldn't turn the dual cell into a range
45210     MsgBox "Error: the cell for storing the shadow prices is invalid. " + _
                  "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
45220     f.refDuals.SetFocus ' Set the focus back to the target text box
45230     DoEvents ' Try to stop RefEdit bugs
45240     Exit Sub
errorHandler:
45250     MsgBox "While constructing the model, OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source & " (frmModel::cmdBuild_Click)", , "OpenSolver Code Error"
45260     DoEvents ' Try to stop RefEdit bugs
45270     Exit Sub
End Sub


'--------------------------------------------------------------------
' cboConRel_Change [event]
' Called whenever the constraint type changes
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub cboConRel_Change()
          frmModel.ModelChangeConRel Me
End Sub

Public Sub ModelChangeConRel(f As UserForm)
45280     If f.cboConRel.Text = "=" Then f.refConRHS.Enabled = True
45290     If f.cboConRel.Text = "<=" Then f.refConRHS.Enabled = True
45300     If f.cboConRel.Text = ">=" Then f.refConRHS.Enabled = True
45310     If f.cboConRel.Text = "int" Or f.cboConRel.Text = "bin" Or f.cboConRel.Text = "alldiff" Then
45320         f.refConRHS.Enabled = False
              'f.refConRHS.Text = ""
45330     End If
          
45340     If ListItem >= 1 And Not model.Constraints Is Nothing Then
              Dim origREL As String
              
45350         origREL = model.Constraints(ListItem).ConstraintType
45360         If f.cboConRel.Text <> origREL Then
45370             Disabler False, f
45380             f.cmdAddCon.Enabled = True
45390             ConChangedMode = True
45400         Else
45410             Disabler True, f
45420             f.cmdAddCon.Enabled = False
45430             ConChangedMode = False
45440         End If
45450     End If
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
          frmModel.ModelAddConstraint Me
End Sub
Public Sub ModelAddConstraint(f As UserForm)
45460     On Error GoTo errorHandler
          
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
45470     TestStringForConstraint f.refConLHS.Text, LHSisRange, LHSisFormula, LHSIsValueWithEqual, LHSIsValueWithoutEqual
          
45480     If LHSisRange = False Then
              ' The string in the LHS refedit does not describe a range
45490         MsgBox "Left-hand-side of constraint must be a range."
45500         Exit Sub
45510     End If
45520     If Range(Trim(f.refConLHS.Text)).Areas.Count > 1 Then
              ' The LHS is multiple areas - not allowed
45530         MsgBox "Left-hand-side of constraint must have only one area."
45540         Exit Sub
45550     End If
45560     Set rngLHS = Range(Trim(f.refConLHS.Text))
          
          '----------------------------------------------------------------
          ' RIGHT HAND SIDE
          Dim RHSisRange As Boolean, RHSisFormula As Boolean, RHSIsValueWithEqual As Boolean, RHSIsValueWithoutEqual As Boolean
          Dim strRel As String
45570     strRel = f.cboConRel.Text
45580     If strRel = "" Then ' Should not happen as of 20/9/2011 (AJM)
45590         MsgBox "Please select a relation such as = or <="
45600         Exit Sub
45610     End If
45620     IsRestrict = Not ((strRel = "=") Or (strRel = "<=") Or (strRel = ">="))
45630     If Not IsRestrict Then
45640         If Trim(f.refConRHS.Text) = "" Then
45650             MsgBox "Please enter a right-hand-side!"
45660             Exit Sub
45670         End If
              
45680         TestStringForConstraint f.refConRHS.Text, RHSisRange, RHSisFormula, RHSIsValueWithEqual, RHSIsValueWithoutEqual
              
45690         If Not RHSisRange And Not RHSisFormula _
              And Not RHSIsValueWithEqual And Not RHSIsValueWithoutEqual Then
45700             MsgBox "The right-hand-side of a constraint can be either:" + vbNewLine + _
                          "A single-cell range (e.g. =A4)" + vbNewLine + _
                          "A multi-cell range of the same size as the LHS (e.g. =A4:B5)" + vbNewLine + _
                          "A single constant value (e.g. =2)" + vbNewLine + _
                          "A formula returning a single value (eg =sin(A4)"
45710             Exit Sub
45720         End If
              
45730         If RHSisRange Then
                  ' If it is single cell, thats OK
                  ' If it is multi cell, it must match cell count for LHS
45740             Set rngRHS = Range(Trim(f.refConRHS.Text))
45750             If rngRHS.Count > 1 Then
45760                 If rngRHS.Count <> rngLHS.Count Then
                          ' Mismatch!
45770                     MsgBox "Right-hand-side of constraint has more than one cell, and does not match the number of cells on the left-hand-side."
45780                     Exit Sub
45790                 End If
45800             End If
45810         End If
              
              ' If not a range then evaluate to see if its legit
              ' Evaluate is not locale-friendly
              ' So we put it in a cell on the internal sheet, then get it back
              ' AJM 20/9/2011: We need to prefix the formula with an "=" otherwise formula such as 'sheet name'!A1 get entered as a string constant (becaused of the leading ')
              Dim internalRHS As String
45820         internalRHS = Trim(f.refConRHS.Text)
              
              ' Turn off dialog display; we do not want try to open a workbook with a name of the worksheet! This happens if the formula comes from a worksheet
              ' whose name contains a space
45830         Application.DisplayAlerts = False
45840         On Error GoTo ErrorHandler_CannotInterpretRHS
45850         OpenSolverSheet.Range("A1").FormulaLocal = IIf(left(internalRHS, 1) = "=", "", "=") & f.refConRHS.Text
45860         internalRHS = OpenSolverSheet.Range("A1").Formula
45870         OpenSolverSheet.Range("A1").Clear ' This must be blank to ensure no risk of dialogs being shown trying to locate a sheet
45880         Application.DisplayAlerts = True
              
45890         If Not RHSisRange Then
                  ' Can we evaluate this function or constant?
                  Dim varReturn As Variant
45900             varReturn = ActiveSheet.Evaluate(internalRHS) ' Must be worksheet.evaluate to get references to names local to the sheet
45910             If VBA.VarType(varReturn) = vbError Then
45920                 MsgBox "The formula or value for the RHS is not valid. Please check and try again."
45930                 f.refConRHS.SetFocus
45940                 DoEvents
45950                 Exit Sub
45960             End If
45970         End If
              
              ' If it isn't a range, lets convert any cell references to absolute
              ' Will fail if refConRHS has a non-English locale number
45980         If left(internalRHS, 1) <> "=" Then
45990             varReturn = Application.ConvertFormula("=" + internalRHS, FromReferenceStyle:=xlA1, ToReferenceStyle:=xlA1, ToAbsolute:=xlAbsolute)
46000         Else
46010             varReturn = Application.ConvertFormula(internalRHS, FromReferenceStyle:=xlA1, ToReferenceStyle:=xlA1, ToAbsolute:=xlAbsolute)
46020         End If
46030         If (VBA.VarType(varReturn) = vbError) Then
                  ' Its valid, but couldn't convert to standard form, probably because not A1... just leave it
46040         Else
                  ' Always comes back with a = at the start
                  ' Unfortunately, return value will have wrong locale...
                  ' But not much can be done with that?
46050             f.refConRHS.Text = Mid(varReturn, 2, Len(varReturn))
46060         End If
              
          
46070     End If
          
46080     Disabler True, f
46090     f.cmdAddCon.Enabled = False
46100     ConChangedMode = False
              
          '================================================================
          ' Update constraint?
46110     If f.cmdAddCon.Caption <> "Add constraint" Then
          
              'With model.Constraints(f.lstConstraints.ListIndex)
46120         With model.Constraints(ListItem)
46130             Set .LHS = rngLHS
46140             Set .Relation = Nothing
46150             .ConstraintType = strRel
46160             If IsRestrict Then
46170                 Set .RHS = Nothing
46180                 .RHSstring = ""
46190             Else
46200                 If RHSisRange Then
46210                     Set .RHS = rngRHS
46220                     .RHSstring = ""
46230                 Else
46240                     Set .RHS = Nothing
46250                     If left(.RHSstring, 1) <> "=" Then
46260                         .RHSstring = "=" + f.refConRHS.Text
46270                     Else
46280                         .RHSstring = f.refConRHS.Text
46290                     End If
46300                 End If
46310             End If
46320         End With

46330         If Not DontRepop Then model.PopulateConstraintListBox f.lstConstraints
46340         Exit Sub
46350     Else
          '================================================================
          ' Add constraint
              Dim NewConstraint As New CConstraint
46360         With NewConstraint
46370             Set .LHS = rngLHS
46380             Set .Relation = Nothing
46390             .ConstraintType = strRel
46400             If IsRestrict Then
46410                 Set .RHS = Nothing
46420                 .RHSstring = ""
46430             Else
46440                 If RHSisRange Then
46450                     Set .RHS = rngRHS
46460                     .RHSstring = ""
46470                 Else
46480                     Set .RHS = Nothing
46490                     If left(.RHSstring, 1) <> "=" Then
46500                         .RHSstring = "=" + f.refConRHS.Text ' This has been converted above into a US-locale formula, value or reference
46510                     Else
46520                         .RHSstring = f.refConRHS.Text ' This has been converted above into a US-locale formula, value or reference
46530                     End If
46540                 End If
46550             End If
46560         End With
              
46570         model.Constraints.Add NewConstraint ', NewConstraint.GetKey
46580         If Not DontRepop Then model.PopulateConstraintListBox f.lstConstraints
46590         Exit Sub
46600     End If

46610     Application.DisplayAlerts = True
46620     OpenSolverSheet.Range("A1").Clear ' This must be blank to ensure no risk of dialogs being shown trying to locate a sheet
46630     Exit Sub

ErrorHandler_CannotInterpretRHS:
46640     Application.DisplayAlerts = True
46650     OpenSolverSheet.Range("A1").Clear ' This must be blank to ensure no risk of dialogs being shown trying to locate a sheet
          ' Couldn't turn the RHS into a formula
46660     MsgBox "The formula or value for the RHS is not valid. Please check and try again."
46670     f.refConRHS.SetFocus
46680     DoEvents ' Try to stop RefEdit bugs
46690     Exit Sub
errorHandler:
46700     Application.DisplayAlerts = True
46710     OpenSolverSheet.Range("A1").FormulaLocal = "" ' This must be blank to ensure no risk of dialogs being shown trying to locate a sheet
46720     MsgBox "While constructing the model, OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source & " (frmModel::cmdBuild_Click)", , "OpenSolver Code Error"
46730     DoEvents ' Try to stop RefEdit bugs
46740     Exit Sub

End Sub


'--------------------------------------------------------------------
' cmdDelSelCon_Click [event]
' Delete selected constraint
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub cmdDelSelCon_Click()
         frmModel.ModelDeleteConstraint Me
End Sub
Public Sub ModelDeleteConstraint(f As UserForm)

46750     If f.lstConstraints.ListIndex = -1 Then
46760         MsgBox "Please select a constraint!", vbOKOnly, "AutoModel"
46770         Exit Sub
46780     End If
            
          ' Remove it
46790     model.Constraints.Remove f.lstConstraints.ListIndex
          
          ' Update form
46800     model.PopulateConstraintListBox f.lstConstraints
End Sub


'--------------------------------------------------------------------
' lstConstraints_Change [event]
' Selection in constraints box changes
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub lstConstraints_Change()
    frmModel.ModelLstConstraintsChange Me
End Sub

Public Sub ModelLstConstraintsChange(f As UserForm)
          
46810     If ConChangedMode = True Then
46820         If f.cmdAddCon.Caption = "Update constraint" Then
46830             If MsgBox("You have made changes to the current constraint." _
                      + vbNewLine + "Do you want to save these changes?", vbYesNo) = vbYes Then
          
                      'Debug.Print "Doing cmdAddCon_Click"
46840                 DontRepop = True
46850                 f.cmdAddCon_Click
46860                 DontRepop = False
                      ''Debug.Print "Done."
46870             End If
46880         Else
46890             If MsgBox("You have entered a constraint." _
                      + vbNewLine + "Do you want to save this as a new constraint?", vbYesNo) = vbYes Then
          
                      ''Debug.Print "Doing cmdAddCon_Click"
46900                 DontRepop = True
46910                 f.cmdAddCon_Click
46920                 DontRepop = False
                      ''Debug.Print "Done."
46930             End If
46940         End If
          
46950         Disabler True, f
46960         f.cmdAddCon.Enabled = False
46970         ConChangedMode = False
              'ListItem = f.lstConstraints.ListIndex
46980         model.PopulateConstraintListBox f.lstConstraints
              'f.lstConstraints.ListIndex = ListItem
46990     End If
          
47000     ListItem = f.lstConstraints.ListIndex
          
47010     If f.lstConstraints.ListIndex = -1 Then
47020         Exit Sub
47030     End If
47040     If f.lstConstraints.ListIndex = 0 Then
              'Add constraint
47050         f.refConLHS.Enabled = True
47060         DoEvents
47070         f.refConLHS.Text = ""
47080         DoEvents
              'refConRHS.Enabled = True
              'DoEvents
47090         f.refConRHS.Text = ""
47100         DoEvents
47110         ModelChangeConRel f ' AJM: Force the RHS to be active only if the current relation is =, < or >, which is set based on the last constraint
47120         f.cmdAddCon.Enabled = True
47130         f.cmdAddCon.Caption = "Add constraint"
47140         f.cmdDelSelCon.Enabled = False
47150         DoEvents
47160         Application.CutCopyMode = False
47170         ActiveCell.Select ' Just select one cell, choosing a cell that should be visible to avoid scrolling
47180         Application.ScreenUpdating = True
              'refConLHS.select
47190         Exit Sub
47200     Else
              ' Update constraint
47210         f.refConLHS.Enabled = True
47220         DoEvents
47230         f.refConRHS.Enabled = True
47240         DoEvents
47250         f.cmdAddCon.Enabled = False
47260         f.cmdAddCon.Caption = "Update constraint"
47270         f.cmdDelSelCon.Enabled = True
47280         DoEvents

              ' Set text and marching ants; the select/copy operations can throw errors if on different sheets
47290         On Error Resume Next
47300         ActiveCell.Select   ' We may fail in the next steps, se we cancel any old highlighting
47310         Application.CutCopyMode = False
              Dim copyRange As Range
47320         With model.Constraints(f.lstConstraints.ListIndex)
47330             f.refConLHS.Text = GetDisplayAddress(.LHS, False)
47340             Set copyRange = .LHS
47350             f.cboConRel.ListIndex = cboPosition(.ConstraintType)
47360             f.refConRHS.Text = ""
47370             If Not .RHS Is Nothing Then
47380                 f.refConRHS.Text = GetDisplayAddress(.RHS, False)
47390                 Set copyRange = ProperUnion(copyRange, .RHS)
47400             ElseIf .RHS Is Nothing And .RHSstring <> "" Then
47410                 If Mid(.RHSstring, 1, 1) = "=" Then
47420                     f.refConRHS.Text = RemoveActiveSheetNameFromString(Mid(.RHSstring, 2, Len(.RHSstring)))
47430                 Else
47440                     f.refConRHS.Text = RemoveActiveSheetNameFromString(.RHSstring)
47450                 End If
47460             End If
47470         End With
47480         ModelChangeConRel f
              
              ' Will fail if LHS and RHS are different shape
              ' Silently fail, nothing that can be done about it
              ' ALSO fails at the Union step if on different shapes
47490         copyRange.Select
47500         copyRange.Copy
47510         Application.ScreenUpdating = True
          
47520         Exit Sub
47530     End If
End Sub

Private Sub lstConstraints_Exit(ByVal Cancel As MSForms.ReturnBoolean)
          ' When the focus leaves this list, we want to remove any highlighting shown by selected cells
47540     ActiveCell.Select
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
47550     TheString = Trim(TheString) ' AJM: Remove any leading/trailing spaces
          
47560     If Len(TheString) = 0 Then Exit Sub
          
          ' Test for RANGE
47570     On Error Resume Next
          Dim r As Range
47580     Set r = Range(TheString)
47590     RefersToRange = (Err.Number = 0)
          
          ' Test for ...
47600     If Not RefersToRange Then
              ' Not a range, might be constant?
47610         If Mid(TheString, 1, 1) <> "=" Then
                  ' Not sure what this is, but assume its OK - if an equal is added
47620             RefersToValueWithoutEqual = True
47630         Else
                  ' Test for a numeric constant, in US format
47640             If IsAmericanNumber(Mid(TheString, 2)) Then
47650                 RefersToValueWithEqual = True
47660             Else
                      'FORMULA
47670                 RefersToFormula = True
47680             End If
47690         End If
47700     End If
          
End Sub
