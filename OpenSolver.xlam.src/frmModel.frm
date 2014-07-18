VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmModel 
   Caption         =   "OpenSolver - Model"
   ClientHeight    =   8265.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9855.001
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

Sub Disabler(TrueIfEnable As Boolean)

42390     lblDescHeader.Enabled = TrueIfEnable
42400     lblDesc.Enabled = TrueIfEnable
42410     cmdRunAutoModel.Enabled = TrueIfEnable
          
42420     frameDiv1.Enabled = False
          
42430     lblStep1.Enabled = TrueIfEnable
42440     refObj.Enabled = TrueIfEnable
42450     optMax.Enabled = TrueIfEnable
42460     optMin.Enabled = TrueIfEnable
42470     optTarget.Enabled = TrueIfEnable
42480     txtObjTarget.Enabled = TrueIfEnable And optTarget.value
          
42490     frameDiv2.Enabled = False
          
42500     lblStep2.Enabled = TrueIfEnable
42510     refDecision.Enabled = TrueIfEnable
          
42520     frameDiv3.Enabled = False
          
42530     chkNonNeg.Enabled = TrueIfEnable
42540     cmdCancelCon.Enabled = Not TrueIfEnable
42550     cmdDelSelCon.Enabled = TrueIfEnable
          
42560     frameDiv4.Enabled = False
          
42570     lblDuals.Enabled = TrueIfEnable

42580     chkGetDuals.Enabled = TrueIfEnable
42590     chkGetDuals2.Enabled = TrueIfEnable
42600     optUpdate.Enabled = chkGetDuals2.value
42610     optNew.Enabled = chkGetDuals2.value
          
42620     refDuals.Enabled = TrueIfEnable And chkGetDuals.value And chkGetDuals.Enabled
          
42630     frameDiv5.Enabled = False
42640     frameDiv6.Enabled = False
42650     frameDiv7.Enabled = False
          
42660     chkShowModel.Enabled = TrueIfEnable
42670     cmdOptions.Enabled = TrueIfEnable
42680     cmdBuild.Enabled = TrueIfEnable
42690     cmdCancel.Enabled = TrueIfEnable

42700     frmOptions.chkLinear.Enabled = True
42710     frmOptions.chkPerformLinearityCheck.Enabled = True
42720     frmOptions.txtTol.Enabled = True
42730     frmOptions.txtMaxIter.Enabled = True
42740     frmOptions.txtPre.Enabled = True
          
        
          Dim Solver As String
42750     If Not GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!OpenSolver_ChosenSolver", Solver) Then
              Solver = "CBC"
              Call SetNameOnSheet("OpenSolver_ChosenSolver", "=" & Solver)
42770     End If
            
          
          If Not SolverHasSensitivityAnalysis(Solver) Then
42780         ' Disable dual options
              chkGetDuals2.Enabled = False
42790         chkGetDuals.Enabled = False
42800         optUpdate.Enabled = False
42810         optNew.Enabled = False
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
Sub UpdateFormFromMemory()
42890     If model.ObjectiveSense = MaximiseObjective Then optMax.value = True
42900     If model.ObjectiveSense = MinimiseObjective Then optMin.value = True
42910     If model.ObjectiveSense = TargetObjective Then optTarget.value = True   ' AJM 20110907
42920     txtObjTarget.Text = CStr(model.ObjectiveTarget)   ' AJM 20110907 Always show the target (which may just be 0)
          
42930     chkNonNeg.value = model.NonNegativityAssumption
         
42940     If Not model.ObjectiveFunctionCell Is Nothing Then refObj.Text = GetDisplayAddress(model.ObjectiveFunctionCell, False)
          
42950     If Not model.DecisionVariables Is Nothing Then refDecision.Text = GetDisplayAddressInCurrentLocale(model.DecisionVariables)
          
42960     chkGetDuals.value = Not model.Duals Is Nothing
42970     If model.Duals Is Nothing Then
42980         refDuals.Text = ""
42990     Else
43000         refDuals.Text = GetDisplayAddress(model.Duals, False)
43010     End If
                    
43020     model.PopulateConstraintListBox lstConstraints
43030     lstConstraints_Change

      '          On Error GoTo nameUndefined
      '          chkGetDuals2.Value = Mid(Names("'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!OpenSolver_DualsNewSheet").Value, 2)
          Dim sheetName As String, value As String, ResetDualsNewSheet As Boolean
43040     sheetName = "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!"
43050     If GetNameValueIfExists(ActiveWorkbook, sheetName & "OpenSolver_DualsNewSheet", value) Then
43060         chkGetDuals2.value = value
              ' If checkbox is null, then the stored value was not 'True' or 'False'. We should reset to false
              If IsNull(chkGetDuals2.value) Then
                  ResetDualsNewSheet = True
              End If
43070     Else
43080         ResetDualsNewSheet = True
43100     End If
          
          If ResetDualsNewSheet Then
              Call SetNameOnSheet("OpenSolver_DualsNewSheet", "=FALSE")
              chkGetDuals2.value = False
          End If
          
43110     optUpdate.Enabled = chkGetDuals2.value
43120     optNew.Enabled = chkGetDuals2.value
43130     If GetNameValueIfExists(ActiveWorkbook, sheetName & "OpenSolver_UpdateSensitivity", value) Then
43140         If value = "TRUE" Then
43150           optUpdate.value = value
43160         Else
43170           optNew.value = True
43180         End If
43190     Else
43200         Call SetNameOnSheet("OpenSolver_UpdateSensitivity", "=TRUE")
43210         optUpdate.value = True
43220     End If
      '          Exit Sub
          
      'nameUndefined:
      '          Call SetNameOnSheet("OpenSolver_DualsNewSheet", "=FALSE")
      '          chkGetDuals2.Value = False 'Mid(Names("'" & Replace(ActiveWorkbook.ActiveSheet.name, "'", "''") & "'!OpenSolver_DualsNewSheet").Value, 2)
          
End Sub

Private Sub chkGetDuals_Click()
43230     refDuals.Enabled = chkGetDuals.value
          'If chkGetDuals.Value Then
          '    MsgBox "This Beta feature will, when the model is solved, produce a 2-column list at a user-chosen location on the current worksheet containing all the constraints and their shadow prices. Please select the top-left cell for this table. Please note that the shadow prices have very little meaning when solving problems with integer or binary variables.", vbOKOnly, "OpenSolver Beta Feature: Shadow Price Listing"
          'End If
End Sub

Private Sub chkGetDuals2_Click()
43240     Call SetNameOnSheet("OpenSolver_DualsNewSheet", "=" & chkGetDuals2.value)
43250     optUpdate.Enabled = chkGetDuals2.value
43260     optNew.Enabled = chkGetDuals2.value
End Sub

Private Sub chkNameRange_Click()
          'Call UpdateFormFromMemory
43270     model.PopulateConstraintListBox lstConstraints
43280     lstConstraints_Change
End Sub

Private Sub cmdCancelCon_Click()
43290     Disabler True
43300     cmdAddCon.Enabled = False
43310     ConChangedMode = False
43320     lstConstraints_Change
End Sub

Private Sub cmdChange_Click()
43330     frmSolverChange.Show
End Sub

Private Sub cmdOptions_Click()
          ' Save the current "Assume Non Negative" option so this is shown in the options dialog.
          ' The saved value gets updated on OK, which we then reflect in our Model dialog
          Dim s As String
43340     SetSolverNameOnSheet "neg", IIf(chkNonNeg.value, "=1", "=2")
              
43350     frmOptions.Show vbModal
43360     If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_neg", s) Then    ' This should always be true
43370         chkNonNeg.value = s = "1"
43380     End If
End Sub

'--------------------------------------------------------------------------------------
'Reset Button
'Deletes the objective function, decision variables and all the constraints in the model
'---------------------------------------------------------------------------------------

Private Sub cmdReset_Click()
                 
          Dim NumConstraints As Single, i As Integer
                  
          'Check the user wants to reset the model
43390     If MsgBox("This will reset the objective function, decision variables and constraints." _
                  & vbCrLf & " Are you sure you want to do this?", vbYesNo, "OpenSolver") = vbNo Then
43400         Exit Sub
43410     End If

          'Reset the objective function and the decision variables
43420     refObj.Text = ""
43430     refDecision.Text = ""
              
          'Find the number of constraints in model
43440     NumConstraints = model.Constraints.Count
          
          ' Remove the constraints
43450     For i = 1 To NumConstraints
43460         model.Constraints.Remove 1
43470     Next i

          ' Update constraints form
43480     model.PopulateConstraintListBox lstConstraints

End Sub

Private Sub optMax_Click()
43490     txtObjTarget.Enabled = optTarget.value
End Sub

Private Sub optMin_Click()
43500     txtObjTarget.Enabled = optTarget.value
End Sub

Private Sub optNew_Click()
43510     Call SetNameOnSheet("OpenSolver_UpdateSensitivity", "=" & optUpdate.value)
End Sub

Private Sub optTarget_Click()
43520     txtObjTarget.Enabled = optTarget.value
End Sub

Private Sub optUpdate_Click()
43530     Call SetNameOnSheet("OpenSolver_UpdateSensitivity", "=" & optUpdate.value)
End Sub

Private Sub refConLHS_Change()
          ' Compare to expected value
43540     If ListItem >= 1 And Not model.Constraints Is Nothing Then
              Dim origLHS As String
              
43550         origLHS = model.Constraints(ListItem).LHS.Address
43560         If refConLHS.Text <> origLHS Then
43570             Disabler False
43580             cmdAddCon.Enabled = True
43590             ConChangedMode = True
43600         Else
43610             Disabler True
43620             cmdAddCon.Enabled = False
43630             ConChangedMode = False
43640         End If
43650     ElseIf ListItem = 0 Then
43660         If refConLHS.Text <> "" Then
43670             Disabler False
43680             cmdAddCon.Enabled = True
43690             ConChangedMode = True
43700         Else
43710             Disabler True
43720             cmdAddCon.Enabled = False
43730             ConChangedMode = False
43740         End If
43750     End If
End Sub

Private Sub refConRHS_Change()
          ' Compare to expected value
43760     If ListItem >= 1 And Not model.Constraints Is Nothing Then
              Dim origRHS As String
43770         If model.Constraints(ListItem).RHS Is Nothing Then
43780             origRHS = model.Constraints(ListItem).RHSstring
43790         Else
43800             origRHS = model.Constraints(ListItem).RHS.Address
43810         End If
43820         If refConRHS.Text <> origRHS Then
43830             Disabler False
43840             cmdAddCon.Enabled = True
43850             ConChangedMode = True
43860         Else
43870             Disabler True
43880             cmdAddCon.Enabled = False
43890             ConChangedMode = False
43900         End If
43910     ElseIf ListItem = 0 Then
43920         If refConLHS.Text <> "" Then
43930             Disabler False
43940             cmdAddCon.Enabled = True
43950             ConChangedMode = True
43960         Else
43970             Disabler True
43980             cmdAddCon.Enabled = False
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
44110     optMax.value = False
44120     optMin.value = False
44130     refObj.Text = ""
44140     refDecision.Text = ""
44150     refConLHS.Text = ""
44160     refConRHS.Text = ""
44170     lstConstraints.Clear
44180     cboConRel.Clear
44190     cboConRel.AddItem "="
44200     cboConRel.AddItem "<="
44210     cboConRel.AddItem ">="
44220     cboConRel.AddItem "int"
44230     cboConRel.AddItem "bin"
44240     cboConRel.AddItem "alldiff"
44250     cboConRel.Text = "="    ' We set an initial value just in case there is no model, and the user goes straight to AddNewConstraint
          
          'Find current solver
          Dim solverName As String
44260     If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!OpenSolver_ChosenSolver", solverName) Then
44270         lblSolver.Caption = "Current Solver Engine: " & UCase(left(solverName, 1)) & Mid(solverName, 2)
44280     Else: lblSolver.Caption = "Current Solver Engine: CBC"
44290     End If
          ' Load the model on the sheet into memory
44300     ListItem = -1
44310     ConChangedMode = False
44320     DontRepop = False
44330     Disabler True
44340     model.LoadFromSheet
44350     DoEvents
44360     UpdateFormFromMemory
44370     DoEvents
          ' Take focus away from refEdits
44380     DoEvents
          'cmdCancel.SetFocus
          'DoEvents
44390     Me.Repaint
          'cmdCancel.SetFocus
44400     DoEvents
End Sub


Private Sub cmdCancel_Click()
44410     DoEvents
44420     On Error Resume Next ' Just to be safe on our select
44430     Application.CutCopyMode = False
44440     ActiveCell.Select ' Just select one cell, choosing a cell that should be visible to avoid scrolling
44450     Application.ScreenUpdating = True
44460     Me.Hide
End Sub


'--------------------------------------------------------------------
' cmdBuild_Click [event]
' Turn the model into a Solver model
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub cmdRunAutoModel_Click()
          ' Try and guess the objective
          Dim status As String
44470     status = model.FindObjective(ActiveSheet)
          
          ' Get it in memory
44480     Load frmAutoModel
          ' Pass it the model reference
44490     Set frmAutoModel.model = model
44500     frmAutoModel.GuessObjStatus = status
          
44510     Select Case status
              Case "NoSense"
44520             frmAutoModel.Show vbModal
44530         Case "SenseNoCell"
44540             frmAutoModel.Show vbModal
44550         Case Else ' Found objective
44560             model.FindVarsAndCons IsFirstTime:=True
44570     End Select

          ' Force the automatically created model to be a linear one, and turn on AssumeNonNegative
44580     model.NonNegativityAssumption = True
44590     SetSolverNameOnSheet "lin", "=1"
          Dim s As String
44600     If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!solver_eng", s) Then SetSolverNameOnSheet "eng", "=2"
                      ' Set this for Solver 2010 models, but only if this name is already defined

44610     UpdateFormFromMemory
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

44710     If Trim(refObj.Text) = "" Then
44720         Set model.ObjectiveFunctionCell = Nothing
44730     Else
44740         Set model.ObjectiveFunctionCell = Range(refObj.Text)
44750     End If
44760     On Error GoTo errorHandler
          
          ' Get the objective sense
44770     If optMax.value = True Then model.ObjectiveSense = MaximiseObjective
44780     If optMin.value = True Then model.ObjectiveSense = MinimiseObjective
44790     If optTarget.value = True Then
44800         model.ObjectiveSense = TargetObjective
44810         On Error GoTo BadObjectiveTarget
44820         model.ObjectiveTarget = CDbl(txtObjTarget.Text)
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
44900     If Trim(refDecision.Text) = "" Then
44910         Set model.DecisionVariables = Nothing
44920     Else
44930         Set model.DecisionVariables = Range(ConvertFromCurrentLocale(refDecision.Text))
44940     End If
44950     On Error GoTo errorHandler

          '----------------------------------------------------------------
          ' Pull possibly updated dual storage cells
44960     On Error GoTo BadDualsRef

44970     If chkGetDuals.value = False Or Trim(refDuals.Text) = "" Then
44980         Set model.Duals = Nothing
44990     Else
45000         Set model.Duals = Range(refDuals.Text)
45010     End If
45020     On Error GoTo errorHandler
          
          '----------------------------------------------------------------
          ' Do it
45030     model.NonNegativityAssumption = chkNonNeg.value
          
45040     model.BuildModel
          
          
          '----------------------------------------------------------------
          ' Display on screen
45050     If chkShowModel.value = True Then OpenSolverVisualizer.ShowSolverModel
45060     Application.Calculate
          '----------------------------------------------------------------
          ' Finish
45070     Me.Hide
45080     Exit Sub
          
          '----------------------------------------------------------------
BadObjRef:
          ' Couldn't turn the objective cell address into a range
45090     MsgBox "Error: the cell address for the objective is invalid. " + _
                  "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
45100     refObj.SetFocus ' Set the focus back to the RefEdit
45110     DoEvents ' Try to stop RefEdit bugs
45120     Exit Sub
          '----------------------------------------------------------------
BadDecRef:
          ' Couldn't turn the decision variable address into a range
45130     MsgBox "Error: the cell address for the decision variables is invalid. " + _
                  "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
45140     refDecision.SetFocus ' Set the focus back to the RefEdit
45150     DoEvents ' Try to stop RefEdit bugs
45160     Exit Sub
BadObjectiveTarget:
          ' Couldn't turn the objective target into a value
45170     MsgBox "Error: the target value for the objective cell is invalid. " + _
                  "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
45180     txtObjTarget.SetFocus ' Set the focus back to the target text box
45190     DoEvents ' Try to stop RefEdit bugs
45200     Exit Sub
BadDualsRef:
          ' Couldn't turn the dual cell into a range
45210     MsgBox "Error: the cell for storing the shadow prices is invalid. " + _
                  "Please correct this and try again.", vbExclamation + vbOKOnly, "OpenSolver"
45220     refDuals.SetFocus ' Set the focus back to the target text box
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
45280     If cboConRel.Text = "=" Then refConRHS.Enabled = True
45290     If cboConRel.Text = "<=" Then refConRHS.Enabled = True
45300     If cboConRel.Text = ">=" Then refConRHS.Enabled = True
45310     If cboConRel.Text = "int" Or cboConRel.Text = "bin" Or cboConRel.Text = "alldiff" Then
45320         refConRHS.Enabled = False
              'refConRHS.Text = ""
45330     End If
          
45340     If ListItem >= 1 And Not model.Constraints Is Nothing Then
              Dim origREL As String
              
45350         origREL = model.Constraints(ListItem).ConstraintType
45360         If cboConRel.Text <> origREL Then
45370             Disabler False
45380             cmdAddCon.Enabled = True
45390             ConChangedMode = True
45400         Else
45410             Disabler True
45420             cmdAddCon.Enabled = False
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
45470     TestStringForConstraint refConLHS.Text, LHSisRange, LHSisFormula, LHSIsValueWithEqual, LHSIsValueWithoutEqual
          
45480     If LHSisRange = False Then
              ' The string in the LHS refedit does not describe a range
45490         MsgBox "Left-hand-side of constraint must be a range."
45500         Exit Sub
45510     End If
45520     If Range(Trim(refConLHS.Text)).Areas.Count > 1 Then
              ' The LHS is multiple areas - not allowed
45530         MsgBox "Left-hand-side of constraint must have only one area."
45540         Exit Sub
45550     End If
45560     Set rngLHS = Range(Trim(refConLHS.Text))
          
          '----------------------------------------------------------------
          ' RIGHT HAND SIDE
          Dim RHSisRange As Boolean, RHSisFormula As Boolean, RHSIsValueWithEqual As Boolean, RHSIsValueWithoutEqual As Boolean
          Dim strRel As String
45570     strRel = cboConRel.Text
45580     If strRel = "" Then ' Should not happen as of 20/9/2011 (AJM)
45590         MsgBox "Please select a relation such as = or <="
45600         Exit Sub
45610     End If
45620     IsRestrict = Not ((strRel = "=") Or (strRel = "<=") Or (strRel = ">="))
45630     If Not IsRestrict Then
45640         If Trim(refConRHS.Text) = "" Then
45650             MsgBox "Please enter a right-hand-side!"
45660             Exit Sub
45670         End If
              
45680         TestStringForConstraint refConRHS.Text, RHSisRange, RHSisFormula, RHSIsValueWithEqual, RHSIsValueWithoutEqual
              
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
45740             Set rngRHS = Range(Trim(refConRHS.Text))
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
45820         internalRHS = Trim(refConRHS.Text)
              
              ' Turn off dialog display; we do not want try to open a workbook with a name of the worksheet! This happens if the formula comes from a worksheet
              ' whose name contains a space
45830         Application.DisplayAlerts = False
45840         On Error GoTo ErrorHandler_CannotInterpretRHS
45850         OpenSolverSheet.Range("A1").FormulaLocal = IIf(left(internalRHS, 1) = "=", "", "=") & refConRHS.Text
45860         internalRHS = OpenSolverSheet.Range("A1").Formula
45870         OpenSolverSheet.Range("A1").Clear ' This must be blank to ensure no risk of dialogs being shown trying to locate a sheet
45880         Application.DisplayAlerts = True
              
45890         If Not RHSisRange Then
                  ' Can we evaluate this function or constant?
                  Dim varReturn As Variant
45900             varReturn = ActiveSheet.Evaluate(internalRHS) ' Must be worksheet.evaluate to get references to names local to the sheet
45910             If VBA.VarType(varReturn) = vbError Then
45920                 MsgBox "The formula or value for the RHS is not valid. Please check and try again."
45930                 refConRHS.SetFocus
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
46050             refConRHS.Text = Mid(varReturn, 2, Len(varReturn))
46060         End If
              
          
46070     End If
          
46080     Disabler True
46090     cmdAddCon.Enabled = False
46100     ConChangedMode = False
              
          '================================================================
          ' Update constraint?
46110     If cmdAddCon.Caption <> "Add constraint" Then
          
              'With model.Constraints(lstConstraints.ListIndex)
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
46260                         .RHSstring = "=" + refConRHS.Text
46270                     Else
46280                         .RHSstring = refConRHS.Text
46290                     End If
46300                 End If
46310             End If
46320         End With

46330         If Not DontRepop Then model.PopulateConstraintListBox lstConstraints
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
46500                         .RHSstring = "=" + refConRHS.Text ' This has been converted above into a US-locale formula, value or reference
46510                     Else
46520                         .RHSstring = refConRHS.Text ' This has been converted above into a US-locale formula, value or reference
46530                     End If
46540                 End If
46550             End If
46560         End With
              
46570         model.Constraints.Add NewConstraint ', NewConstraint.GetKey
46580         If Not DontRepop Then model.PopulateConstraintListBox lstConstraints
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
46670     refConRHS.SetFocus
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
46750     If lstConstraints.ListIndex = -1 Then
46760         MsgBox "Please select a constraint!", vbOKOnly, "AutoModel"
46770         Exit Sub
46780     End If
            
          ' Remove it
46790     model.Constraints.Remove lstConstraints.ListIndex
          
          ' Update form
46800     model.PopulateConstraintListBox lstConstraints
End Sub


'--------------------------------------------------------------------
' lstConstraints_Change [event]
' Selection in constraints box changes
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub lstConstraints_Change()
          
46810     If ConChangedMode = True Then
46820         If cmdAddCon.Caption = "Update constraint" Then
46830             If MsgBox("You have made changes to the current constraint." _
                      + vbNewLine + "Do you want to save these changes?", vbYesNo) = vbYes Then
          
                      'Debug.Print "Doing cmdAddCon_Click"
46840                 DontRepop = True
46850                 cmdAddCon_Click
46860                 DontRepop = False
                      ''Debug.Print "Done."
46870             End If
46880         Else
46890             If MsgBox("You have entered a constraint." _
                      + vbNewLine + "Do you want to save this as a new constraint?", vbYesNo) = vbYes Then
          
                      ''Debug.Print "Doing cmdAddCon_Click"
46900                 DontRepop = True
46910                 cmdAddCon_Click
46920                 DontRepop = False
                      ''Debug.Print "Done."
46930             End If
46940         End If
          
46950         Disabler True
46960         cmdAddCon.Enabled = False
46970         ConChangedMode = False
              'ListItem = lstConstraints.ListIndex
46980         model.PopulateConstraintListBox lstConstraints
              'lstConstraints.ListIndex = ListItem
46990     End If
          
47000     ListItem = lstConstraints.ListIndex
          
47010     If lstConstraints.ListIndex = -1 Then
47020         Exit Sub
47030     End If
47040     If lstConstraints.ListIndex = 0 Then
              'Add constraint
47050         refConLHS.Enabled = True
47060         DoEvents
47070         refConLHS.Text = ""
47080         DoEvents
              'refConRHS.Enabled = True
              'DoEvents
47090         refConRHS.Text = ""
47100         DoEvents
47110         cboConRel_Change ' AJM: Force the RHS to be active only if the current relation is =, < or >, which is set based on the last constraint
47120         cmdAddCon.Enabled = True
47130         cmdAddCon.Caption = "Add constraint"
47140         cmdDelSelCon.Enabled = False
47150         DoEvents
47160         Application.CutCopyMode = False
47170         ActiveCell.Select ' Just select one cell, choosing a cell that should be visible to avoid scrolling
47180         Application.ScreenUpdating = True
              'refConLHS.select
47190         Exit Sub
47200     Else
              ' Update constraint
47210         refConLHS.Enabled = True
47220         DoEvents
47230         refConRHS.Enabled = True
47240         DoEvents
47250         cmdAddCon.Enabled = False
47260         cmdAddCon.Caption = "Update constraint"
47270         cmdDelSelCon.Enabled = True
47280         DoEvents

              ' Set text and marching ants; the select/copy operations can throw errors if on different sheets
47290         On Error Resume Next
47300         ActiveCell.Select   ' We may fail in the next steps, se we cancel any old highlighting
47310         Application.CutCopyMode = False
              Dim copyRange As Range
47320         With model.Constraints(lstConstraints.ListIndex)
47330             refConLHS.Text = GetDisplayAddress(.LHS, False)
47340             Set copyRange = .LHS
47350             cboConRel.Text = .ConstraintType
47360             refConRHS.Text = ""
47370             If Not .RHS Is Nothing Then
47380                 refConRHS.Text = GetDisplayAddress(.RHS, False)
47390                 Set copyRange = ProperUnion(copyRange, .RHS)
47400             ElseIf .RHS Is Nothing And .RHSstring <> "" Then
47410                 If Mid(.RHSstring, 1, 1) = "=" Then
47420                     refConRHS.Text = RemoveActiveSheetNameFromString(Mid(.RHSstring, 2, Len(.RHSstring)))
47430                 Else
47440                     refConRHS.Text = RemoveActiveSheetNameFromString(.RHSstring)
47450                 End If
47460             End If
47470         End With
47480         cboConRel_Change
              
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
