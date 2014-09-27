VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAutoModel 
   Caption         =   "OpenSolver - AutoModel"
   ClientHeight    =   4480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8805
   OleObjectBlob   =   "frmAutoModel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAutoModel"
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
' frmAutoModel2
' Userform around the functionality in CModel
' Allows user to manually guide the auto-model process, and explains
' how it works.
'
' Created by:       IRD
'--------------------------------------------------------------------

Option Explicit

' The handle to the AutoModel instance
Public model As CModel
Public GuessObjStatus As String


Private Sub cmdCancel_Click()
          DoEvents
          Unload frmAutoModel
41310     Unload Me
End Sub

'--------------------------------------------------------------------
' UserForm_Activate [event]
' Called when the form is shown.
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub UserForm_Activate()
    frmAutoModel.AutoModelActivate Me
End Sub

Public Sub AutoModelActivate(f As UserForm)
          ' Reset enabled flags
41320     ResetEverything f
          ' Make sure sheet is up to date
41330     On Error Resume Next
41340     Application.Calculate
41350     On Error GoTo 0
          ' Remove the 'marching ants' showing if a range is copied.
          ' Otherwise, the ants stay visible, and visually conflict with
          ' our cell selection. The ants are also left behind on the
          ' screen. This works around an apparent bug (?) in Excel 2007.
41360     Application.CutCopyMode = False
          ' Get ready to process
41370     DoEvents
          ' Show results of finding objective
41380     GuessObj f
End Sub

'--------------------------------------------------------------------
' ResetEverything
' Resets everything to a fresh state
'
' Written by:       IRD
'--------------------------------------------------------------------
Public Sub ResetEverything(f As UserForm)

41390     f.optMax.value = False
41400     f.optMin.value = False
41410     f.refObj.Text = ""

End Sub


'--------------------------------------------------------------------
' GuessObj
' Called after showing the window, attempts to find objective
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub GuessObj(f As UserForm)
              
41420     Select Case GuessObjStatus
              Case "NoSense"
                  ' Didn't find anything
41430             f.lblStatus.Caption = _
                  "AutoModel was unable to guess anything." + vbNewLine + _
                  "Please enter the objective sense and the objective function cell."
                  'lblStatus.ForeColor = vbRed
41440         Case "SenseNoCell"
41450             If model.ObjectiveSense = MaximiseObjective Then f.optMax.value = True
41460             If model.ObjectiveSense = MinimiseObjective Then f.optMin.value = True
                  'lblStatus.ForeColor = vbBlue
41470             f.lblStatus.Caption = _
                  "AutoModel found the objective sense, but couldn't find the objective cell." + vbNewLine + _
                  "Please check the objective sense and enter the objective function cell."
41480     End Select
          
41490     f.Repaint
41500     DoEvents
End Sub


'--------------------------------------------------------------------
' cmdFinish_Click [event]
' Validate form input, update model, and continue to next step
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub cmdFinish_Click()
         frmAutoModel.AutoModelFinish Me
End Sub

Public Sub AutoModelFinish(f As UserForm)
          ' Check if user changed objective cell
41510     On Error GoTo BadObjRef
41520     Set model.ObjectiveFunctionCell = ActiveSheet.Range(f.refObj.Text)
          
          ' Get the objective sense
41530     If f.optMax.value = True Then model.ObjectiveSense = MaximiseObjective
41540     If f.optMin.value = True Then model.ObjectiveSense = MinimiseObjective
41550     If model.ObjectiveSense = UnknownObjectiveSense Then
41560         MsgBox "Error: Please select an objective sense (minimise or maximise)!", vbExclamation + vbOKOnly, "AutoModel"
              'frmModel.Show
41570         Exit Sub
41580     End If
          
          ' Find the vars, cons
          Dim result As Boolean
41590     result = model.FindVarsAndCons(IsFirstTime:=True)

41600     If result = False Then
              ' Didn't work/error
41610         MsgBox "An unknown error occurred while trying to find the model.", vbOKOnly
41620     End If
          
          'frmModel.Show
41630     Unload Me
          Unload f
          DoEvents
41640     Exit Sub
          
BadObjRef:
          ' Couldn't turn the objective cell address into a range
41650     MsgBox "Error: the cell address for the objective is invalid. Please correct " + _
                  "and click 'Finish AutoModel' again.", vbExclamation + vbOKOnly, "AutoModel"
41660     f.refObj.SetFocus ' Set the focus back to the RefEdit
41670     DoEvents ' Try to stop RefEdit bugs
41680     Exit Sub
End Sub
