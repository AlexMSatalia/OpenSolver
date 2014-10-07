VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAutoModel 
   Caption         =   "OpenSolver - AutoModel"
   ClientHeight    =   4485
   ClientLeft      =   42
   ClientTop       =   378
   ClientWidth     =   8806
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
4045      DoEvents
4046      Unload frmAutoModel
4047      Unload Me
End Sub

'--------------------------------------------------------------------
' UserForm_Activate [event]
' Called when the form is shown.
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub UserForm_Activate()
4048      frmAutoModel.AutoModelActivate Me
End Sub

Public Sub AutoModelActivate(f As UserForm)
          ' Reset enabled flags
4049      ResetEverything f
          ' Make sure sheet is up to date
4050      On Error Resume Next
4051      Application.Calculate
4052      On Error GoTo 0
          ' Remove the 'marching ants' showing if a range is copied.
          ' Otherwise, the ants stay visible, and visually conflict with
          ' our cell selection. The ants are also left behind on the
          ' screen. This works around an apparent bug (?) in Excel 2007.
4053      Application.CutCopyMode = False
          ' Get ready to process
4054      DoEvents
          ' Show results of finding objective
4055      GuessObj f
End Sub

'--------------------------------------------------------------------
' ResetEverything
' Resets everything to a fresh state
'
' Written by:       IRD
'--------------------------------------------------------------------
Public Sub ResetEverything(f As UserForm)

4056      f.optMax.value = False
4057      f.optMin.value = False
4058      f.refObj.Text = ""

End Sub


'--------------------------------------------------------------------
' GuessObj
' Called after showing the window, attempts to find objective
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub GuessObj(f As UserForm)
              
4059      Select Case GuessObjStatus
              Case "NoSense"
                  ' Didn't find anything
4060              f.lblStatus.Caption = _
                  "AutoModel was unable to guess anything." + vbNewLine + _
                  "Please enter the objective sense and the objective function cell."
                  'lblStatus.ForeColor = vbRed
4061          Case "SenseNoCell"
4062              If model.ObjectiveSense = MaximiseObjective Then f.optMax.value = True
4063              If model.ObjectiveSense = MinimiseObjective Then f.optMin.value = True
                  'lblStatus.ForeColor = vbBlue
4064              f.lblStatus.Caption = _
                  "AutoModel found the objective sense, but couldn't find the objective cell." + vbNewLine + _
                  "Please check the objective sense and enter the objective function cell."
4065      End Select
          
4066      f.Repaint
4067      DoEvents
End Sub


'--------------------------------------------------------------------
' cmdFinish_Click [event]
' Validate form input, update model, and continue to next step
'
' Written by:       IRD
'--------------------------------------------------------------------
Private Sub cmdFinish_Click()
4068           frmAutoModel.AutoModelFinish Me
End Sub

Public Sub AutoModelFinish(f As UserForm)
          ' Check if user changed objective cell
4069      On Error GoTo BadObjRef
4070      Set model.ObjectiveFunctionCell = ActiveSheet.Range(f.refObj.Text)
          
          ' Get the objective sense
4071      If f.optMax.value = True Then model.ObjectiveSense = MaximiseObjective
4072      If f.optMin.value = True Then model.ObjectiveSense = MinimiseObjective
4073      If model.ObjectiveSense = UnknownObjectiveSense Then
4074          MsgBox "Error: Please select an objective sense (minimise or maximise)!", vbExclamation + vbOKOnly, "AutoModel"
              'frmModel.Show
4075          Exit Sub
4076      End If
          
          ' Find the vars, cons
          Dim result As Boolean
4077      result = model.FindVarsAndCons(IsFirstTime:=True)

4078      If result = False Then
              ' Didn't work/error
4079          MsgBox "An unknown error occurred while trying to find the model.", vbOKOnly
4080      End If
          
          'frmModel.Show
4081      Unload Me
4082      Unload f
4083      DoEvents
4084      Exit Sub
          
BadObjRef:
          ' Couldn't turn the objective cell address into a range
4085      MsgBox "Error: the cell address for the objective is invalid. Please correct " + _
                  "and click 'Finish AutoModel' again.", vbExclamation + vbOKOnly, "AutoModel"
4086      f.refObj.SetFocus ' Set the focus back to the RefEdit
4087      DoEvents ' Try to stop RefEdit bugs
4088      Exit Sub
End Sub
