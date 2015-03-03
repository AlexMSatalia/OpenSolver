Attribute VB_Name = "OpenSolverAutoModel"
Option Explicit

Public Function RunAutoModel(Optional MinimiseUserInteraction As Boolean = False, Optional ByRef InputModel As CModel) As Boolean

    If Not CheckWorksheetAvailable Then Exit Function
    Dim model As CModel, AskedToShow As Boolean, ShowModel As Boolean, DoBuild As Boolean
    If InputModel Is Nothing Then
        Set model = New CModel
        DoBuild = True
    Else
        Set model = InputModel
        DoBuild = False
    End If
    ShowModel = False
    AskedToShow = False
    
    model.FindObjective ActiveSheet
    If model.ObjectiveFunctionCell Is Nothing Then
        If Not MinimiseUserInteraction Then
            Load frmAutoModel
            Set frmAutoModel.ObjectiveCell = model.ObjectiveFunctionCell
            frmAutoModel.ObjectiveSense = model.ObjectiveSense
            frmAutoModel.chkShow.value = DoBuild
            frmAutoModel.chkShow.Visible = DoBuild
            
            frmAutoModel.Show
            
            If frmAutoModel.Tag = "Cancelled" Then Exit Function
            
            Set model.ObjectiveFunctionCell = frmAutoModel.ObjectiveCell
            model.ObjectiveSense = frmAutoModel.ObjectiveSense
            ShowModel = frmAutoModel.chkShow.value
            AskedToShow = True
            Unload frmAutoModel
        End If
    End If
    
    If Not model.FindVarsAndCons(True) Then
        If Not MinimiseUserInteraction Then MsgBox "Error while looking for variables and constraints"
        RunAutoModel = False
        Exit Function
    End If
    
    model.NonNegativityAssumption = True
    
    If DoBuild Then
        model.BuildModel
        
        If MinimiseUserInteraction Then
            ShowModel = True
        ElseIf Not AskedToShow Then
            If MsgBox("Automodel done! Show model?", vbYesNo, "OpenSolver - AutoModel") = vbYes Then ShowModel = True
        End If
    
        If ShowModel Then
            OpenSolverVisualizer.ShowSolverModel
        Else
            OpenSolverVisualizer.HideSolverModel
        End If
    End If
    
    RunAutoModel = True
End Function
