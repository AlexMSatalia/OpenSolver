Attribute VB_Name = "OpenSolverMain"
' OpenSolver
' Copyright Andrew Mason 2010
' http://www.OpenSolver.org
' This software is distributed under the terms of the GNU General Public License
'
'
' This file is part of OpenSolver.
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

' Dim m_Ribbon As IRibbonUI

Option Explicit

' Version information (as displayed in the About box...)
Public Const sOpenSolverVersion As String = "2.5.4 alpha"
Public Const sOpenSolverDate As String = "2014.07.03"

' Used for the 2003 menu code
Private Const strAddInName As String = "OpenSolver"
Private Const strMenuName  As String = "&OpenSolver"

Dim OpenSolver As COpenSolver

Sub OpenSolver_SolveClickHandler(Optional Control)
27810     If Not CheckWorksheetAvailable Then Exit Sub
27820     RunOpenSolver False
End Sub

Sub OpenSolver_ModelOptions(Optional Control)
27830     If Not CheckWorksheetAvailable Then Exit Sub
27840     frmOptions.Show vbModal
End Sub

Sub OpenSolver_SolverOptions(Optional Control)
27850     If Not CheckWorksheetAvailable Then Exit Sub
27860     frmSolverChange.Show
End Sub

Sub OpenSolver_SolveRelaxationClickHandler(Optional Control)
27870     If Not CheckWorksheetAvailable Then Exit Sub
27880     RunOpenSolver True
End Sub

Sub OpenSolver_LaunchCBCCommandLine(Optional Control)
          ' Open the CBC solver with our last model loaded.
          ' If we have a worksheet open with a model, then we pass the solver options (max runtime etc) from this model to CBC. Otherwise, we don't pass any options.
27890     On Error GoTo errorHandler
          Dim errorPrefix  As String
27900     errorPrefix = ""
            
          Dim WorksheetAvailable As Boolean
27910     WorksheetAvailable = CheckWorksheetAvailable(SuppressDialogs:=True)
          
          Dim ExternalSolverPathName As String, Solver As String
27920     Solver = "cbc.exe"
27930     GetExternalSolverPathName ExternalSolverPathName, Solver ' May throw an error if no solver can be found
          
          Dim ModelFileName As String
27940     ModelFileName = GetModelFileName
          Dim ModelFilePathName As String
27950     ModelFilePathName = GetModelFullPath
          
          ' Get all the options that we pass to CBC when we solve the problem and pass them here as well
          ' Get the Solver Options, stored in named ranges with values such as "=0.12"
          ' Because these are NAMEs, they are always in English, not the local language, so get their value using Val
          Dim SolveOptions As SolveOptionsType, ErrorString As String, SolveOptionsString As String
27960     If WorksheetAvailable Then
27970         GetSolveOptions "'" & Replace(ActiveSheet.Name, "'", "''") & "'!", SolveOptions, ErrorString ' NB: We have to double any ' when we quote the sheet name
27980         If ErrorString = "" Then
27990            SolveOptionsString = " -ratioGap " & str(SolveOptions.Tolerance) & " -seconds " & str(SolveOptions.maxTime)
28000         End If
28010     End If
          
          Dim CBCExtraParametersString As String
28020     If WorksheetAvailable Then
28030         CBCExtraParametersString = GetCBCExtraParametersString(ActiveSheet, ErrorString)
28040         If ErrorString <> "" Then CBCExtraParametersString = ""
28050     End If
             
          Dim CBCRunString As String
28060     CBCRunString = " -directory " & GetTempFolder _
                           & " -import " & ModelFileName _
                           & SolveOptionsString _
                           & CBCExtraParametersString _
                           & " -" ' Force CBC to accept commands from the command line
28070     OSSolveSync ExternalSolverPathName, CBCRunString, "", SW_SHOWNORMAL, False

ExitSub:
28080     Exit Sub
errorHandler:
28090     MsgBox "OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source, , "OpenSolver Code Error"
28100     Resume ExitSub

End Sub

Sub OpenSolver_ShowHideModelClickHandler(Optional Control)
28110     If Not CheckWorksheetAvailable Then Exit Sub
          Dim sheet As Worksheet
28120     On Error GoTo ExitSub
28130     Set sheet = ActiveSheet
          'If SheetHasOpenSolverDataHighlighting(sheet) Then
          '    HideSolverModel ' Hide the OpenSolverStudio data highlighting, and then show the model
          '    ShowSolverModel
          'ElseIf SheetHasOpenSolverModelHighlighting(sheet) Then
28140     If SheetHasOpenSolverHighlighting(sheet) Then
28150         HideSolverModel
28160     Else
28170         ShowSolverModel
28180     End If
ExitSub:
End Sub

Sub OpenSolver_SetQuickSolveParametersClickHandler(Optional Control)
28190     If Not CheckWorksheetAvailable Then Exit Sub
28200     If UserSetQuickSolveParameterRange Then
28210         Set OpenSolver = Nothing ' Was: OpenSolver.ClearQuickSolve  ' Reset any pre-initialized quicksolve data
28220     End If
End Sub

Sub OpenSolver_InitQuickSolveClickHandler(Optional Control)
28230     If Not CheckWorksheetAvailable Then Exit Sub
28240     InitializeQuickSolve
End Sub

Sub OpenSolver_QuickSolveClickHandler(Optional Control)
28250     If Not CheckWorksheetAvailable Then Exit Sub
28260     RunQuickSolve
End Sub

Sub OpenSolver_ViewLastModelClickHandler(Optional Control)
28270     On Error GoTo errorHandler
28280     If FileExists(GetModelFullPath) = "" Then
28290         MsgBox "Error: There is no model .lp file (" & GetModelFullPath & ") to open. Please solve the model using one of the linear solvers within OpenSolver, and then try again.", , "OpenSolver" & sOpenSolverVersion & " Error"
28300     Else
              ' Check that there is no workbook open with the same name
              Dim w As Workbook
28310         On Error Resume Next
28320         Set w = Workbooks(GetModelFileName)
28330         On Error GoTo errorHandler
              ' If w Is Nothing Then
28340             Workbooks.Open FileName:=GetModelFullPath, ReadOnly:=True ' , Format:=Tabs
                  'Workbooks.OpenText Filename:=GetModelFullPath, Origin:=xlMSDOS _
                      , StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array(0, 2), TrailingMinusNumbers:=True, Read
              'Else
              '    msgbox "Error: A model file is already open. Please close the '" & GetModelFileName & "' workbook and try again."
              'End If
28350     End If
ExitSub:
28360     Exit Sub
errorHandler:
28370     MsgBox "OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source, , "OpenSolver Code Error"
28380     Resume ExitSub
End Sub

Sub OpenSolver_ViewLogFile(Optional Control)
28390     On Error GoTo errorHandler
          Dim logPath As String
28400     logPath = GetTempFolder & "log1.tmp"
28410     If FileExists(logPath) = "" Then
28420         MsgBox "Error: There is no log file (" & logPath & ") to open. Please re-solve the OpenSolver model, and then try again.", , "OpenSolver" & sOpenSolverVersion & " Error"
28430     Else
              ' Check that there is no workbook open with the same name
              Dim w As Workbook
28440         On Error Resume Next
28450         Set w = Workbooks("log1.tmp")
28460         On Error GoTo errorHandler
              ' If w Is Nothing Then
28470         Workbooks.Open FileName:=logPath, ReadOnly:=True ' , Format:=Tabs
                  'Workbooks.OpenText Filename:=GetModelFullPath, Origin:=xlMSDOS _
                      , StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array(0, 2), TrailingMinusNumbers:=True, Read
              'Else
              '    msgbox "Error: A model file is already open. Please close the '" & GetModelFileName & "' workbook and try again."
              'End If
28480     End If
ExitSub:
28490     Exit Sub
errorHandler:
28500     MsgBox "OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source, , "OpenSolver Code Error"
End Sub

Sub OpenSolver_ViewLastSolutionClickHandler(Optional Control)
28510     On Error GoTo errorHandler
28520     If FileExists(GetSolutionFullPath) = "" Then
28530         MsgBox "Error: There is no solution file (" & GetSolutionFullPath & ") to open. Please solve the model using the CBC solver for OpenSolver, and then try again. Or if you solved your model using a different solver try opening that file instead.", , "OpenSolver" & sOpenSolverVersion & " Error"
28540     Else
              ' Check that there is no workbook open with the same name
              Dim w As Workbook
28550         On Error Resume Next
28560         Set w = Workbooks(GetSolutionFileName)
28570         On Error GoTo errorHandler
              'If w Is Nothing Then
28580             Workbooks.Open FileName:=GetSolutionFullPath, ReadOnly:=True
                  ' Workbooks.OpenText Filename:=GetSolutionFullPath, Origin:=xlMSDOS _
                      , StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array(0, 2), TrailingMinusNumbers:=True
              'Else
              '    msgbox "Error: A solution file is already open. Please close the '" & GetSolutionFileName & "' workbook and try again."
              'End If
28590     End If
ExitSub:
28600     Exit Sub
errorHandler:
28610     MsgBox "OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source, , "OpenSolver Code Error"
28620     Resume ExitSub
End Sub

Sub OpenSolver_ViewLastGurobiSolutionClickHandler(Optional Control)
28630     On Error GoTo errorHandler
          Dim GurobiSolutionPath As String
28640     GurobiSolutionPath = Replace(GetSolutionFullPath, ".txt", ".sol")
28650     If FileExists(GurobiSolutionPath) = "" Then
28660         MsgBox "Error: There is no solution file (" & GurobiSolutionPath & ") to open. Please solve the model using the Gurobi solver for OpenSolver, and then try again. Or if you solved your model using a different solver try opening that file instead.", , "OpenSolver" & sOpenSolverVersion & " Error"
28670     Else
               ' Check that there is no workbook open with the same name
               Dim w As Workbook
28680         On Error Resume Next
              Dim GurobiSolutionName As String
28690         GurobiSolutionName = Replace(GetSolutionFileName, ".txt", ".sol")
28700         Set w = Workbooks(GurobiSolutionPath)
28710         On Error GoTo errorHandler
               'If w Is Nothing Then
28720             Workbooks.Open FileName:=GurobiSolutionPath, ReadOnly:=True
                   ' Workbooks.OpenText Filename:=GetSolutionFullPath, Origin:=xlMSDOS _
                       , StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array(0, 2), TrailingMinusNumbers:=True
               'Else
               '    msgbox "Error: A solution file is already open. Please close the '" & GetSolutionFileName & "' workbook and try again."
               'End If
28730     End If
ExitSub:
28740     Exit Sub
errorHandler:
28750     MsgBox "OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source, , "OpenSolver Code Error"
28760     Resume ExitSub
End Sub

Sub OpenSolver_OnlineHelp(Optional Control)
28770     Call OpenURL("http://help.opensolver.org")
End Sub

Sub OpenSolver_AboutClickHandler(Optional Control)
28780     UserFormAbout.Show
End Sub

Sub OpenSolver_AboutCoinOR(Optional Control)
28790     MsgBox "COIN-OR" & vbCrLf & _
                 "http://www.Coin-OR.org" & vbCrLf & _
                 vbCrLf & _
                 "The Computational Infrastructure for Operations Research (COIN-OR, or simply COIN)  project is an initiative to spur the development of open-source software for the operations research community." & vbCrLf & _
                 vbCrLf & _
                 "OpenSolver uses the Coin-OR CBC optimization engine. CBC is licensed under the Common Public License 1.0. Visit the web sites for more information."
End Sub

Sub OpenSolver_VisitOpenSolverOrg(Optional Control)
28800     Call OpenURL("http://www.opensolver.org")
End Sub

Sub OpenSolver_VisitCoinOROrg(Optional Control)
28810     Call OpenURL("http://www.coin-or.org")
End Sub

Sub AutoOpenSolver()
    'If Not solver.Solver1.AutoOpened Then
    'solver.Solver2.Auto_Open
    'End If
End Sub

Function RunOpenSolver(Optional SolveRelaxation As Boolean = False, Optional MinimiseUserInteraction As Boolean = False) As OpenSolverResult
28820     On Error GoTo errorHandler

          'Save iterative calcalation state
          Dim oldIterationMode As Boolean
          oldIterationMode = Application.Iteration

28830     RunOpenSolver = OpenSolverResult.Unsolved
28840     Set OpenSolver = New COpenSolver
28850     OpenSolver.BuildModelFromSolverData
          If OpenSolver.Solver Like "PuLP" Or OpenSolver.Solver Like "*Cou*" Or OpenSolver.Solver Like "*Bon*" Then
              GoTo Tokeniser
          End If
28860     RunOpenSolver = OpenSolver.SolveModel(SolveRelaxation)
28870     If Not MinimiseUserInteraction Then OpenSolver.ReportAnySolutionSubOptimality
28880     Set OpenSolver = Nothing    ' Free any OpenSolver memory used
          Application.Iteration = oldIterationMode
28890     Exit Function
Tokeniser:
    On Error GoTo errHandle
    
    Dim TokenSolver As New CModel2
    TokenSolver.Setup ActiveWorkbook, ActiveSheet
    TokenSolver.ProcessSolverModel
    modPuLP.GenerateFile TokenSolver, OpenSolver.Solver, True
    Application.Iteration = oldIterationMode
    Exit Function
    
errHandle:
    MsgBox "An error occurred while trying build the model:" + vbNewLine _
            + "Description: " + Err.Description, vbOKOnly
    Application.Iteration = oldIterationMode
    Exit Function
errorHandler:
28900     Set OpenSolver = Nothing    ' Free any OpenSolver memory used
          Application.Iteration = oldIterationMode
28910     RunOpenSolver = OpenSolverResult.ErrorOccurred
28920     If Err.Number <> OpenSolver_UserCancelledError And Not MinimiseUserInteraction Then
28930         MsgBox "OpenSolver" & sOpenSolverVersion & " encountered an error:" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & vbCrLf & "Source = " & Err.Source & ", ErrNumber=" & Err.Number, , "OpenSolver" & sOpenSolverVersion & " Error"
28940     End If
End Function

'Sub BuildOpenSolverModel()
'    Set OpenSolver = New COpenSolver
'    OpenSolver.BuildModelFromSolverData
'End Sub

Sub InitializeQuickSolve()
28950     On Error GoTo errorHandler
28960     If CheckModelHasParameterRange Then
28970         Set OpenSolver = New COpenSolver
28980         OpenSolver.BuildModelFromSolverData
28990         OpenSolver.InitializeQuickSolve
29000     End If
29010     Exit Sub
errorHandler:
29020     MsgBox "OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source, , "OpenSolver" & sOpenSolverVersion & " Error"
End Sub

Sub RunQuickSolve()
29030     On Error GoTo errorHandler
29040     If OpenSolver Is Nothing Then
29050         MsgBox "Error: There is no model to solve, and so the quick solve cannot be completed. Please select the Initialize Quick Solve command.", , "OpenSolver" & sOpenSolverVersion & " Error"
29060     ElseIf OpenSolver.CanDoQuickSolveForActiveSheet Then    ' This will report any errors
29070         OpenSolver.DoQuickSolve
29080     End If
29090     Exit Sub
errorHandler:
29100     MsgBox "OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source, , "OpenSolver" & sOpenSolverVersion & " Error"
End Sub

Sub OpenSolver_ModelClick(Optional Control)
          'frmAutoModel.Show
          'frmAutoModel2.Show
29110     If Not CheckWorksheetAvailable Then Exit Sub
29120     frmModel.Show
29130     DoEvents
End Sub

Sub OpenSolver_QuickAutoModelClick(Optional Control)
29140     If Not CheckWorksheetAvailable Then Exit Sub
          Dim model As New CModel
29150     If Not CheckWorksheetAvailable Then Exit Sub
29160     If Not model.FindObjective(ActiveSheet) = "OK" Then
29170         MsgBox "Couldn't find objective, and couldn't finish as a result."
29180         Exit Sub
29190     End If
29200     If Not model.FindVarsAndCons(True) Then
29210         MsgBox "Error while looking for variables and constraints"
29220         Exit Sub
29230     End If
29240     model.NonNegativityAssumption = True
29250     model.BuildModel
29260     If MsgBox("Done! Show model?", vbYesNo, "Quick AutoModel") = vbYes Then
29270         OpenSolverVisualizer.ShowSolverModel
29280     Else
29290         OpenSolverVisualizer.HideSolverModel
29300     End If
End Sub

Sub OpenSolver_AutoModelAndSolveClick(Optional Control)
29310     If Not CheckWorksheetAvailable Then Exit Sub
          Dim model As New CModel
          Dim status As String
          
29320     status = model.FindObjective(ActiveSheet)
                
          ' Pass it the model reference
29330     Set frmAutoModel.model = model
29340     frmAutoModel.GuessObjStatus = status

      '--------------No Longer requires an objective to solve------------------------------
      '    If Not model.FindObjective(ActiveSheet) = "OK" Then
      ''        frmAutoModel.Show vbModal
      '         MsgBox "Couldn't find objective, and couldn't finish as a result. Check you have used the key words 'min', 'max' or 'target'."
      '         Exit Sub
      '    End If
      '-------------------------------------------------------------------------------------
          
29350     If Not model.FindVarsAndCons(True) Then
29360         MsgBox "Error while looking for variables and constraints"
29370         GoTo Viewer
29380     End If
29390     model.NonNegativityAssumption = True
29400     model.BuildModel
          
29410     RunOpenSolver False
         
Viewer:
29420     OpenSolverVisualizer.ShowSolverModel

End Sub
'====================================================================
' Excel 2003 Menu Code
' Provided by Paul Becker of Eclipse Engineering (www.eclipseeng.com)
'====================================================================
Public Sub AddMenuItems()
         
         Dim intHelpMenu      As Integer
         Dim objMainMenuBar   As CommandBar
         Dim objCustomMenu    As CommandBarControl
         
29430    Call DelMenuItems
         
29440    Set objMainMenuBar = Application.CommandBars("Worksheet Menu Bar")
         
         'intHelpMenu = objMainMenuBar.Controls("Help").Index
         
         'Set objCustomMenu = objMainMenuBar.Controls.Add(Type:=msoControlPopup, Before:=intHelpMenu)
29450    Set objCustomMenu = objMainMenuBar.Controls.Add(Type:=msoControlPopup)
         
29460    objCustomMenu.Caption = strMenuName
         
         
         'Main menu items
29470    With objCustomMenu.Controls.Add(Type:=msoControlButton)
29480       .Caption = "&Model"
29490       .OnAction = "OpenSolver_ModelClick"
29500       .FaceId = 0
29510    End With
         
29520    With objCustomMenu.Controls.Add(Type:=msoControlButton)
29530       .Caption = "&Solve"
29540       .OnAction = "OpenSolver_SolveClickHandler"
29550       .FaceId = 0
29560    End With
         
29570    With objCustomMenu.Controls.Add(Type:=msoControlButton)
29580       .Caption = "Show/&Hide Model"
29590       .OnAction = "OpenSolver_ShowHideModelClickHandler"
29600       .FaceId = 0
29610    End With
         
29620    With objCustomMenu.Controls.Add(Type:=msoControlButton)
29630       .Caption = "&Quick Solve"
29640       .OnAction = "OpenSolver_QuickSolveClickHandler"
29650       .FaceId = 0
29660    End With
         
         
         'Sub menu items
29670    Set objCustomMenu = objCustomMenu.Controls.Add(Type:=msoControlPopup)
29680    objCustomMenu.Caption = strAddInName
29690    objCustomMenu.BeginGroup = True
         
29700    With objCustomMenu.Controls.Add(Type:=msoControlButton)
29710       .Caption = "Set Quick Solve Parameters..."
29720       .OnAction = "OpenSolver_SetQuickSolveParametersClickHandler"
29730       .FaceId = 0
29740       .BeginGroup = True
29750    End With
         
29760    With objCustomMenu.Controls.Add(Type:=msoControlButton)
29770       .Caption = "Initialize Quick Solve"
29780       .OnAction = "OpenSolver_InitQuickSolveClickHandler"
29790       .FaceId = 0
29800    End With
         
         
29810    With objCustomMenu.Controls.Add(Type:=msoControlButton)
29820       .Caption = "Solve LP Relaxation"
29830       .OnAction = "OpenSolver_SolveRelaxationClickHandler"
29840       .FaceId = 0
29850       .BeginGroup = True
29860    End With
         
29870    With objCustomMenu.Controls.Add(Type:=msoControlButton)
29880       .Caption = "View Last Model .lp File"
29890       .OnAction = "OpenSolver_ViewLastModelClickHandler"
29900       .FaceId = 0
29910    End With
         
29920    With objCustomMenu.Controls.Add(Type:=msoControlButton)
29930       .Caption = "View Last CBC Solution File"
29940       .OnAction = "OpenSolver_ViewLastSolutionClickHandler"
29950       .FaceId = 0
29960    End With
         
29970    With objCustomMenu.Controls.Add(Type:=msoControlButton)
29980       .Caption = "Open Last Model in CBC..."
29990       .OnAction = "OpenSolver_LaunchCBCCommandLine"
30000       .FaceId = 0
30010    End With
         
         
         
30020    With objCustomMenu.Controls.Add(Type:=msoControlButton)
30030       .Caption = "Online Help..."
30040       .OnAction = "OpenSolver_OnlineHelp"
30050       .FaceId = 0 '984 '49
30060       .BeginGroup = True
30070    End With
         
         
30080    With objCustomMenu.Controls.Add(Type:=msoControlButton)
30090       .Caption = "About " & strAddInName & "..."
30100       .OnAction = "OpenSolver_AboutClickHandler"
30110       .FaceId = 0
30120       .BeginGroup = True
30130    End With
         
30140    With objCustomMenu.Controls.Add(Type:=msoControlButton)
30150       .Caption = "About COIN-OR..."
30160       .OnAction = "OpenSolver_AboutCoinOr"
30170       .FaceId = 0
30180    End With
         
         
30190    With objCustomMenu.Controls.Add(Type:=msoControlButton)
30200       .Caption = "Open " & strAddInName & ".org..."
30210       .OnAction = "OpenSolver_VisitOpenSolverOrg"
30220       .FaceId = 0
30230       .BeginGroup = True
30240    End With
         
30250    With objCustomMenu.Controls.Add(Type:=msoControlButton)
30260       .Caption = "Open COIN-OR.org..."
30270       .OnAction = "OpenSolver_VisitCoinOrOrg"
30280       .FaceId = 0
30290    End With
         
End Sub

Public Sub DelMenuItems()
30300     On Error Resume Next
30310     Application.CommandBars("Worksheet Menu Bar").Controls(strMenuName).Delete
End Sub
'====================================================================
' Excel 2003 Menu Code
' Provided by Paul Becker of Eclipse Engineering (www.eclipseeng.com)
'====================================================================
