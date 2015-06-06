Attribute VB_Name = "OpenSolverMenu"
Option Explicit

' Used for legacy menu title
Private Const strAddInName As String = "OpenSolver"
Private Const strMenuName  As String = "&OpenSolver"

Sub AlterMenuItems(AddItems As Boolean)
          Dim NeedToAdd As Boolean
5         NeedToAdd = Application.Version = "11.0"
          #If Mac Then
6             NeedToAdd = True
          #End If
7         If NeedToAdd Then
8             If AddItems Then
9                 AddMenuItems
10            Else
11                DelMenuItems
12            End If
13        End If
End Sub

' Menu/ribbon click handlers
Sub OpenSolver_SolveClickHandler(Optional Control)
2754      If Not CheckWorksheetAvailable Then Exit Sub
2755      RunOpenSolver False, False, 0
          AutoUpdateCheck
End Sub

Sub OpenSolver_ModelOptions(Optional Control)
2756      If Not CheckWorksheetAvailable Then Exit Sub
          Dim frmOptions As FOptions
          Set frmOptions = New FOptions
2757      frmOptions.Show
          Unload frmOptions
          AutoUpdateCheck
End Sub

Sub OpenSolver_SolverOptions(Optional Control)
2759      If Not CheckWorksheetAvailable Then Exit Sub
          Dim frmSolverChange As FSolverChange
          Set frmSolverChange = New FSolverChange
2761      frmSolverChange.Show
          Unload frmSolverChange
          AutoUpdateCheck
End Sub

Sub OpenSolver_SolveRelaxationClickHandler(Optional Control)
2762      If Not CheckWorksheetAvailable Then Exit Sub
2763      RunOpenSolver True, False, 0
          AutoUpdateCheck
End Sub

Sub OpenSolver_LaunchCBCCommandLine(Optional Control)
2764      LaunchCommandLine_CBC
          AutoUpdateCheck
End Sub

Sub OpenSolver_ShowHideModelClickHandler(Optional Control)
2765      If Not CheckWorksheetAvailable Then Exit Sub
          Dim sheet As Worksheet
2766      On Error GoTo ExitSub
2767      Set sheet = ActiveSheet
2768      If SheetHasOpenSolverHighlighting(sheet) Then
2769          HideSolverModel
2770      Else
2771          ShowSolverModel
2772      End If
          AutoUpdateCheck
ExitSub:
End Sub

Sub OpenSolver_SetQuickSolveParametersClickHandler(Optional Control)
2773      If Not CheckWorksheetAvailable Then Exit Sub
2774      If SetQuickSolveParameterRange Then
2775          ClearQuickSolve
2776      End If
          AutoUpdateCheck
End Sub

Sub OpenSolver_InitQuickSolveClickHandler(Optional Control)
2777      If Not CheckWorksheetAvailable Then Exit Sub
2778      InitializeQuickSolve
          AutoUpdateCheck
End Sub

Sub OpenSolver_QuickSolveClickHandler(Optional Control)
2779      If Not CheckWorksheetAvailable Then Exit Sub
2780      RunQuickSolve
          AutoUpdateCheck
End Sub

Sub OpenSolver_ViewLastModelClickHandler(Optional Control)
          Dim notFoundMessage As String, FilePath As String
2781      GetLPFilePath FilePath
2782      notFoundMessage = "Error: There is no LP file (" & FilePath & ") to open. Please solve the model using one of the linear solvers within OpenSolver, and then try again."
2783      OpenFile FilePath, notFoundMessage
          AutoUpdateCheck
End Sub

Sub OpenSolver_ViewLastAmplClickHandler(Optional Control)
          Dim notFoundMessage As String, FilePath As String
2784      GetAMPLFilePath FilePath
2785      notFoundMessage = "Error: There is no AMPL file (" & FilePath & ") to open. Please solve the model using one of the NEOS solvers within OpenSolver, and then try again."
2786      OpenFile FilePath, notFoundMessage
          AutoUpdateCheck
End Sub

Sub OpenSolver_ViewLogFile(Optional Control)
          Dim notFoundMessage As String, FilePath As String
2787      GetLogFilePath FilePath
2788      notFoundMessage = "Error: There is no log file (" & FilePath & ") to open. Please re-solve the OpenSolver model, and then try again."
2789      OpenFile FilePath, notFoundMessage
          AutoUpdateCheck
End Sub

Sub OpenSolver_ViewLastSolutionClickHandler(Optional Control)
          Dim notFoundMessage As String, FilePath As String, CbcSolver As CSolverCbc
          Set CbcSolver = New CSolverCbc
2790      FilePath = CbcSolver.SolutionFilePath()
2791      notFoundMessage = "Error: There is no solution file (" & FilePath & ") to open. Please solve the model using the CBC solver for OpenSolver, and then try again. Or if you solved your model using a different solver try opening that file instead."
2792      OpenFile FilePath, notFoundMessage
          AutoUpdateCheck
End Sub

Sub OpenSolver_ViewLastGurobiSolutionClickHandler(Optional Control)
          Dim notFoundMessage As String, FilePath As String, GurobiSolver As CSolverGurobi
          Set GurobiSolver = New CSolverGurobi
2790      FilePath = GurobiSolver.SolutionFilePath()
2794      notFoundMessage = "Error: There is no solution file (" & FilePath & ") to open. Please solve the model using the Gurobi solver for OpenSolver, and then try again. Or if you solved your model using a different solver try opening that file instead."
2795      OpenFile FilePath, notFoundMessage
          AutoUpdateCheck
End Sub

Sub OpenSolver_OnlineHelp(Optional Control)
2796      OpenURL "http://help.opensolver.org"
          AutoUpdateCheck
End Sub

Sub OpenSolver_AboutClickHandler(Optional Control)
          Dim frmAbout As FAbout
          Set frmAbout = New FAbout
2798      frmAbout.Show
          Unload frmAbout
          AutoUpdateCheck
End Sub

Sub OpenSolver_AboutCoinOR(Optional Control)
2799      MsgBox "COIN-OR" & vbCrLf & _
                 "http://www.Coin-OR.org" & vbCrLf & _
                 vbCrLf & _
                 "The Computational Infrastructure for Operations Research (COIN-OR, or simply COIN)  project is an initiative to spur the development of open-source software for the operations research community." & vbCrLf & _
                 vbCrLf & _
                 "OpenSolver uses the Coin-OR CBC optimization engine. CBC is licensed under the Common Public License 1.0. Visit the web sites for more information."
          AutoUpdateCheck
End Sub

Sub OpenSolver_VisitOpenSolverOrg(Optional Control)
2800      OpenURL "http://www.opensolver.org"
          AutoUpdateCheck
End Sub

Sub OpenSolver_VisitCoinOROrg(Optional Control)
2801      OpenURL "http://www.coin-or.org"
          AutoUpdateCheck
End Sub
Sub OpenSolver_ModelClick(Optional Control)
2851      If Not CheckWorksheetAvailable Then Exit Sub
          Dim frmModel As FModel
          Set frmModel = New FModel
2853      frmModel.Show
          Unload frmModel
2854      DoEvents
          AutoUpdateCheck
End Sub

Sub OpenSolver_QuickAutoModelClick(Optional Control)
2855      RunAutoModel False
          AutoUpdateCheck
End Sub

Sub OpenSolver_AutoModelAndSolveClick(Optional Control)
          If Not RunAutoModel(False) Then Exit Sub
2882      RunOpenSolver False, False, 0
          AutoUpdateCheck
End Sub

'====================================================================
' Excel 2003 Menu Code
' Provided by Paul Becker of Eclipse Engineering (www.eclipseeng.com)
'====================================================================
Sub AddItemToMenu(Menu As CommandBarControl, Caption As String, Action As String, Optional BeginGroup As Boolean = False)
          With Menu.Controls.Add(Type:=msoControlButton)
             .Caption = Caption
             .OnAction = Action
             .BeginGroup = BeginGroup
             .FaceId = 0
          End With
End Sub

Public Sub AddMenuItems()
         
          Dim intHelpMenu       As Long
          Dim objMainMenuBar    As CommandBar
          Dim objCustomMenu     As CommandBarControl
          Dim objCustomSubMenu  As CommandBarControl
          
2884      DelMenuItems
         
2885      Set objMainMenuBar = Application.CommandBars("Worksheet Menu Bar")
2886      Set objCustomMenu = objMainMenuBar.Controls.Add(Type:=msoControlPopup)
2887      objCustomMenu.Caption = strMenuName
         
          'Model menu items
2888      Set objCustomSubMenu = objCustomMenu.Controls.Add(Type:=msoControlPopup)
2889      objCustomSubMenu.Caption = "&Model"
2890      objCustomSubMenu.BeginGroup = True
         
          AddItemToMenu objCustomSubMenu, "&Model...", "OpenSolver_ModelClick"
          AddItemToMenu objCustomSubMenu, "Quick AutoModel", "OpenSolver_QuickAutoModelClick"
          AddItemToMenu objCustomSubMenu, "&AutoModel and Solve", "OpenSolver_AutoModelAndSolveClick"
          AddItemToMenu objCustomSubMenu, "&Solver Engine...", "OpenSolver_SolverOptions"
          AddItemToMenu objCustomSubMenu, "&Options...", "OpenSolver_ModelOptions"

          ' Main menu items
          AddItemToMenu objCustomMenu, "&Solve", "OpenSolver_SolveClickHandler"
          AddItemToMenu objCustomMenu, "Show/&Hide Model", "OpenSolver_ShowHideModelClickHandler"
          AddItemToMenu objCustomMenu, "&Quick Solve", "OpenSolver_QuickSolveClickHandler"
         
          'OpenSolver menu items
2931      Set objCustomSubMenu = objCustomMenu.Controls.Add(Type:=msoControlPopup)
2932      objCustomSubMenu.Caption = strAddInName
2933      objCustomSubMenu.BeginGroup = True

          AddItemToMenu objCustomSubMenu, "Set Quick Solve Parameters...", "OpenSolver_SetQuickSolveParametersClickHandler", True
          AddItemToMenu objCustomSubMenu, "Initialize Quick Solve", "OpenSolver_InitQuickSolveClickHandler"
          
          AddItemToMenu objCustomSubMenu, "Solve LP Relaxation", "OpenSolver_SolveRelaxationClickHandler", True
          AddItemToMenu objCustomSubMenu, "View Last Model .lp File", "OpenSolver_ViewLastModelClickHandler"
          AddItemToMenu objCustomSubMenu, "View Last AMPL File", "OpenSolver_ViewLastAmplClickHandler"
          AddItemToMenu objCustomSubMenu, "View Last Log File", "OpenSolver_ViewLogFile"
          
          AddItemToMenu objCustomSubMenu, "View Last CBC Solution File", "OpenSolver_ViewLastSolutionClickHandler", True
          AddItemToMenu objCustomSubMenu, "Open Last Model in CBC...", "OpenSolver_LaunchCBCCommandLine"
          
          AddItemToMenu objCustomSubMenu, "View Last Gurobi Solution File", "OpenSolver_ViewLastGurobiSolutionClickHandler", True
          
          AddItemToMenu objCustomSubMenu, "Online Help...", "OpenSolver_OnlineHelp", True
          
          AddItemToMenu objCustomSubMenu, "About " & strAddInName & "...", "OpenSolver_AboutClickHandler", True
          AddItemToMenu objCustomSubMenu, "About COIN-OR...", "OpenSolver_AboutCoinOr"
          
          AddItemToMenu objCustomSubMenu, "Open " & strAddInName & ".org...", "OpenSolver_VisitOpenSolverOrg", True
          AddItemToMenu objCustomSubMenu, "Open COIN-OR.org...", "OpenSolver_VisitCoinOrOrg"
End Sub

Public Sub DelMenuItems()
3011      On Error Resume Next
3012      Application.CommandBars("Worksheet Menu Bar").Controls(strMenuName).Delete
End Sub
'====================================================================
' Excel 2003 Menu Code
' Provided by Paul Becker of Eclipse Engineering (www.eclipseeng.com)
'====================================================================

