Attribute VB_Name = "OpenSolverMain"
Option Explicit

' Version information (as displayed in the About box...)
Public Const sOpenSolverVersion As String = "2.6.1"
Public Const sOpenSolverDate As String = "2015.02.15"

' Used for the 2003 and Mac menu code
Private Const strAddInName As String = "OpenSolver"
Private Const strMenuName  As String = "&OpenSolver"

Dim OpenSolver As COpenSolver

Sub OpenSolver_SolveClickHandler(Optional Control)
2754      If Not CheckWorksheetAvailable Then Exit Sub
2755      RunOpenSolver False
End Sub

Sub OpenSolver_ModelOptions(Optional Control)
2756      If Not CheckWorksheetAvailable Then Exit Sub
2757      frmOptions.Show
End Sub

Sub OpenSolver_SolverOptions(Optional Control)
2759      If Not CheckWorksheetAvailable Then Exit Sub
2761      frmSolverChange.Show
End Sub

Sub OpenSolver_SolveRelaxationClickHandler(Optional Control)
2762      If Not CheckWorksheetAvailable Then Exit Sub
2763      RunOpenSolver True
End Sub

Sub OpenSolver_LaunchCBCCommandLine(Optional Control)
2764      LaunchCommandLine_CBC
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
ExitSub:
End Sub

Sub OpenSolver_SetQuickSolveParametersClickHandler(Optional Control)
2773      If Not CheckWorksheetAvailable Then Exit Sub
2774      If UserSetQuickSolveParameterRange Then
2775          Set OpenSolver = Nothing
2776      End If
End Sub

Sub OpenSolver_InitQuickSolveClickHandler(Optional Control)
2777      If Not CheckWorksheetAvailable Then Exit Sub
2778      InitializeQuickSolve
End Sub

Sub OpenSolver_QuickSolveClickHandler(Optional Control)
2779      If Not CheckWorksheetAvailable Then Exit Sub
2780      RunQuickSolve
End Sub

Sub OpenSolver_ViewLastModelClickHandler(Optional Control)
          Dim notFoundMessage As String, FilePath As String
2781      FilePath = GetTempFilePath(LPFileName)
2782      notFoundMessage = "Error: There is no LP file (" & FilePath & ") to open. Please solve the model using one of the linear solvers within OpenSolver, and then try again."
2783      OpenFile FilePath, notFoundMessage
End Sub

Sub OpenSolver_ViewLastAmplClickHandler(Optional Control)
          Dim notFoundMessage As String, FilePath As String
2784      FilePath = GetTempFilePath(AMPLFileName)
2785      notFoundMessage = "Error: There is no AMPL file (" & FilePath & ") to open. Please solve the model using one of the NEOS solvers within OpenSolver, and then try again."
2786      OpenFile FilePath, notFoundMessage
End Sub

Sub OpenSolver_ViewLogFile(Optional Control)
          Dim notFoundMessage As String, FilePath As String
2787      FilePath = GetTempFilePath("log1.tmp")
2788      notFoundMessage = "Error: There is no log file (" & FilePath & ") to open. Please re-solve the OpenSolver model, and then try again."
2789      OpenFile FilePath, notFoundMessage
End Sub

Sub OpenSolver_ViewLastSolutionClickHandler(Optional Control)
          Dim notFoundMessage As String, FilePath As String
2790      FilePath = SolutionFilePath_CBC()
2791      notFoundMessage = "Error: There is no solution file (" & FilePath & ") to open. Please solve the model using the CBC solver for OpenSolver, and then try again. Or if you solved your model using a different solver try opening that file instead."
2792      OpenFile FilePath, notFoundMessage
End Sub

Sub OpenSolver_ViewLastGurobiSolutionClickHandler(Optional Control)
          Dim notFoundMessage As String, FilePath As String
2793      FilePath = SolutionFilePath_Gurobi()
2794      notFoundMessage = "Error: There is no solution file (" & FilePath & ") to open. Please solve the model using the Gurobi solver for OpenSolver, and then try again. Or if you solved your model using a different solver try opening that file instead."
2795      OpenFile FilePath, notFoundMessage
End Sub

Sub OpenSolver_OnlineHelp(Optional Control)
2796      Call OpenURL("http://help.opensolver.org")
End Sub

Sub OpenSolver_AboutClickHandler(Optional Control)
2798      frmAbout.Show
End Sub

Sub OpenSolver_AboutCoinOR(Optional Control)
2799      MsgBox "COIN-OR" & vbCrLf & _
                 "http://www.Coin-OR.org" & vbCrLf & _
                 vbCrLf & _
                 "The Computational Infrastructure for Operations Research (COIN-OR, or simply COIN)  project is an initiative to spur the development of open-source software for the operations research community." & vbCrLf & _
                 vbCrLf & _
                 "OpenSolver uses the Coin-OR CBC optimization engine. CBC is licensed under the Common Public License 1.0. Visit the web sites for more information."
End Sub

Sub OpenSolver_VisitOpenSolverOrg(Optional Control)
2800      Call OpenURL("http://www.opensolver.org")
End Sub

Sub OpenSolver_VisitCoinOROrg(Optional Control)
2801      Call OpenURL("http://www.coin-or.org")
End Sub

Function RunOpenSolver(Optional SolveRelaxation As Boolean = False, Optional MinimiseUserInteraction As Boolean = False, Optional LinearityCheckOffset As Double = 0) As OpenSolverResult
          On Error GoTo ErrorHandler

          'Save iterative calcalation state
          Dim oldIterationMode As Boolean
2803      oldIterationMode = Application.Iteration

2804      RunOpenSolver = OpenSolverResult.Unsolved
2805      Set OpenSolver = New COpenSolver
2806      OpenSolver.BuildModelFromSolverData LinearityCheckOffset, MinimiseUserInteraction, SolveRelaxation

          ' Run appropriate solve routine
          Dim OpenSolverParsed As COpenSolverParsed
2807      If UsesParsedModel(OpenSolver.Solver) Then
              Set OpenSolverParsed = New COpenSolverParsed
              OpenSolverParsed.SolveModel OpenSolver, SolveRelaxation, MinimiseUserInteraction
              RunOpenSolver = OpenSolver.SolveStatus
2809      Else
2810          RunOpenSolver = OpenSolver.SolveModel(SolveRelaxation, MinimiseUserInteraction)
2811      End If

          If Not MinimiseUserInteraction Then OpenSolver.ReportAnySolutionSubOptimality

ExitFunction:
          Set OpenSolver = Nothing    ' Free any OpenSolver memory used
          Set OpenSolverParsed = Nothing
          Application.Iteration = oldIterationMode
          Exit Function

ErrorHandler:
          ReportError "OpenSolverMain", "RunOpenSolver", True, MinimiseUserInteraction
          If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then
              RunOpenSolver = AbortedThruUserAction
          Else
              RunOpenSolver = OpenSolverResult.ErrorOccurred
          End If
          GoTo ExitFunction
End Function

Sub InitializeQuickSolve()
          On Error GoTo ErrorHandler

          If UsesParsedModel(GetChosenSolver) Then
              Err.Raise OpenSolver_ModelError, Description:="The selected solver does not support QuickSolve"
          End If

2832      If CheckModelHasParameterRange Then
2833          Set OpenSolver = New COpenSolver
2834          OpenSolver.BuildModelFromSolverData
2835          OpenSolver.InitializeQuickSolve
2836      End If

ExitSub:
          Exit Sub

ErrorHandler:
          ReportError "OpenSolverMain", "InitializeQuickSolve", True
          GoTo ExitSub
End Sub

Function RunQuickSolve(Optional MinimiseUserInteraction As Boolean = False) As Long
          On Error GoTo ErrorHandler

2840      If OpenSolver Is Nothing Then
2841          MsgBox "Error: There is no model to solve, and so the quick solve cannot be completed. Please select the Initialize Quick Solve command.", , "OpenSolver" & sOpenSolverVersion & " Error"
              'MsgBoxEx "Help_Error: There is no model to solve, and so the quick solve cannot be completed. Please select the Initialize Quick Solve command."
2842          RunQuickSolve = OpenSolverResult.ErrorOccurred
2843      ElseIf OpenSolver.CanDoQuickSolveForActiveSheet Then    ' This will report any errors
2844          RunQuickSolve = OpenSolver.DoQuickSolve(MinimiseUserInteraction)
2845      End If

ExitFunction:
          Exit Function

ErrorHandler:
          ReportError "OpenSolverMain", "RunQuickSolve", True, MinimiseUserInteraction
          If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then
              RunQuickSolve = AbortedThruUserAction
          Else
              RunQuickSolve = OpenSolverResult.ErrorOccurred
          End If
          GoTo ExitFunction
End Function

Sub OpenSolver_ModelClick(Optional Control)
2851      If Not CheckWorksheetAvailable Then Exit Sub
2853      frmModel.Show
2854      DoEvents
End Sub

Sub OpenSolver_QuickAutoModelClick(Optional Control)
2855      RunAutoModel False
End Sub

Sub OpenSolver_AutoModelAndSolveClick(Optional Control)
          If Not RunAutoModel(False) Then Exit Sub
2882      RunOpenSolver False
End Sub
'====================================================================
' Excel 2003 Menu Code
' Provided by Paul Becker of Eclipse Engineering (www.eclipseeng.com)
'====================================================================
Public Sub AddMenuItems()
         
          Dim intHelpMenu      As Long
          Dim objMainMenuBar   As CommandBar
          Dim objCustomMenu    As CommandBarControl
          Dim objCustomMenu2   As CommandBarControl
          
2884      Call DelMenuItems
         
2885      Set objMainMenuBar = Application.CommandBars("Worksheet Menu Bar")
2886      Set objCustomMenu = objMainMenuBar.Controls.Add(Type:=msoControlPopup)
         
2887      objCustomMenu.Caption = strMenuName
         
         
          'Model menu items
2888      Set objCustomMenu2 = objCustomMenu.Controls.Add(Type:=msoControlPopup)
2889      objCustomMenu2.Caption = "&Model"
2890      objCustomMenu2.BeginGroup = True
         
2891      With objCustomMenu2.Controls.Add(Type:=msoControlButton)
2892         .Caption = "&Model..."
2893         .OnAction = "OpenSolver_ModelClick"
2894         .FaceId = 0
2895      End With

2896      With objCustomMenu2.Controls.Add(Type:=msoControlButton)
2897         .Caption = "Quick AutoModel"
2898         .OnAction = "OpenSolver_QuickAutoModelClick"
2899         .FaceId = 0
2900      End With

2901      With objCustomMenu2.Controls.Add(Type:=msoControlButton)
2902         .Caption = "&AutoModel and Solve"
2903         .OnAction = "OpenSolver_AutoModelAndSolveClick"
2904         .FaceId = 0
2905      End With

2906      With objCustomMenu2.Controls.Add(Type:=msoControlButton)
2907         .Caption = "&Solver Engine..."
2908         .OnAction = "OpenSolver_SolverOptions"
2909         .FaceId = 0
2910      End With

2911      With objCustomMenu2.Controls.Add(Type:=msoControlButton)
2912         .Caption = "&Options..."
2913         .OnAction = "OpenSolver_ModelOptions"
2914         .FaceId = 0
2915      End With
         
          ' Main menu items
2916      With objCustomMenu.Controls.Add(Type:=msoControlButton)
2917         .Caption = "&Solve"
2918         .OnAction = "OpenSolver_SolveClickHandler"
2919         .FaceId = 0
2920      End With
         
2921      With objCustomMenu.Controls.Add(Type:=msoControlButton)
2922         .Caption = "Show/&Hide Model"
2923         .OnAction = "OpenSolver_ShowHideModelClickHandler"
2924         .FaceId = 0
2925      End With
         
2926      With objCustomMenu.Controls.Add(Type:=msoControlButton)
2927         .Caption = "&Quick Solve"
2928         .OnAction = "OpenSolver_QuickSolveClickHandler"
2929         .FaceId = 0
2930      End With
         
         
          'OpenSolver menu items
2931      Set objCustomMenu = objCustomMenu.Controls.Add(Type:=msoControlPopup)
2932      objCustomMenu.Caption = strAddInName
2933      objCustomMenu.BeginGroup = True
         
2934      With objCustomMenu.Controls.Add(Type:=msoControlButton)
2935         .Caption = "Set Quick Solve Parameters..."
2936         .OnAction = "OpenSolver_SetQuickSolveParametersClickHandler"
2937         .FaceId = 0
2938         .BeginGroup = True
2939      End With
         
2940      With objCustomMenu.Controls.Add(Type:=msoControlButton)
2941         .Caption = "Initialize Quick Solve"
2942         .OnAction = "OpenSolver_InitQuickSolveClickHandler"
2943         .FaceId = 0
2944      End With
         
         
2945      With objCustomMenu.Controls.Add(Type:=msoControlButton)
2946         .Caption = "Solve LP Relaxation"
2947         .OnAction = "OpenSolver_SolveRelaxationClickHandler"
2948         .FaceId = 0
2949         .BeginGroup = True
2950      End With
         
2951      With objCustomMenu.Controls.Add(Type:=msoControlButton)
2952         .Caption = "View Last Model .lp File"
2953         .OnAction = "OpenSolver_ViewLastModelClickHandler"
2954         .FaceId = 0
2955      End With
         
2956      With objCustomMenu.Controls.Add(Type:=msoControlButton)
2957         .Caption = "View Last AMPL File"
2958         .OnAction = "OpenSolver_ViewLastAmplClickHandler"
2959         .FaceId = 0
2960      End With
         
2961      With objCustomMenu.Controls.Add(Type:=msoControlButton)
2962         .Caption = "View Last Log File"
2963         .OnAction = "OpenSolver_ViewLogFile"
2964         .FaceId = 0
2965      End With
         
2966      With objCustomMenu.Controls.Add(Type:=msoControlButton)
2967         .Caption = "View Last CBC Solution File"
2968         .OnAction = "OpenSolver_ViewLastSolutionClickHandler"
2969         .FaceId = 0
2970         .BeginGroup = True
2971      End With
         
2972      With objCustomMenu.Controls.Add(Type:=msoControlButton)
2973         .Caption = "Open Last Model in CBC..."
2974         .OnAction = "OpenSolver_LaunchCBCCommandLine"
2975         .FaceId = 0
2976      End With
         
2977      With objCustomMenu.Controls.Add(Type:=msoControlButton)
2978         .Caption = "View Last Gurobi Solution File"
2979         .OnAction = "OpenSolver_ViewLastGurobiSolutionClickHandler"
2980         .FaceId = 0
2981         .BeginGroup = True
2982      End With
         
2983      With objCustomMenu.Controls.Add(Type:=msoControlButton)
2984         .Caption = "Online Help..."
2985         .OnAction = "OpenSolver_OnlineHelp"
2986         .FaceId = 0 '984 '49
2987         .BeginGroup = True
2988      End With
         
         
2989      With objCustomMenu.Controls.Add(Type:=msoControlButton)
2990         .Caption = "About " & strAddInName & "..."
2991         .OnAction = "OpenSolver_AboutClickHandler"
2992         .FaceId = 0
2993         .BeginGroup = True
2994      End With
         
2995      With objCustomMenu.Controls.Add(Type:=msoControlButton)
2996         .Caption = "About COIN-OR..."
2997         .OnAction = "OpenSolver_AboutCoinOr"
2998         .FaceId = 0
2999      End With
         
         
3000      With objCustomMenu.Controls.Add(Type:=msoControlButton)
3001         .Caption = "Open " & strAddInName & ".org..."
3002         .OnAction = "OpenSolver_VisitOpenSolverOrg"
3003         .FaceId = 0
3004         .BeginGroup = True
3005      End With
         
3006      With objCustomMenu.Controls.Add(Type:=msoControlButton)
3007         .Caption = "Open COIN-OR.org..."
3008         .OnAction = "OpenSolver_VisitCoinOrOrg"
3009         .FaceId = 0
3010      End With
End Sub

Public Sub DelMenuItems()
3011      On Error Resume Next
3012      Application.CommandBars("Worksheet Menu Bar").Controls(strMenuName).Delete
End Sub
'====================================================================
' Excel 2003 Menu Code
' Provided by Paul Becker of Eclipse Engineering (www.eclipseeng.com)
'====================================================================
