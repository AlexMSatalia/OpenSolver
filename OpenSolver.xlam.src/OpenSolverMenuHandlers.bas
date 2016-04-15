Attribute VB_Name = "OpenSolverMenuHandlers"
Option Explicit

Sub OpenSolver_SolveClickHandler(Optional Control)
          Dim sheet As Worksheet
          GetActiveSheetIfMissing sheet
2755      RunOpenSolver SolveRelaxation:=False, MinimiseUserInteraction:=False, sheet:=sheet
          AutoUpdateCheck
End Sub

Sub OpenSolver_ModelOptions(Optional Control)
          Dim frmOptions As FOptions
          Set frmOptions = New FOptions
2757      frmOptions.Show
          Unload frmOptions
          AutoUpdateCheck
End Sub

Sub OpenSolver_SolverOptions(Optional Control)
          Dim frmSolverChange As FSolverChange
          Set frmSolverChange = New FSolverChange
2761      frmSolverChange.Show
          Unload frmSolverChange
          AutoUpdateCheck
End Sub

Sub OpenSolver_SolveRelaxationClickHandler(Optional Control)
          Dim sheet As Worksheet
          GetActiveSheetIfMissing sheet
2763      RunOpenSolver SolveRelaxation:=True, MinimiseUserInteraction:=False, sheet:=sheet
          AutoUpdateCheck
End Sub

Sub OpenSolver_LaunchCBCCommandLine(Optional Control)
          If Len(LastUsedSolver) = 0 Then
              MsgBox "Cannot open the last model in CBC as the model has not been solved yet."
          Else
              Dim Solver As ISolver
              Set Solver = CreateSolver(LastUsedSolver)
              If TypeOf Solver Is ISolverFile Then
                  Dim FileSolver As ISolverFile
                  Set FileSolver = Solver
                  If FileSolver.FileType = LP Then
2764                  LaunchCommandLine_CBC
                  Else
                      GoTo NotLPSolver
                  End If
              Else
NotLPSolver:
                  MsgBox "The last used solver (" & DisplayName(Solver) & ") does not use .lp model files, so CBC cannot load the model. " & _
                         "Please solve the model using a solver that uses .lp files, such as CBC or Gurobi, and try again."
              End If
          End If
          AutoUpdateCheck
End Sub

Sub OpenSolver_ShowHideModelClickHandler(Optional Control)
          Dim sheet As Worksheet
          GetActiveSheetIfMissing sheet
          
2766      On Error GoTo ExitSub
2768      If SheetHasOpenSolverHighlighting(sheet) Then
2769          HideSolverModel sheet
2770      Else
2771          ShowSolverModel sheet, HandleError:=True
2772      End If
          AutoUpdateCheck
ExitSub:
End Sub

Sub OpenSolver_SetQuickSolveParametersClickHandler(Optional Control)
2774      If SetQuickSolveParameterRange Then
2775          ClearQuickSolve
2776      End If
          AutoUpdateCheck
End Sub

Sub OpenSolver_InitQuickSolveClickHandler(Optional Control)
          Dim sheet As Worksheet
          GetActiveSheetIfMissing sheet
2778      InitializeQuickSolve sheet:=sheet
          AutoUpdateCheck
End Sub

Sub OpenSolver_QuickSolveClickHandler(Optional Control)
2780      RunQuickSolve
          AutoUpdateCheck
End Sub

Sub OpenSolver_ViewLastModelClickHandler(Optional Control)
          If Len(LastUsedSolver) = 0 Then
              MsgBox "Cannot open the last model file as the model has not been solved yet."
          Else
              Dim Solver As ISolver
              Set Solver = CreateSolver(LastUsedSolver)
              If TypeOf Solver Is ISolverFile Then
                  Dim NotFoundMessage As String, FilePath As String
                  FilePath = GetModelFilePath(Solver)
                  NotFoundMessage = "Error: There is no model file (" & FilePath & ") to open. Please solve the OpenSolver model and then try again."
                  OpenFile FilePath, NotFoundMessage
              Else
                  MsgBox "The last used solver (" & DisplayName(Solver) & ") does not use a model file."
              End If
          End If
          AutoUpdateCheck
End Sub

Sub OpenSolver_ViewSolverLogFileClickHandler(Optional Control)
          Dim NotFoundMessage As String, FilePath As String
2787      GetLogFilePath FilePath
2788      NotFoundMessage = "Error: There is no solver log file (" & FilePath & ") to open. Please solve the OpenSolver model and then try again."
2789      OpenFile FilePath, NotFoundMessage
          AutoUpdateCheck
End Sub

Sub OpenSolver_ViewErrorLogFileClickHandler(Optional Control)
          Dim NotFoundMessage As String, FilePath As String
2787      FilePath = GetErrorLogFilePath()
2788      NotFoundMessage = "Error: There is no error log file (" & FilePath & ") to open."
2789      OpenFile FilePath, NotFoundMessage
          AutoUpdateCheck
End Sub

Sub OpenSolver_ViewLastSolutionClickHandler(Optional Control)
          If Len(LastUsedSolver) = 0 Then
              MsgBox "Cannot open the last solution file as the model has not been solved yet."
          Else
              Dim Solver As ISolver
              Set Solver = CreateSolver(LastUsedSolver)
              If TypeOf Solver Is ISolverLocalExec Then
                  Dim NotFoundMessage As String, FilePath As String
2790              GetSolutionFilePath FilePath
2791              NotFoundMessage = "Error: There is no solution file (" & FilePath & ") to open. Please solve the OpenSolver model and then try again."
2792              OpenFile FilePath, NotFoundMessage
              Else
                  MsgBox "The last used solver (" & DisplayName(Solver) & ") does not produce a solution file. Please check the log file for any solution information."
              End If
          End If
          AutoUpdateCheck
End Sub

Sub OpenSolver_ViewTempFolderClickHandler(Optional Control)
    Dim NotFoundMessage As String, FolderPath As String
    FolderPath = GetTempFolder()
    NotFoundMessage = "Error: The OpenSolver temporary files folder (" & FolderPath & ") doesn't exist."
    OpenFolder FolderPath, NotFoundMessage
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
          Static ShownFormBefore As Boolean, ShowNamedRangesState As Boolean, ShowModelAfterSavingState As Boolean
          ' Set the checkboxes default to true
          If Not ShownFormBefore Then
              ShownFormBefore = True
              ShowNamedRangesState = True
              ShowModelAfterSavingState = True
          End If

          Dim frmModel As FModel
          Set frmModel = New FModel
          frmModel.ShowModelAfterSavingState = ShowModelAfterSavingState
          frmModel.ShowNamedRangesState = ShowNamedRangesState
2853      frmModel.Show
          ShowModelAfterSavingState = frmModel.ShowModelAfterSavingState
          ShowNamedRangesState = frmModel.ShowNamedRangesState
          Unload frmModel
          
2854      DoEvents
          AutoUpdateCheck
End Sub

Sub OpenSolver_QuickAutoModelClick(Optional Control)
          Dim sheet As Worksheet
          GetActiveSheetIfMissing sheet
          
          Dim AutoModel As CAutoModel
          Set AutoModel = New CAutoModel
          AutoModel.BuildModel sheet, MinimiseUserInteraction:=False, SaveAfterBuilding:=True
          
          AutoUpdateCheck
End Sub

Sub OpenSolver_AutoModelAndSolveClick(Optional Control)
          Dim sheet As Worksheet
          GetActiveSheetIfMissing sheet
          
          Dim AutoModel As CAutoModel
          Set AutoModel = New CAutoModel
          If Not AutoModel.BuildModel(sheet, MinimiseUserInteraction:=False, SaveAfterBuilding:=True) Then Exit Sub

2882      RunOpenSolver SolveRelaxation:=False, MinimiseUserInteraction:=False, sheet:=sheet

          AutoUpdateCheck
End Sub

