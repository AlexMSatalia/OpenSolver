Attribute VB_Name = "OpenSolverMenuHandlers"
Option Explicit

Function CheckForActiveSheet() As Boolean
    On Error GoTo ErrorHandler
    Dim sheet As Worksheet
    Set sheet = ActiveSheetWithValidation
    CheckForActiveSheet = True
    Exit Function
    
ErrorHandler:
    CheckForActiveSheet = False
    MsgBox "No active workbook available", Title:="OpenSolver Error"
End Function

Sub OpenSolver_SolveClickHandler(Optional Control)
          If Not CheckForActiveSheet() Then Exit Sub
          
          Dim sheet As Worksheet
1         GetActiveSheetIfMissing sheet
2         RunOpenSolver SolveRelaxation:=False, MinimiseUserInteraction:=False, sheet:=sheet
3         AutoUpdateCheck
End Sub

Sub OpenSolver_ModelOptions(Optional Control)
          If Not CheckForActiveSheet() Then Exit Sub
              
          Dim frmOptions As FOptions
1         Set frmOptions = New FOptions
2         frmOptions.Show
3         Unload frmOptions
4         AutoUpdateCheck
End Sub

Sub OpenSolver_SolverOptions(Optional Control)
          If Not CheckForActiveSheet() Then Exit Sub
              
          Dim frmSolverChange As FSolverChange
1         Set frmSolverChange = New FSolverChange
2         frmSolverChange.Show
3         Unload frmSolverChange
4         AutoUpdateCheck
End Sub

Sub OpenSolver_SolveRelaxationClickHandler(Optional Control)
          If Not CheckForActiveSheet() Then Exit Sub
              
          Dim sheet As Worksheet
1         GetActiveSheetIfMissing sheet
2         RunOpenSolver SolveRelaxation:=True, MinimiseUserInteraction:=False, sheet:=sheet
3         AutoUpdateCheck
End Sub

Sub OpenSolver_LaunchCBCCommandLine(Optional Control)
1         If Len(LastUsedSolver) = 0 Then
2             MsgBox "Cannot open the last model in CBC as the model has not been solved yet."
3         Else
              Dim Solver As ISolver
4             Set Solver = CreateSolver(LastUsedSolver)
5             If TypeOf Solver Is ISolverFile Then
                  Dim FileSolver As ISolverFile
6                 Set FileSolver = Solver
7                 If FileSolver.FileType = LP Then
8                     LaunchCommandLine_CBC
9                 Else
10                    GoTo NotLPSolver
11                End If
12            Else
NotLPSolver:
13                MsgBox "The last used solver (" & DisplayName(Solver) & ") does not use .lp model files, so CBC cannot load the model. " & _
                         "Please solve the model using a solver that uses .lp files, such as CBC or Gurobi, and try again."
14            End If
15        End If
16        AutoUpdateCheck
End Sub

Sub OpenSolver_ShowHideModelClickHandler(Optional Control)
          If Not CheckForActiveSheet() Then Exit Sub
              
          Dim sheet As Worksheet
1         GetActiveSheetIfMissing sheet
          
2         On Error GoTo ExitSub
3         If SheetHasOpenSolverHighlighting(sheet) Then
4             HideSolverModel sheet
5         Else
6             ShowSolverModel sheet, HandleError:=True
7         End If
8         AutoUpdateCheck
ExitSub:
End Sub

Sub OpenSolver_SetQuickSolveParametersClickHandler(Optional Control)
          If Not CheckForActiveSheet() Then Exit Sub
              
1         If SetQuickSolveParameterRange Then
2             ClearQuickSolve
3         End If
4         AutoUpdateCheck
End Sub

Sub OpenSolver_InitQuickSolveClickHandler(Optional Control)
          If Not CheckForActiveSheet() Then Exit Sub
              
          Dim sheet As Worksheet
1         GetActiveSheetIfMissing sheet
2         InitializeQuickSolve sheet:=sheet
3         AutoUpdateCheck
End Sub

Sub OpenSolver_QuickSolveClickHandler(Optional Control)
1         RunQuickSolve
2         AutoUpdateCheck
End Sub

Sub OpenSolver_ViewLastModelClickHandler(Optional Control)
1               If Len(LastUsedSolver) = 0 Then
2                   MsgBox "Cannot open the last model file as the model has not been solved yet."
3               Else
                    Dim Solver As ISolver
4                   Set Solver = CreateSolver(LastUsedSolver)
5                   If TypeOf Solver Is ISolverFile Then
                        Dim NotFoundMessage As String, FilePath As String
6                       FilePath = GetModelFilePath(Solver)
7                       NotFoundMessage = "Error: There is no model file (" & FilePath & ") to open. Please solve the OpenSolver model and then try again."
8                       OpenFile FilePath, NotFoundMessage
9                   Else
10                      MsgBox "The last used solver (" & DisplayName(Solver) & ") does not use a model file."
11                  End If
12              End If
13              AutoUpdateCheck
End Sub

Sub OpenSolver_ViewSolverLogFileClickHandler(Optional Control)
          Dim NotFoundMessage As String, FilePath As String
1         GetLogFilePath FilePath
2         NotFoundMessage = "Error: There is no solver log file (" & FilePath & ") to open. Please solve the OpenSolver model and then try again."
3         OpenFile FilePath, NotFoundMessage
4         AutoUpdateCheck
End Sub

Sub OpenSolver_ViewErrorLogFileClickHandler(Optional Control)
          Dim NotFoundMessage As String, FilePath As String
1         FilePath = GetErrorLogFilePath()
2         NotFoundMessage = "Error: There is no error log file (" & FilePath & ") to open."
3         OpenFile FilePath, NotFoundMessage
4         AutoUpdateCheck
End Sub

Sub OpenSolver_ViewLastSolutionClickHandler(Optional Control)
1         If Len(LastUsedSolver) = 0 Then
2             MsgBox "Cannot open the last solution file as the model has not been solved yet."
3         Else
              Dim Solver As ISolver
4             Set Solver = CreateSolver(LastUsedSolver)
5             If TypeOf Solver Is ISolverLocalExec Then
                  Dim NotFoundMessage As String, FilePath As String
6                 GetSolutionFilePath FilePath
7                 NotFoundMessage = "Error: There is no solution file (" & FilePath & ") to open. Please solve the OpenSolver model and then try again."
8                 OpenFile FilePath, NotFoundMessage
9             Else
10                MsgBox "The last used solver (" & DisplayName(Solver) & ") does not produce a solution file. Please check the log file for any solution information."
11            End If
12        End If
13        AutoUpdateCheck
End Sub

Sub OpenSolver_ViewTempFolderClickHandler(Optional Control)
          Dim NotFoundMessage As String, FolderPath As String
1         FolderPath = GetTempFolder()
2         NotFoundMessage = "Error: The OpenSolver temporary files folder (" & FolderPath & ") doesn't exist."
3         OpenFolder FolderPath, NotFoundMessage
4         AutoUpdateCheck
End Sub

Sub OpenSolver_OnlineHelp(Optional Control)
1         OpenURL "http://help.opensolver.org"
2         AutoUpdateCheck
End Sub

Sub OpenSolver_AboutClickHandler(Optional Control)
          Dim frmAbout As FAbout
1         Set frmAbout = New FAbout
2         frmAbout.Show
3         Unload frmAbout
4         AutoUpdateCheck
End Sub

Sub OpenSolver_AboutCoinOR(Optional Control)
1         MsgBox "COIN-OR" & vbCrLf & _
                 "http://www.Coin-OR.org" & vbCrLf & _
                 vbCrLf & _
                 "The Computational Infrastructure for Operations Research (COIN-OR, or simply COIN)  project is an initiative to spur the development of open-source software for the operations research community." & vbCrLf & _
                 vbCrLf & _
                 "OpenSolver uses the Coin-OR CBC optimization engine. CBC is licensed under the Common Public License 1.0. Visit the web sites for more information."
2         AutoUpdateCheck
End Sub

Sub OpenSolver_VisitOpenSolverOrg(Optional Control)
1         OpenURL "http://www.opensolver.org"
2         AutoUpdateCheck
End Sub

Sub OpenSolver_VisitCoinOROrg(Optional Control)
1         OpenURL "http://www.coin-or.org"
2         AutoUpdateCheck
End Sub
Sub OpenSolver_ModelClick(Optional Control)
          If Not CheckForActiveSheet() Then Exit Sub
              
          Static ShownFormBefore As Boolean, ShowNamedRangesState As Boolean, ShowModelAfterSavingState As Boolean
          ' Set the checkboxes default to true
1         If Not ShownFormBefore Then
2             ShownFormBefore = True
3             ShowNamedRangesState = True
4             ShowModelAfterSavingState = True
5         End If

          Dim frmModel As FModel
6         Set frmModel = New FModel
7         frmModel.ShowModelAfterSavingState = ShowModelAfterSavingState
8         frmModel.ShowNamedRangesState = ShowNamedRangesState
9         frmModel.Show
10        ShowModelAfterSavingState = frmModel.ShowModelAfterSavingState
11        ShowNamedRangesState = frmModel.ShowNamedRangesState
12        Unload frmModel
          
13        DoEvents
14        AutoUpdateCheck
End Sub

Sub OpenSolver_QuickAutoModelClick(Optional Control)
          If Not CheckForActiveSheet() Then Exit Sub
              
          Dim sheet As Worksheet
1         GetActiveSheetIfMissing sheet
                
          Dim AutoModel As CAutoModel
2         Set AutoModel = New CAutoModel
3         AutoModel.BuildModel sheet, MinimiseUserInteraction:=False, SaveAfterBuilding:=True
                
4         AutoUpdateCheck
End Sub

Sub OpenSolver_AutoModelAndSolveClick(Optional Control)
          If Not CheckForActiveSheet() Then Exit Sub
              
          Dim sheet As Worksheet
1         GetActiveSheetIfMissing sheet
          
          Dim AutoModel As CAutoModel
2         Set AutoModel = New CAutoModel
3         If Not AutoModel.BuildModel(sheet, MinimiseUserInteraction:=False, SaveAfterBuilding:=True) Then Exit Sub

4         RunOpenSolver SolveRelaxation:=False, MinimiseUserInteraction:=False, sheet:=sheet

5         AutoUpdateCheck
End Sub

