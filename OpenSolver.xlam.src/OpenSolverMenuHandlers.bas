Attribute VB_Name = "OpenSolverMenuHandlers"
Option Explicit

Function CheckForActiveSheet() As Boolean
1         On Error GoTo ErrorHandler
          Dim sheet As Worksheet
2         Set sheet = ActiveSheetWithValidation
3         CheckForActiveSheet = True
4         Exit Function
          
ErrorHandler:
5         CheckForActiveSheet = False
6         MsgBox "No active workbook available", Title:="OpenSolver Error"
End Function

Sub OpenSolver_SolveClickHandler(Optional Control)
1         If Not CheckForActiveSheet() Then Exit Sub
          
          Dim sheet As Worksheet
2         GetActiveSheetIfMissing sheet
3         RunOpenSolver SolveRelaxation:=False, MinimiseUserInteraction:=False, sheet:=sheet
4         AutoUpdateCheck
End Sub

Sub OpenSolver_ModelOptions(Optional Control)
1         If Not CheckForActiveSheet() Then Exit Sub
              
          Dim frmOptions As FOptions
2         Set frmOptions = New FOptions
3         frmOptions.Show
4         Unload frmOptions
5         AutoUpdateCheck
End Sub

Sub OpenSolver_SolverOptions(Optional Control)
1         If Not CheckForActiveSheet() Then Exit Sub
              
          Dim frmSolverChange As FSolverChange
2         Set frmSolverChange = New FSolverChange
3         frmSolverChange.Show
4         Unload frmSolverChange
5         AutoUpdateCheck
End Sub

Sub OpenSolver_SolveRelaxationClickHandler(Optional Control)
1         If Not CheckForActiveSheet() Then Exit Sub
              
          Dim sheet As Worksheet
2         GetActiveSheetIfMissing sheet
3         RunOpenSolver SolveRelaxation:=True, MinimiseUserInteraction:=False, sheet:=sheet
4         AutoUpdateCheck
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
1         If Not CheckForActiveSheet() Then Exit Sub
              
          Dim sheet As Worksheet
2         GetActiveSheetIfMissing sheet
          
3         On Error GoTo ExitSub
4         If SheetHasOpenSolverHighlighting(sheet) Then
5             HideSolverModel sheet
6         Else
7             ShowSolverModel sheet, HandleError:=True
8         End If
9         AutoUpdateCheck
ExitSub:
End Sub

Sub OpenSolver_SetQuickSolveParametersClickHandler(Optional Control)
1         If Not CheckForActiveSheet() Then Exit Sub
              
2         If SetQuickSolveParameterRange Then
3             ClearQuickSolve
4         End If
5         AutoUpdateCheck
End Sub

Sub OpenSolver_InitQuickSolveClickHandler(Optional Control)
1         If Not CheckForActiveSheet() Then Exit Sub
              
          Dim sheet As Worksheet
2         GetActiveSheetIfMissing sheet
3         InitializeQuickSolve sheet:=sheet
4         AutoUpdateCheck
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
1         If Not CheckForActiveSheet() Then Exit Sub
              
          Static ShownFormBefore As Boolean, ShowNamedRangesState As Boolean, ShowModelAfterSavingState As Boolean
          ' Set the checkboxes default to true
2         If Not ShownFormBefore Then
3             ShownFormBefore = True
4             ShowNamedRangesState = True
5             ShowModelAfterSavingState = True
6         End If

          Dim frmModel As FModel
7         Set frmModel = New FModel
8         frmModel.ShowModelAfterSavingState = ShowModelAfterSavingState
9         frmModel.ShowNamedRangesState = ShowNamedRangesState
10        frmModel.Show
11        ShowModelAfterSavingState = frmModel.ShowModelAfterSavingState
12        ShowNamedRangesState = frmModel.ShowNamedRangesState
13        Unload frmModel
          
14        DoEvents
15        AutoUpdateCheck
End Sub

Sub OpenSolver_QuickAutoModelClick(Optional Control)
1         If Not CheckForActiveSheet() Then Exit Sub
              
          Dim sheet As Worksheet
2         GetActiveSheetIfMissing sheet
                
          Dim AutoModel As CAutoModel
3         Set AutoModel = New CAutoModel
4         AutoModel.BuildModel sheet, MinimiseUserInteraction:=False, SaveAfterBuilding:=True
                
5         AutoUpdateCheck
End Sub

Sub OpenSolver_AutoModelAndSolveClick(Optional Control)
1         If Not CheckForActiveSheet() Then Exit Sub
              
          Dim sheet As Worksheet
2         GetActiveSheetIfMissing sheet
          
          Dim AutoModel As CAutoModel
3         Set AutoModel = New CAutoModel
4         If Not AutoModel.BuildModel(sheet, MinimiseUserInteraction:=False, SaveAfterBuilding:=True) Then Exit Sub

5         RunOpenSolver SolveRelaxation:=False, MinimiseUserInteraction:=False, sheet:=sheet

6         AutoUpdateCheck
End Sub

Sub OpenSolver_ImportLPClick(Optional Control)
1             If Not CheckForActiveSheet() Then Exit Sub
              
              Dim Path As String
2             With Application.FileDialog(msoFileDialogOpen)
3                 .AllowMultiSelect = False
4                 .Title = "Select LP File"
5                 .Filters.Clear
6                 .Filters.Add "LP File", "*.lp"
7                 .Show
8                 If .SelectedItems.Count = 1 Then
9                     Path = .SelectedItems(1)
10                End If
11            End With
              
12            If Path <> "" Then
                  Dim ws As Worksheet
13                RunImportLP Path, ws
14            End If
          
15            AutoUpdateCheck
End Sub
