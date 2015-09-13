Attribute VB_Name = "OpenSolverMenu"
Option Explicit

' Used for legacy menu title
Private Const MenuName As String = "&OpenSolver"
Private Const MenuBarName As String = "Worksheet Menu Bar"

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
2755      RunOpenSolver False, False, 0
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
2763      RunOpenSolver True, False, 0
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
2766      On Error GoTo ExitSub
2768      If SheetHasOpenSolverHighlighting Then
2769          HideSolverModel
2770      Else
2771          ShowSolverModel
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
2778      InitializeQuickSolve
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
          If Len(LastUsedSolver) = 0 Then
              MsgBox "Cannot open the log file as the model has not been solved yet."
          Else
              Dim NotFoundMessage As String, FilePath As String
2787          GetLogFilePath FilePath
2788          NotFoundMessage = "Error: There is no solver log file (" & FilePath & ") to open. Please solve the OpenSolver model and then try again."
2789          OpenFile FilePath, NotFoundMessage
          End If
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
2790              FilePath = GetSolutionFilePath()
2791              NotFoundMessage = "Error: There is no solution file (" & FilePath & ") to open. Please solve the OpenSolver model and then try again."
2792              OpenFile FilePath, NotFoundMessage
              Else
                  MsgBox "The last used solver (" & DisplayName(Solver) & ") does not produce a solution file. Please check the log file for any solution information."
              End If
          End If
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
          Dim frmModel As FModel
          Set frmModel = New FModel
2853      frmModel.Show
          Unload frmModel
2854      DoEvents
          AutoUpdateCheck
End Sub

Sub OpenSolver_QuickAutoModelClick(Optional Control)
          Dim sheet As Worksheet
          GetActiveSheetIfMissing sheet
          
2855      RunAutoModel sheet, False
          AutoUpdateCheck
End Sub

Sub OpenSolver_AutoModelAndSolveClick(Optional Control)
          Dim sheet As Worksheet
          GetActiveSheetIfMissing sheet
          If Not RunAutoModel(sheet, False) Then Exit Sub
2882      RunOpenSolver False, False, 0, sheet
          AutoUpdateCheck
End Sub

'====================================================================
' Adapted from Excel 2003 Menu Code originally provided by Paul Becker of Eclipse Engineering (www.eclipseeng.com)
'====================================================================

Public Sub AddMenuItems()
2884      DelMenuItems

          Dim MainMenuBar As CommandBar, OpenSolverMenu As CommandBarControl
2885      Set MainMenuBar = Application.CommandBars(MenuBarName)
2886      Set OpenSolverMenu = MainMenuBar.Controls.Add(Type:=msoControlPopup)
2887      OpenSolverMenu.Caption = MenuName

          Dim Item As MenuItem
          For Each Item In GenerateMenuItems()
              AddToMenu OpenSolverMenu, Item
          Next Item
End Sub

Public Sub DelMenuItems()
3011      On Error Resume Next
3012      Application.CommandBars(MenuBarName).Controls(MenuName).Delete
End Sub
'====================================================================
' Adapted from Excel 2003 Menu Code originally provided by Paul Becker of Eclipse Engineering (www.eclipseeng.com)
'====================================================================

Function GenerateMenuItems() As Collection
    Dim Items As Collection
    Set Items = New Collection
    
    ' Model sub-menu
    Dim ModelScreenTip As String, ModelSuperTip As String, ModelOnAction As String
    ModelScreenTip = "Build and edit Solver models"
    ModelSuperTip = "Build or edit your optimization model. Will detect and load any pre-existing model built with Solver, and will save the results to the sheet in a Solver-friendly way."
    ModelOnAction = "OpenSolver_ModelClick"
    
    Dim ModelSB As MenuItem
    Set ModelSB = NewMenuItem("splitButton", "OpenSolverModelSB", Size:="large")
    ModelSB.Children.Add NewMenuItem("button", "OpenSolverModel", "&Model", ModelOnAction, ModelScreenTip, ModelSuperTip, "model")
    
    Dim ModelMenu As MenuItem
    Set ModelMenu = NewMenuItem("menu", "OpenSolverModelMenu", Size:="normal")
    With ModelMenu.Children
        .Add NewMenuItem("button", "OpenSolverModel2", "&Model...", ModelOnAction, ModelScreenTip, ModelSuperTip)
        .Add NewMenuItem("button", "OpenSolverQuickAutomodel", "&Quick AutoModel", _
                         "OpenSolver_QuickAutoModelClick", "Run AutoModel with default options", _
                         "Run AutoModel with default options. No dialog menu will appear, so for " & _
                         "well-structured sheets it is unnecessary to run through the full AutoModel " & _
                         "procedure step-by-step for small changes.")
        .Add NewMenuItem("button", "OpenSolverModelAutoModel", "&AutoModel And Solve", _
                         "OpenSolver_AutoModelAndSolveClick", "Run AutoModel with default options and then solve problem", _
                         "Run AutoModel with default options and then solve the problem. No dialog menu will appear, so for " & _
                         "well-structured sheets it is unnecessary to run through the full AutoModel procedure step-by-step for small changes.")
        .Add NewMenuItem("button", "OpenSolverChosenSolver", "&Solver Engine...", _
                         "OpenSolver_SolverOptions", "Choose your solver", _
                         "Choose your solver: CBC (default), Gurobi, or NOMAD (non-linear).")
        .Add NewMenuItem("button", "OpenSolverModelOptions", "&Options...", _
                         "OpenSolver_ModelOptions", "Set solve options", _
                         "Set options: linearity, non-negativity, max solve time, tolerance.")
    End With
    ModelSB.Children.Add ModelMenu
    Items.Add ModelSB
    
    ' Main menu items
    Items.Add NewMenuItem("button", "OpenSolverSolve", "&Solve", "OpenSolver_SolveClickHandler", "Solve optimization model", _
                          "Solve an existing Solver model on the active worksheet by constructing the model's equations and " & _
                          "then calling the current chosen optimization engine.", "solve", "large")
    Items.Add NewMenuItem("button", "OpenSolverShowModel", "Show/&Hide Model", "OpenSolver_ShowHideModelClickHandler", _
                          "Show or hide the optimization model on this sheet", _
                          "OpenSolver will analyse an existing model on the active sheet, and add coloured annotations " & _
                          "to the sheet that indicate the variable cells, the objective cell, and the constraints.")
    Items.Add NewMenuItem("button", "OpenSolverQuickSolve", "&Quick Solve", "OpenSolver_QuickSolveClickHandler", _
                          "Quickly re-solve a model after changing the parameter values", _
                          "OpenSolver can re-solve problems very quickly if it is first told about the cells (termed parameters) " & _
                          "that change between solves. The Quick Solve menu items below can be used to set this up for your model.")
                               
    ' OpenSolver submenu
    Dim OpenSolverMenu As MenuItem
    Set OpenSolverMenu = NewMenuItem("menu", "menu", "&OpenSolver")
    With OpenSolverMenu.Children
        .Add NewMenuItem("menuSeparator", "separator0", "QuickSolve Options")
        .Add NewMenuItem("button", "OpenSolverInitParameters", "Set QuickSolve Parameters...", _
                         "OpenSolver_SetQuickSolveParametersClickHandler", _
                         "Define the parameter cells for QuickSolve", _
                         "OpenSolver can re-solve problems very quickly if it is first told about the cells (termed parameters) " & _
                         "that change between solves. This menu item lets you define the parameter cells. Note that OpenSolver " & _
                         "assumes these parameters change the model's constraint right hand sides in a linear fashion.")
        .Add NewMenuItem("button", "OpenSolverInitQuicksolve", "Initialize QuickSolve", _
                         "OpenSolver_InitQuickSolveClickHandler", _
                         "Construct the model's equations and prepare for QuickSolve", _
                         "OpenSolver can re-solve problems very quickly if it is first told about the cells (termed parameters) " & _
                         "that change between solves. After you have defined the parameter cells, this menu item constructs the " & _
                         "model's equations ready for quick solving.")
                              
        .Add NewMenuItem("menuSeparator", "separator1", "Temporary Files")
        .Add NewMenuItem("button", "OpenSolverSolveRelaxation", "Solve Relaxation", "OpenSolver_SolveRelaxationClickHandler", _
                         "Solve a modified problem without any integer or binary constraints", _
                         "Relaxes any integer or binary requirements on the variables, and solves the resulting linear program, " & _
                         "typically giving an answer with fractional variables.")
        .Add NewMenuItem("button", "OpenSolverViewModel", "View Last Model File", "OpenSolver_ViewLastModelClickHandler", _
                         "View the model file created when OpenSolver last solved a model", _
                         "For some solvers, OpenSolver writes the model to a temporary text file that is read by the solver. " & _
                         "It is often useful to load and view this file.")
        .Add NewMenuItem("button", "OpenSolverViewSolution", "View Last Solution File", "OpenSolver_ViewLastSolutionClickHandler", _
                         "View the last solution file", _
                         "When a model is solved, a temporary solution file can be created by the solver containing the solution " & _
                         "to the optimization problem. It is sometimes useful to load and view this file.")
        .Add NewMenuItem("button", "OpenSolverViewSolverLogFile", "View Last Solver Log File", "OpenSolver_ViewSolverLogFileClickHandler", _
                         "View the last solver log file", _
                         "OpenSolver creates a log file every time a model is solved. The log file contains output from the solver, " & _
                         "such as iteration details during the solve, which can give you more information about your model.")
        .Add NewMenuItem("button", "OpenSolverViewErrorLogFile", "View Last Error Log File", "OpenSolver_ViewErrorLogFileClickHandler", _
                         "View the last error log file", _
                         "OpenSolver creates a log file every time an error occurs with detailed information about the error.")
        .Add NewMenuItem("button", "OpenSolverLaunchCBC", "Open Last Model in CBC...", "OpenSolver_LaunchCBCCommandLine", _
                         "Open the CBC command line, and load in the last model.", _
                         "Open the CBC optimizer at the command line, and load in the last model solved by OpenSolver. " & _
                         "Type '?' at the CBC command line to get help on the CBC commands, and 'exit' to quit CBC. " & _
                         "Note that any solutions generated are discarded; they are not loaded back into your spreadsheet.")
                              
        .Add NewMenuItem("menuSeparator", "separator2", "About")
        .Add NewMenuItem("button", "OpenSolverAbout", "About OpenSolver...", "OpenSolver_AboutClickHandler")
        .Add NewMenuItem("button", "OpenSolverAboutCoinOR", "About COIN-OR...", "OpenSolver_AboutCoinOR")
                              
        .Add NewMenuItem("menuSeparator", "separator3", "Help")
        .Add NewMenuItem("button", "OpenSolverVisitOpenSolverOrg", "Open OpenSolver.org", "OpenSolver_VisitOpenSolverOrg")
        .Add NewMenuItem("button", "OpenSolverVisitCoinOROrg", "Open COIN-OR.org", "OpenSolver_VisitCoinOROrg")
    End With
    Items.Add OpenSolverMenu
    
    Set GenerateMenuItems = Items
End Function

Function NewMenuItem(Tag As String, Id As String, Optional Label As String, _
        Optional OnAction As String, Optional ScreenTip As String, _
        Optional SuperTip As String, Optional Image As String, _
        Optional Size As String) As MenuItem
    Set NewMenuItem = New MenuItem
    With NewMenuItem
        .Tag = Tag
        .Id = Id
        .Label = Label
        .OnAction = OnAction
        .ScreenTip = ScreenTip
        .SuperTip = SuperTip
        .Image = Image
        .Size = Size
        Set .Children = New Collection
    End With
End Function

Sub AddToMenu(Menu As CommandBarControl, Item As MenuItem)
    Static sBeginGroup As Boolean

    Select Case Item.Tag
    Case "button"
        With Menu.Controls.Add(Type:=msoControlButton)
          .Caption = Item.Label
          .OnAction = Item.OnAction
          .BeginGroup = sBeginGroup
          .FaceId = 0
        End With
        sBeginGroup = False
    Case "menuSeparator"
        sBeginGroup = True
    Case "menu", "splitButton"
        Dim SubMenu As CommandBarControl
        Set SubMenu = Menu.Controls.Add(Type:=msoControlPopup)
        SubMenu.BeginGroup = True  ' TODO may not always be this
        
        Dim Children As Collection
        If Item.Tag = "menu" Then
            SubMenu.Caption = Item.Label
            Set Children = Item.Children
        Else
            SubMenu.Caption = Item.Children.Item(1).Label
            Set Children = Item.Children.Item(2).Children
        End If
        
        Dim SubItem As MenuItem
        For Each SubItem In Children
            AddToMenu SubMenu, SubItem
        Next SubItem
    End Select
End Sub

Sub CreateRibbonXML()
    Dim CustomXMLFile As String
    GetExistingFilePathName JoinPaths(ThisWorkbook.Path, "RibbonX", "customUI"), "customUI.xml", CustomXMLFile
    
    Dim FileNum As Integer
    FileNum = FreeFile()
    Open CustomXMLFile For Output As #FileNum
        OutputRibbonHeader FileNum
        
        Dim Item As MenuItem
        For Each Item In GenerateMenuItems()
            OutputRibbonXML FileNum, Item, 10
        Next Item
        
        OutputRibbonFooter FileNum
    Close #FileNum
End Sub

Sub OutputRibbonHeader(FileNum As Integer)
    Print #FileNum, Spc(0); "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
    Print #FileNum, Spc(0); "<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">"
    Print #FileNum, Spc(2); "<ribbon startFromScratch=""false"">"
    Print #FileNum, Spc(4); "<tabs>"
    Print #FileNum, Spc(6); "<tab idMso=""TabData"">"
    Print #FileNum, Spc(8); "<group id=""GroupOpenSolver"" label=""OpenSolver"">"
End Sub

Sub OutputRibbonFooter(FileNum As Integer)
    Print #FileNum, Spc(8); "</group>"
    Print #FileNum, Spc(6); "</tab>"
    Print #FileNum, Spc(4); "</tabs>"
    Print #FileNum, Spc(2); "</ribbon>"
    Print #FileNum, Spc(0); "</customUI>"
End Sub

Sub OutputRibbonXML(FileNum As Integer, Item As MenuItem, Indent As Integer)
    Print #FileNum, Spc(Indent); "<" & Item.Tag & " id=" & Quote(Item.Id);
    OutputIfExists FileNum, Replace(Item.Label, "&", "&amp;"), IIf(Item.Tag = "menuSeparator", "title", "label")
    OutputIfExists FileNum, Item.OnAction, "onAction"
    OutputIfExists FileNum, Item.ScreenTip, "screentip"
    OutputIfExists FileNum, Item.SuperTip, "supertip"
    OutputIfExists FileNum, Item.Size, IIf(Item.Tag = "menu", "itemSize", "size")
    OutputIfExists FileNum, Item.Image, "image"
        
    Select Case Item.Tag
    Case "button", "menuSeparator"
        Print #FileNum, "/>"
    Case "menu", "splitButton"
        Print #FileNum, ">"
        
        Dim SubItem As MenuItem
        For Each SubItem In Item.Children
            OutputRibbonXML FileNum, SubItem, Indent + 2
        Next SubItem
        Print #FileNum, Spc(Indent); "</" & Item.Tag & ">"
    End Select
End Sub

Sub OutputIfExists(FileNum As Integer, value As String, Tag As String)
    If Len(value) <> 0 Then Print #FileNum, " " & Tag & "=" & Quote(value);
End Sub

