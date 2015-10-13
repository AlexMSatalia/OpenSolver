Attribute VB_Name = "OpenSolverMenuCreation"
Option Explicit

' Used for legacy menu title
Private Const MenuName As String = "&OpenSolver"
Private Const MenuBarName As String = "Worksheet Menu Bar"

Sub AlterMenuItems(AddItems As Boolean)
          ' Add if we are on Mac or prior to Excel 2007
7         If IsMac Or Val(Application.Version) < 12 Then
8             If AddItems Then
9                 AddMenuItems
10            Else
11                DelMenuItems
12            End If
13        End If
End Sub

Private Sub AddMenuItems()
    ' If we are on Mac 2016 we have a toolbar style menu
    If IsMac And Int(Val(Application.Version)) = 15 Then
        AddMenuItems_Toolbar
    Else
        AddMenuItems_MenuBar
    End If
End Sub

Private Sub DelMenuItems()
    ' If we are on Mac 2016 we have a toolbar style menu
    If IsMac And Int(Val(Application.Version)) = 15 Then
        DelMenuItems_Toolbar
    Else
        DelMenuItems_MenuBar
    End If
End Sub

'====================================================================
' Menu content configuration
'====================================================================

' All the menus are created using the same collection of MenuItem objects

Private Function NewMenuItem(Tag As String, Id As String, Optional Label As String, Optional OnAction As String, Optional ScreenTip As String, Optional SuperTip As String, Optional Image As String, Optional Size As String, Optional NewGroup = False) As MenuItem
    Set NewMenuItem = New MenuItem
    With NewMenuItem
        .Tag = Tag                      ' The type of item
        .Id = Id                        ' The unique id for the item
        .Label = Label                  ' The user-facing label for the item
        .OnAction = OnAction            ' The macro called when clicked
        .ScreenTip = ScreenTip          ' The tooltip heading on hover in the ribbon
        .SuperTip = SuperTip            ' The tooltip body text on hover in the ribbon
        .Image = Image                  ' The name of the image to include in the ribbon (must be in the .xlam file)
        .Size = Size                    ' The size of the item in the ribbon
        .NewGroup = NewGroup            ' Whether the item starts a new group in the menu
        Set .Children = New Collection  ' All children of the item
    End With
End Function

Private Function GenerateMenuItems() As Collection
    Dim Items As Collection
    Set Items = New Collection
    
    ' Model sub-menu
    Dim ModelScreenTip As String, ModelSuperTip As String, ModelOnAction As String
    ModelScreenTip = "Build and edit OpenSolver models"
    ModelSuperTip = "Build or edit your optimization model. Will detect and load any pre-existing model and will save the results to the sheet."
    ModelOnAction = "OpenSolver_ModelClick"
    
    Dim ModelSB As MenuItem
    Set ModelSB = NewMenuItem("splitButton", "OpenSolverModelSB", Size:="large", NewGroup:=True)
    ModelSB.Children.Add NewMenuItem("button", "OpenSolverModel", "&Model", ModelOnAction, ModelScreenTip, ModelSuperTip, "model")
    
    Dim ModelMenu As MenuItem
    Set ModelMenu = NewMenuItem("menu", "OpenSolverModelMenu", Size:="normal")
    With ModelMenu.Children
        .Add NewMenuItem("button", "OpenSolverModel2", "&Model...", ModelOnAction, ModelScreenTip, ModelSuperTip)
        .Add NewMenuItem("button", "OpenSolverQuickAutomodel", "&Quick AutoModel", _
                         "OpenSolver_QuickAutoModelClick", "Run AutoModel with default options", _
                         "OpenSolver's AutoModel will analyze the current sheet and use the sheet's structure " & _
                         "and formulae to build an optimization model automatically.")
        .Add NewMenuItem("button", "OpenSolverModelAutoModel", "&AutoModel And Solve", _
                         "OpenSolver_AutoModelAndSolveClick", "Run AutoModel and then solve problem", _
                         "Run AutoModel with the default options to build the problem automatically " & _
                         "and then solve the resulting problem with the currently selected solver engine.")
        .Add NewMenuItem("button", "OpenSolverChosenSolver", "&Solver Engine...", _
                         "OpenSolver_SolverOptions", "Select your solver", _
                         "Choose your solver from the list of supported solvers.")
        .Add NewMenuItem("button", "OpenSolverModelOptions", "&Options...", _
                         "OpenSolver_ModelOptions", "Set solve options", _
                         "Set options such as linearity, non-negativity, max solve time, and tolerance.")
    End With
    ModelSB.Children.Add ModelMenu
    Items.Add ModelSB
    
    ' Solve sub-menu
    Dim SolveScreenTip As String, SolveSuperTip As String, SolveOnAction As String
    SolveScreenTip = "Solve optimization model"
    SolveSuperTip = "Solve an existing model on the active worksheet by constructing the model's equations and " & _
                    "then calling the current chosen optimization engine."
    SolveOnAction = "OpenSolver_SolveClickHandler"
    
    Dim SolveSB As MenuItem
    Set SolveSB = NewMenuItem("splitButton", "OpenSolverSolveSB", Size:="large")
    SolveSB.Children.Add NewMenuItem("button", "OpenSolverSolve", "&Solve", SolveOnAction, SolveScreenTip, SolveSuperTip, "solve")
    
    Dim SolveMenu As MenuItem
    Set SolveMenu = NewMenuItem("menu", "OpenSolverSolveMenu", Size:="normal")
    With SolveMenu.Children
        .Add NewMenuItem("button", "OpenSolverSolve2", "&Solve", SolveOnAction, SolveScreenTip, SolveSuperTip)
        .Add NewMenuItem("button", "OpenSolverSolveRelaxation", "Solve &Relaxation", "OpenSolver_SolveRelaxationClickHandler", _
                         "Solve a modified problem without any integer or binary constraints", _
                         "Relaxes any integer or binary requirements on the variables, and solves the resulting linear program, " & _
                         "typically giving an answer with fractional variables.")
    End With
    SolveSB.Children.Add SolveMenu
    Items.Add SolveSB
    
    ' Single-level main menu items
    Items.Add NewMenuItem("button", "OpenSolverShowModel", "Show/&Hide Model", "OpenSolver_ShowHideModelClickHandler", _
                          "Show or hide the optimization model on this sheet", _
                          "OpenSolver will analyze an existing model on the active sheet, and add coloured annotations " & _
                          "to the sheet that indicate the variable cells, the objective cell, and the constraints.")
    Items.Add NewMenuItem("button", "OpenSolverQuickSolve", "&Quick Solve", "OpenSolver_QuickSolveClickHandler", _
                          "Quickly re-solve a model after changing the parameter values", _
                          "OpenSolver can re-solve problems very quickly if it is first told about the cells (termed parameters) " & _
                          "that change between solves. The Quick Solve menu items below can be used to set this up for your model.")
                               
    ' OpenSolver sub-menu
    Dim OpenSolverMenu As MenuItem
    Set OpenSolverMenu = NewMenuItem("menu", "menu", "&OpenSolver", NewGroup:=True)
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
        .Add NewMenuItem("button", "OpenSolverViewModel", "View Last Model File", "OpenSolver_ViewLastModelClickHandler", _
                         "View the model file created when OpenSolver last solved a model", _
                         "For some solvers, OpenSolver writes the model to a temporary text file that is read by the solver. " & _
                         "It is often useful to load and view this file.")
        .Add NewMenuItem("button", "OpenSolverViewSolution", "View Last Solution File", "OpenSolver_ViewLastSolutionClickHandler", _
                         "View the last solution file", _
                         "When a model is solved, a temporary solution file can be created by the solver containing the solution " & _
                         "to the optimization problem. It is sometimes useful to load and view this file.")
        .Add NewMenuItem("button", "OpenSolverViewSolverLogFile", "View Last Solve Log File", "OpenSolver_ViewSolverLogFileClickHandler", _
                         "View the last solve log file", _
                         "OpenSolver creates a log file every time a model is solved. The log file contains output from the solver, " & _
                         "such as iteration details during the solve, which can give you more information about your model.")
        .Add NewMenuItem("button", "OpenSolverViewErrorLogFile", "View Last Error Log File", "OpenSolver_ViewErrorLogFileClickHandler", _
                         "View the last error log file", _
                         "OpenSolver creates a log file every time an error occurs with detailed information about the error.")
        .Add NewMenuItem("button", "OpenSolverViewTempFolder", "View All OpenSolver Files...", _
                         "OpenSolver_ViewTempFolderClickHandler", "View all temporary OpenSolver files", _
                         "Opens the folder containing all files created while using OpenSolver.")
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

'====================================================================
' Code for making the MenuBar menu
' Adapted from Excel 2003 Menu Code originally provided by Paul Becker of Eclipse Engineering (www.eclipseeng.com)
'====================================================================

Private Sub AddMenuItems_MenuBar()
2884      DelMenuItems_MenuBar

          Dim MainMenuBar As CommandBar, OpenSolverMenu As CommandBarControl
2885      Set MainMenuBar = Application.CommandBars(MenuBarName)
2886      Set OpenSolverMenu = MainMenuBar.Controls.Add(Type:=msoControlPopup)
2887      OpenSolverMenu.Caption = MenuName

          Dim Item As MenuItem
          For Each Item In GenerateMenuItems()
              AddToMenu_MenuBar OpenSolverMenu, Item
          Next Item
End Sub

Private Sub AddToMenu_MenuBar(Menu As CommandBarControl, Item As MenuItem)
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
        SubMenu.BeginGroup = Item.NewGroup
        
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
            AddToMenu_MenuBar SubMenu, SubItem
        Next SubItem
    End Select
End Sub

Private Sub DelMenuItems_MenuBar()
3011      On Error Resume Next
3012      Application.CommandBars(MenuBarName).Controls(MenuName).Delete
End Sub

'====================================================================
' Code for making the Toolbar menu
' Adapted from http://peltiertech.com/office-2016-for-mac-is-here/#comment-693017
'====================================================================

Public Sub AddMenuItems_Toolbar(Optional RootItem As String)
    ' Creates a menu using the children of the MenuItem with the id specified by RootItem
    DelMenuItems_Toolbar
    
    Dim OpenSolverMenu As CommandBar
    Set OpenSolverMenu = Application.CommandBars.Add(MenuName)
    OpenSolverMenu.Visible = True
    
    ' Get the children of RootItem
    Dim MenuItems As Collection
    Set MenuItems = FindChildren(GenerateMenuItems(), RootItem)
    
    ' Add a 'Back to main menu' item if we are in a sub-menu
    If Len(RootItem) <> 0 Then
        AddToMenu_Toolbar OpenSolverMenu, _
            NewMenuItem("button", "back", ChrW(&H25C2) & " Back to Main Menu", "AddMenuItems_Toolbar")
    End If
    
    Dim Item As MenuItem
    For Each Item In MenuItems
        AddToMenu_Toolbar OpenSolverMenu, Item
    Next Item
End Sub

Private Function FindChildren(Items As Collection, RootItem As String) As Collection
    If Len(RootItem) = 0 Then
        Set FindChildren = Items
        Exit Function
    End If
    
    Dim Item As MenuItem
    For Each Item In Items
        If Item.Id = RootItem Then
            Set FindChildren = Item.Children
            Exit Function
        End If
        Set FindChildren = FindChildren(Item.Children, RootItem)
        If FindChildren.Count <> 0 Then Exit Function
    Next Item
    ' No match, return empty collection
    Set FindChildren = New Collection
End Function

Private Sub AddToMenu_Toolbar(Menu As CommandBar, Item As MenuItem)
    Select Case Item.Tag
    Case "button"
        ' Add button for the item
        With Menu.Controls.Add(Type:=msoControlButton)
          .Style = msoButtonCaption
          .Caption = Item.Label
          .OnAction = Item.OnAction
          .Enabled = True
        End With
    Case "splitButton", "menu"
        ' Add button to open the specified sub-menu
        With Menu.Controls.Add(Type:=msoControlButton)
          .Style = msoButtonCaption
          .Caption = IIf(Item.Tag = "splitButton", Item.Children(1).Label, Item.Label)
          .OnAction = "AddMenuItems_Toolbar" & Replace(.Caption, "&", "")
          .Caption = .Caption & " " & ChrW(&H25BE)
          .Enabled = True
        End With
    End Select
End Sub

' Click handlers for opening sub-menus
Public Sub AddMenuItems_ToolbarModel()
    AddMenuItems_Toolbar "OpenSolverModelMenu"
End Sub

Public Sub AddMenuItems_ToolbarSolve()
    AddMenuItems_Toolbar "OpenSolverSolveMenu"
End Sub

Public Sub AddMenuItems_ToolbarOpenSolver()
    AddMenuItems_Toolbar "menu"
End Sub

Private Sub DelMenuItems_Toolbar()
          On Error Resume Next
          Application.CommandBars(MenuName).Delete
End Sub

'====================================================================
' Code for making the Ribbon XML
'====================================================================

Public Sub CreateRibbonXML()
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

Private Sub OutputRibbonHeader(FileNum As Integer)
    Print #FileNum, Spc(0); "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
    Print #FileNum, Spc(0); "<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">"
    Print #FileNum, Spc(2); "<ribbon startFromScratch=""false"">"
    Print #FileNum, Spc(4); "<tabs>"
    Print #FileNum, Spc(6); "<tab idMso=""TabData"">"
    Print #FileNum, Spc(8); "<group id=""GroupOpenSolver"" label=""OpenSolver"">"
End Sub

Private Sub OutputRibbonFooter(FileNum As Integer)
    Print #FileNum, Spc(8); "</group>"
    Print #FileNum, Spc(6); "</tab>"
    Print #FileNum, Spc(4); "</tabs>"
    Print #FileNum, Spc(2); "</ribbon>"
    Print #FileNum, Spc(0); "</customUI>"
End Sub

Private Sub OutputRibbonXML(FileNum As Integer, Item As MenuItem, Indent As Integer)
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

Private Sub OutputIfExists(FileNum As Integer, value As String, Tag As String)
    If Len(value) <> 0 Then Print #FileNum, " " & Tag & "=" & Quote(value);
End Sub

