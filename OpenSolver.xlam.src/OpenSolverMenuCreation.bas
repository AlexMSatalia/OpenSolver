Attribute VB_Name = "OpenSolverMenuCreation"
Option Explicit

' Used for legacy menu title
Private Const MenuName As String = "&OpenSolver"
Private Const MenuBarName As String = "Worksheet Menu Bar"

Sub AlterMenuItems(AddItems As Boolean)
          ' Add if we are on Mac 2011 or Windows prior to Excel 2007
1         If (IsMac And Val(Application.Version) < 15) Or Val(Application.Version) < 12 Then
2             If AddItems Then
3                 AddMenuItems_MenuBar
4             Else
5                 DelMenuItems_MenuBar
6             End If
7         End If
End Sub

'====================================================================
' Menu content configuration
'====================================================================

' All the menus are created using the same collection of MenuItem objects

Private Function NewMenuItem(Tag As String, Id As String, Optional Label As String, Optional OnAction As String, Optional ScreenTip As String, Optional SuperTip As String, Optional Image As String, Optional Size As String, Optional NewGroup = False) As MenuItem
1         Set NewMenuItem = New MenuItem
2         With NewMenuItem
3             .Tag = Tag                      ' The type of item
4             .Id = Id                        ' The unique id for the item
5             .Label = Label                  ' The user-facing label for the item
6             .OnAction = OnAction            ' The macro called when clicked
7             .ScreenTip = ScreenTip          ' The tooltip heading on hover in the ribbon
8             .SuperTip = SuperTip            ' The tooltip body text on hover in the ribbon
9             .Image = Image                  ' The name of the image to include in the ribbon (must be in the .xlam file)
10            .Size = Size                    ' The size of the item in the ribbon
11            .NewGroup = NewGroup            ' Whether the item starts a new group in the menu
12            Set .Children = New Collection  ' All children of the item
13        End With
End Function

Private Function GenerateMenuItems() As Collection
          Dim Items As Collection
1         Set Items = New Collection
          
          ' Model sub-menu
          Dim ModelScreenTip As String, ModelSuperTip As String, ModelOnAction As String
2         ModelScreenTip = "Build and edit OpenSolver models"
3         ModelSuperTip = "Build or edit your optimization model. Will detect and load any pre-existing model and will save the results to the sheet."
4         ModelOnAction = "OpenSolver_ModelClick"
          
          Dim ModelSB As MenuItem
5         Set ModelSB = NewMenuItem("splitButton", "OpenSolverModelSB", Size:="large", NewGroup:=True)
6         ModelSB.Children.Add NewMenuItem("button", "OpenSolverModel", "&Model", ModelOnAction, ModelScreenTip, ModelSuperTip, "model")
          
          Dim ModelMenu As MenuItem
7         Set ModelMenu = NewMenuItem("menu", "OpenSolverModelMenu", Size:="normal")
8         With ModelMenu.Children
9             .Add NewMenuItem("button", "OpenSolverModel2", "&Model...", ModelOnAction, ModelScreenTip, ModelSuperTip)
10            .Add NewMenuItem("button", "OpenSolverQuickAutomodel", "&Quick AutoModel", _
                               "OpenSolver_QuickAutoModelClick", "Run AutoModel with default options", _
                               "OpenSolver's AutoModel will analyze the current sheet and use the sheet's structure " & _
                               "and formulae to build an optimization model automatically.")
11            .Add NewMenuItem("button", "OpenSolverModelAutoModel", "&AutoModel And Solve", _
                               "OpenSolver_AutoModelAndSolveClick", "Run AutoModel and then solve problem", _
                               "Run AutoModel with the default options to build the problem automatically " & _
                               "and then solve the resulting problem with the currently selected solver engine.")
12            .Add NewMenuItem("button", "OpenSolverChosenSolver", "&Solver Engine...", _
                               "OpenSolver_SolverOptions", "Select your solver", _
                               "Choose your solver from the list of supported solvers.")
13            .Add NewMenuItem("button", "OpenSolverModelOptions", "&Options...", _
                               "OpenSolver_ModelOptions", "Set solve options", _
                               "Set options such as linearity, non-negativity, max solve time, and tolerance.")
14        End With
15        ModelSB.Children.Add ModelMenu
16        Items.Add ModelSB
          
          ' Solve sub-menu
          Dim SolveScreenTip As String, SolveSuperTip As String, SolveOnAction As String
17        SolveScreenTip = "Solve optimization model"
18        SolveSuperTip = "Solve an existing model on the active worksheet by constructing the model's equations and " & _
                          "then calling the current chosen optimization engine."
19        SolveOnAction = "OpenSolver_SolveClickHandler"
          
          Dim SolveSB As MenuItem
20        Set SolveSB = NewMenuItem("splitButton", "OpenSolverSolveSB", Size:="large")
21        SolveSB.Children.Add NewMenuItem("button", "OpenSolverSolve", "&Solve", SolveOnAction, SolveScreenTip, SolveSuperTip, "solve")
          
          Dim SolveMenu As MenuItem
22        Set SolveMenu = NewMenuItem("menu", "OpenSolverSolveMenu", Size:="normal")
23        With SolveMenu.Children
24            .Add NewMenuItem("button", "OpenSolverSolve2", "&Solve", SolveOnAction, SolveScreenTip, SolveSuperTip)
25            .Add NewMenuItem("button", "OpenSolverSolveRelaxation", "Solve &Relaxation", "OpenSolver_SolveRelaxationClickHandler", _
                               "Solve a modified problem without any integer or binary constraints", _
                               "Relaxes any integer or binary requirements on the variables, and solves the resulting linear program, " & _
                               "typically giving an answer with fractional variables.")
26        End With
27        SolveSB.Children.Add SolveMenu
28        Items.Add SolveSB
          
          ' Single-level main menu items
29        Items.Add NewMenuItem("button", "OpenSolverShowModel", "Show/&Hide Model", "OpenSolver_ShowHideModelClickHandler", _
                                "Show or hide the optimization model on this sheet", _
                                "OpenSolver will analyze an existing model on the active sheet, and add coloured annotations " & _
                                "to the sheet that indicate the variable cells, the objective cell, and the constraints.")
30        Items.Add NewMenuItem("button", "OpenSolverQuickSolve", "&Quick Solve", "OpenSolver_QuickSolveClickHandler", _
                                "Quickly re-solve a model after changing the parameter values", _
                                "OpenSolver can re-solve problems very quickly if it is first told about the cells (termed parameters) " & _
                                "that change between solves. The Quick Solve menu items below can be used to set this up for your model.")
                                     
          ' OpenSolver sub-menu
          Dim OpenSolverMenu As MenuItem
31        Set OpenSolverMenu = NewMenuItem("menu", "menu", "&OpenSolver", NewGroup:=True)
32        With OpenSolverMenu.Children
33            .Add NewMenuItem("menuSeparator", "separator0", "QuickSolve Options")
34            .Add NewMenuItem("button", "OpenSolverInitParameters", "Set QuickSolve Parameters...", _
                               "OpenSolver_SetQuickSolveParametersClickHandler", _
                               "Define the parameter cells for QuickSolve", _
                               "OpenSolver can re-solve problems very quickly if it is first told about the cells (termed parameters) " & _
                               "that change between solves. This menu item lets you define the parameter cells. Note that OpenSolver " & _
                               "assumes these parameters change the model's constraint right hand sides in a linear fashion.")
35            .Add NewMenuItem("button", "OpenSolverInitQuicksolve", "Initialize QuickSolve", _
                               "OpenSolver_InitQuickSolveClickHandler", _
                               "Construct the model's equations and prepare for QuickSolve", _
                               "OpenSolver can re-solve problems very quickly if it is first told about the cells (termed parameters) " & _
                               "that change between solves. After you have defined the parameter cells, this menu item constructs the " & _
                               "model's equations ready for quick solving.")
                                    
36            .Add NewMenuItem("menuSeparator", "separator1", "Temporary Files")
37            .Add NewMenuItem("button", "OpenSolverViewModel", "View Last Model File", "OpenSolver_ViewLastModelClickHandler", _
                               "View the model file created when OpenSolver last solved a model", _
                               "For some solvers, OpenSolver writes the model to a temporary text file that is read by the solver. " & _
                               "It is often useful to load and view this file.")
38            .Add NewMenuItem("button", "OpenSolverViewSolution", "View Last Solution File", "OpenSolver_ViewLastSolutionClickHandler", _
                               "View the last solution file", _
                               "When a model is solved, a temporary solution file can be created by the solver containing the solution " & _
                               "to the optimization problem. It is sometimes useful to load and view this file.")
39            .Add NewMenuItem("button", "OpenSolverViewSolverLogFile", "View Last Solve Log File", "OpenSolver_ViewSolverLogFileClickHandler", _
                               "View the last solve log file", _
                               "OpenSolver creates a log file every time a model is solved. The log file contains output from the solver, " & _
                               "such as iteration details during the solve, which can give you more information about your model.")
40            .Add NewMenuItem("button", "OpenSolverViewErrorLogFile", "View Last Error Log File", "OpenSolver_ViewErrorLogFileClickHandler", _
                               "View the last error log file", _
                               "OpenSolver creates a log file every time an error occurs with detailed information about the error.")
41            .Add NewMenuItem("button", "OpenSolverViewTempFolder", "View All OpenSolver Files...", _
                               "OpenSolver_ViewTempFolderClickHandler", "View all temporary OpenSolver files", _
                               "Opens the folder containing all files created while using OpenSolver.")
42            .Add NewMenuItem("button", "OpenSolverLaunchCBC", "Open Last Model in CBC...", "OpenSolver_LaunchCBCCommandLine", _
                               "Open the CBC command line, and load in the last model.", _
                               "Open the CBC optimizer at the command line, and load in the last model solved by OpenSolver. " & _
                               "Type '?' at the CBC command line to get help on the CBC commands, and 'exit' to quit CBC. " & _
                               "Note that any solutions generated are discarded; they are not loaded back into your spreadsheet.")
                                    
              .Add NewMenuItem("menuSeparator", "separator2", "Import")
              .Add NewMenuItem("button", "OpenSolverModelImportLP", "Import LP File...", _
                                "OpenSolver_ImportLPClick", "Import an existing LP file", _
                                "Parse an existing CPLEX LP file on to a new sheet and load it into the model.")
                
43            .Add NewMenuItem("menuSeparator", "separator3", "About")
44            .Add NewMenuItem("button", "OpenSolverAbout", "About OpenSolver...", "OpenSolver_AboutClickHandler")
45            .Add NewMenuItem("button", "OpenSolverAboutCoinOR", "About COIN-OR...", "OpenSolver_AboutCoinOR")
                                    
46            .Add NewMenuItem("menuSeparator", "separator4", "Help")
47            .Add NewMenuItem("button", "OpenSolverVisitOpenSolverOrg", "Open OpenSolver.org", "OpenSolver_VisitOpenSolverOrg")
48            .Add NewMenuItem("button", "OpenSolverVisitCoinOROrg", "Open COIN-OR.org", "OpenSolver_VisitCoinOROrg")

49        End With
50        Items.Add OpenSolverMenu
          
51        Set GenerateMenuItems = Items
End Function

'====================================================================
' Code for making the MenuBar menu
' Adapted from Excel 2003 Menu Code originally provided by Paul Becker of Eclipse Engineering (www.eclipseeng.com)
'====================================================================

Private Sub AddMenuItems_MenuBar()
1         DelMenuItems_MenuBar

          Dim MainMenuBar As CommandBar, OpenSolverMenu As CommandBarControl
2         Set MainMenuBar = Application.CommandBars(MenuBarName)
3         Set OpenSolverMenu = MainMenuBar.Controls.Add(Type:=msoControlPopup)
4         OpenSolverMenu.Caption = MenuName

          Dim Item As MenuItem
5         For Each Item In GenerateMenuItems()
6             AddToMenu_MenuBar OpenSolverMenu, Item
7         Next Item
End Sub

Private Sub AddToMenu_MenuBar(Menu As CommandBarControl, Item As MenuItem)
          Static sBeginGroup As Boolean

1         Select Case Item.Tag
          Case "button"
2             With Menu.Controls.Add(Type:=msoControlButton)
3               .Caption = Item.Label
4               .OnAction = Item.OnAction
5               .BeginGroup = sBeginGroup
6               .FaceId = 0
7             End With
8             sBeginGroup = False
9         Case "menuSeparator"
10            sBeginGroup = True
11        Case "menu", "splitButton"
              Dim SubMenu As CommandBarControl
12            Set SubMenu = Menu.Controls.Add(Type:=msoControlPopup)
13            SubMenu.BeginGroup = Item.NewGroup
              
              Dim Children As Collection
14            If Item.Tag = "menu" Then
15                SubMenu.Caption = Item.Label
16                Set Children = Item.Children
17            Else
18                SubMenu.Caption = Item.Children.Item(1).Label
19                Set Children = Item.Children.Item(2).Children
20            End If
              
              Dim SubItem As MenuItem
21            For Each SubItem In Children
22                AddToMenu_MenuBar SubMenu, SubItem
23            Next SubItem
24        End Select
End Sub

Private Sub DelMenuItems_MenuBar()
1         On Error Resume Next
2         Application.CommandBars(MenuBarName).Controls(MenuName).Delete
End Sub

'====================================================================
' Code for making the Ribbon XML
'====================================================================

Public Sub CreateRibbonXML()
          Dim CustomXMLFile As String
1         GetExistingFilePathName JoinPaths(ThisWorkbook.Path, "RibbonX", "customUI"), "customUI.xml", CustomXMLFile
          
          Dim FileNum As Integer
2         FileNum = FreeFile()
3         Open CustomXMLFile For Output As #FileNum
4             OutputRibbonHeader FileNum
              
              Dim Item As MenuItem
5             For Each Item In GenerateMenuItems()
6                 OutputRibbonXML FileNum, Item, 10
7             Next Item
              
8             OutputRibbonFooter FileNum
9         Close #FileNum
End Sub

Private Sub OutputRibbonHeader(FileNum As Integer)
1         Print #FileNum, Spc(0); "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
2         Print #FileNum, Spc(0); "<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">"
3         Print #FileNum, Spc(2); "<ribbon startFromScratch=""false"">"
4         Print #FileNum, Spc(4); "<tabs>"
5         Print #FileNum, Spc(6); "<tab idMso=""TabData"">"
6         Print #FileNum, Spc(8); "<group id=""GroupOpenSolver"" label=""OpenSolver"">"
End Sub

Private Sub OutputRibbonFooter(FileNum As Integer)
1         Print #FileNum, Spc(8); "</group>"
2         Print #FileNum, Spc(6); "</tab>"
3         Print #FileNum, Spc(4); "</tabs>"
4         Print #FileNum, Spc(2); "</ribbon>"
5         Print #FileNum, Spc(0); "</customUI>"
End Sub

Private Sub OutputRibbonXML(FileNum As Integer, Item As MenuItem, Indent As Integer)
1         Print #FileNum, Spc(Indent); "<" & Item.Tag & " id=" & Quote(Item.Id);
2         OutputIfExists FileNum, Replace(Item.Label, "&", "&amp;"), IIf(Item.Tag = "menuSeparator", "title", "label")
3         OutputIfExists FileNum, Item.OnAction, "onAction"
4         OutputIfExists FileNum, Item.ScreenTip, "screentip"
5         OutputIfExists FileNum, Item.SuperTip, "supertip"
6         OutputIfExists FileNum, Item.Size, IIf(Item.Tag = "menu", "itemSize", "size")
7         OutputIfExists FileNum, Item.Image, "image"
              
8         Select Case Item.Tag
          Case "button", "menuSeparator"
9             Print #FileNum, "/>"
10        Case "menu", "splitButton"
11            Print #FileNum, ">"
              
              Dim SubItem As MenuItem
12            For Each SubItem In Item.Children
13                OutputRibbonXML FileNum, SubItem, Indent + 2
14            Next SubItem
15            Print #FileNum, Spc(Indent); "</" & Item.Tag & ">"
16        End Select
End Sub

Private Sub OutputIfExists(FileNum As Integer, value As String, Tag As String)
1         If Len(value) <> 0 Then Print #FileNum, " " & Tag & "=" & Quote(value);
End Sub

