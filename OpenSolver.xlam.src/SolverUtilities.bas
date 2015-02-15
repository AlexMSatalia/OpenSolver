Attribute VB_Name = "SolverUtilities"
' Functions that relate to multiple solvers, or delegate to internal solver functions.
Option Explicit
Public Const LPFileName As String = "model.lp"
Public Const AMPLFileName As String = "model.ampl"
Public Const NLFileName As String = "model.nl"
Public Const NLSolutionFileName As String = "model.sol"
Public Const PuLPFileName As String = "opensolver.py"

Public Const SolverDir As String = "Solvers"
Public Const SolverDirMac As String = "osx"
Public Const SolverDirWin32 As String = "win32"
Public Const SolverDirWin64 As String = "win64"

Function SolverAvailable(Solver As String, Optional SolverPath As String, Optional errorString As String) As Boolean
      ' Delegated function returns True if solver is available and sets SolverPath to location of solver
          
          'All Neos solvers always available
6570      If RunsOnNeos(Solver) Then
6571          SolverAvailable = True
6572          SolverPath = ""
6573          Exit Function
6574      End If

6575      Select Case Solver
          Case "CBC"
6576          SolverAvailable = SolverAvailable_CBC(SolverPath, errorString)
6577      Case "Gurobi"
6578          SolverAvailable = SolverAvailable_Gurobi(SolverPath, errorString)
6579      Case "NOMAD"
6580          SolverAvailable = SolverAvailable_NOMAD(errorString)
6581      Case "Bonmin"
6582          SolverAvailable = SolverAvailable_Bonmin(SolverPath, errorString)
6583      Case "Couenne"
6584          SolverAvailable = SolverAvailable_Couenne(SolverPath, errorString)
6585      Case Else
6586          SolverAvailable = False
6587          SolverPath = ""
6588      End Select
End Function

Function SolverFilePath_Default(Solver As String, Optional errorString As String) As String
          Dim SearchPath As String
6589      SearchPath = JoinPaths(ThisWorkbook.Path, SolverDir)
          
#If Mac Then
6590      If GetExistingFilePathName(JoinPaths(SearchPath, SolverDirMac), SolverName(Solver), SolverFilePath_Default) Then Exit Function ' Found a mac solver
6591      errorString = "Unable to find Mac version of " & Solver & " ('" & SolverName(Solver) & "') in the Solvers folder."
6592      SolverFilePath_Default = ""
6593      Exit Function
#Else
          ' Look for the 64 bit version
6594      If SystemIs64Bit Then
6595          If GetExistingFilePathName(JoinPaths(SearchPath, SolverDirWin64), SolverName(Solver), SolverFilePath_Default) Then Exit Function ' Found a 64 bit solver
6596      End If
          ' Look for the 32 bit version
6597      If GetExistingFilePathName(JoinPaths(SearchPath, SolverDirWin32), SolverName(Solver), SolverFilePath_Default) Then
6598          If SystemIs64Bit Then
6599              errorString = "Unable to find 64-bit " & Solver & " in the Solvers folder. A 32-bit version will be used instead."
6600          End If
6601          Exit Function
6602      End If
          ' Fail
6603      SolverFilePath_Default = ""
6604      errorString = "Unable to find " & Solver & " ('" & SolverName(Solver) & "') in the Solvers folder."
#End If
End Function

Function SolverType(Solver As String) As String
6605      Select Case Solver
          Case "CBC"
6606          SolverType = SolverType_CBC
6607      Case "Gurobi"
6608          SolverType = SolverType_Gurobi
6609      Case "NeosCBC"
6610          SolverType = SolverType_NeosCBC
6611      Case "NOMAD"
6612          SolverType = SolverType_NOMAD
6613      Case "NeosBon"
6614          SolverType = SolverType_NeosBon
6615      Case "NeosCou"
6616          SolverType = SolverType_NeosCou
6617      Case "Bonmin"
6618          SolverType = SolverType_Bonmin
6619      Case "Couenne"
6620          SolverType = SolverType_Couenne
6621      Case Else
6622          SolverType = OpenSolver_SolverType.Unknown
6623      End Select
End Function

Function SolverName(Solver As String) As String
6624      Select Case Solver
          Case "CBC"
6625          SolverName = SolverName_CBC
6626      Case "Gurobi"
6627          SolverName = SolverName_Gurobi
6628      Case "Bonmin"
6629          SolverName = SolverName_Bonmin
6630      Case "Couenne"
6631          SolverName = SolverName_Couenne
6632      Case Else
6633          SolverName = ""
6634      End Select
End Function

Function SolutionFilePath(Solver As String) As String
6635      Select Case Solver
          Case "CBC"
6636          SolutionFilePath = SolutionFilePath_CBC
6637      Case "Gurobi"
6638          SolutionFilePath = SolutionFilePath_Gurobi
6639      Case "Bonmin"
6640          SolutionFilePath = SolutionFilePath_Bonmin
6641      Case "Couenne"
6642          SolutionFilePath = SolutionFilePath_Couenne
6643      Case "PuLP"
6644          SolutionFilePath = SolutionFilePath_PuLP
6645      Case Else
6646          SolutionFilePath = ""
6647      End Select
End Function

Sub CleanFiles(errorPrefix)
6648      CleanFiles_CBC (errorPrefix)
6649      CleanFiles_Gurobi (errorPrefix)
End Sub

Function ReadModel(Solver As String, SolutionFilePathName As String, errorString As String, s As COpenSolver) As Boolean
6650      Select Case Solver
          Case "CBC"
6651          ReadModel = ReadModel_CBC(SolutionFilePathName, errorString, s)
6652      Case "Gurobi"
6653          ReadModel = ReadModel_Gurobi(SolutionFilePathName, errorString, s)
6654      Case Else
6655          ReadModel = False
6656          errorString = "The solver " & Solver & " has not yet been incorporated fully into OpenSolver."
6657      End Select
End Function

Function ModelFile(Solver As String) As String
6666      Select Case Solver
          Case "CBC", "Gurobi"
              ' output the model to an LP format text file
              ' See http://lpsolve.sourceforge.net/5.5/CPLEX-format.htm
6667          ModelFile = LPFileName
6668      Case "NeosCBC", "NeosCou", "NeosBon"
6669          ModelFile = AMPLFileName
6670      Case "PuLP"
6671          ModelFile = PuLPFileName
6672      Case "Couenne", "Bonmin"
6673          ModelFile = NLFileName
6674      Case Else
6675          ModelFile = ""
6676      End Select
End Function

Function ModelFilePath(Solver As String) As String
6677      ModelFilePath = ModelFile(Solver)
          ' If model file is empty, then don't return anything
6678      If ModelFilePath <> "" Then
6679          ModelFilePath = GetTempFilePath(ModelFilePath)
6680      End If
End Function

Function RunsOnNeos(Solver As String) As Boolean
6681      If Solver Like "Neos*" Then
6682          RunsOnNeos = True
6683      Else
6684          RunsOnNeos = False
6685      End If
End Function

Function SolverUsesUpperBounds(Solver As String) As Boolean
6686      Select Case Solver
          Case "NOMAD"
6687          SolverUsesUpperBounds = True
6688      Case Else
6689          SolverUsesUpperBounds = False
6690      End Select
End Function

Sub GetNeosValues(Solver As String, Category As String, SolverType As String)
6691      Select Case Solver
          Case "NeosCBC"
6692          Category = "milp"
6693          SolverType = "cbc"
6694      Case "NeosBon"
6695          Category = "minco"
6696          SolverType = "Bonmin"
6697      Case "NeosCou"
6698          Category = "minco"
6699          SolverType = "Couenne"
6700      End Select
End Sub

Function GetAmplSolverValues(Solver As String) As String
6701      Select Case Solver
          Case "NeosCBC"
6702          GetAmplSolverValues = "cbc"
6703      Case "NeosBon"
6704          GetAmplSolverValues = "bonmin"
6705      Case "NeosCou"
6706          GetAmplSolverValues = "couenne"
6707      End Select
End Function

Function UsesParsedModel(Solver As String) As Boolean
6708      Select Case Solver
          Case "PuLP", "NeosBon", "NeosCou", "Couenne", "Bonmin"
6709          UsesParsedModel = True
6710      Case Else
6711          UsesParsedModel = False
6712      End Select
End Function

Function DoBackSubstitution(Solver As String) As Boolean
          ' Enable when .nl files are added
6713      DoBackSubstitution = False
End Function


Function GetExtraParameters(Solver As String, sheet As Worksheet, errorString As String) As String
6714      Select Case Solver
          Case "CBC"
6715          GetExtraParameters = GetExtraParameters_CBC(sheet, errorString)
6716      Case Else
6717          GetExtraParameters = ""
6718      End Select
End Function

Function CreateSolveScript(Solver As String, SolutionFilePathName As String, ExtraParametersString As String, SolveOptions As SolveOptionsType, s As COpenSolver) As String
6719      Select Case Solver
          Case "CBC"
6720          CreateSolveScript = CreateSolveScript_CBC(SolutionFilePathName, ExtraParametersString, SolveOptions, s)
6721      Case "Gurobi"
6722          CreateSolveScript = CreateSolveScript_Gurobi(SolutionFilePathName, ExtraParametersString, SolveOptions)
6723      End Select
End Function

Function CreateSolveScriptParsed(Solver As String, SolutionFilePathName As String, SolveOptions As SolveOptionsType) As String
6724      Select Case Solver
          Case "Bonmin"
6725          CreateSolveScriptParsed = CreateSolveScript_Bonmin(SolutionFilePathName, SolveOptions)
6726      Case "Couenne"
6727          CreateSolveScriptParsed = CreateSolveScript_Couenne(SolutionFilePathName, SolveOptions)
6728      End Select
End Function

Function ScriptFilePath(Solver As String) As String
6724      Select Case Solver
          Case "Bonmin"
6725          ScriptFilePath = ScriptFilePath_Bonmin()
6726      Case "Couenne"
6727          ScriptFilePath = ScriptFilePath_Couenne()
6728      End Select
End Function

Function SolverTitle(Solver As String) As String
6729      Select Case Solver
          Case "CBC"
6730          SolverTitle = SolverTitle_CBC
6731      Case "Gurobi"
6732          SolverTitle = SolverTitle_Gurobi
6733      Case "Bonmin"
6734          SolverTitle = SolverTitle_Bonmin
6735      Case "Couenne"
6736          SolverTitle = SolverTitle_Couenne
6737      Case "NOMAD"
6738          SolverTitle = SolverTitle_NOMAD
6739      Case "NeosCBC"
6740          SolverTitle = SolverTitle_NeosCBC
6741      Case "NeosCou"
6742          SolverTitle = SolverTitle_NeosCou
6743      Case "NeosBon"
6744          SolverTitle = SolverTitle_NeosBon
6745      End Select
End Function

Function ReverseSolverTitle(SolverTitle As String) As String
6746      Select Case SolverTitle
          Case SolverTitle_CBC
6747          ReverseSolverTitle = "CBC"
6748      Case SolverTitle_Gurobi
6749          ReverseSolverTitle = "Gurobi"
6750      Case SolverTitle_NOMAD
6751          ReverseSolverTitle = "NOMAD"
6752      Case SolverTitle_Bonmin
6753          ReverseSolverTitle = "Bonmin"
6754      Case SolverTitle_Couenne
6755          ReverseSolverTitle = "Couenne"
6756      Case SolverTitle_NeosCBC
6757          ReverseSolverTitle = "NeosCBC"
6758      Case SolverTitle_NeosCou
6759          ReverseSolverTitle = "NeosCou"
6760      Case SolverTitle_NeosBon
6761          ReverseSolverTitle = "NeosBon"
6762      End Select
End Function

Function SolverDesc(Solver As String) As String
6763      Select Case Solver
          Case "CBC"
6764          SolverDesc = SolverDesc_CBC
6765      Case "Gurobi"
6766          SolverDesc = SolverDesc_Gurobi
6767      Case "NOMAD"
6768          SolverDesc = SolverDesc_NOMAD
6769      Case "Bonmin"
6770          SolverDesc = SolverDesc_Bonmin
6771      Case "Couenne"
6772          SolverDesc = SolverDesc_Couenne
6773      Case "NeosCBC"
6774          SolverDesc = SolverDesc_NeosCBC
6775      Case "NeosCou"
6776          SolverDesc = SolverDesc_NeosCou
6777      Case "NeosBon"
6778          SolverDesc = SolverDesc_NeosBon
6779      End Select
End Function

Function SolverLink(Solver As String) As String
6780      Select Case Solver
          Case "CBC"
6781          SolverLink = SolverLink_CBC
6782      Case "Gurobi"
6783          SolverLink = SolverLink_Gurobi
6784      Case "NOMAD"
6785          SolverLink = SolverLink_NOMAD
6786      Case "Bonmin"
6787          SolverLink = SolverLink_Bonmin
6788      Case "Couenne"
6789          SolverLink = SolverLink_Couenne
6790      Case "NeosCBC"
6791          SolverLink = SolverLink_NeosCBC
6792      Case "NeosCou"
6793          SolverLink = SolverLink_NeosCou
6794      Case "NeosBon"
6795          SolverLink = SolverLink_NeosBon
6796      End Select
End Function

Function SolverHasSensitivityAnalysis(Solver As String) As Boolean
          ' Non-linear solvers don't have sensitivity analysis
6797      If SolverType(Solver) = OpenSolver_SolverType.NonLinear Then
6798          SolverHasSensitivityAnalysis = False
6799      End If
          
6800      Select Case Solver
          Case "CBC", "Gurobi"
6801          SolverHasSensitivityAnalysis = True
6802      Case Else
6803          SolverHasSensitivityAnalysis = False
6804      End Select
End Function

Function UsesPrecision(Solver As String) As String
    Select Case Solver
    Case "CBC"
        UsesPrecision = UsesPrecision_CBC
    Case "Gurobi"
        UsesPrecision = UsesPrecision_Gurobi
    Case "Bonmin"
        UsesPrecision = UsesPrecision_Bonmin
    Case "Couenne"
        UsesPrecision = UsesPrecision_Couenne
    Case "NOMAD"
        UsesPrecision = UsesPrecision_NOMAD
    Case "NeosCBC"
        UsesPrecision = UsesPrecision_NeosCBC
    Case "NeosCou"
        UsesPrecision = UsesPrecision_NeosCou
    Case "NeosBon"
        UsesPrecision = UsesPrecision_NeosBon
    End Select
End Function

Function UsesTimeLimit(Solver As String) As String
    Select Case Solver
    Case "CBC"
        UsesTimeLimit = UsesTimeLimit_CBC
    Case "Gurobi"
        UsesTimeLimit = UsesTimeLimit_Gurobi
    Case "Bonmin"
        UsesTimeLimit = UsesTimeLimit_Bonmin
    Case "Couenne"
        UsesTimeLimit = UsesTimeLimit_Couenne
    Case "NOMAD"
        UsesTimeLimit = UsesTimeLimit_NOMAD
    Case "NeosCBC"
        UsesTimeLimit = UsesTimeLimit_NeosCBC
    Case "NeosCou"
        UsesTimeLimit = UsesTimeLimit_NeosCou
    Case "NeosBon"
        UsesTimeLimit = UsesTimeLimit_NeosBon
    End Select
End Function

Function UsesIterationLimit(Solver As String) As String
    Select Case Solver
    Case "CBC"
        UsesIterationLimit = UsesIterationLimit_CBC
    Case "Gurobi"
        UsesIterationLimit = UsesIterationLimit_Gurobi
    Case "Bonmin"
        UsesIterationLimit = UsesIterationLimit_Bonmin
    Case "Couenne"
        UsesIterationLimit = UsesIterationLimit_Couenne
    Case "NOMAD"
        UsesIterationLimit = UsesIterationLimit_NOMAD
    Case "NeosCBC"
        UsesIterationLimit = UsesIterationLimit_NeosCBC
    Case "NeosCou"
        UsesIterationLimit = UsesIterationLimit_NeosCou
    Case "NeosBon"
        UsesIterationLimit = UsesIterationLimit_NeosBon
    End Select
End Function

Function UsesTolerance(Solver As String) As String
    Select Case Solver
    Case "CBC"
        UsesTolerance = UsesTolerance_CBC
    Case "Gurobi"
        UsesTolerance = UsesTolerance_Gurobi
    Case "Bonmin"
        UsesTolerance = UsesTolerance_Bonmin
    Case "Couenne"
        UsesTolerance = UsesTolerance_Couenne
    Case "NOMAD"
        UsesTolerance = UsesTolerance_NOMAD
    Case "NeosCBC"
        UsesTolerance = UsesTolerance_NeosCBC
    Case "NeosCou"
        UsesTolerance = UsesTolerance_NeosCou
    Case "NeosBon"
        UsesTolerance = UsesTolerance_NeosBon
    End Select
End Function

Function OptionsFilePath(Solver As String) As String
    Select Case Solver
    Case "Bonmin"
        OptionsFilePath = OptionsFilePath_Bonmin
    Case "Couenne"
        OptionsFilePath = OptionsFilePath_Couenne
    End Select
End Function
