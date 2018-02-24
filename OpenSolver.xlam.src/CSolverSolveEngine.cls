VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSolverSolveEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISolver
Implements ISolverLinear
Implements ISolverFile

Private Const pName = "SolveEngine"
Private Const pShortName = "SolveEngine"
Private Const pTitle = "Satalia using SolveEngine"
Private Const pDesc = "SolveEngine provides optimization as a commercial online service. It uses algorithms from industry and academia to deliver rapid problem solving. To use the SolveEngine, register for an API key using the link below, choose a plan, and enter the API key when prompted."
Private Const pLink = "https://opensolver.org/solveengine/"

Private Const pSensitivityAnalysisAvailable As Boolean = False

Private Const pFileType = OpenSolver_FileType.LP
Private Const pModelType = OpenSolver_ModelType.Diff

' ISolver

Private Property Get ISolver_Name() As String:                      ISolver_Name = pName:           End Property
Private Property Get ISolver_ShortName() As String:                 ISolver_ShortName = pShortName: End Property
Private Property Get ISolver_Title() As String:                     ISolver_Title = pTitle:         End Property
Private Property Get ISolver_Desc() As String:                      ISolver_Desc = pDesc:           End Property
Private Property Get ISolver_Link() As String:                      ISolver_Link = pLink:           End Property
Private Property Get ISolver_IterationLimitName() As String:                                        End Property
Private Property Get ISolver_PrecisionName() As String:                                             End Property
Private Property Get ISolver_TimeLimitName() As String:                                             End Property
Private Property Get ISolver_ToleranceName() As String:                                             End Property
Private Property Get ISolver_ModelType() As OpenSolver_ModelType:   ISolver_ModelType = pModelType: End Property

Private Sub ISolver_CleanFiles()
End Sub

Private Sub ISolver_CheckLog(s As COpenSolver)
End Sub

' ISolverFile

Private Sub ISolverFile_ReadResults(s As COpenSolver, Optional solution As String)
1         ReadResults s, solution
End Sub

Private Property Get ISolverFile_FileType() As OpenSolver_FileType
1         ISolverFile_FileType = pFileType
End Property

' ISolverLinear

Private Property Get ISolverLinear_SensitivityAnalysisAvailable() As Boolean
1         ISolverLinear_SensitivityAnalysisAvailable = pSensitivityAnalysisAvailable
End Property

Private Sub ReadResults(s As COpenSolver, solution As String)
          Dim resp As Dictionary
1         Set resp = ParseJson(solution)
          
2         If resp.Exists(strJsonKeyJobId) Then
              Dim result As Dictionary
3             Set result = resp(strJsonKeyResult)
              
              Dim status As String
4             status = result(strJsonKeyStatus)
              
5             Select Case status
                  Case strSolveStatusSEOptimal:
6                     s.SolveStatus = OpenSolverResult.Optimal
7                     s.SolveStatusString = "Optimal"
8                 Case strSolveStatusSEInfeasible:
9                     s.SolveStatus = OpenSolverResult.Infeasible
10                    s.SolveStatusString = "No Feasible Solution"
11                Case strSolveStatusSEUnbounded:
12                    s.SolveStatus = OpenSolverResult.Unbounded
13                    s.SolveStatusString = "No Solution Found (Unbounded)"
14                Case strSolveStatusSEInterrupted
15                    s.SolveStatus = OpenSolverResult.LimitedSubOptimal
16                    s.SolveStatusString = "Stopped on Time Limit"
17                Case strSolveStatusSEFailed
18                    s.SolveStatus = OpenSolverResult.LimitedSubOptimal
19                    s.SolveStatusString = "Stopped on solver difficulties"
20            End Select
              
21            If result.Exists(strJsonKeyVariables) Then
                  Dim variables As Collection
22                Set variables = result(strJsonKeyVariables)
                    
                  Dim variable As Dictionary, i As Long
23                i = 1
24                For Each variable In variables
25                    s.VarFinalValue(i) = variable(strJsonKeyValue)
26                    s.VarCellName(i) = GetVarNameFromLPName(variable(strJsonKeyName))
27                    i = i + 1
28                Next variable
                  
29                If variables.Count > 0 Then s.SolutionWasLoaded = True
30            End If
31        Else
32            RaiseGeneralError "Unexpected error while getting solution: " & GetErrorMessage(resp)
33        End If

End Sub

