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

Private Const pName = "Solve Engine"
Private Const pShortName = "SolveEngine"
Private Const pTitle = "Satalia using Solve Engine"
Private Const pDesc = "Solve Engine provides optimization as a service. It uses the world's best algorithms from industry and academia to deliver rapid problem solving. To use the Solve Engine, register for an API key and enter the key when prompted."
Private Const pLink = "https://solve.satalia.com/register"

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
    Set resp = ParseJson(solution)
    
    If resp.Exists(strJsonKeyJobId) Then
        Dim result As Dictionary
        Set result = resp(strJsonKeyResult)
        
        Dim status As String
        status = result(strJsonKeyStatus)
        
        Select Case status
            Case strSolveStatusSEOptimal:
                s.SolveStatus = OpenSolverResult.Optimal
                s.SolveStatusString = "Optimal"
            Case strSolveStatusSEInfeasible:
                s.SolveStatus = OpenSolverResult.Infeasible
                s.SolveStatusString = "No Feasible Solution"
            Case strSolveStatusSEUnbounded:
                s.SolveStatus = OpenSolverResult.Unbounded
                s.SolveStatusString = "No Solution Found (Unbounded)"
            Case strSolveStatusSEInterrupted
                s.SolveStatus = OpenSolverResult.LimitedSubOptimal
                s.SolveStatusString = "Stopped on Time Limit"
            Case strSolveStatusSEFailed
                s.SolveStatus = OpenSolverResult.LimitedSubOptimal
                s.SolveStatusString = "Stopped on solver difficulties"
        End Select
        
        If result.Exists(strJsonKeyVariables) Then
            Dim variables As Collection
            Set variables = result(strJsonKeyVariables)
              
            Dim variable As Dictionary, i As Long
            i = 1
            For Each variable In variables
                s.VarFinalValue(i) = variable(strJsonKeyValue)
                s.VarCellName(i) = GetVarNameFromLPName(variable(strJsonKeyName))
                i = i + 1
            Next variable
            
            If variables.Count > 0 Then s.SolutionWasLoaded = True
        End If
    Else
        RaiseGeneralError "Unexpected error while getting solution: " & GetErrorMessage(resp)
    End If

End Sub

