VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSolverNeosCplex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISolver
Implements ISolverNeos
Implements ISolverLinear
Implements ISolverFile
Implements ISolverFileNL

Private Const pTitle = "NEOS using CPLEX (Linear solver)"
Private Const pDesc = "The NEOS (Network-Enabled Optimization System) Server is a free Internet-based service for solving optimization problems. The IBM ILOG CPLEX Optimization Studio (CPLEX) is an optimization software package." _
                        & vbNewLine & vbNewLine & "CPLEX requires an email address, which must be supplied under Model > Options. An email will be sent containing the model results, in addition to being loaded on to the sheet."
Private Const pLink = NeosTermsAndConditionsLink
Private Const pName = "CPLEX"
Private Const pShortName = "NeosCplex"

Private Const pSensitivityAnalysisAvailable As Boolean = False

Private Const pNeosSolverName As String = "CPLEX"
Private Const pNeosSolverCategory As String = "milp"
Private Const pOptionFile As String = vbNullString
Private Const pFileType = OpenSolver_FileType.NL
Private Const pModelType = OpenSolver_ModelType.Parsed

' ISolver

Private Property Get ISolver_Name() As String
1         ISolver_Name = pName
End Property

Private Property Get ISolver_ShortName() As String
1         ISolver_ShortName = pShortName
End Property

Private Property Get ISolver_Title() As String
1         ISolver_Title = pTitle
End Property

Private Property Get ISolver_Desc() As String
1         ISolver_Desc = pDesc
End Property

Private Property Get ISolver_Link() As String
1         ISolver_Link = pLink
End Property

Private Property Get ISolver_IterationLimitName() As String

End Property

Private Property Get ISolver_PrecisionName() As String

End Property

Private Property Get ISolver_TimeLimitName() As String

End Property

Private Property Get ISolver_ToleranceName() As String

End Property

Private Property Get ISolver_ModelType() As OpenSolver_ModelType
1         ISolver_ModelType = pModelType
End Property

Private Sub ISolver_CleanFiles()
End Sub
Private Sub ISolver_CheckLog(s As COpenSolver)
End Sub

' ISolverNeos

Private Property Get ISolverNeos_NeosSolverName() As String
1         ISolverNeos_NeosSolverName = pNeosSolverName
End Property

Private Property Get ISolverNeos_NeosSolverCategory() As String
1         ISolverNeos_NeosSolverCategory = pNeosSolverCategory
End Property

Private Property Get ISolverNeos_OptionFile() As String
1         ISolverNeos_OptionFile = pOptionFile
End Property

' ISolverLinear

Private Property Get ISolverLinear_SensitivityAnalysisAvailable() As Boolean
1         ISolverLinear_SensitivityAnalysisAvailable = pSensitivityAnalysisAvailable
End Property

' ISolverFile

Private Sub ISolverFile_ReadResults(s As COpenSolver, Optional solution As String)
1         ReadResults_NL s
End Sub

Private Property Get ISolverFile_FileType() As OpenSolver_FileType
1         ISolverFile_FileType = pFileType
End Property

' ISolverFileNL

Private Property Get ISolverFileNL_OptionFile() As String
1         ISolverFileNL_OptionFile = pOptionFile
End Property

