VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSolverNeosCou"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISolver
Implements ISolverNeos
Implements ISolverFile
Implements ISolverFileAMPL

Private pISolver As ISolver  ' Couenne solver to inherit properties from

Private Const pTitle = "NEOS using Couenne (Non-linear solver)"
Private Const pDesc = "The NEOS (Network-Enabled Optimization System) Server is a free Internet-based service for solving optimization problems. Couenne (Convex Over and Under ENvelopes for Nonlinear Estimation) is a branch & bound algorithm to solve Mixed-Integer Nonlinear Programming (MINLP) problems of specific forms."
Private Const pLink = NeosTermsAndConditionsLink
Private Const pShortName = "NeosCou"

Private Const pNeosSolverName = "Couenne"
Private Const pNeosSolverCategory = "minco"
Private Const pUsesOptionFile = True
Private Const pFileType = OpenSolver_FileType.AMPL
Private Const pModelType = OpenSolver_ModelType.Parsed

Private Const pAmplSolverName As String = "couenne"

' Class

Private Sub Class_Initialize()
    Set pISolver = New CSolverCouenne
End Sub

' ISolver

Private Property Get ISolver_Name() As String
    ISolver_Name = pISolver.Name
End Property

Private Property Get ISolver_ShortName() As String
    ISolver_ShortName = pShortName
End Property

Private Property Get ISolver_Title() As String
    ISolver_Title = pTitle
End Property

Private Property Get ISolver_Desc() As String
    ISolver_Desc = pDesc
End Property

Private Property Get ISolver_Link() As String
    ISolver_Link = pLink
End Property

Private Property Get ISolver_IterationLimitName() As String
    ISolver_IterationLimitName = pISolver.IterationLimitName
End Property

Private Property Get ISolver_PrecisionName() As String
    ISolver_PrecisionName = pISolver.PrecisionName
End Property

Private Property Get ISolver_TimeLimitName() As String
    ISolver_TimeLimitName = pISolver.TimeLimitName
End Property

Private Property Get ISolver_ToleranceName() As String
    ISolver_ToleranceName = pISolver.ToleranceName
End Property

Private Property Get ISolver_ModelType() As OpenSolver_ModelType
    ISolver_ModelType = pModelType
End Property

Private Sub ISolver_CleanFiles()
End Sub

Private Sub ISolver_CheckLog(s As COpenSolver)
    pISolver.CheckLog s
End Sub

' ISolverNeos

Private Property Get ISolverNeos_NeosSolverName() As String
    ISolverNeos_NeosSolverName = pNeosSolverName
End Property

Private Property Get ISolverNeos_NeosSolverCategory() As String
    ISolverNeos_NeosSolverCategory = pNeosSolverCategory
End Property

Private Property Get ISolverNeos_UsesOptionFile() As Boolean
    ISolverNeos_UsesOptionFile = pUsesOptionFile
End Property

' ISolverFile

Private Sub ISolverFile_ReadResults(s As COpenSolver, Optional solution As String)
    ReadResults_AMPL s, solution
End Sub

Private Property Get ISolverFile_FileType() As OpenSolver_FileType
    ISolverFile_FileType = pFileType
End Property

' ISolverFileAMPL

Private Property Get ISolverFileAMPL_AmplSolverName() As String
    ISolverFileAMPL_AmplSolverName = pAmplSolverName
End Property

