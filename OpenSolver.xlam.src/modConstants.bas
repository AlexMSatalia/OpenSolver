Attribute VB_Name = "modConstants"
'==============================================================================
' OpenSolver
' Copyright Andrew Mason, Iain Dunning, 2011
' http://www.opensolver.org
'==============================================================================
' modConstants
' OpenSolver-wide constants and enumerations
' Ideally, almost all strings and numbers should be here - no "magic" numbers
' floating in the code.
' It is also the home of the error constants - the associated descriptions
' are in the RaiseOSErr function. Allows for easier localisation too.
'==============================================================================
Option Explicit

Public Enum LpStatus
    LpStatusOptimal = 1
    LpStatusInfeasible = 2
    LpStatusUnbounded = 3
    LpStatusNotSolved = 4
    LpStatusIntegerInfeasible = 5
End Enum

Public Enum SolverConstraintType
    ConIsSingleCell
    ConIsMultipleCell
    ConIsValueOrFormula
End Enum

Public Const NameAdjCells As String = "solver_adj"
Public Const NameObjSense As String = "solver_typ"
Public Const NameObjTarget As String = "solver_val"
Public Const NameObjCell As String = "solver_opt"
Public Const NameConsCount As String = "solver_num"
Public Const NameLHS As String = "solver_lhs"
Public Const NameREL As String = "solver_rel"
Public Const NameRHS As String = "solver_rhs"
Public Const NameNonNegative As String = "solver_neg"
Public Const NameLinear As String = "solver_lin"
Public Const NameEngine As String = "solver_eng"
Public Const NameRelaxation As String = "solver_rlx"

Public Enum OpenSolverErrors
    GeneralExcelError
    OSErrObjectiveMissing
    OSErrObjectiveError
    OSErrObjectiveNAN
    
    OSErrAdjCellMissing
    OSErrAdjCellBadMerge
    OSErrAdjCellNone
    
    OSErrConsMissing
    OSErrConsLHSMissing
    OSErrConsLHSError
    OSErrConsLHSNotRange
    OSErrConsIntBinOnNonAdjustable
    OSErrConsUnknownRel
    OSErrConsRHSMissing
    OSErrConsRHSError
    OSErrConsBadCellCount
    
    OSErrPulpTokenErrText
    
    OSErrNoWorkbook
    OSErrActiveNotSheet
    
    OSErrUserCancel
    OSErrRelaxationSet
    OSErrDeleteFile
    OSErrTargetObjConstant
    OSErrUnsatisfiableConstraint
End Enum

'==============================================================================
' RaiseOSErr
' OpenSolver-wide error reporting system
' Whenever an error is caught/occurs, call this function with appropriate code
' from above, and the associated string will be returned here.
' TODO: Make this function check a global boolean that controls whether
' errors are silent or not.
Sub RaiseOSErr(intErrNumber As Integer, Optional varArg As Variant)
    Dim strDesc As String
    
    Select Case intErrNumber
        Case GeneralExcelError
            strDesc = "General Excel Error: " + varArg
        Case OSErrObjectiveMissing
            strDesc = "Couldn't find objective cell. Name 'solver_opt' is missing."
        Case OSErrObjectiveError
            strDesc = "Objective is marked #REF!. Cell may have been deleted."
        Case OSErrObjectiveNAN
            strDesc = "Objective cell doesn't contain a numeric value."
        
        Case OSErrAdjCellMissing
            strDesc = "Adjustable cell name could not be converted to a Range."
        Case OSErrAdjCellBadMerge
            strDesc = "Adjustable cell " & varArg.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & _
                      " that is inaccessible as it is within the merged range " & varArg.MergeArea.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "."
        Case OSErrAdjCellNone
            strDesc = "No valid adjustable cells in this model."
        
        Case OSErrConsMissing
            strDesc = "Contraint count missing ('solver_num')."
        Case OSErrConsLHSMissing
            strDesc = "LHS of a constraint is missing ('solver_lhs" & varArg & "')"
        Case OSErrConsLHSError
            strDesc = "LHS of a constraint is marked #REF!. Cell may have been deleted ('solver_lhs" & varArg & "')"
        Case OSErrConsLHSNotRange
            strDesc = "LHS of a constraint is not a range ('solver_lhs" & varArg & "')"
            
        Case OSErrConsIntBinOnNonAdjustable
            strDesc = "Binary or Integer constraint on a cell that is not an AdjustableCell. ('solver_rel" & varArg & "')"
        Case OSErrConsUnknownRel
            strDesc = "Unrecognised relationship for a constraint. ('solver_rel" & varArg & "')"
        
        Case OSErrConsRHSMissing
            strDesc = "RHS of a constraint is missing ('solver_rhs" & varArg & "')"
        Case OSErrConsRHSError
            strDesc = "RHS of a constraint is marked #REF!. Cell may have been deleted ('solver_rhs" & varArg & "')"
        Case OSErrConsBadCellCount
            strDesc = "Constraint has a different cell count on the left and the right. ('solver_rhs" & varArg & "')"
            
        Case OSErrPulpTokenErrText
            strDesc = "An error token was found while parsing the formulae."
            
        Case OSErrNoWorkbook
            strDesc = "No active workbook available."
        Case OSErrActiveNotSheet
            strDesc = "Active sheet is not a worksheet."
        
        Case OSErrUserCancel
            strDesc = "User chose to cancel solve."
        Case OSErrRelaxationSet
            strDesc = "The Solve Relaxation option is set. This must be off to use OpenSolver. To solve the relaxation, please use the OpenSolver Relaxation option to do so."
        Case OSErrDeleteFile
            strDesc = "Unable to delete file " & varArg
        Case OSErrTargetObjConstant
            strDesc = "The objective cell does not depend on the adjustable cells, so can not achieve target value."
        Case OSErrUnsatisfiableConstraint
            strDesc = "Unsatisfiable constraint"
    End Select
    
    Err.Raise vbObjectError + intErrNumber, "OpenSolver", strDesc
End Sub



