VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ObjectiveSense As ObjectiveSenseType
Public ObjectiveFunctionCell As Range
Public ObjectiveTarget As Double

Public DecisionVariables As Range

Public Constraints As Collection

Public Duals As Range   ' The range where we store the duals; specific to OpenSolver

Public NonNegativityAssumption As Boolean

Function BuildModel() As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          ' Check int/bin constraints not set on non-decision variables before we start saving things
          Dim curCon As CConstraint, ShowWarning As Boolean
          ShowWarning = False
3924      For Each curCon In Constraints
3925          If Not RelationHasRHS(curCon.RelationType) Then
3927              If Not SetDifference(curCon.LHS, DecisionVariables) Is Nothing Then
                      If MsgBox("This model has specified that a non-decision cell must take an integer/binary value. " & _
                                "This is a valid model, but not one that OpenSolver can solve. " & _
                                "Do you wish to continue with saving this model?", _
                                vbQuestion + vbYesNo, "OpenSolver - Warning") = vbYes Then
                          GoTo ContinueSaving
                      Else
                          BuildModel = False
                          GoTo ExitFunction
                      End If
                  End If
3934          End If
3935      Next curCon

ContinueSaving:
3906      SetDecisionVariables DecisionVariables
3911      SetObjectiveFunctionCell ObjectiveFunctionCell
          
          ' Objective type
3916      SetObjectiveSense ObjectiveSense
3918      If ObjectiveSense = TargetObjective Then SetObjectiveTargetValue ObjectiveTarget
          
          ' Constraints
3936      SetNumConstraints Constraints.Count
          
          Dim c As Long
3937      For c = 1 To Constraints.Count
              Set curCon = Constraints(c)
              With curCon
                  UpdateConstraint c, .LHS, .RelationType, .RHS, .RHSstring
              End With
3968      Next c
          
3969      SetDuals Duals
3974      SetNonNegativity NonNegativityAssumption

          BuildModel = True

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CModel", "BuildModel") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Public Sub PopulateConstraintListBox(ByRef lst As MSForms.ListBox, showNameRange As Boolean)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim DisplayString As String, curCon As CConstraint
          Dim oldLI As Long
          
3984      If showNameRange Then SearchRangeName_DestroyCache
          
3985      oldLI = lst.ListIndex
3986      lst.Clear
3987      lst.AddItem "<Add new constraint>"
3988      For Each curCon In Constraints
3989          With curCon
3990              DisplayString = GetDisplayAddress(.LHS, showNameRange) & " " & RelationEnumToString(.RelationType)
3991              If Not .RHS Is Nothing Then
                      ' This is not a type restriction, so it has a RHS
3992                  DisplayString = DisplayString + " " + GetDisplayAddress(.RHS, showNameRange)
3993              ElseIf .RHSstring <> "" Then
                      Dim DisplayStringRHS As String
                      DisplayStringRHS = ConvertToCurrentLocale(.RHSstring)
3994                  If left(DisplayStringRHS, 1) = "=" Then DisplayString = Mid(DisplayStringRHS, 2)
3995                  DisplayString = DisplayString + " " + RemoveActiveSheetNameFromString(DisplayStringRHS)
3999              End If
4000          End With
4001          lst.AddItem DisplayString
4002      Next curCon

          ' Restore selected index, forcing a valid selection if needed
4006      lst.ListIndex = Max(Min(oldLI, lst.ListCount - 1), 0)

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("CModel", "PopulateConstraintListBox") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Public Sub LoadFromSheet()
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

4009      Set Constraints = New Collection

          Dim sheetName As String
4010      sheetName = EscapeSheetName(ActiveWorkbook.ActiveSheet)

          Set DecisionVariables = GetDecisionVariablesWithDefault()
          
4012      NonNegativityAssumption = GetNonNegativity()
           
4014      ObjectiveSense = GetObjectiveSense()
4016      ObjectiveTarget = GetObjectiveTargetValue()
4018      Set ObjectiveFunctionCell = GetObjectiveFunctionCell()
          
          Dim constraint As Long, NumConstraints As Long
4020      NumConstraints = GetNumConstraints()
4021      For constraint = 1 To NumConstraints
              
              Dim rLHS As Range
4022          On Error Resume Next
              Set rLHS = GetConstraintLhs(constraint)
              If Err.Number <> 0 Then GoTo SkipCon
              On Error GoTo ErrorHandler
              
              Dim rel As RelationConsts
4024          rel = GetConstraintRel(constraint)
                            
4025          If rel < RelationConsts.[_First] Or rel > RelationConsts.[_Last] Then GoTo SkipCon
              
              If Not RelationHasRHS(rel) Then
4026              AddConstraintToModel Constraints, rLHS, rel
              Else
                  Dim valRHS As Double, rRHS As Range, RHSRefersToFormula As Boolean, sRefersToRHS As String
4030              On Error Resume Next
                  Set rRHS = GetConstraintRhs(constraint, sRefersToRHS, valRHS, RHSRefersToFormula)
4031              If Err.Number <> 0 Then GoTo SkipCon
                  On Error GoTo ErrorHandler

4033              AddConstraintToModel Constraints, rLHS, rel, Nothing, rRHS, sRefersToRHS
              End If
SkipCon:
4043      Next constraint
          
          Set Duals = GetDuals()
          
ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("CModel", "LoadFromSheet") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub
