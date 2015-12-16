VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CModelParsed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Takes a Solver model and
' 1. Validates it according to the same rules Solver uses
' 2. Converts it to a form that will be used by the output model builders
'==============================================================================
Option Explicit


Public Formulae As Collection
Public VarTypeMap As Collection
Public AdjCellNameMap As Collection
Public LHSKeys As Collection
Public RHSKeys As Collection
Public RELs As Collection

Public lngMaxDepth As Long

Private s As COpenSolver

'==============================================================================
Public Function ProcessSolverModel(OpenSolver As COpenSolver) As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Set s = OpenSolver
          
          Set SolverModelParsed.SheetNameMap = New Collection
          Set SolverModelParsed.SheetNameMapReverse = New Collection

          Dim c As Range
4793      For Each c In s.AdjustableCells
              ' Reindex lower bound collection by local address
              ' TODO why not just use absolute?
4794          If TestKeyExists(s.VarLowerBounds, c.Address) Then
                  Dim bound As Double
                  bound = s.VarLowerBounds(c.Address)
                  s.VarLowerBounds.Remove c.Address
4795              s.VarLowerBounds.Add bound, c.Address(RowAbsolute:=False, ColumnAbsolute:=False)
4796          End If
4797      Next c

          ' =========================================
          ' Conduct parsed-specific processing
4816      ProcessVarTypes
4817      ProcessSolverModel = ParseFormulae()

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CModelParsed", "ProcessSolverModel") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Private Sub ProcessVarTypes()
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          ' Get var types of constraints
4818      Set VarTypeMap = New Collection
4819      Set AdjCellNameMap = New Collection
          Dim curVarType As Long, lngArea As Long, cellName As String, c As Range
4820      For lngArea = 1 To s.AdjustableCells.Areas.Count
4821          For Each c In s.AdjustableCells.Areas(lngArea)
4822              cellName = ConvertCellToStandardName(c)
4823              If Not TestKeyExists(AdjCellNameMap, cellName) Then
4824                  AdjCellNameMap.Add c, cellName
4825                  curVarType = VarContinuous
4826                  If Not (s.IntegerCellsRange Is Nothing) Then
4827                      If Not (Intersect(c, s.IntegerCellsRange) Is Nothing) Then curVarType = VarInteger
4828                  End If
4829                  If Not (s.BinaryCellsRange Is Nothing) Then
4830                      If Not (Intersect(c, s.BinaryCellsRange) Is Nothing) Then curVarType = VarBinary
4831                  End If
4832                  VarTypeMap.Add curVarType, cellName
4833              End If
4834          Next
4835      Next

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("CModelParsed", "ProcessVarTypes") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Private Function ParseFormulae() As Boolean
' Parses all constraints and objective formulae to create formulae that depend only on variable cells and constants
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          
          ' Create a collection for the spreadsheet's formulae (DAG)
4837      Set Formulae = New Collection
          ' Remember the greatest depth of the DAG
4838      lngMaxDepth = 0

          '==========================================================================
          ' STEP 1. Parse objective cell
          Dim objresult As String
          
4839      If Not s.ObjRange Is Nothing Then
4840          objresult = ExpandFormula(s.ObjRange.Formula, s.ObjRange, 0)
              
4841          If objresult Like "*Error*" Then
4842              Err.Raise OpenSolver_BuildError, Description:="The objective contains an error"
4843          End If
4844          CleanFormulae 1, Formulae.Count
4845      End If
          
          '==========================================================================
          ' STEP 2. Parse constraints
4846      Set LHSKeys = New Collection
4847      Set RHSKeys = New Collection
4848      Set RELs = New Collection

          Dim i As Long
          Dim Count As Long
4849      Count = 1
4850      For i = 1 To s.NumConstraints
              Dim lngFormulaeCountBefore As Long
              Dim cRow As Long, cCol As Long, cRowCount As Long, cColCount As Long
              Dim varLHSFormulae As Variant, rngLHS As Range, curLHSFormulae As String
4851          Set rngLHS = s.LHSRange(i)
              
              Dim rel As RelationConsts
4852          rel = s.RELs(i)
              If Not RelationHasRHS(rel) Then GoTo NextCons

              Dim strRHS As String, rngRHS As Range
              Set rngRHS = s.RHSRange(i)
              strRHS = s.RHSFormula(i)
              
4857          lngFormulaeCountBefore = Formulae.Count
              
              ' Possibilities
              ' LHS is one cell, RHS is one cell/formula/value
              ' LHS is a range (any shape), RHS is one cell/formula/value
              ' LHS is a range (any shape), RHS is a range of same size, not necessarily same shape
              ' Because mismatched shapes are cruel, we will just throw an error if we find them
              ' OpenSolver1 just threw a general error.
              
4858          cRowCount = rngLHS.Rows.Count
4859          cColCount = rngLHS.Columns.Count
4861          varLHSFormulae = rngLHS.Formula
                      
              Dim singleRngLHS As Range
4862          For cRow = 1 To cRowCount
4863              For cCol = 1 To cColCount
4864                  Count = Count + 1
                      
                      ' Get LHS
4865                  If cRowCount <> 1 Or cColCount <> 1 Then
4866                      curLHSFormulae = varLHSFormulae(cRow, cCol)
4867                      Set singleRngLHS = rngLHS(cRow, cCol)
4868                  Else
4869                      curLHSFormulae = varLHSFormulae
4870                      Set singleRngLHS = rngLHS
4871                  End If
                      
                      ' Get RHS
                      Dim rngRHSCell As Range, strRHSCell As String
4872                  Select Case s.RHSType(i)
                      ' 1.1 RHS is a single cell
                      Case SingleCellRange
4874                      Set rngRHSCell = rngRHS
4873                      strRHSCell = rngRHSCell.Formula
                      ' 1.2 RHS is multiple cells
4875                  Case MultiCellRange
4877                      Set rngRHSCell = rngRHS.Cells(cRow, cCol)
4876                      strRHSCell = rngRHSCell.Formula
                      ' 1.3 RHS is value or formula
4878                  Case constant, Formula
4879                      strRHSCell = "=" & strRHS
4880                      Set rngRHSCell = Nothing
4881                  End Select

                      ' If LHS is constant, then check if RHS is constant
4882                  If CellIsConstant(singleRngLHS) Then
                          Dim isConst As Boolean
4883                      isConst = False
                           
4884                      Select Case s.RHSType(i)
                          Case constant
4885                          isConst = True
4886                      Case Formula
4887                          isConst = StringFormulaIsConstant(strRHSCell)
4888                      Case SingleCellRange, MultiCellRange
4889                          isConst = CellIsConstant(rngRHSCell)
4890                      End Select
                          
                          ' If both sides are constant we need to see if the constraint is valid
4891                      If isConst Then
                              Dim isValid As Boolean, LHSValue As Double, RHSValue As Double
4892                          isValid = True
                              
4893                          LHSValue = singleRngLHS.Value2
4894                          If strRHSCell = "" Then
4895                              RHSValue = 0
4896                          Else
4897                              RHSValue = s.sheet.Evaluate(strRHSCell)
4898                          End If
                              
                              ' Test LHS and RHS in the constraint
4899                          Select Case rel
                              Case RelationGE
4900                              If LHSValue < RHSValue Then
4901                                  isValid = False
4902                              End If
4903                          Case RelationLE
4904                              If LHSValue > RHSValue Then
4905                                  isValid = False
4906                              End If
4907                          Case RelationEQ
4908                              If LHSValue <> RHSValue Then
4909                                  isValid = False
4910                              End If
4911                          End Select
                              
                              ' If constraint is not satisfied, the model is infeasible
4912                          If Not isValid Then
4913                              s.SolveStatus = OpenSolverResult.Infeasible
4914                              s.SolveStatusString = "Infeasible"
4915                              s.SolveStatusComment = _
                                      "The model contains an Excel constraint which does not depend on the decision variables and is not satisfied." & vbCrLf & _
                                      "Constraint specifies: " & _
                                      "LHS=" & singleRngLHS.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "=" & LHSValue & _
                                      "   " & RelationEnumToString(rel) & "   " & _
                                      "RHS=" & "'" & strRHSCell & "'" & "=" & RHSValue
4916                              ParseFormulae = False
4917                              GoTo ExitFunction
4918                          Else
                                  ' Otherwise the constraint is trivially satisfied, we can ignore it
4919                              GoTo NextCol
4920                          End If
4921                      End If
4922                  End If
                      
                      ' Parse LHS
4923                  LHSKeys.Add ExpandFormula(curLHSFormulae, rngLHS(cRow, cCol), i)
                      
                      ' Parse RHS
4924                  RHSKeys.Add ExpandFormula(strRHSCell, rngRHSCell, i)
                      
                      ' Determine appropriate relation
4925                  RELs.Add rel
NextCol:
4926              Next cCol
4927          Next cRow
              
              ' Clean Formulae
4928          CleanFormulae lngFormulaeCountBefore + 1, Formulae.Count, i
              
NextCons:
4929      Next i
          
          '==========================================================================
          ' STEP 3. Make all possible substitutions in formula tree if required by solver
4931      If DoBackSubstitution(s.Solver) Then
          
              Dim lngCurDepth As Long
              Dim lngIndex As Long, objCurNode As CFormula, strParentAdr As Variant
              
4932          For lngCurDepth = lngMaxDepth To 0 Step -1
4933              For lngIndex = 1 To Formulae.Count
4934                  If Formulae(lngIndex).lngDepth = lngCurDepth Then
4935                      Set objCurNode = Formulae(lngIndex)
                          
4936                      For Each strParentAdr In objCurNode.Parents
4937                          With Formulae(strParentAdr)
4938                              .strFormulaParsed = Replace(.strFormulaParsed, objCurNode.strAddress, "(" & objCurNode.strFormulaParsed & ")")
4939                          End With
4940                      Next
4941                  End If
4942              Next lngIndex
4943          Next lngCurDepth
4944      End If
4945      ParseFormulae = True

ExitFunction:
          Application.StatusBar = False
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CModelParsed", "ParseFormulae") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function DoBackSubstitution(Solver As ISolver) As Boolean
6713      DoBackSubstitution = False
End Function

' Checks whether a cell is constant i.e. not a decision variable and does not depend on a decision variable
Function CellIsConstant(cell As Range) As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          ' Check if cell is a variable cell
          If cell.Parent.Name = s.AdjustableCells.Parent.Name Then
4964          If Not Intersect(cell, s.AdjustableCells) Is Nothing Then
4965              CellIsConstant = False
4966              GoTo ExitFunction
4967          End If
          End If
          
          ' Check if dependent on decision variable
4968      If cell.HasFormula Then
              
              ' Try and get cell precedents. If the precedents are in another sheet we get stuck
              ' TODO: We could try this: http://www.mrexcel.com/forum/excel-questions/371863-test-if-cell-has-dependents-precedents.html#post1851288
              Dim CellPrecedents As Range
              On Error Resume Next
              Set CellPrecedents = cell.Precedents
              If Err.Number = 1004 Then
                  CellIsConstant = True
                  GoTo ExitFunction
              End If
              On Error GoTo ErrorHandler
              
              If cell.Precedents.Parent.Name = s.AdjustableCells.Parent.Name Then
4969              If Not Intersect(cell.Precedents, s.AdjustableCells) Is Nothing Then
4970                  CellIsConstant = False
4971                  GoTo ExitFunction
                  End If
4972          End If
4973      End If
4974      CellIsConstant = True

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CModelParsed", "CellIsConstant") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Checks whether a string formula depends on variable cells or is simply constant.
' If constant, we can evaluate it directly
Function StringFormulaIsConstant(strFormula As String) As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim tksFormula As Tokens
4975      Set tksFormula = ParseFormula(strFormula)
          
          Dim i As Long, c As Range, tkn As Token, CheckForColon As Boolean
          Dim LookingForEndOfRef As Boolean, strSheetPrefix As String, BuildingRef As String, rngBuildingRef As Range
          
          ' We loop through tokens looking for cell references.
          ' If we find any, we test them to see if they depend on the variable cells
4976      For i = 1 To tksFormula.Count
4977          Set tkn = tksFormula.Item(i)
              
4978          If CheckForColon = True And tkn.Text <> ":" Then
                  ' We don't hit colon right after the reference, we abandon
                  ' search for another reference
4979              LookingForEndOfRef = False
4980              CheckForColon = False
4981              If strSheetPrefix <> "" Then BuildingRef = strSheetPrefix + "!" + BuildingRef
4982              Set rngBuildingRef = s.sheet.Range(BuildingRef)
                  
                  ' Cell reference found, check if constant
4983              If rngBuildingRef.Count > 1 Then
4984                  For Each c In rngBuildingRef
4985                      If Not CellIsConstant(c) Then GoTo NotConstant
4986                  Next c
4987              Else
4988                  If Not CellIsConstant(rngBuildingRef) Then GoTo NotConstant
4989              End If
                  
                  ' Reset the sheet we are tracking
4990              strSheetPrefix = ""
4991          End If
          
              ' Decide what to insert based on token type
4992          Select Case tkn.TokenType
                  Case TokenType.ErrorText
                      ' TODO: Can't handle that, throw error
4993                  Err.Raise OpenSolver_BuildError, "Parsing model formulae", "An error token was found while parsing the formulae."
                      
4994              Case TokenType.Reference
                      ' Are we trying to complete a range?
4995                  If LookingForEndOfRef Then
                          ' We were, so finish building it...
4996                      BuildingRef = BuildingRef + tkn.Text
                          ' ... and stop trying to build it more
4997                      LookingForEndOfRef = False
4998                      CheckForColon = False
                          
                          ' Make sure this range is on the right sheet
4999                      If strSheetPrefix <> "" Then BuildingRef = strSheetPrefix + "!" + BuildingRef
                          
                          ' Cell reference found, check if constant
5000                      Set rngBuildingRef = s.sheet.Range(BuildingRef)
5001                      For Each c In rngBuildingRef
5002                          If Not Intersect(c.Precedents, s.AdjustableCells) Is Nothing Then GoTo NotConstant
5003                      Next c

                          ' Reset the sheet we are tracking
5004                      strSheetPrefix = ""
                          
5005                  Else
                          ' This is a new range, look for : and another cell ref
5006                      BuildingRef = tkn.Text
5007                      CheckForColon = True
5008                      LookingForEndOfRef = True
5009                  End If
                    
5010              Case TokenType.RangeOperator
                      ' Colon. The range we were building is about to be completed
                      ' and should come up after this
5011                  BuildingRef = BuildingRef + ":"
5012                  CheckForColon = False
                      
5013              Case TokenType.ReferenceQualifier
5014                  strSheetPrefix = "'" + tkn.Text + "'"
                      
5015              Case TokenType.ExternalReferenceOperator
                      ' We'll insert the ! ourselves manually later

5016          End Select
5017      Next i

          ' Are we still looking for another reference?
5018      If LookingForEndOfRef = True Then
              ' It never came - probably a solitary cell reference at the end
              ' of a formula. Check if we need a new node, add token to parsed string
5019          If strSheetPrefix <> "" Then BuildingRef = strSheetPrefix + "!" + BuildingRef
5020          Set rngBuildingRef = s.sheet.Range(BuildingRef)
              
              ' Cell reference found, check if constant
5021          If rngBuildingRef.Count > 1 Then
5022              For Each c In rngBuildingRef
5023                  If Not Intersect(c.Precedents, s.AdjustableCells) Is Nothing Then GoTo NotConstant
5024              Next c
5025          Else
5026              If Not Intersect(rngBuildingRef.Precedents, s.AdjustableCells) Is Nothing Then GoTo NotConstant
5027          End If
              
5028          strSheetPrefix = ""
5029      End If
5030      StringFormulaIsConstant = True
          
ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CModelParsed", "StringFormulaIsConstant") Then Resume
          RaiseError = True
          GoTo ExitFunction

NotConstant:
5032      StringFormulaIsConstant = False
          GoTo ExitFunction
End Function

Function ExpandFormula(ByVal strFormula As String, _
                       ByRef rngSourceCell As Range, Optional ConstraintNumber As Long) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          '==========================================================================
          ' STEP 0. Misc Setup
          ' The parsed string, this is what we store at the node
          Dim strParsed As String
          ' The current node we are considering
          Dim objCurNode As CFormula
          ' The sheet the next reference token is on
          Dim strSheetPrefix As String
          ' The sheet the adjustable cells are on
          Dim strAdjCellSheet As String
5034      strAdjCellSheet = s.AdjustableCells.Parent.Name
          ' Avoid unnecessary evaluations of AddNodeIfNew by checking if whole ranges
          ' are decision variables
          Dim rngBuildingRef As Range
          
          ' The following are used for converting SumProduct to AMPL
          Dim FunctionName As String
          Dim FunctionCount As Long
          Dim start As Long
          Dim output As String, node As String
          Dim Count As Long
5035      FunctionName = ""
5036      FunctionCount = 0
          
          '==========================================================================
          ' STEP 1. Setup root node for search
          Dim strRootAddress As String
5037      If Not (rngSourceCell Is Nothing) Then
5038          strRootAddress = ConvertCellToStandardName(rngSourceCell)
              ' If root node exists already, don't need to do this again
5039          If TestKeyExists(Formulae, strRootAddress) Then
5040              ExpandFormula = strRootAddress
5041              GoTo ExitFunction
5042          End If
              ' Check that the source cell (e.g. objective, LHS, RHS) is not an
              ' adjustable cell itself
5043          If rngSourceCell.Parent.Name = strAdjCellSheet Then
5044              If Not (Intersect(rngSourceCell, s.AdjustableCells) Is Nothing) Then
                      ' Its an adjustable cell, just return standard name and stop
5045                  ExpandFormula = strRootAddress
5046                  GoTo ExitFunction
5047              End If
5048          End If
5049      End If
          ' It doesn't exist, it isn't an adjustable cell
          ' Is it even a formula?
5050      If Len(strFormula) = 0 Then
5051          ExpandFormula = "0": GoTo ExitFunction
5052      End If
5053      If Left(strFormula, 1) <> "=" Then
5054          If IsAmericanNumber(strFormula) Then
5055              ExpandFormula = strFormula
5056          Else
5057              ExpandFormula = "'" + strFormula + "'"
5058          End If
5059          GoTo ExitFunction
5060      End If
          ' Store the last index of the list-representation of the DAG
          Dim lngBaseIndex As Long
5061      lngBaseIndex = Formulae.Count
          ' Create the root node and add it
5062      Set objCurNode = New CFormula
5063      If Not (rngSourceCell Is Nothing) Then
5064          objCurNode.strAddress = strRootAddress
5065          objCurNode.initialValue = CDbl(rngSourceCell.Value2)
5066          Set objCurNode.rngAddress = rngSourceCell
5067          ExpandFormula = strRootAddress
5068      Else
              ' Its a constant or formula - return the parsed formula directly
5069          objCurNode.strAddress = "_" + CStr(Formulae.Count)
5070          Set objCurNode.rngAddress = Nothing
5071      End If
5072      objCurNode.strFormula = strFormula  ' Could also get from rngSourceCell
5073      objCurNode.lngDepth = 0             ' Root node depth is 0
5074      objCurNode.boolIsRoot = True
5075      Formulae.Add objCurNode, objCurNode.strAddress ' Add it to the DAG
          ' Start at the root node
          Dim lngIndex As Long
5076      lngIndex = lngBaseIndex + 1
          
          ' tracking variables
          Dim TotalCount As Long, curProgress As Long
          
          '==========================================================================
          ' STEP 2. Keep processing nodes until nothing interesting left
5077      Do While lngIndex <= Formulae.Count
5078          UpdateStatusBar "OpenSolver: Constraint progress: " & ConstraintNumber & "/" & s.NumConstraints & ". " & _
                              "Processing formula " & lngIndex & "/" & Formulae.Count & ". "
              
5082          Set objCurNode = Formulae(lngIndex)
5083          strParsed = ""
5084          strSheetPrefix = ""
          
              '======================================================================
              ' STEP 2.C. Tokenise if we can't evaluate
5085          If objCurNode.boolCanEval Then GoTo skipTokenising
              
              ' Tokenize the formula
              Dim tksFormula As Tokens
5086          Set tksFormula = ParseFormula(objCurNode.strFormula)
          
              ' If we find a reference to a cell name, it could be the beginning
              ' of a multi-cell range. So always assume thats going to happen, and
              ' look for another reference (LookingForEndOfRef) after a colon :
              ' (CheckForColon). If a colon isn't there, stop looking.
              Dim LookingForEndOfRef As Boolean, CheckForColon As Boolean
5087          LookingForEndOfRef = False
5088          CheckForColon = False
              
              ' The candidate multicell range we are building
              Dim BuildingRef As String
5089          BuildingRef = ""
                      
              ' Take a walk through the tokens
              Dim i As Long, c As Range, tkn As Token
5090          For i = 1 To tksFormula.Count
5091              Set tkn = tksFormula.Item(i)
                  
                  ' CheckForColon means we hit a cell reference, and want to see if
                  ' its actually referring to multiple cells
5092              If CheckForColon = True And tkn.Text <> ":" Then
                      ' We don't hit colon right after the reference, we abandon
                      ' search for another reference
5093                  LookingForEndOfRef = False
5094                  CheckForColon = False
5095                  If strSheetPrefix <> "" Then BuildingRef = strSheetPrefix + "!" + BuildingRef
5096                  Set rngBuildingRef = s.sheet.Range(BuildingRef)
5097                  If rngBuildingRef.Count > 1 Then
5098                      strParsed = strParsed + "["
5099                      For Each c In rngBuildingRef
5100                          strParsed = strParsed + AddNodeIfNew(objCurNode, c) + ","
5101                      Next c
5102                      strParsed = Left(strParsed, Len(strParsed) - 1) + "]"
5103                  Else
5104                      strParsed = strParsed + AddNodeIfNew(objCurNode, rngBuildingRef)
5105                  End If
                      ' Reset the sheet we are tracking
5106                  strSheetPrefix = ""
5107              End If
          
                  ' Decide what to insert based on token type
5108              Select Case tkn.TokenType
                      Case TokenType.Text
                          ' Output with quotes, tokenizer turns "=""test""" -> test
5109                      strParsed = strParsed + "'" + tkn.Text + "'"
                              
5110                  Case TokenType.Number
                          ' TODO: Scientific notation for small/large numbers?
5111                      strParsed = strParsed + StrExNoPlus(Val(tkn.Text))
                          
5112                  Case TokenType.Bool
5113                      strParsed = strParsed + IIf(tkn.Text = "TRUE", "True", "False")
                          
5114                  Case TokenType.ErrorText
                          ' TODO: Can't handle that, throw error
5115                      Err.Raise OpenSolver_BuildError, "Parsing model formulae", "An error token was found while parsing the formulae."
                          
5116                  Case TokenType.Reference
                          ' Are we trying to complete a range?
5117                      If LookingForEndOfRef Then
                              ' We were, so finish building it...
5118                          BuildingRef = BuildingRef + tkn.Text
                              ' ... and stop trying to build it more
5119                          LookingForEndOfRef = False
5120                          CheckForColon = False
                              
                              ' Make sure this range is on the right sheet
5121                          If strSheetPrefix <> "" Then BuildingRef = strSheetPrefix + "!" + BuildingRef
                              
5122                          Set rngBuildingRef = s.sheet.Range(BuildingRef)
                              
                              'Replace preceding comma with + if summing a second argument
5123                          If FunctionName = "SUM" And FunctionCount > 0 Then
5124                              strParsed = Left(strParsed, Len(strParsed) - 1) + "+"
5125                          End If
                              
                              'Replace preceding comma with + if summing a second argument
                              If FunctionName = "PRODUCT" And FunctionCount > 0 Then
                                  strParsed = Left(strParsed, Len(strParsed) - 1) + "*"
                              End If

                              'If we are starting a SUMIF, we need to get the next argument now
                              If FunctionName = "SUMIF" And FunctionCount = 0 Then
                                  Dim NewToken As Token, IfOperator As String
                                  Set NewToken = tksFormula.Item(i + 1)
                                  If NewToken.Text <> "," Then
                                      MsgBox "Misformed SUMIF"
                                  End If
                                  
                                  Set NewToken = tksFormula.Item(i + 2)
                                  IfOperator = "="
                                  If NewToken.TokenType = Text Then
                                      Select Case Left(NewToken.Text, 2)
                                          Case "<=", ">=", "<>"
                                              IfOperator = ""
                                      End Select
                                      Select Case Left(NewToken.Text, 1)
                                          Case "<", ">", "="
                                              IfOperator = ""
                                      End Select
                                  End If
                                  If NewToken.TokenType = TokenType.Reference Then
                                      IfOperator = IfOperator + AddNodeIfNew(objCurNode, s.sheet.Range(NewToken.Text))
                                  Else
                                      IfOperator = IfOperator + CStr(NewToken.Text)
                                  End If
                                  i = i + 2
                              End If
                              
                              ' If we are building up a function, save the start of the function so we don't overwrite earlier work
5126                          If FunctionCount = 0 Then
5127                              start = Len(strParsed)
5128                          End If
                              
5129                          Count = 0
5130                          TotalCount = rngBuildingRef.Count
5131                          curProgress = 0
5132                          For Each c In rngBuildingRef
                                  ' Update status
5133                              curProgress = curProgress + 1
5134                              UpdateStatusBar "OpenSolver: Constraint progress: " & ConstraintNumber & "/" & s.NumConstraints & ". " & _
                                                  "Processing formula " & lngIndex & "/" & Formulae.Count & ". " & _
                                                  "Current formula progress: " & curProgress & "/" & TotalCount
                                  
5138                              node = AddNodeIfNew(objCurNode, c)
                                  
5139                              If FunctionName = "SUMPRODUCT" And Not FunctionCount = 0 Then
5140                                  Count = Count + 1
5141                                  If FunctionCount > 1 And Count < rngBuildingRef.Count Then
5142                                      strParsed = Replace(strParsed, "+", "*" & node & "+", FindPosition(strParsed, "+", Count), 1)
5143                                  Else
5144                                      strParsed = Replace(strParsed, ",", "*" & node & "+", 1, 1)
5145                                  End If
                                  ElseIf FunctionName = "SUMIF" Then
                                      If FunctionCount = 0 Then
                                          strParsed = strParsed + "if(" + node + IfOperator + ","
                                      Else
                                          Dim Position As Long
                                          Count = Count + 1
                                          Position = FindPosition(strParsed, ",", 2 * Count, start)
                                          strParsed = Left(strParsed, Position) + node + ",0)+" + Mid(strParsed, Position + 1)
                                      End If
5146                              ElseIf FunctionName = "SUM" Then
5147                                  strParsed = strParsed + node + "+"
                                  ElseIf FunctionName = "PRODUCT" Then
                                      strParsed = strParsed + node + "*"
5148                              Else
                                      ' Check if we need a new node, add token to parsed string
5149                                  strParsed = strParsed + node + ","
5150                              End If
5151                          Next c
5152                          FunctionCount = FunctionCount + 1
                              ' Reset the sheet we are tracking
5153                          strSheetPrefix = ""
                              ' Kill the extra char added by last cell in range
5154                          strParsed = Left(strParsed, Len(strParsed) - 1)
                              
5155                      Else
                              ' This is a new range, look for : and another cell ref
5156                          BuildingRef = tkn.Text
5157                          CheckForColon = True
5158                          LookingForEndOfRef = True
5159                      End If
                          
5160                  Case TokenType.whitespace
                          ' Do nothing
                          
5161                  Case TokenType.UnaryOperator
5162                      If tkn.Text = "+" Then
5163                          strParsed = strParsed + "+"
5164                      ElseIf tkn.Text = "-" Then
5165                          strParsed = strParsed + "-"
5166                      Else
5167                          MsgBox "Unary Operator that isn't + or -! Develop code to handle this: " + tkn.Text
5168                      End If
                          
5169                  Case TokenType.ArithmeticOperator
5170                      strParsed = strParsed + tkn.Text
                          
5171                  Case TokenType.ComparisonOperator
5172                      If tkn.Text = "=" Then
5173                          strParsed = strParsed + "=="
5174                      Else 'TODO (maybe): <>?
5175                          strParsed = strParsed + tkn.Text
5176                      End If
                          
5177                  Case TokenType.TextOperator
                          ' Text concatenation &
5178                      strParsed = strParsed + " + "
                          
5179                  Case TokenType.RangeOperator
                          ' Colon. The range we were building is about to be completed
                          ' and should come up after this
5180                      BuildingRef = BuildingRef + ":"
5181                      CheckForColon = False
                          
5182                  Case TokenType.ReferenceQualifier
5183                      strSheetPrefix = "'" + tkn.Text + "'"
                          
5184                  Case TokenType.ExternalReferenceOperator
                          ' We'll insert the ! ourselves manually later
                          
5185                  Case TokenType.PostfixOperator
                          ' Only percentage sign it seems
5186                      strParsed = strParsed + "/100.0"
                          
5187                  Case TokenType.FunctionOpen
5188                      FunctionName = tkn.Text
5189                      output = tkn.Text
5190                      Select Case output
                          Case "SUMPRODUCT"
5191                          output = "("
5192                      Case "SUM", "PRODUCT"
5193                          output = "("
5194                      Case Else
                              ' Delegate to solver specific conversion code
5195                          output = ConvertFormula(output)
5196                      End Select
5197                      strParsed = strParsed + output
                          
5198                  Case TokenType.ParameterSeparator
5199                      strParsed = strParsed + tkn.Text
                          
5200                  Case TokenType.FunctionClose
5201                      strParsed = strParsed + ")"
5202                      FunctionCount = 0
                          
5203                  Case TokenType.SubExpressionOpen
5204                      strParsed = strParsed + tkn.Text
                      
5205                  Case TokenType.SubExpressionClose
5206                      strParsed = strParsed + tkn.Text
                          
5207                  Case Else
                          ' TODO: table things
5208              End Select
5209          Next i

              ' Are we still looking for another reference?
5210          If LookingForEndOfRef = True Then
                  ' It never came - probably a solitary cell reference at the end
                  ' of a formula. Check if we need a new node, add token to parsed string
5211              If strSheetPrefix <> "" Then BuildingRef = strSheetPrefix + "!" + BuildingRef
5212              Set rngBuildingRef = s.sheet.Range(BuildingRef)
5213              If rngBuildingRef.Count > 1 Then
5214                  strParsed = strParsed + "["
5215                  For Each c In rngBuildingRef
5216                      strParsed = strParsed + AddNodeIfNew(objCurNode, c) + ","
5217                  Next c
5218                  strParsed = Left(strParsed, Len(strParsed) - 1) + "]"
5219              Else
5220                  strParsed = strParsed + AddNodeIfNew(objCurNode, rngBuildingRef)
5221              End If
5222              strSheetPrefix = ""
5223          End If
              
              
              '======================================================================
              ' STEP 2.C. Store parsed formula in the DAG
skipTokenising:
5224          If objCurNode.boolIsConstant = False Then objCurNode.strFormulaParsed = strParsed
5225          If Left(objCurNode.strAddress, 1) = "_" Then
                  ' Root node is actually a formula or constant, so return the parsed formula
5226              ExpandFormula = strParsed
5227          End If
5228          If objCurNode.boolIsRoot And objCurNode.boolIsConstant Then
                  ' Was just a simple constant, no need to create anything in Formulae
5229              ExpandFormula = objCurNode.strFormulaParsed
5230              Formulae.Remove lngIndex
5231              GoTo ExitFunction
5232          End If
              
              '======================================================================
              ' On to the next node
5233          lngIndex = lngIndex + 1
5234      Loop
          
ExitFunction:
          Application.StatusBar = False
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CModelParsed", "ExpandFormula") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function FindPosition(Word As String, FindString As String, Count As Long, Optional start As Long = 1) As Long
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim It As Long
5250      It = 1
          
5251      FindPosition = start
          
          ' Loop until count reached
5252      While It < Count
5253          FindPosition = InStr(FindPosition + 1, Word, FindString)
5254          It = It + 1
5255      Wend

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CModelParsed", "FindPosition") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Sub CleanFormulae(ByVal lngFStart As Long, ByVal lngFEnd As Long, Optional ConstraintNumber As Long = 0)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim objCurNode As CFormula, lngIndex As Long, varReturn As Variant

          ' Any AdjCellUnknown can now be decided one way or the other
          Dim lngCurDepth As Long, strChildAddr As Variant
5257      For lngCurDepth = lngMaxDepth To 0 Step -1
5258          For lngIndex = lngFStart To lngFEnd
5259              UpdateStatusBar "OpenSolver: Constraint progress: " & ConstraintNumber & "/" & s.NumConstraints & ". " & _
                                  "Cleaning decision variables in formula " & lngIndex & "/" & Formulae.Count & ". "
                  
5263              Set objCurNode = Formulae(lngIndex)
5264              If objCurNode.lngDepth = lngCurDepth Then
5265                  If objCurNode.intRefsAdjCell = AdjCellUnknown And Not objCurNode.boolIsRoot Then
                          ' Look through children - if any are adj cell depedendant,
                          ' this is too. Otherwise, its not!
                          
5266                      objCurNode.intRefsAdjCell = AdjCellIndependent
5267                      For Each strChildAddr In objCurNode.Children
5268                          If Formulae(strChildAddr).intRefsAdjCell = AdjCellDependent Then
5269                              objCurNode.intRefsAdjCell = AdjCellDependent
5270                          End If
5271                      Next
                          ' If its AdjCellIndependent, eval it!
5272                      If objCurNode.intRefsAdjCell = AdjCellIndependent Then
5273                          objCurNode.boolCanEval = True
                              ' NOTE: Application.Evaluate has a 255 character limit
                              ' See also: http://dutchgemini.wordpress.com/2009/08/07/error-2015-using-application-evaluate-in-excel-vba/
5274                          varReturn = s.sheet.Evaluate(objCurNode.strFormula)
5275                          If (VBA.VarType(varReturn) = vbError) Then
                                  ' Fall back to calculating the cell
5276                              objCurNode.rngAddress.Calculate
5277                              objCurNode.strFormulaParsed = CStr(objCurNode.rngAddress.value)
5278                              objCurNode.boolEvaledWithCalculate = True
5279                          Else
5280                              objCurNode.strFormulaParsed = varReturn
5281                          End If
5282                          objCurNode.boolIsConstant = True
5283                      End If
5284                  End If
5285              End If
5286          Next lngIndex
5287      Next lngCurDepth
          
          ' Pull up constants where possible
          Dim strParentAdr As Variant
5288      For lngCurDepth = lngMaxDepth To 0 Step -1
5289          For lngIndex = lngFStart To lngFEnd
5290              UpdateStatusBar "OpenSolver: Constraint progress: " & ConstraintNumber & "/" & s.NumConstraints & ". " & _
                                  "Cleaning constants in formula " & lngIndex & "/" & Formulae.Count & ". "
                  
5294              If Formulae(lngIndex).lngDepth = lngCurDepth Then
5295                  If Formulae(lngIndex).boolIsConstant Then
                          ' Constant - so can pull up into its parents
5296                      Set objCurNode = Formulae(lngIndex)
                          
5297                      For Each strParentAdr In objCurNode.Parents
5298                          With Formulae(strParentAdr)
5299                              .strFormulaParsed = Replace(.strFormulaParsed, objCurNode.strAddress, objCurNode.strFormulaParsed)
5300                          End With
5301                      Next
                          
5302                      objCurNode.boolCanBeRemoved = Not objCurNode.boolIsRoot
5303                  End If
5304              End If
5305          Next lngIndex
5306      Next lngCurDepth
          
          ' Remove any unneeded nodes
5307      For lngIndex = lngFEnd To lngFStart Step -1
5308          UpdateStatusBar "OpenSolver: Constraint progress: " & ConstraintNumber & "/" & s.NumConstraints & ". " & _
                              "Removing unnecessary formulae: " & lngIndex & "/" & Formulae.Count & ". "
          
5312          If Formulae(lngIndex).boolCanBeRemoved Then Formulae.Remove lngIndex
5313      Next lngIndex

ExitSub:
          Application.StatusBar = False
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("CModelParsed", "CleanFormulae") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Sub PUSHDOWN(ByRef objNode As CFormula, ByVal lngNewDepth As Long)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

5329      If lngNewDepth > lngMaxDepth Then lngMaxDepth = lngNewDepth
5330      If lngNewDepth > objNode.lngDepth Then objNode.lngDepth = lngNewDepth
          'Dim child As Long
          'For child = 1 To objNode.Children.Count
          '    PUSHDOWN Formulae(objNode.Children(child)), objNode.lngDepth + 1, Formulae, lngMaxDepth
          'Next
          Dim child As Variant
5331      For Each child In objNode.Children
5332          If TestKeyExists(Formulae, CStr(child)) Then PUSHDOWN Formulae(child), objNode.lngDepth + 1
5333      Next

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("CModelParsed", "PUSHDOWN") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Function AddNodeIfNew(ByRef objCurNode As CFormula, ByRef c As Range) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim cell As Range
          
          ' 1.1 Get the standard name of this cell
          Dim strStdName As String
5335      strStdName = ConvertCellToStandardName(c)
          
          ' Check if cell is an adjustable cell
5336      If c.Parent.Name = s.sheet.Name Then
5337          If Not Intersect(c, s.AdjustableCells) Is Nothing Then
5338              AddNodeIfNew = strStdName
5339              GoTo finishedNode
5340          End If
5341      End If
          '==========================================================================
          ' 0. Is Simple Cell?
          ' 0.1 Get the standard name of this cell
          Dim strCFormula As String
5342      strCFormula = c.Formula
          ' 0.2 Is blank cell (TODO: breaks some IF statements)
5343      If Len(strCFormula) = 0 Then
5344          AddNodeIfNew = "0"
5345          GoTo finishedNode
5346      End If
          ' 0.3 Is constant?
5347      If Left(strCFormula, 1) <> "=" Then
              ' Number?
5348          If IsAmericanNumber(strCFormula) Then
5349              AddNodeIfNew = strCFormula
5350          Else
                  ' Not a number, assume text
5351              AddNodeIfNew = """" + strCFormula + """"
5352          End If
5353          GoTo finishedNode
5354      End If
                  
          '==========================================================================
          ' 1. Is Existing Node?
          ' 1.2 Does the node exist yet?
5355      If TestKeyExists(Formulae, strStdName) Then
              ' Node already exists
              Dim objExistNode As CFormula
5356          Set objExistNode = Formulae(strStdName)
              ' Has it been evaluated?
5357          If objExistNode.boolIsConstant Then
                  ' Its a constant, so we can just fold it straight in
5358              AddNodeIfNew = objExistNode.strFormulaParsed
5359              GoTo finishedNode
5360          End If
              ' Is it dependent on a decision variable?
5361          If objExistNode.intRefsAdjCell = AdjCellDependent Then
                  ' That means this cell is also dependent on adj. cells, which
                  ' we may not have known
5362              objCurNode.intRefsAdjCell = AdjCellDependent
5363          End If
              ' Add this node to the current nodes Children list
5364          If Not TestKeyExists(objCurNode.Children, strStdName) Then
5365              objCurNode.Children.Add strStdName, strStdName
5366          End If
              ' Tell it about its new parent
5367          If Not TestKeyExists(objExistNode.Parents, objCurNode.strAddress) Then
5368              objExistNode.Parents.Add objCurNode.strAddress, objCurNode.strAddress
5369          End If
              ' Ensure it is at the correct depth
5370          PUSHDOWN objExistNode, objCurNode.lngDepth + 1
              ' Return the standardised name of this existing node
5371          AddNodeIfNew = strStdName
5372          GoTo finishedNode
5373      End If
          
          '==========================================================================
          ' 2. Is Adjustable Cell?
          Dim varReturn As Variant
          Dim objNewNode As New CFormula
5374      Set objNewNode.rngAddress = c
5375      objNewNode.strAddress = strStdName
          ' Is on same sheet?
5376      If objNewNode.GetSheet = s.AdjustableCells.Parent.Name Then
              ' This formula is on the same sheet as the adjacent cells
              ' Thus we can safely check if this formula's cell is depedent on
              ' a decision variable
5377          If objNewNode.IsDependentOn(s.AdjustableCells.Dependents) Then
                  ' This formula depends on the value of the adjustable cells
                  ' This means we can NOT evaluate it
5378              objNewNode.boolCanEval = False
5379              objNewNode.boolIsConstant = False
5380              objNewNode.intRefsAdjCell = AdjCellDependent
5381          Else
                  ' This formula is NOT dependent on an adjustable cell, so we
                  ' can just evaluate its formula (TODO: or take its value?)
                  ' We know its not a simple constant, or we would of got
                  ' it earlier
                  ' NOTE: Application.Evaluate has a 255 character limit
                  
5382              varReturn = s.sheet.Evaluate(strCFormula)
5383              If (VBA.VarType(varReturn) = vbError) Then
                      ' Fall back to calculating the cell
5384                  c.Calculate
5385                  AddNodeIfNew = StrExNoPlus(c.value)
5386                  GoTo finishedNode
                  ElseIf IsArray(varReturn) Then
                      AddNodeIfNew = StrExNoPlus(CDbl(varReturn(1)))
                      GoTo finishedNode
5387              Else
                      On Error Resume Next
5388                  AddNodeIfNew = StrExNoPlus(CDbl(varReturn))
                      If Err.Number <> 0 Then
                          On Error GoTo ErrorHandler
                          Err.Raise OpenSolver_ModelError, Description:= _
                              "The constant in cell " & c.Address & " has value `" & varReturn & "` which is not numeric. " & _
                              "Make sure this constant isn't being included in a formula unnecessarily."
                      End If
                      On Error GoTo ErrorHandler
5389                  GoTo finishedNode
5390              End If
5391          End If
5392      Else
              ' This formula is on a different sheet to the adjacent cells.
              ' This means we can't check if its depedent on an adjacent cell,
              ' because Range. Dependents only returns the cells for the same
              ' sheet as Range.Parent
              'objNewNode.boolCanEval = False
              'objNewNode.boolIsConstant = False
              'objNewNode.intRefsAdjCell = AdjCellUnknown
5393          varReturn = s.sheet.Evaluate(strCFormula)
5394              If (VBA.VarType(varReturn) = vbError) Then
                      ' Fall back to calculating the cell
5395                  c.Calculate
5396                  AddNodeIfNew = StrExNoPlus(c.value)
5397                  GoTo finishedNode
5398              Else
5399                  AddNodeIfNew = StrExNoPlus(CDbl(varReturn))
5400                  GoTo finishedNode
5401              End If
5402      End If
          
          '==========================================================================
          ' 3. Create new node
5403      With objNewNode
5404          .strFormula = strCFormula
              If c.Value2 = "" Then
                  .initialValue = 0
              Else
5405              .initialValue = CDbl(c.Value2)
              End If
5406          .lngDepth = objCurNode.lngDepth + 1
5407          If .lngDepth > lngMaxDepth Then lngMaxDepth = .lngDepth
5408          .Parents.Add objCurNode.strAddress, objCurNode.strAddress
5409      End With
          ' Add this new node to the child list of current node
5410      objCurNode.Children.Add strStdName, strStdName
          ' Add new node to the DAG
5411      Formulae.Add objNewNode, strStdName
          ' Return the standardised name of this new node
5412      AddNodeIfNew = strStdName

finishedNode:
          ' Done!
ExitFunction:
          Application.StatusBar = False
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CModelParsed", "AddNodeIfNew") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function ConvertFormula(tokenText As String) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim FileSolver As ISolverFile
          Set FileSolver = s.Solver

5428      Select Case FileSolver.FileType
          Case AMPL
5431          ConvertFormula = ConvertFormula_AMPL(tokenText)
5432      Case NL
5433          ConvertFormula = ConvertFormula_NL(tokenText)
5434      Case Else
5435          ConvertFormula = tokenText + "("
5436      End Select

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CModelParsed", "ConvertFormula") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function ConvertFormula_AMPL(tokenText As String) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

5437      tokenText = LCase(tokenText)
5438      Select Case tokenText
          ' RADIANS
          Case "radians"
5439          ConvertFormula_AMPL = "pi/180*("
5440      Case "pi"
5441          ConvertFormula_AMPL = "(pi"
          Case "sumif"
              Err.Raise OpenSolver_BuildError, "Parsing model formulae", "The operator " & UCase(tokenText) & " is not yet supported by OpenSolver when using the NEOS solvers."
              ConvertFormula_AMPL = tokenText + "("
5447      Case Else
5448          ConvertFormula_AMPL = tokenText + "("
5449      End Select

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CModelParsed", "ConvertFormula_AMPL") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function ConvertFormula_NL(tokenText As String) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

5453      Select Case tokenText
          ' RADIANS
          Case "RADIANS"
5454          ConvertFormula_NL = "4*ATAN(1)/180*("
          Case "SUMIF"
              ConvertFormula_NL = "("
5455      Case "PI"
5456          ConvertFormula_NL = "(4*ATAN(1)"
5462      Case Else
5463          ConvertFormula_NL = tokenText + "("
5464      End Select

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CModelParsed", "ConvertFormula_NL") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function


