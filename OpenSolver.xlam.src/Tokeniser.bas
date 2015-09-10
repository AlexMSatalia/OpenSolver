Attribute VB_Name = "Tokeniser"
'==============================================================================
' OpenSolver
' Formula tokenizer functionality is from http://www.dailydoseofexcel.com
' Code is written by Rob van Gelder
' http://www.dailydoseofexcel.com/archives/2009/12/05/formula-tokenizer/
' This file is unmodified.
'==============================================================================

Option Explicit

Public Enum ParsingState
    ParsingError
    Expression1
    Expression2
    LeadingName1
    LeadingName2
    LeadingName3
    LeadingNameE
    whitespace
    Text1
    Text2
    Number1
    Number2
    Number3
    Number4
    NumberE
    Bool
    ErrorX
    MinusSign               'for ambiguity between unary minus and sign
    PrefixOperator
    ArithmeticOperator
    ComparisonOperator1
    ComparisonOperator2
    ComparisonOperatorE
    TextOperator
    PostfixOperator
    RangeOperator
    ReferenceQualifier
    ListSeparator           'for ambiguity between function parameter separator and reference union operator
    ArrayRowSeparator
    ArrayColumnSeparator
    FunctionX
    ParameterSeparator
    SubExpression
    BracketClose
    ArrayOpen
    ArrayConstant1
    ArrayConstant2
    ArrayClose
    SquareBracketOpen
    Table1
    Table2
    Table3
    Table4
    Table5
    Table6
    TableQ
    TableE
    R1C1Reference           'todo: implement R1C1Reference
End Enum

Public Enum TokenType
    Text
    Number
    Bool
    ErrorText
    Reference
    whitespace
    UnaryOperator
    ArithmeticOperator
    ComparisonOperator
    TextOperator
    RangeOperator
    ReferenceQualifier
    ExternalReferenceOperator
    PostfixOperator
    FunctionOpen
    ParameterSeparator
    FunctionClose
    SubExpressionOpen
    SubExpressionClose
    ArrayOpen
    ArrayRowSeparator
    ArrayColumnSeparator
    ArrayClose
    TableOpen
    TableSection
    TableColumn
    TableItemSeparator
    TableColumnSeparator
    TableClose
End Enum

Public Function ParseFormula(strFormula As String) As Tokens
          Dim i As Long, str As String, str2 As String, c As String, bln As Boolean, lng As Long, strError As String
          Dim varState As ParsingState, varReturnState As ParsingState, blnLoop As Boolean, lngFormulaLen As Long
          Dim lngTokenIndex As Long, lngPrevIndex As Long, lngStuckCount As Long

          Dim strDecimalSeparator As String, strListSeparator As String
          Dim strArrayRowSeparator As String, strArrayColumnSeparator As String
          Dim strLeftBrace As String, strRightBrace As String
          Dim strLeftBracket As String, strRightBracket As String
          Dim strLeftRoundBracket As String, strRightRoundBracket As String
          Dim strBooleanTrue As String, strBooleanFalse As String
          Dim strDoubleQuote As String

          Dim strErrorRef As String, strErrorDiv0 As String, strErrorNA As String, strErrorName As String
          Dim strErrorNull As String, strErrorNum As String, strErrorValue As String, strErrorGettingData As String

          Dim objToken As Token, objTokens As Tokens, objTokenStack As TokenStack

5490      Set objTokens = New Tokens
5491      Set objTokenStack = New TokenStack

5492      strDecimalSeparator = "." ' Application.International(xlDecimalSeparator)
5493      strListSeparator = "," ' Application.International(xlListSeparator)
5494      strArrayRowSeparator = ";" ' Application.International(xlRowSeparator)
'5495      If Application.International(xlColumnSeparator) = Application.International(xlDecimalSeparator) Then
'5496          strArrayColumnSeparator = Application.International(xlAlternateArraySeparator)
'5497      Else
5498          strArrayColumnSeparator = "," ' Application.International(xlColumnSeparator)
'5499      End If
5500      strLeftBrace = "{" ' Application.International(xlLeftBrace)
5501      strRightBrace = "}" ' Application.International(xlRightBrace)
5502      strLeftBracket = "[" ' Application.International(xlLeftBracket)       'todo: implement ' check if this applies to tableopen and tableclose symbols
5503      strRightBracket = "]" ' Application.International(xlRightBracket)     'todo: implement
5504      strLeftRoundBracket = "("
5505      strRightRoundBracket = ")"
5506      strBooleanTrue = "TRUE"
5507      strBooleanFalse = "FALSE"
5508      strDoubleQuote = """"

5509      strErrorRef = "#REF!"
5510      strErrorDiv0 = "#DIV/0!"
5511      strErrorNA = "#N/A"
5512      strErrorName = "#NAME?"
5513      strErrorNull = "#NULL!"
5514      strErrorNum = "#NUM!"
5515      strErrorValue = "#VALUE!"
5516      strErrorGettingData = "#GETTING_DATA"

5517      lngFormulaLen = Len(strFormula)

5518      If lngFormulaLen <= 1 Then GoTo e
5519      If Left(strFormula, 1) <> "=" Then GoTo e

5520      varState = ParsingState.Expression1
5521      i = 2
5522      lngPrevIndex = 1

5523      blnLoop = True

5524      Do
5525          If i <= lngFormulaLen Then c = Mid(strFormula, i, 1) Else c = ""

      ''' -------------------- -------------------- -------------------- '''

5526          If varState = ParsingState.ParsingError Then
5527              MsgBox strError, vbCritical, "Error"
5528              blnLoop = False
      ''' -------------------- -------------------- -------------------- '''

5529          ElseIf varState = ParsingState.Expression1 Then
5530              If c = " " Or c = Chr(10) Then
5531                  varReturnState = varState
5532                  varState = ParsingState.whitespace
5533                  lngTokenIndex = i

5534              ElseIf c = strDoubleQuote Then
5535                  varReturnState = ParsingState.Expression2
5536                  varState = ParsingState.Text1
5537                  lngTokenIndex = i
5538                  i = i + 1

5539              ElseIf c = "0" Or c = "1" Or c = "2" Or c = "3" Or c = "4" Or _
                         c = "5" Or c = "6" Or c = "7" Or c = "8" Or c = "9" Then
5540                  varReturnState = ParsingState.Expression2
5541                  varState = ParsingState.Number1
5542                  lngTokenIndex = i

5543              ElseIf c = "-" Then
5544                  str = c
5545                  varReturnState = ParsingState.Expression2
5546                  varState = ParsingState.MinusSign
5547                  lngTokenIndex = i
5548                  i = i + 1

5549              ElseIf c = "+" Then
5550                  varState = ParsingState.PrefixOperator
5551                  lngTokenIndex = i

5552              ElseIf c = strLeftRoundBracket Then
5553                  varState = ParsingState.SubExpression

5554              ElseIf c = strLeftBrace Then
5555                  varState = ParsingState.ArrayOpen

5556              ElseIf c = "'" Then
5557                  varState = ParsingState.LeadingName2
5558                  lngTokenIndex = i
5559                  i = i + 1

5560              ElseIf c = "#" Then
5561                  varReturnState = ParsingState.Expression2
5562                  varState = ParsingState.ErrorX
5563                  lngTokenIndex = i

5564              Else
5565                  varState = ParsingState.LeadingName1
5566                  lngTokenIndex = i

5567              End If

      ''' -------------------- -------------------- -------------------- '''

5568          ElseIf varState = ParsingState.Expression2 Then
5569              If c = "" Then
5570                  blnLoop = False

5571              ElseIf c = " " Or c = Chr(10) Then
5572                  varReturnState = varState
5573                  varState = ParsingState.whitespace
5574                  lngTokenIndex = i

5575              ElseIf c = strRightRoundBracket Then
5576                  varState = ParsingState.BracketClose

5577              ElseIf c = "+" Or c = "-" Or c = "*" Or c = "/" Or c = "^" Then
5578                  varState = ParsingState.ArithmeticOperator
5579                  lngTokenIndex = i

5580              ElseIf c = "=" Or c = "<" Or c = ">" Then
5581                  varState = ParsingState.ComparisonOperator1
5582                  lngTokenIndex = i

5583              ElseIf c = "&" Then
5584                  varState = ParsingState.TextOperator
5585                  lngTokenIndex = i

5586              ElseIf c = "%" Then
5587                  varState = ParsingState.PostfixOperator
5588                  lngTokenIndex = i

5589              ElseIf c = ":" Then
5590                  varState = ParsingState.RangeOperator
5591                  lngTokenIndex = i

5592              ElseIf c = strListSeparator Then
5593                  varState = ParsingState.ListSeparator
5594                  lngTokenIndex = i

5595              Else
                      'Check if whitespace is actually union operator
5596                  Set objToken = objTokens(objTokens.Count)
5597                  bln = False
5598                  If objToken.TokenType = TokenType.whitespace Then
5599                      lng = InStr(1, objToken.Text, " ")
5600                      If lng > 0 Then
5601                          str = Mid(objToken.Text, lng + 1)
5602                          If lng = 1 Then
5603                              objToken.TokenType = TokenType.RangeOperator
5604                              objToken.Text = " "
5605                              objToken.FormulaLength = 1
5606                          Else
5607                              objToken.Text = Left(objToken.Text, lng - 1)
5608                              objToken.FormulaLength = lng - 1
5609                              objTokens.Add objTokens.NewToken(" ", TokenType.RangeOperator, objToken.FormulaIndex + lng - 1, 1)
5610                          End If

5611                          If str <> "" Then
5612                              objTokens.Add objTokens.NewToken(str, TokenType.whitespace, objToken.FormulaIndex + lng, Len(str))
5613                              str = ""
5614                          End If

5615                          varState = ParsingState.Expression1
5616                          bln = True
5617                      End If

5618                  ElseIf objToken.TokenType = TokenType.ErrorText And objToken.Text = strErrorRef Then
5619                      varState = ParsingState.Expression1
5620                      bln = True
5621                  End If
5622                  If Not bln Then
5623                      strError = "Expected Operator, but got " & c & " at position " & i
5624                      varState = ParsingState.ParsingError
5625                  End If
5626              End If

      ''' -------------------- -------------------- -------------------- '''

5627          ElseIf varState = ParsingState.ArrayConstant1 Then
5628              If c = strDoubleQuote Then
5629                  varReturnState = ParsingState.ArrayConstant2
5630                  varState = ParsingState.Text1
5631                  lngTokenIndex = i
5632                  i = i + 1

5633              ElseIf c = "0" Or c = "1" Or c = "2" Or c = "3" Or c = "4" Or _
                         c = "5" Or c = "6" Or c = "7" Or c = "8" Or c = "9" Then
5634                  varReturnState = ParsingState.ArrayConstant2
5635                  varState = ParsingState.Number1
5636                  lngTokenIndex = i

5637              ElseIf c = "-" Then
5638                  str = c
5639                  varReturnState = ParsingState.ArrayConstant2
5640                  varState = ParsingState.Number1
5641                  lngTokenIndex = i
5642                  i = i + 1

5643              ElseIf c = "#" Then
5644                  varReturnState = ParsingState.ArrayConstant2
5645                  varState = ParsingState.ErrorX
5646                  lngTokenIndex = i

5647              Else
5648                  varReturnState = ParsingState.ArrayConstant2
5649                  varState = ParsingState.Bool
5650                  lngTokenIndex = i

5651              End If

      ''' -------------------- -------------------- -------------------- '''

5652          ElseIf varState = ParsingState.ArrayConstant2 Then
5653              If c = strRightBrace Then
5654                  varState = ParsingState.ArrayClose

5655              ElseIf c = strArrayRowSeparator Then
5656                  varState = ParsingState.ArrayRowSeparator
5657                  lngTokenIndex = i

5658              ElseIf c = strArrayColumnSeparator Then
5659                  varState = ParsingState.ArrayColumnSeparator
5660                  lngTokenIndex = i

5661              Else
5662                  strError = "Expected " & strRightBrace & " " & strArrayRowSeparator & " " & ParsingState.ArrayColumnSeparator & " but got " & c & " at position " & i
5663                  varState = ParsingState.ParsingError
5664              End If

      ''' -------------------- -------------------- -------------------- '''

5665          ElseIf varState = ParsingState.ArrayOpen Then
5666              Set objToken = objTokens.NewToken(c, TokenType.ArrayOpen, i, 1)
5667              objTokens.Add objToken
5668              objTokenStack.Push objToken

5669              varState = ParsingState.ArrayConstant1
5670              i = i + 1

5671          ElseIf varState = ParsingState.ArrayClose Then
5672              Set objToken = objTokenStack.Pop
5673              If objToken.TokenType = TokenType.ArrayOpen Then
5674                  objTokens.Add objTokens.NewToken(c, TokenType.ArrayClose, i, 1)
5675                  varState = ParsingState.Expression2
5676                  i = i + 1
5677              Else
5678                  strError = "Encountered " & strRightBrace & " without matching " & strLeftBrace & " at position " & i
5679                  varState = ParsingState.ParsingError
5680              End If

5681          ElseIf varState = ParsingState.ArrayColumnSeparator Then
5682              i = i + 1
5683              objTokens.Add objTokens.NewToken(c, TokenType.ArrayColumnSeparator, lngTokenIndex, i - lngTokenIndex)
5684              varState = ParsingState.ArrayConstant1

5685          ElseIf varState = ParsingState.ArrayRowSeparator Then
5686              i = i + 1
5687              objTokens.Add objTokens.NewToken(c, TokenType.ArrayRowSeparator, lngTokenIndex, i - lngTokenIndex)
5688              varState = ParsingState.ArrayConstant1

      ''' -------------------- -------------------- -------------------- '''

5689          ElseIf varState = ParsingState.Bool Then
5690              str = str & c
5691              lng = Len(str)
5692              i = i + 1

5693              str2 = Chr(0) & strBooleanTrue & Chr(0) & strBooleanFalse & Chr(0)

5694              If InStr(1, str2, Chr(0) & str) > 0 Then
5695                  If InStr(1, str2, Chr(0) & str & Chr(0)) > 0 Then
5696                      objTokens.Add objTokens.NewToken(str, TokenType.Bool, lngTokenIndex, i - lngTokenIndex)
5697                      str = ""
5698                      varState = varReturnState
5699                  End If
5700              Else
5701                  strError = "Expected Array Constant at position " & lngTokenIndex
5702                  varState = ParsingState.ParsingError
5703              End If

      ''' -------------------- -------------------- -------------------- '''

5704          ElseIf varState = ParsingState.whitespace Then
5705              If c = " " Or c = Chr(10) Then
5706                  str = str & c
5707                  i = i + 1
5708              Else
5709                  objTokens.Add objTokens.NewToken(str, TokenType.whitespace, lngTokenIndex, i - lngTokenIndex)
5710                  str = ""
5711                  varState = varReturnState

5712              End If

      ''' -------------------- -------------------- -------------------- '''

5713          ElseIf varState = ParsingState.SubExpression Then
5714              Set objToken = objTokens.NewToken(c, TokenType.SubExpressionOpen, i, 1)
5715              objTokens.Add objToken
5716              objTokenStack.Push objToken

5717              varState = ParsingState.Expression1
5718              i = i + 1

5719          ElseIf varState = ParsingState.BracketClose Then
5720              Set objToken = objTokenStack.Pop
5721              If objToken.TokenType = TokenType.FunctionOpen Then
5722                  Set objToken = objTokens.NewToken(c, TokenType.FunctionClose, i, 1)
5723              ElseIf objToken.TokenType = TokenType.SubExpressionOpen Then
5724                  Set objToken = objTokens.NewToken(c, TokenType.SubExpressionClose, i, 1)
5725              Else
5726                  strError = "Encountered " & strRightRoundBracket & " without matching " & strLeftRoundBracket & " at position " & i
5727                  varState = ParsingState.ParsingError
5728              End If
5729              objTokens.Add objToken

5730              varState = ParsingState.Expression2
5731              i = i + 1

      ''' -------------------- -------------------- -------------------- '''
      ''' Decide: Leading Name for Function, Table, R1C1Reference, Reference Qualifier or Cell Reference

5732          ElseIf varState = ParsingState.LeadingName1 Then
5733              If str <> "" And c = strLeftRoundBracket Then
5734                  varState = ParsingState.FunctionX
5735                  bln = True

5736              ElseIf str <> "" And c = "[" Then                   'todo move symbol to variable
5737                  varState = ParsingState.SquareBracketOpen

5738              ElseIf str <> "" And c = "!" Then
5739                  varState = ParsingState.ReferenceQualifier

5740              ElseIf c = "" Or c = strRightRoundBracket Or c = strRightBrace Or _
                         c = " " Or c = Chr(10) Or _
                         c = "+" Or c = "-" Or c = "*" Or c = "/" Or c = "^" Or c = "%" Or _
                         c = "=" Or c = "<" Or c = ">" Or _
                         c = "&" Or c = ":" Or c = strListSeparator Then
5741                  varState = ParsingState.LeadingNameE
5742              Else
5743                  str = str & c
5744                  i = i + 1
5745              End If

5746          ElseIf varState = ParsingState.LeadingName2 Then
5747              If c = "'" Then
5748                  varState = ParsingState.LeadingName3
5749              Else
5750                  str = str & c
5751              End If
5752              i = i + 1

5753          ElseIf varState = ParsingState.LeadingName3 Then
5754              If c = "'" Then
5755                  varState = ParsingState.LeadingName2   'Was escape sequence
5756                  str = str & c
5757                  i = i + 1
5758              Else
5759                  varState = ParsingState.LeadingName1
5760              End If

5761          ElseIf varState = ParsingState.LeadingNameE Then
5762              If str <> "" Then
5763                  If str = strBooleanTrue Or str = strBooleanFalse Then
5764                      Set objToken = objTokens.NewToken(str, TokenType.Bool, lngTokenIndex, i - lngTokenIndex)
5765                  Else
5766                      Set objToken = objTokens.NewToken(str, TokenType.Reference, lngTokenIndex, i - lngTokenIndex)
5767                  End If
5768                  objTokens.Add objToken
5769                  str = ""
5770              End If

5771              varState = ParsingState.Expression2

      ''' -------------------- -------------------- -------------------- '''
      ''' Function

5772          ElseIf varState = ParsingState.FunctionX Then
5773              If Left(str, 1) = "@" Then str = Mid(str, 2)
5774              Set objToken = objTokens.NewToken(str, TokenType.FunctionOpen, lngTokenIndex, i - lngTokenIndex + 1)
5775              objTokens.Add objToken
5776              objTokenStack.Push objToken
5777              str = ""
5778              varState = ParsingState.Expression1
5779              i = i + 1

      ''' -------------------- -------------------- -------------------- '''
      ''' Decide: Table or R1C1 Reference

5780          ElseIf varState = ParsingState.SquareBracketOpen Then
5781              If str = "R" Or str = "C" Then                          'todo: use International characters
5782                  varState = ParsingState.R1C1Reference
5783              Else
5784                  Set objToken = objTokens.NewToken(str, TokenType.TableOpen, lngTokenIndex, i - lngTokenIndex + 1)
5785                  objTokens.Add objToken
5786                  objTokenStack.Push objToken
5787                  str = ""
5788                  varState = ParsingState.Table1
5789                  i = i + 1
5790              End If

      ''' -------------------- -------------------- -------------------- '''
      ''' Table

          'todo: escape character

5791          ElseIf varState = ParsingState.Table1 Then
5792              lngTokenIndex = i
5793              If c = " " Then
5794                  varReturnState = ParsingState.Table1
5795                  varState = ParsingState.whitespace
5796              ElseIf c = "[" Then                                     'todo, possibly use international bracket
5797                  i = i + 1
5798                  varState = ParsingState.Table6
5799              Else
5800                  varState = ParsingState.Table5
5801              End If

5802          ElseIf varState = ParsingState.Table2 Then
5803              lngTokenIndex = i
5804              If c = " " Then
5805                  varReturnState = ParsingState.Table2
5806                  varState = ParsingState.whitespace
5807              ElseIf c = "," Then                                     'todo possibly use list separator?
5808                  varState = ParsingState.Table3
5809              ElseIf c = ":" Then
5810                  varState = ParsingState.Table4
5811              ElseIf c = "]" Then                                     'todo, possibly use international bracket
5812                  varState = ParsingState.TableE
5813              Else
5814                  varState = ParsingState.Table5
5815              End If

5816          ElseIf varState = ParsingState.Table3 Then
5817              i = i + 1

5818              objTokens.Add objTokens.NewToken(c, TokenType.TableItemSeparator, lngTokenIndex, i - lngTokenIndex)
5819              str = ""
5820              varState = ParsingState.Table1

5821          ElseIf varState = ParsingState.Table4 Then
5822              i = i + 1

                  'todo: decide whether to make this a tablecolumns token type, or leave separated
5823              objTokens.Add objTokens.NewToken(c, TokenType.TableColumnSeparator, lngTokenIndex, i - lngTokenIndex)
5824              str = ""
5825              varState = ParsingState.Table1

5826          ElseIf varState = ParsingState.Table5 Then
5827              If c = "]" Then                             'todo, possibly use international bracket
5828                  If Left(str, 1) = "#" Then
5829                      objTokens.Add objTokens.NewToken(str, TokenType.TableSection, lngTokenIndex, i - lngTokenIndex)
5830                  Else
5831                      objTokens.Add objTokens.NewToken(str, TokenType.TableColumn, lngTokenIndex, i - lngTokenIndex)
5832                  End If
5833                  str = ""
5834                  varState = ParsingState.TableE
5835              ElseIf c = "'" Then
5836                  i = i + 1
5837                  varReturnState = ParsingState.Table5
5838                  varState = ParsingState.TableQ
5839              Else
5840                  str = str & c
5841                  i = i + 1
5842              End If

5843          ElseIf varState = ParsingState.Table6 Then
5844              If c = "]" Then                             'todo, possibly use international bracket
5845                  i = i + 1
5846                  If Left(str, 1) = "#" Then
5847                      objTokens.Add objTokens.NewToken(str, TokenType.TableSection, lngTokenIndex, i - lngTokenIndex)
5848                  Else
5849                      objTokens.Add objTokens.NewToken(str, TokenType.TableColumn, lngTokenIndex, i - lngTokenIndex)
5850                  End If
5851                  str = ""
5852                  varState = ParsingState.Table2
5853              ElseIf c = "'" Then
5854                  i = i + 1
5855                  varReturnState = ParsingState.Table6
5856                  varState = ParsingState.TableQ
5857              Else
5858                  str = str & c
5859                  i = i + 1
5860              End If

5861          ElseIf varState = ParsingState.TableQ Then
5862              str = str & c
5863              i = i + 1
5864              varState = varReturnState

5865          ElseIf varState = ParsingState.TableE Then
5866              Set objToken = objTokenStack.Pop
5867              If objToken.TokenType = TokenType.TableOpen Then
5868                  objTokens.Add objTokens.NewToken(c, TokenType.TableClose, i, 1)
5869                  varState = ParsingState.Expression2
5870                  i = i + 1
5871              Else
5872                  strError = "Encountered " & "]" & " without matching " & "[" & " at position " & i 'todo, possibly use international bracket
5873                  varState = ParsingState.ParsingError
5874              End If


      ''' -------------------- -------------------- -------------------- '''
      ''' Text

5875          ElseIf varState = ParsingState.Text1 Then
5876              If c = strDoubleQuote Then
5877                  varState = ParsingState.Text2
5878              Else
5879                  str = str & c
5880              End If
5881              i = i + 1

5882          ElseIf varState = ParsingState.Text2 Then
5883              If c = strDoubleQuote Then
5884                  varState = ParsingState.Text1   'Was escape sequence
5885                  str = str & c
5886                  i = i + 1
5887              Else
5888                  objTokens.Add objTokens.NewToken(str, TokenType.Text, lngTokenIndex, i - lngTokenIndex)
5889                  str = ""
5890                  varState = varReturnState
5891              End If

      ''' -------------------- -------------------- -------------------- '''
      ''' Number

5892          ElseIf varState = ParsingState.Number1 Or varState = ParsingState.Number2 Then
5893              If c = "0" Or c = "1" Or c = "2" Or c = "3" Or c = "4" Or _
                     c = "5" Or c = "6" Or c = "7" Or c = "8" Or c = "9" Then
5894                  str = str & c
5895                  i = i + 1

5896              ElseIf c = strDecimalSeparator And varState = ParsingState.Number1 Then
5897                  str = str & c
5898                  varState = ParsingState.Number2
5899                  i = i + 1

5900              ElseIf c = "E" Then
5901                  str = str & c
5902                  varState = ParsingState.Number3
5903                  i = i + 1

5904              Else
5905                  varState = ParsingState.NumberE
5906              End If

5907          ElseIf varState = ParsingState.Number3 Then
5908              If c = "+" Or c = "-" Then
5909                  str = str & c
5910                  varState = ParsingState.Number4
5911                  i = i + 1
5912              Else
5913                  strError = "Expected + or - at position " & i
5914                  varState = ParsingState.ParsingError
5915              End If

5916          ElseIf varState = ParsingState.Number4 Then
5917              If c = "0" Or c = "1" Or c = "2" Or c = "3" Or c = "4" Or _
                     c = "5" Or c = "6" Or c = "7" Or c = "8" Or c = "9" Then
5918                  str = str & c
5919                  i = i + 1
5920              Else
5921                  varState = ParsingState.NumberE
5922              End If

5923          ElseIf varState = ParsingState.NumberE Then
5924              objTokens.Add objTokens.NewToken(str, TokenType.Number, lngTokenIndex, i - lngTokenIndex)
5925              str = ""
5926              varState = varReturnState

      ''' -------------------- -------------------- -------------------- '''
      ''' Error

5927          ElseIf varState = ParsingState.ErrorX Then
5928              str = str & c
5929              lng = Len(str)
5930              i = i + 1

5931              str2 = Chr(0) & strErrorRef & Chr(0) & strErrorDiv0 & Chr(0) & strErrorNA & Chr(0) & strErrorName & _
                         Chr(0) & strErrorNull & Chr(0) & strErrorNum & Chr(0) & strErrorValue & Chr(0) & strErrorGettingData & Chr(0)

5932              If InStr(1, str2, Chr(0) & str) > 0 Then
5933                  If InStr(1, str2, Chr(0) & str & Chr(0)) > 0 Then
5934                      objTokens.Add objTokens.NewToken(str, TokenType.ErrorText, lngTokenIndex, i - lngTokenIndex)
5935                      str = ""
5936                      varState = varReturnState
5937                  End If
5938              Else
5939                  strError = "Expected Error Constant at position " & lngTokenIndex
5940                  varState = ParsingState.ParsingError
5941              End If

      ''' -------------------- -------------------- -------------------- '''

5942          ElseIf varState = ParsingState.ReferenceQualifier Then
5943              objTokens.Add objTokens.NewToken(str, TokenType.ReferenceQualifier, lngTokenIndex, i - lngTokenIndex)
5944              objTokens.Add objTokens.NewToken(c, TokenType.ExternalReferenceOperator, i, 1)

5945              i = i + 1
5946              str = ""
5947              varState = ParsingState.Expression1

5948          ElseIf varState = ParsingState.MinusSign Then
5949              If c = "0" Or c = "1" Or c = "2" Or c = "3" Or c = "4" Or _
                     c = "5" Or c = "6" Or c = "7" Or c = "8" Or c = "9" Then
5950                  varState = ParsingState.Number1
5951              Else
5952                  varState = ParsingState.PrefixOperator
5953                  i = i - 1
5954              End If

5955          ElseIf varState = ParsingState.PrefixOperator Then
5956              i = i + 1

5957              objTokens.Add objTokens.NewToken(c, TokenType.UnaryOperator, lngTokenIndex, i - lngTokenIndex)
5958              str = ""
5959              varState = ParsingState.Expression1

      ''' -------------------- -------------------- -------------------- '''

              'ListSeparator can be either function parameter separator, or a union operator
5960          ElseIf varState = ParsingState.ListSeparator Then
5961              bln = False
5962              If objTokenStack.Count > 0 Then
5963                  If objTokenStack.Peek.TokenType = TokenType.FunctionOpen Then
5964                      varState = ParsingState.ParameterSeparator
5965                      bln = True
5966                  End If
5967              End If
5968              If Not bln Then varState = ParsingState.RangeOperator

5969          ElseIf varState = ParsingState.RangeOperator Then
5970              str = c
5971              i = i + 1

5972              objTokens.Add objTokens.NewToken(str, TokenType.RangeOperator, lngTokenIndex, i - lngTokenIndex)
5973              str = ""
5974              varState = ParsingState.Expression1

      ''' -------------------- -------------------- -------------------- '''

5975          ElseIf varState = ParsingState.ArithmeticOperator Then
5976              str = c
5977              i = i + 1

5978              objTokens.Add objTokens.NewToken(c, TokenType.ArithmeticOperator, lngTokenIndex, i - lngTokenIndex)
5979              str = ""
5980              varState = ParsingState.Expression1

      ''' -------------------- -------------------- -------------------- '''

5981          ElseIf varState = ParsingState.ComparisonOperator1 Then
5982              str = c
5983              i = i + 1

5984              If c = "<" Or c = ">" Then
5985                  varState = ParsingState.ComparisonOperator2
5986              Else
5987                  varState = ParsingState.ComparisonOperatorE
5988              End If

5989          ElseIf varState = ParsingState.ComparisonOperator2 Then
5990              If c = "=" Or (str = "<" And c = ">") Then
5991                  str = str & c
5992                  i = i + 1
5993              End If

5994              varState = ParsingState.ComparisonOperatorE

5995          ElseIf varState = ParsingState.ComparisonOperatorE Then
5996              objTokens.Add objTokens.NewToken(str, TokenType.ComparisonOperator, lngTokenIndex, i - lngTokenIndex)
5997              str = ""
5998              varState = ParsingState.Expression1

      ''' -------------------- -------------------- -------------------- '''

5999          ElseIf varState = ParsingState.TextOperator Then
6000              str = c
6001              i = i + 1

6002              objTokens.Add objTokens.NewToken(c, TokenType.TextOperator, lngTokenIndex, i - lngTokenIndex)
6003              str = ""
6004              varState = ParsingState.Expression1

      ''' -------------------- -------------------- -------------------- '''

6005          ElseIf varState = ParsingState.PostfixOperator Then
6006              str = c
6007              i = i + 1

6008              objTokens.Add objTokens.NewToken(str, TokenType.PostfixOperator, lngTokenIndex, i - lngTokenIndex)
6009              str = ""
6010              varState = ParsingState.Expression2

      ''' -------------------- -------------------- -------------------- '''

6011          ElseIf varState = ParsingState.ParameterSeparator Then
6012              str = c
6013              i = i + 1

6014              objTokens.Add objTokens.NewToken(str, TokenType.ParameterSeparator, lngTokenIndex, i - lngTokenIndex)
6015              str = ""
6016              varState = ParsingState.Expression1

6017          End If

6018          lngPrevIndex = i
6019      Loop While blnLoop

          'todo: error if token stack not empty

          'post processing
          '1. Scan and detect !, and join if the 3 tokens prior are CellReference-Colon-CellReference
          '2. Scan and detect CellReference-Colon-CellReference, because they should be joined into an AreaReference

6020 e:    Set ParseFormula = objTokens
End Function

