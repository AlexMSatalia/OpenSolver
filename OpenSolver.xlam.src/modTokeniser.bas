Attribute VB_Name = "modTokeniser"
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

    Set objTokens = New Tokens
    Set objTokenStack = New TokenStack

    strDecimalSeparator = Application.International(xlDecimalSeparator)
    strListSeparator = Application.International(xlListSeparator)
    strArrayRowSeparator = Application.International(xlRowSeparator)
    If Application.International(xlColumnSeparator) = Application.International(xlDecimalSeparator) Then
        strArrayColumnSeparator = Application.International(xlAlternateArraySeparator)
    Else
        strArrayColumnSeparator = Application.International(xlColumnSeparator)
    End If
    strLeftBrace = Application.International(xlLeftBrace)
    strRightBrace = Application.International(xlRightBrace)
    strLeftBracket = Application.International(xlLeftBracket)       'todo: implement ' check if this applies to tableopen and tableclose symbols
    strRightBracket = Application.International(xlRightBracket)     'todo: implement
    strLeftRoundBracket = "("
    strRightRoundBracket = ")"
    strBooleanTrue = "TRUE"
    strBooleanFalse = "FALSE"
    strDoubleQuote = """"

    strErrorRef = "#REF!"
    strErrorDiv0 = "#DIV/0!"
    strErrorNA = "#N/A"
    strErrorName = "#NAME?"
    strErrorNull = "#NULL!"
    strErrorNum = "#NUM!"
    strErrorValue = "#VALUE!"
    strErrorGettingData = "#GETTING_DATA"

    lngFormulaLen = Len(strFormula)

    If lngFormulaLen <= 1 Then GoTo e
    If left(strFormula, 1) <> "=" Then GoTo e

    varState = ParsingState.Expression1
    i = 2
    lngPrevIndex = 1

    blnLoop = True

    Do
        If i <= lngFormulaLen Then c = Mid(strFormula, i, 1) Else c = ""

''' -------------------- -------------------- -------------------- '''

        If varState = ParsingState.ParsingError Then
            MsgBox strError, vbCritical, "Error"
            blnLoop = False
''' -------------------- -------------------- -------------------- '''

        ElseIf varState = ParsingState.Expression1 Then
            If c = " " Or c = Chr(10) Then
                varReturnState = varState
                varState = ParsingState.whitespace
                lngTokenIndex = i

            ElseIf c = strDoubleQuote Then
                varReturnState = ParsingState.Expression2
                varState = ParsingState.Text1
                lngTokenIndex = i
                i = i + 1

            ElseIf c = "0" Or c = "1" Or c = "2" Or c = "3" Or c = "4" Or _
                   c = "5" Or c = "6" Or c = "7" Or c = "8" Or c = "9" Then
                varReturnState = ParsingState.Expression2
                varState = ParsingState.Number1
                lngTokenIndex = i

            ElseIf c = "-" Then
                str = c
                varReturnState = ParsingState.Expression2
                varState = ParsingState.MinusSign
                lngTokenIndex = i
                i = i + 1

            ElseIf c = "+" Then
                varState = ParsingState.PrefixOperator
                lngTokenIndex = i

            ElseIf c = strLeftRoundBracket Then
                varState = ParsingState.SubExpression

            ElseIf c = strLeftBrace Then
                varState = ParsingState.ArrayOpen

            ElseIf c = "'" Then
                varState = ParsingState.LeadingName2
                lngTokenIndex = i
                i = i + 1

            ElseIf c = "#" Then
                varReturnState = ParsingState.Expression2
                varState = ParsingState.ErrorX
                lngTokenIndex = i

            Else
                varState = ParsingState.LeadingName1
                lngTokenIndex = i

            End If

''' -------------------- -------------------- -------------------- '''

        ElseIf varState = ParsingState.Expression2 Then
            If c = "" Then
                blnLoop = False

            ElseIf c = " " Or c = Chr(10) Then
                varReturnState = varState
                varState = ParsingState.whitespace
                lngTokenIndex = i

            ElseIf c = strRightRoundBracket Then
                varState = ParsingState.BracketClose

            ElseIf c = "+" Or c = "-" Or c = "*" Or c = "/" Or c = "^" Then
                varState = ParsingState.ArithmeticOperator
                lngTokenIndex = i

            ElseIf c = "=" Or c = "<" Or c = ">" Then
                varState = ParsingState.ComparisonOperator1
                lngTokenIndex = i

            ElseIf c = "&" Then
                varState = ParsingState.TextOperator
                lngTokenIndex = i

            ElseIf c = "%" Then
                varState = ParsingState.PostfixOperator
                lngTokenIndex = i

            ElseIf c = ":" Then
                varState = ParsingState.RangeOperator
                lngTokenIndex = i

            ElseIf c = strListSeparator Then
                varState = ParsingState.ListSeparator
                lngTokenIndex = i

            Else
                'Check if whitespace is actually union operator
                Set objToken = objTokens(objTokens.Count)
                bln = False
                If objToken.TokenType = TokenType.whitespace Then
                    lng = InStr(1, objToken.Text, " ")
                    If lng > 0 Then
                        str = Mid(objToken.Text, lng + 1)
                        If lng = 1 Then
                            objToken.TokenType = TokenType.RangeOperator
                            objToken.Text = " "
                            objToken.FormulaLength = 1
                        Else
                            objToken.Text = left(objToken.Text, lng - 1)
                            objToken.FormulaLength = lng - 1
                            objTokens.Add objTokens.NewToken(" ", TokenType.RangeOperator, objToken.FormulaIndex + lng - 1, 1)
                        End If

                        If str <> "" Then
                            objTokens.Add objTokens.NewToken(str, TokenType.whitespace, objToken.FormulaIndex + lng, Len(str))
                            str = ""
                        End If

                        varState = ParsingState.Expression1
                        bln = True
                    End If

                ElseIf objToken.TokenType = TokenType.ErrorText And objToken.Text = strErrorRef Then
                    varState = ParsingState.Expression1
                    bln = True
                End If
                If Not bln Then
                    strError = "Expected Operator, but got " & c & " at position " & i
                    varState = ParsingState.ParsingError
                End If
            End If

''' -------------------- -------------------- -------------------- '''

        ElseIf varState = ParsingState.ArrayConstant1 Then
            If c = strDoubleQuote Then
                varReturnState = ParsingState.ArrayConstant2
                varState = ParsingState.Text1
                lngTokenIndex = i
                i = i + 1

            ElseIf c = "0" Or c = "1" Or c = "2" Or c = "3" Or c = "4" Or _
                   c = "5" Or c = "6" Or c = "7" Or c = "8" Or c = "9" Then
                varReturnState = ParsingState.ArrayConstant2
                varState = ParsingState.Number1
                lngTokenIndex = i

            ElseIf c = "-" Then
                str = c
                varReturnState = ParsingState.ArrayConstant2
                varState = ParsingState.Number1
                lngTokenIndex = i
                i = i + 1

            ElseIf c = "#" Then
                varReturnState = ParsingState.ArrayConstant2
                varState = ParsingState.ErrorX
                lngTokenIndex = i

            Else
                varReturnState = ParsingState.ArrayConstant2
                varState = ParsingState.Bool
                lngTokenIndex = i

            End If

''' -------------------- -------------------- -------------------- '''

        ElseIf varState = ParsingState.ArrayConstant2 Then
            If c = strRightBrace Then
                varState = ParsingState.ArrayClose

            ElseIf c = strArrayRowSeparator Then
                varState = ParsingState.ArrayRowSeparator
                lngTokenIndex = i

            ElseIf c = strArrayColumnSeparator Then
                varState = ParsingState.ArrayColumnSeparator
                lngTokenIndex = i

            Else
                strError = "Expected " & strRightBrace & " " & strArrayRowSeparator & " " & ParsingState.ArrayColumnSeparator & " but got " & c & " at position " & i
                varState = ParsingState.ParsingError
            End If

''' -------------------- -------------------- -------------------- '''

        ElseIf varState = ParsingState.ArrayOpen Then
            Set objToken = objTokens.NewToken(c, TokenType.ArrayOpen, i, 1)
            objTokens.Add objToken
            objTokenStack.Push objToken

            varState = ParsingState.ArrayConstant1
            i = i + 1

        ElseIf varState = ParsingState.ArrayClose Then
            Set objToken = objTokenStack.Pop
            If objToken.TokenType = TokenType.ArrayOpen Then
                objTokens.Add objTokens.NewToken(c, TokenType.ArrayClose, i, 1)
                varState = ParsingState.Expression2
                i = i + 1
            Else
                strError = "Encountered " & strRightBrace & " without matching " & strLeftBrace & " at position " & i
                varState = ParsingState.ParsingError
            End If

        ElseIf varState = ParsingState.ArrayColumnSeparator Then
            i = i + 1
            objTokens.Add objTokens.NewToken(c, TokenType.ArrayColumnSeparator, lngTokenIndex, i - lngTokenIndex)
            varState = ParsingState.ArrayConstant1

        ElseIf varState = ParsingState.ArrayRowSeparator Then
            i = i + 1
            objTokens.Add objTokens.NewToken(c, TokenType.ArrayRowSeparator, lngTokenIndex, i - lngTokenIndex)
            varState = ParsingState.ArrayConstant1

''' -------------------- -------------------- -------------------- '''

        ElseIf varState = ParsingState.Bool Then
            str = str & c
            lng = Len(str)
            i = i + 1

            str2 = Chr(0) & strBooleanTrue & Chr(0) & strBooleanFalse & Chr(0)

            If InStr(1, str2, Chr(0) & str) > 0 Then
                If InStr(1, str2, Chr(0) & str & Chr(0)) > 0 Then
                    objTokens.Add objTokens.NewToken(str, TokenType.Bool, lngTokenIndex, i - lngTokenIndex)
                    str = ""
                    varState = varReturnState
                End If
            Else
                strError = "Expected Array Constant at position " & lngTokenIndex
                varState = ParsingState.ParsingError
            End If

''' -------------------- -------------------- -------------------- '''

        ElseIf varState = ParsingState.whitespace Then
            If c = " " Or c = Chr(10) Then
                str = str & c
                i = i + 1
            Else
                objTokens.Add objTokens.NewToken(str, TokenType.whitespace, lngTokenIndex, i - lngTokenIndex)
                str = ""
                varState = varReturnState

            End If

''' -------------------- -------------------- -------------------- '''

        ElseIf varState = ParsingState.SubExpression Then
            Set objToken = objTokens.NewToken(c, TokenType.SubExpressionOpen, i, 1)
            objTokens.Add objToken
            objTokenStack.Push objToken

            varState = ParsingState.Expression1
            i = i + 1

        ElseIf varState = ParsingState.BracketClose Then
            Set objToken = objTokenStack.Pop
            If objToken.TokenType = TokenType.FunctionOpen Then
                Set objToken = objTokens.NewToken(c, TokenType.FunctionClose, i, 1)
            ElseIf objToken.TokenType = TokenType.SubExpressionOpen Then
                Set objToken = objTokens.NewToken(c, TokenType.SubExpressionClose, i, 1)
            Else
                strError = "Encountered " & strRightRoundBracket & " without matching " & strLeftRoundBracket & " at position " & i
                varState = ParsingState.ParsingError
            End If
            objTokens.Add objToken

            varState = ParsingState.Expression2
            i = i + 1

''' -------------------- -------------------- -------------------- '''
''' Decide: Leading Name for Function, Table, R1C1Reference, Reference Qualifier or Cell Reference

        ElseIf varState = ParsingState.LeadingName1 Then
            If str <> "" And c = strLeftRoundBracket Then
                varState = ParsingState.FunctionX
                bln = True

            ElseIf str <> "" And c = "[" Then                   'todo move symbol to variable
                varState = ParsingState.SquareBracketOpen

            ElseIf str <> "" And c = "!" Then
                varState = ParsingState.ReferenceQualifier

            ElseIf c = "" Or c = strRightRoundBracket Or c = strRightBrace Or _
                   c = " " Or c = Chr(10) Or _
                   c = "+" Or c = "-" Or c = "*" Or c = "/" Or c = "^" Or c = "%" Or _
                   c = "=" Or c = "<" Or c = ">" Or _
                   c = "&" Or c = ":" Or c = strListSeparator Then
                varState = ParsingState.LeadingNameE
            Else
                str = str & c
                i = i + 1
            End If

        ElseIf varState = ParsingState.LeadingName2 Then
            If c = "'" Then
                varState = ParsingState.LeadingName3
            Else
                str = str & c
            End If
            i = i + 1

        ElseIf varState = ParsingState.LeadingName3 Then
            If c = "'" Then
                varState = ParsingState.LeadingName2   'Was escape sequence
                str = str & c
                i = i + 1
            Else
                varState = ParsingState.LeadingName1
            End If

        ElseIf varState = ParsingState.LeadingNameE Then
            If str <> "" Then
                If str = strBooleanTrue Or str = strBooleanFalse Then
                    Set objToken = objTokens.NewToken(str, TokenType.Bool, lngTokenIndex, i - lngTokenIndex)
                Else
                    Set objToken = objTokens.NewToken(str, TokenType.Reference, lngTokenIndex, i - lngTokenIndex)
                End If
                objTokens.Add objToken
                str = ""
            End If

            varState = ParsingState.Expression2

''' -------------------- -------------------- -------------------- '''
''' Function

        ElseIf varState = ParsingState.FunctionX Then
            If left(str, 1) = "@" Then str = Mid(str, 2)
            Set objToken = objTokens.NewToken(str, TokenType.FunctionOpen, lngTokenIndex, i - lngTokenIndex + 1)
            objTokens.Add objToken
            objTokenStack.Push objToken
            str = ""
            varState = ParsingState.Expression1
            i = i + 1

''' -------------------- -------------------- -------------------- '''
''' Decide: Table or R1C1 Reference

        ElseIf varState = ParsingState.SquareBracketOpen Then
            If str = "R" Or str = "C" Then                          'todo: use International characters
                varState = ParsingState.R1C1Reference
            Else
                Set objToken = objTokens.NewToken(str, TokenType.TableOpen, lngTokenIndex, i - lngTokenIndex + 1)
                objTokens.Add objToken
                objTokenStack.Push objToken
                str = ""
                varState = ParsingState.Table1
                i = i + 1
            End If

''' -------------------- -------------------- -------------------- '''
''' Table

    'todo: escape character

        ElseIf varState = ParsingState.Table1 Then
            lngTokenIndex = i
            If c = " " Then
                varReturnState = ParsingState.Table1
                varState = ParsingState.whitespace
            ElseIf c = "[" Then                                     'todo, possibly use international bracket
                i = i + 1
                varState = ParsingState.Table6
            Else
                varState = ParsingState.Table5
            End If

        ElseIf varState = ParsingState.Table2 Then
            lngTokenIndex = i
            If c = " " Then
                varReturnState = ParsingState.Table2
                varState = ParsingState.whitespace
            ElseIf c = "," Then                                     'todo possibly use list separator?
                varState = ParsingState.Table3
            ElseIf c = ":" Then
                varState = ParsingState.Table4
            ElseIf c = "]" Then                                     'todo, possibly use international bracket
                varState = ParsingState.TableE
            Else
                varState = ParsingState.Table5
            End If

        ElseIf varState = ParsingState.Table3 Then
            i = i + 1

            objTokens.Add objTokens.NewToken(c, TokenType.TableItemSeparator, lngTokenIndex, i - lngTokenIndex)
            str = ""
            varState = ParsingState.Table1

        ElseIf varState = ParsingState.Table4 Then
            i = i + 1

            'todo: decide whether to make this a tablecolumns token type, or leave separated
            objTokens.Add objTokens.NewToken(c, TokenType.TableColumnSeparator, lngTokenIndex, i - lngTokenIndex)
            str = ""
            varState = ParsingState.Table1

        ElseIf varState = ParsingState.Table5 Then
            If c = "]" Then                             'todo, possibly use international bracket
                If left(str, 1) = "#" Then
                    objTokens.Add objTokens.NewToken(str, TokenType.TableSection, lngTokenIndex, i - lngTokenIndex)
                Else
                    objTokens.Add objTokens.NewToken(str, TokenType.TableColumn, lngTokenIndex, i - lngTokenIndex)
                End If
                str = ""
                varState = ParsingState.TableE
            ElseIf c = "'" Then
                i = i + 1
                varReturnState = ParsingState.Table5
                varState = ParsingState.TableQ
            Else
                str = str & c
                i = i + 1
            End If

        ElseIf varState = ParsingState.Table6 Then
            If c = "]" Then                             'todo, possibly use international bracket
                i = i + 1
                If left(str, 1) = "#" Then
                    objTokens.Add objTokens.NewToken(str, TokenType.TableSection, lngTokenIndex, i - lngTokenIndex)
                Else
                    objTokens.Add objTokens.NewToken(str, TokenType.TableColumn, lngTokenIndex, i - lngTokenIndex)
                End If
                str = ""
                varState = ParsingState.Table2
            ElseIf c = "'" Then
                i = i + 1
                varReturnState = ParsingState.Table6
                varState = ParsingState.TableQ
            Else
                str = str & c
                i = i + 1
            End If

        ElseIf varState = ParsingState.TableQ Then
            str = str & c
            i = i + 1
            varState = varReturnState

        ElseIf varState = ParsingState.TableE Then
            Set objToken = objTokenStack.Pop
            If objToken.TokenType = TokenType.TableOpen Then
                objTokens.Add objTokens.NewToken(c, TokenType.TableClose, i, 1)
                varState = ParsingState.Expression2
                i = i + 1
            Else
                strError = "Encountered " & "]" & " without matching " & "[" & " at position " & i 'todo, possibly use international bracket
                varState = ParsingState.ParsingError
            End If


''' -------------------- -------------------- -------------------- '''
''' Text

        ElseIf varState = ParsingState.Text1 Then
            If c = strDoubleQuote Then
                varState = ParsingState.Text2
            Else
                str = str & c
            End If
            i = i + 1

        ElseIf varState = ParsingState.Text2 Then
            If c = strDoubleQuote Then
                varState = ParsingState.Text1   'Was escape sequence
                str = str & c
                i = i + 1
            Else
                objTokens.Add objTokens.NewToken(str, TokenType.Text, lngTokenIndex, i - lngTokenIndex)
                str = ""
                varState = varReturnState
            End If

''' -------------------- -------------------- -------------------- '''
''' Number

        ElseIf varState = ParsingState.Number1 Or varState = ParsingState.Number2 Then
            If c = "0" Or c = "1" Or c = "2" Or c = "3" Or c = "4" Or _
               c = "5" Or c = "6" Or c = "7" Or c = "8" Or c = "9" Then
                str = str & c
                i = i + 1

            ElseIf c = strDecimalSeparator And varState = ParsingState.Number1 Then
                str = str & c
                varState = ParsingState.Number2
                i = i + 1

            ElseIf c = "E" Then
                str = str & c
                varState = ParsingState.Number3
                i = i + 1

            Else
                varState = ParsingState.NumberE
            End If

        ElseIf varState = ParsingState.Number3 Then
            If c = "+" Or c = "-" Then
                str = str & c
                varState = ParsingState.Number4
                i = i + 1
            Else
                strError = "Expected + or - at position " & i
                varState = ParsingState.ParsingError
            End If

        ElseIf varState = ParsingState.Number4 Then
            If c = "0" Or c = "1" Or c = "2" Or c = "3" Or c = "4" Or _
               c = "5" Or c = "6" Or c = "7" Or c = "8" Or c = "9" Then
                str = str & c
                i = i + 1
            Else
                varState = ParsingState.NumberE
            End If

        ElseIf varState = ParsingState.NumberE Then
            objTokens.Add objTokens.NewToken(str, TokenType.Number, lngTokenIndex, i - lngTokenIndex)
            str = ""
            varState = varReturnState

''' -------------------- -------------------- -------------------- '''
''' Error

        ElseIf varState = ParsingState.ErrorX Then
            str = str & c
            lng = Len(str)
            i = i + 1

            str2 = Chr(0) & strErrorRef & Chr(0) & strErrorDiv0 & Chr(0) & strErrorNA & Chr(0) & strErrorName & _
                   Chr(0) & strErrorNull & Chr(0) & strErrorNum & Chr(0) & strErrorValue & Chr(0) & strErrorGettingData & Chr(0)

            If InStr(1, str2, Chr(0) & str) > 0 Then
                If InStr(1, str2, Chr(0) & str & Chr(0)) > 0 Then
                    objTokens.Add objTokens.NewToken(str, TokenType.ErrorText, lngTokenIndex, i - lngTokenIndex)
                    str = ""
                    varState = varReturnState
                End If
            Else
                strError = "Expected Error Constant at position " & lngTokenIndex
                varState = ParsingState.ParsingError
            End If

''' -------------------- -------------------- -------------------- '''

        ElseIf varState = ParsingState.ReferenceQualifier Then
            objTokens.Add objTokens.NewToken(str, TokenType.ReferenceQualifier, lngTokenIndex, i - lngTokenIndex)
            objTokens.Add objTokens.NewToken(c, TokenType.ExternalReferenceOperator, i, 1)

            i = i + 1
            str = ""
            varState = ParsingState.Expression1

        ElseIf varState = ParsingState.MinusSign Then
            If c = "0" Or c = "1" Or c = "2" Or c = "3" Or c = "4" Or _
               c = "5" Or c = "6" Or c = "7" Or c = "8" Or c = "9" Then
                varState = ParsingState.Number1
            Else
                varState = ParsingState.PrefixOperator
                i = i - 1
            End If

        ElseIf varState = ParsingState.PrefixOperator Then
            i = i + 1

            objTokens.Add objTokens.NewToken(c, TokenType.UnaryOperator, lngTokenIndex, i - lngTokenIndex)
            str = ""
            varState = ParsingState.Expression1

''' -------------------- -------------------- -------------------- '''

        'ListSeparator can be either function parameter separator, or a union operator
        ElseIf varState = ParsingState.ListSeparator Then
            bln = False
            If objTokenStack.Count > 0 Then
                If objTokenStack.Peek.TokenType = TokenType.FunctionOpen Then
                    varState = ParsingState.ParameterSeparator
                    bln = True
                End If
            End If
            If Not bln Then varState = ParsingState.RangeOperator

        ElseIf varState = ParsingState.RangeOperator Then
            str = c
            i = i + 1

            objTokens.Add objTokens.NewToken(str, TokenType.RangeOperator, lngTokenIndex, i - lngTokenIndex)
            str = ""
            varState = ParsingState.Expression1

''' -------------------- -------------------- -------------------- '''

        ElseIf varState = ParsingState.ArithmeticOperator Then
            str = c
            i = i + 1

            objTokens.Add objTokens.NewToken(c, TokenType.ArithmeticOperator, lngTokenIndex, i - lngTokenIndex)
            str = ""
            varState = ParsingState.Expression1

''' -------------------- -------------------- -------------------- '''

        ElseIf varState = ParsingState.ComparisonOperator1 Then
            str = c
            i = i + 1

            If c = "<" Or c = ">" Then
                varState = ParsingState.ComparisonOperator2
            Else
                varState = ParsingState.ComparisonOperatorE
            End If

        ElseIf varState = ParsingState.ComparisonOperator2 Then
            If c = "=" Or (str = "<" And c = ">") Then
                str = str & c
                i = i + 1
            End If

            varState = ParsingState.ComparisonOperatorE

        ElseIf varState = ParsingState.ComparisonOperatorE Then
            objTokens.Add objTokens.NewToken(str, TokenType.ComparisonOperator, lngTokenIndex, i - lngTokenIndex)
            str = ""
            varState = ParsingState.Expression1

''' -------------------- -------------------- -------------------- '''

        ElseIf varState = ParsingState.TextOperator Then
            str = c
            i = i + 1

            objTokens.Add objTokens.NewToken(c, TokenType.TextOperator, lngTokenIndex, i - lngTokenIndex)
            str = ""
            varState = ParsingState.Expression1

''' -------------------- -------------------- -------------------- '''

        ElseIf varState = ParsingState.PostfixOperator Then
            str = c
            i = i + 1

            objTokens.Add objTokens.NewToken(str, TokenType.PostfixOperator, lngTokenIndex, i - lngTokenIndex)
            str = ""
            varState = ParsingState.Expression2

''' -------------------- -------------------- -------------------- '''

        ElseIf varState = ParsingState.ParameterSeparator Then
            str = c
            i = i + 1

            objTokens.Add objTokens.NewToken(str, TokenType.ParameterSeparator, lngTokenIndex, i - lngTokenIndex)
            str = ""
            varState = ParsingState.Expression1

        End If

        lngPrevIndex = i
    Loop While blnLoop

    'todo: error if token stack not empty

    'post processing
    '1. Scan and detect !, and join if the 3 tokens prior are CellReference-Colon-CellReference
    '2. Scan and detect CellReference-Colon-CellReference, because they should be joined into an AreaReference

e:  Set ParseFormula = objTokens
End Function

