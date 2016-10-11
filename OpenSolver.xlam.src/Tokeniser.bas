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

1         Set objTokens = New Tokens
2         Set objTokenStack = New TokenStack

3         strDecimalSeparator = "." ' Application.International(xlDecimalSeparator)
4         strListSeparator = "," ' Application.International(xlListSeparator)
5         strArrayRowSeparator = ";" ' Application.International(xlRowSeparator)
'5495      If Application.International(xlColumnSeparator) = Application.International(xlDecimalSeparator) Then
'5496          strArrayColumnSeparator = Application.International(xlAlternateArraySeparator)
'5497      Else
6             strArrayColumnSeparator = "," ' Application.International(xlColumnSeparator)
'5499      End If
7         strLeftBrace = "{" ' Application.International(xlLeftBrace)
8         strRightBrace = "}" ' Application.International(xlRightBrace)
9         strLeftBracket = "[" ' Application.International(xlLeftBracket)       'todo: implement ' check if this applies to tableopen and tableclose symbols
10        strRightBracket = "]" ' Application.International(xlRightBracket)     'todo: implement
11        strLeftRoundBracket = "("
12        strRightRoundBracket = ")"
13        strBooleanTrue = "TRUE"
14        strBooleanFalse = "FALSE"
15        strDoubleQuote = """"

16        strErrorRef = "#REF!"
17        strErrorDiv0 = "#DIV/0!"
18        strErrorNA = "#N/A"
19        strErrorName = "#NAME?"
20        strErrorNull = "#NULL!"
21        strErrorNum = "#NUM!"
22        strErrorValue = "#VALUE!"
23        strErrorGettingData = "#GETTING_DATA"

24        lngFormulaLen = Len(strFormula)

25        If lngFormulaLen <= 1 Then GoTo e
26        If Left(strFormula, 1) <> "=" Then GoTo e

27        varState = ParsingState.Expression1
28        i = 2
29        lngPrevIndex = 1

30        blnLoop = True

31        Do
32            If i <= lngFormulaLen Then c = Mid(strFormula, i, 1) Else c = ""

      ''' -------------------- -------------------- -------------------- '''

33            If varState = ParsingState.ParsingError Then
34                MsgBox strError, vbCritical, "Error"
35                blnLoop = False
      ''' -------------------- -------------------- -------------------- '''

36            ElseIf varState = ParsingState.Expression1 Then
37                If c = " " Or c = Chr(10) Then
38                    varReturnState = varState
39                    varState = ParsingState.whitespace
40                    lngTokenIndex = i

41                ElseIf c = strDoubleQuote Then
42                    varReturnState = ParsingState.Expression2
43                    varState = ParsingState.Text1
44                    lngTokenIndex = i
45                    i = i + 1

46                ElseIf c = "0" Or c = "1" Or c = "2" Or c = "3" Or c = "4" Or _
                         c = "5" Or c = "6" Or c = "7" Or c = "8" Or c = "9" Then
47                    varReturnState = ParsingState.Expression2
48                    varState = ParsingState.Number1
49                    lngTokenIndex = i

50                ElseIf c = "-" Then
51                    str = c
52                    varReturnState = ParsingState.Expression2
53                    varState = ParsingState.MinusSign
54                    lngTokenIndex = i
55                    i = i + 1

56                ElseIf c = "+" Then
57                    varState = ParsingState.PrefixOperator
58                    lngTokenIndex = i

59                ElseIf c = strLeftRoundBracket Then
60                    varState = ParsingState.SubExpression

61                ElseIf c = strLeftBrace Then
62                    varState = ParsingState.ArrayOpen

63                ElseIf c = "'" Then
64                    varState = ParsingState.LeadingName2
65                    lngTokenIndex = i
66                    i = i + 1

67                ElseIf c = "#" Then
68                    varReturnState = ParsingState.Expression2
69                    varState = ParsingState.ErrorX
70                    lngTokenIndex = i

71                Else
72                    varState = ParsingState.LeadingName1
73                    lngTokenIndex = i

74                End If

      ''' -------------------- -------------------- -------------------- '''

75            ElseIf varState = ParsingState.Expression2 Then
76                If c = "" Then
77                    blnLoop = False

78                ElseIf c = " " Or c = Chr(10) Then
79                    varReturnState = varState
80                    varState = ParsingState.whitespace
81                    lngTokenIndex = i

82                ElseIf c = strRightRoundBracket Then
83                    varState = ParsingState.BracketClose

84                ElseIf c = "+" Or c = "-" Or c = "*" Or c = "/" Or c = "^" Then
85                    varState = ParsingState.ArithmeticOperator
86                    lngTokenIndex = i

87                ElseIf c = "=" Or c = "<" Or c = ">" Then
88                    varState = ParsingState.ComparisonOperator1
89                    lngTokenIndex = i

90                ElseIf c = "&" Then
91                    varState = ParsingState.TextOperator
92                    lngTokenIndex = i

93                ElseIf c = "%" Then
94                    varState = ParsingState.PostfixOperator
95                    lngTokenIndex = i

96                ElseIf c = ":" Then
97                    varState = ParsingState.RangeOperator
98                    lngTokenIndex = i

99                ElseIf c = strListSeparator Then
100                   varState = ParsingState.ListSeparator
101                   lngTokenIndex = i

102               Else
                      'Check if whitespace is actually union operator
103                   Set objToken = objTokens(objTokens.Count)
104                   bln = False
105                   If objToken.TokenType = TokenType.whitespace Then
106                       lng = InStr(1, objToken.Text, " ")
107                       If lng > 0 Then
108                           str = Mid(objToken.Text, lng + 1)
109                           If lng = 1 Then
110                               objToken.TokenType = TokenType.RangeOperator
111                               objToken.Text = " "
112                               objToken.FormulaLength = 1
113                           Else
114                               objToken.Text = Left(objToken.Text, lng - 1)
115                               objToken.FormulaLength = lng - 1
116                               objTokens.Add objTokens.NewToken(" ", TokenType.RangeOperator, objToken.FormulaIndex + lng - 1, 1)
117                           End If

118                           If str <> "" Then
119                               objTokens.Add objTokens.NewToken(str, TokenType.whitespace, objToken.FormulaIndex + lng, Len(str))
120                               str = ""
121                           End If

122                           varState = ParsingState.Expression1
123                           bln = True
124                       End If

125                   ElseIf objToken.TokenType = TokenType.ErrorText And objToken.Text = strErrorRef Then
126                       varState = ParsingState.Expression1
127                       bln = True
128                   End If
129                   If Not bln Then
130                       strError = "Expected Operator, but got " & c & " at position " & i
131                       varState = ParsingState.ParsingError
132                   End If
133               End If

      ''' -------------------- -------------------- -------------------- '''

134           ElseIf varState = ParsingState.ArrayConstant1 Then
135               If c = strDoubleQuote Then
136                   varReturnState = ParsingState.ArrayConstant2
137                   varState = ParsingState.Text1
138                   lngTokenIndex = i
139                   i = i + 1

140               ElseIf c = "0" Or c = "1" Or c = "2" Or c = "3" Or c = "4" Or _
                         c = "5" Or c = "6" Or c = "7" Or c = "8" Or c = "9" Then
141                   varReturnState = ParsingState.ArrayConstant2
142                   varState = ParsingState.Number1
143                   lngTokenIndex = i

144               ElseIf c = "-" Then
145                   str = c
146                   varReturnState = ParsingState.ArrayConstant2
147                   varState = ParsingState.Number1
148                   lngTokenIndex = i
149                   i = i + 1

150               ElseIf c = "#" Then
151                   varReturnState = ParsingState.ArrayConstant2
152                   varState = ParsingState.ErrorX
153                   lngTokenIndex = i

154               Else
155                   varReturnState = ParsingState.ArrayConstant2
156                   varState = ParsingState.Bool
157                   lngTokenIndex = i

158               End If

      ''' -------------------- -------------------- -------------------- '''

159           ElseIf varState = ParsingState.ArrayConstant2 Then
160               If c = strRightBrace Then
161                   varState = ParsingState.ArrayClose

162               ElseIf c = strArrayRowSeparator Then
163                   varState = ParsingState.ArrayRowSeparator
164                   lngTokenIndex = i

165               ElseIf c = strArrayColumnSeparator Then
166                   varState = ParsingState.ArrayColumnSeparator
167                   lngTokenIndex = i

168               Else
169                   strError = "Expected " & strRightBrace & " " & strArrayRowSeparator & " " & ParsingState.ArrayColumnSeparator & " but got " & c & " at position " & i
170                   varState = ParsingState.ParsingError
171               End If

      ''' -------------------- -------------------- -------------------- '''

172           ElseIf varState = ParsingState.ArrayOpen Then
173               Set objToken = objTokens.NewToken(c, TokenType.ArrayOpen, i, 1)
174               objTokens.Add objToken
175               objTokenStack.Push objToken

176               varState = ParsingState.ArrayConstant1
177               i = i + 1

178           ElseIf varState = ParsingState.ArrayClose Then
179               Set objToken = objTokenStack.Pop
180               If objToken.TokenType = TokenType.ArrayOpen Then
181                   objTokens.Add objTokens.NewToken(c, TokenType.ArrayClose, i, 1)
182                   varState = ParsingState.Expression2
183                   i = i + 1
184               Else
185                   strError = "Encountered " & strRightBrace & " without matching " & strLeftBrace & " at position " & i
186                   varState = ParsingState.ParsingError
187               End If

188           ElseIf varState = ParsingState.ArrayColumnSeparator Then
189               i = i + 1
190               objTokens.Add objTokens.NewToken(c, TokenType.ArrayColumnSeparator, lngTokenIndex, i - lngTokenIndex)
191               varState = ParsingState.ArrayConstant1

192           ElseIf varState = ParsingState.ArrayRowSeparator Then
193               i = i + 1
194               objTokens.Add objTokens.NewToken(c, TokenType.ArrayRowSeparator, lngTokenIndex, i - lngTokenIndex)
195               varState = ParsingState.ArrayConstant1

      ''' -------------------- -------------------- -------------------- '''

196           ElseIf varState = ParsingState.Bool Then
197               str = str & c
198               lng = Len(str)
199               i = i + 1

200               str2 = Chr(0) & strBooleanTrue & Chr(0) & strBooleanFalse & Chr(0)

201               If InStr(1, str2, Chr(0) & str) > 0 Then
202                   If InStr(1, str2, Chr(0) & str & Chr(0)) > 0 Then
203                       objTokens.Add objTokens.NewToken(str, TokenType.Bool, lngTokenIndex, i - lngTokenIndex)
204                       str = ""
205                       varState = varReturnState
206                   End If
207               Else
208                   strError = "Expected Array Constant at position " & lngTokenIndex
209                   varState = ParsingState.ParsingError
210               End If

      ''' -------------------- -------------------- -------------------- '''

211           ElseIf varState = ParsingState.whitespace Then
212               If c = " " Or c = Chr(10) Then
213                   str = str & c
214                   i = i + 1
215               Else
216                   objTokens.Add objTokens.NewToken(str, TokenType.whitespace, lngTokenIndex, i - lngTokenIndex)
217                   str = ""
218                   varState = varReturnState

219               End If

      ''' -------------------- -------------------- -------------------- '''

220           ElseIf varState = ParsingState.SubExpression Then
221               Set objToken = objTokens.NewToken(c, TokenType.SubExpressionOpen, i, 1)
222               objTokens.Add objToken
223               objTokenStack.Push objToken

224               varState = ParsingState.Expression1
225               i = i + 1

226           ElseIf varState = ParsingState.BracketClose Then
227               Set objToken = objTokenStack.Pop
228               If objToken.TokenType = TokenType.FunctionOpen Then
229                   Set objToken = objTokens.NewToken(c, TokenType.FunctionClose, i, 1)
230               ElseIf objToken.TokenType = TokenType.SubExpressionOpen Then
231                   Set objToken = objTokens.NewToken(c, TokenType.SubExpressionClose, i, 1)
232               Else
233                   strError = "Encountered " & strRightRoundBracket & " without matching " & strLeftRoundBracket & " at position " & i
234                   varState = ParsingState.ParsingError
235               End If
236               objTokens.Add objToken

237               varState = ParsingState.Expression2
238               i = i + 1

      ''' -------------------- -------------------- -------------------- '''
      ''' Decide: Leading Name for Function, Table, R1C1Reference, Reference Qualifier or Cell Reference

239           ElseIf varState = ParsingState.LeadingName1 Then
240               If str <> "" And c = strLeftRoundBracket Then
241                   varState = ParsingState.FunctionX
242                   bln = True

243               ElseIf str <> "" And c = "[" Then                   'todo move symbol to variable
244                   varState = ParsingState.SquareBracketOpen

245               ElseIf str <> "" And c = "!" Then
246                   varState = ParsingState.ReferenceQualifier

247               ElseIf c = "" Or c = strRightRoundBracket Or c = strRightBrace Or _
                         c = " " Or c = Chr(10) Or _
                         c = "+" Or c = "-" Or c = "*" Or c = "/" Or c = "^" Or c = "%" Or _
                         c = "=" Or c = "<" Or c = ">" Or _
                         c = "&" Or c = ":" Or c = strListSeparator Then
248                   varState = ParsingState.LeadingNameE
249               Else
250                   str = str & c
251                   i = i + 1
252               End If

253           ElseIf varState = ParsingState.LeadingName2 Then
254               If c = "'" Then
255                   varState = ParsingState.LeadingName3
256               Else
257                   str = str & c
258               End If
259               i = i + 1

260           ElseIf varState = ParsingState.LeadingName3 Then
261               If c = "'" Then
262                   varState = ParsingState.LeadingName2   'Was escape sequence
263                   str = str & c
264                   i = i + 1
265               Else
266                   varState = ParsingState.LeadingName1
267               End If

268           ElseIf varState = ParsingState.LeadingNameE Then
269               If str <> "" Then
270                   If str = strBooleanTrue Or str = strBooleanFalse Then
271                       Set objToken = objTokens.NewToken(str, TokenType.Bool, lngTokenIndex, i - lngTokenIndex)
272                   Else
273                       Set objToken = objTokens.NewToken(str, TokenType.Reference, lngTokenIndex, i - lngTokenIndex)
274                   End If
275                   objTokens.Add objToken
276                   str = ""
277               End If

278               varState = ParsingState.Expression2

      ''' -------------------- -------------------- -------------------- '''
      ''' Function

279           ElseIf varState = ParsingState.FunctionX Then
280               If Left(str, 1) = "@" Then str = Mid(str, 2)
281               Set objToken = objTokens.NewToken(str, TokenType.FunctionOpen, lngTokenIndex, i - lngTokenIndex + 1)
282               objTokens.Add objToken
283               objTokenStack.Push objToken
284               str = ""
285               varState = ParsingState.Expression1
286               i = i + 1

      ''' -------------------- -------------------- -------------------- '''
      ''' Decide: Table or R1C1 Reference

287           ElseIf varState = ParsingState.SquareBracketOpen Then
288               If str = "R" Or str = "C" Then                          'todo: use International characters
289                   varState = ParsingState.R1C1Reference
290               Else
291                   Set objToken = objTokens.NewToken(str, TokenType.TableOpen, lngTokenIndex, i - lngTokenIndex + 1)
292                   objTokens.Add objToken
293                   objTokenStack.Push objToken
294                   str = ""
295                   varState = ParsingState.Table1
296                   i = i + 1
297               End If

      ''' -------------------- -------------------- -------------------- '''
      ''' Table

          'todo: escape character

298           ElseIf varState = ParsingState.Table1 Then
299               lngTokenIndex = i
300               If c = " " Then
301                   varReturnState = ParsingState.Table1
302                   varState = ParsingState.whitespace
303               ElseIf c = "[" Then                                     'todo, possibly use international bracket
304                   i = i + 1
305                   varState = ParsingState.Table6
306               Else
307                   varState = ParsingState.Table5
308               End If

309           ElseIf varState = ParsingState.Table2 Then
310               lngTokenIndex = i
311               If c = " " Then
312                   varReturnState = ParsingState.Table2
313                   varState = ParsingState.whitespace
314               ElseIf c = "," Then                                     'todo possibly use list separator?
315                   varState = ParsingState.Table3
316               ElseIf c = ":" Then
317                   varState = ParsingState.Table4
318               ElseIf c = "]" Then                                     'todo, possibly use international bracket
319                   varState = ParsingState.TableE
320               Else
321                   varState = ParsingState.Table5
322               End If

323           ElseIf varState = ParsingState.Table3 Then
324               i = i + 1

325               objTokens.Add objTokens.NewToken(c, TokenType.TableItemSeparator, lngTokenIndex, i - lngTokenIndex)
326               str = ""
327               varState = ParsingState.Table1

328           ElseIf varState = ParsingState.Table4 Then
329               i = i + 1

                  'todo: decide whether to make this a tablecolumns token type, or leave separated
330               objTokens.Add objTokens.NewToken(c, TokenType.TableColumnSeparator, lngTokenIndex, i - lngTokenIndex)
331               str = ""
332               varState = ParsingState.Table1

333           ElseIf varState = ParsingState.Table5 Then
334               If c = "]" Then                             'todo, possibly use international bracket
335                   If Left(str, 1) = "#" Then
336                       objTokens.Add objTokens.NewToken(str, TokenType.TableSection, lngTokenIndex, i - lngTokenIndex)
337                   Else
338                       objTokens.Add objTokens.NewToken(str, TokenType.TableColumn, lngTokenIndex, i - lngTokenIndex)
339                   End If
340                   str = ""
341                   varState = ParsingState.TableE
342               ElseIf c = "'" Then
343                   i = i + 1
344                   varReturnState = ParsingState.Table5
345                   varState = ParsingState.TableQ
346               Else
347                   str = str & c
348                   i = i + 1
349               End If

350           ElseIf varState = ParsingState.Table6 Then
351               If c = "]" Then                             'todo, possibly use international bracket
352                   i = i + 1
353                   If Left(str, 1) = "#" Then
354                       objTokens.Add objTokens.NewToken(str, TokenType.TableSection, lngTokenIndex, i - lngTokenIndex)
355                   Else
356                       objTokens.Add objTokens.NewToken(str, TokenType.TableColumn, lngTokenIndex, i - lngTokenIndex)
357                   End If
358                   str = ""
359                   varState = ParsingState.Table2
360               ElseIf c = "'" Then
361                   i = i + 1
362                   varReturnState = ParsingState.Table6
363                   varState = ParsingState.TableQ
364               Else
365                   str = str & c
366                   i = i + 1
367               End If

368           ElseIf varState = ParsingState.TableQ Then
369               str = str & c
370               i = i + 1
371               varState = varReturnState

372           ElseIf varState = ParsingState.TableE Then
373               Set objToken = objTokenStack.Pop
374               If objToken.TokenType = TokenType.TableOpen Then
375                   objTokens.Add objTokens.NewToken(c, TokenType.TableClose, i, 1)
376                   varState = ParsingState.Expression2
377                   i = i + 1
378               Else
379                   strError = "Encountered " & "]" & " without matching " & "[" & " at position " & i 'todo, possibly use international bracket
380                   varState = ParsingState.ParsingError
381               End If


      ''' -------------------- -------------------- -------------------- '''
      ''' Text

382           ElseIf varState = ParsingState.Text1 Then
383               If c = strDoubleQuote Then
384                   varState = ParsingState.Text2
385               Else
386                   str = str & c
387               End If
388               i = i + 1

389           ElseIf varState = ParsingState.Text2 Then
390               If c = strDoubleQuote Then
391                   varState = ParsingState.Text1   'Was escape sequence
392                   str = str & c
393                   i = i + 1
394               Else
395                   objTokens.Add objTokens.NewToken(str, TokenType.Text, lngTokenIndex, i - lngTokenIndex)
396                   str = ""
397                   varState = varReturnState
398               End If

      ''' -------------------- -------------------- -------------------- '''
      ''' Number

399           ElseIf varState = ParsingState.Number1 Or varState = ParsingState.Number2 Then
400               If c = "0" Or c = "1" Or c = "2" Or c = "3" Or c = "4" Or _
                     c = "5" Or c = "6" Or c = "7" Or c = "8" Or c = "9" Then
401                   str = str & c
402                   i = i + 1

403               ElseIf c = strDecimalSeparator And varState = ParsingState.Number1 Then
404                   str = str & c
405                   varState = ParsingState.Number2
406                   i = i + 1

407               ElseIf c = "E" Then
408                   str = str & c
409                   varState = ParsingState.Number3
410                   i = i + 1

411               Else
412                   varState = ParsingState.NumberE
413               End If

414           ElseIf varState = ParsingState.Number3 Then
415               If c = "+" Or c = "-" Then
416                   str = str & c
417                   varState = ParsingState.Number4
418                   i = i + 1
419               Else
420                   strError = "Expected + or - at position " & i
421                   varState = ParsingState.ParsingError
422               End If

423           ElseIf varState = ParsingState.Number4 Then
424               If c = "0" Or c = "1" Or c = "2" Or c = "3" Or c = "4" Or _
                     c = "5" Or c = "6" Or c = "7" Or c = "8" Or c = "9" Then
425                   str = str & c
426                   i = i + 1
427               Else
428                   varState = ParsingState.NumberE
429               End If

430           ElseIf varState = ParsingState.NumberE Then
431               objTokens.Add objTokens.NewToken(str, TokenType.Number, lngTokenIndex, i - lngTokenIndex)
432               str = ""
433               varState = varReturnState

      ''' -------------------- -------------------- -------------------- '''
      ''' Error

434           ElseIf varState = ParsingState.ErrorX Then
435               str = str & c
436               lng = Len(str)
437               i = i + 1

438               str2 = Chr(0) & strErrorRef & Chr(0) & strErrorDiv0 & Chr(0) & strErrorNA & Chr(0) & strErrorName & _
                         Chr(0) & strErrorNull & Chr(0) & strErrorNum & Chr(0) & strErrorValue & Chr(0) & strErrorGettingData & Chr(0)

439               If InStr(1, str2, Chr(0) & str) > 0 Then
440                   If InStr(1, str2, Chr(0) & str & Chr(0)) > 0 Then
441                       objTokens.Add objTokens.NewToken(str, TokenType.ErrorText, lngTokenIndex, i - lngTokenIndex)
442                       str = ""
443                       varState = varReturnState
444                   End If
445               Else
446                   strError = "Expected Error Constant at position " & lngTokenIndex
447                   varState = ParsingState.ParsingError
448               End If

      ''' -------------------- -------------------- -------------------- '''

449           ElseIf varState = ParsingState.ReferenceQualifier Then
450               objTokens.Add objTokens.NewToken(str, TokenType.ReferenceQualifier, lngTokenIndex, i - lngTokenIndex)
451               objTokens.Add objTokens.NewToken(c, TokenType.ExternalReferenceOperator, i, 1)

452               i = i + 1
453               str = ""
454               varState = ParsingState.Expression1

455           ElseIf varState = ParsingState.MinusSign Then
456               If c = "0" Or c = "1" Or c = "2" Or c = "3" Or c = "4" Or _
                     c = "5" Or c = "6" Or c = "7" Or c = "8" Or c = "9" Then
457                   varState = ParsingState.Number1
458               Else
459                   varState = ParsingState.PrefixOperator
460                   i = i - 1
461               End If

462           ElseIf varState = ParsingState.PrefixOperator Then
463               i = i + 1

464               objTokens.Add objTokens.NewToken(c, TokenType.UnaryOperator, lngTokenIndex, i - lngTokenIndex)
465               str = ""
466               varState = ParsingState.Expression1

      ''' -------------------- -------------------- -------------------- '''

              'ListSeparator can be either function parameter separator, or a union operator
467           ElseIf varState = ParsingState.ListSeparator Then
468               bln = False
469               If objTokenStack.Count > 0 Then
470                   If objTokenStack.Peek.TokenType = TokenType.FunctionOpen Then
471                       varState = ParsingState.ParameterSeparator
472                       bln = True
473                   End If
474               End If
475               If Not bln Then varState = ParsingState.RangeOperator

476           ElseIf varState = ParsingState.RangeOperator Then
477               str = c
478               i = i + 1

479               objTokens.Add objTokens.NewToken(str, TokenType.RangeOperator, lngTokenIndex, i - lngTokenIndex)
480               str = ""
481               varState = ParsingState.Expression1

      ''' -------------------- -------------------- -------------------- '''

482           ElseIf varState = ParsingState.ArithmeticOperator Then
483               str = c
484               i = i + 1

485               objTokens.Add objTokens.NewToken(c, TokenType.ArithmeticOperator, lngTokenIndex, i - lngTokenIndex)
486               str = ""
487               varState = ParsingState.Expression1

      ''' -------------------- -------------------- -------------------- '''

488           ElseIf varState = ParsingState.ComparisonOperator1 Then
489               str = c
490               i = i + 1

491               If c = "<" Or c = ">" Then
492                   varState = ParsingState.ComparisonOperator2
493               Else
494                   varState = ParsingState.ComparisonOperatorE
495               End If

496           ElseIf varState = ParsingState.ComparisonOperator2 Then
497               If c = "=" Or (str = "<" And c = ">") Then
498                   str = str & c
499                   i = i + 1
500               End If

501               varState = ParsingState.ComparisonOperatorE

502           ElseIf varState = ParsingState.ComparisonOperatorE Then
503               objTokens.Add objTokens.NewToken(str, TokenType.ComparisonOperator, lngTokenIndex, i - lngTokenIndex)
504               str = ""
505               varState = ParsingState.Expression1

      ''' -------------------- -------------------- -------------------- '''

506           ElseIf varState = ParsingState.TextOperator Then
507               str = c
508               i = i + 1

509               objTokens.Add objTokens.NewToken(c, TokenType.TextOperator, lngTokenIndex, i - lngTokenIndex)
510               str = ""
511               varState = ParsingState.Expression1

      ''' -------------------- -------------------- -------------------- '''

512           ElseIf varState = ParsingState.PostfixOperator Then
513               str = c
514               i = i + 1

515               objTokens.Add objTokens.NewToken(str, TokenType.PostfixOperator, lngTokenIndex, i - lngTokenIndex)
516               str = ""
517               varState = ParsingState.Expression2

      ''' -------------------- -------------------- -------------------- '''

518           ElseIf varState = ParsingState.ParameterSeparator Then
519               str = c
520               i = i + 1

521               objTokens.Add objTokens.NewToken(str, TokenType.ParameterSeparator, lngTokenIndex, i - lngTokenIndex)
522               str = ""
523               varState = ParsingState.Expression1

524           End If

525           lngPrevIndex = i
526       Loop While blnLoop

          'todo: error if token stack not empty

          'post processing
          '1. Scan and detect !, and join if the 3 tokens prior are CellReference-Colon-CellReference
          '2. Scan and detect CellReference-Colon-CellReference, because they should be joined into an AreaReference

527 e:      Set ParseFormula = objTokens
End Function

