Attribute VB_Name = "JsonConverter"
''
' VBA-JSON v2.2.3
' (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
'
' JSON Converter for VBA
'
' Errors:
' 10001 - JSON parse error
'
' @class JsonConverter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Based originally on vba-json (with extensive changes)
' BSD license included below
'
' JSONLib, http://code.google.com/p/vba-json/
'
' Copyright (c) 2013, Ryo Yokoyama
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' === VBA-UTC Headers
#If Mac Then

#If VBA7 Then

' 64-bit Mac (2016)
Private Declare PtrSafe Function utc_popen Lib "libc.dylib" Alias "popen" _
    (ByVal utc_Command As String, ByVal utc_Mode As String) As LongPtr
Private Declare PtrSafe Function utc_pclose Lib "libc.dylib" Alias "pclose" _
    (ByVal utc_File As Long) As LongPtr
Private Declare PtrSafe Function utc_fread Lib "libc.dylib" Alias "fread" _
    (ByVal utc_Buffer As String, ByVal utc_Size As LongPtr, ByVal utc_Number As LongPtr, ByVal utc_File As LongPtr) As LongPtr
Private Declare PtrSafe Function utc_feof Lib "libc.dylib" Alias "feof" _
    (ByVal utc_File As LongPtr) As LongPtr

#Else

' 32-bit Mac
Private Declare Function utc_popen Lib "libc.dylib" Alias "popen" _
    (ByVal utc_Command As String, ByVal utc_Mode As String) As Long
Private Declare Function utc_pclose Lib "libc.dylib" Alias "pclose" _
    (ByVal utc_File As Long) As Long
Private Declare Function utc_fread Lib "libc.dylib" Alias "fread" _
    (ByVal utc_Buffer As String, ByVal utc_Size As Long, ByVal utc_Number As Long, ByVal utc_File As Long) As Long
Private Declare Function utc_feof Lib "libc.dylib" Alias "feof" _
    (ByVal utc_File As Long) As Long

#End If

#ElseIf VBA7 Then

' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724421.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724949.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms725485.aspx
Private Declare PtrSafe Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare PtrSafe Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare PtrSafe Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

#Else

Private Declare Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

#End If

#If Mac Then

#If VBA7 Then
Private Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As LongPtr
End Type

#Else

Private Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As Long
End Type

#End If

#Else

Private Type utc_SYSTEMTIME
    utc_wYear As Integer
    utc_wMonth As Integer
    utc_wDayOfWeek As Integer
    utc_wDay As Integer
    utc_wHour As Integer
    utc_wMinute As Integer
    utc_wSecond As Integer
    utc_wMilliseconds As Integer
End Type

Private Type utc_TIME_ZONE_INFORMATION
    utc_Bias As Long
    utc_StandardName(0 To 31) As Integer
    utc_StandardDate As utc_SYSTEMTIME
    utc_StandardBias As Long
    utc_DaylightName(0 To 31) As Integer
    utc_DaylightDate As utc_SYSTEMTIME
    utc_DaylightBias As Long
End Type

#End If
' === End VBA-UTC

#If Mac Then
#ElseIf VBA7 Then

Private Declare PtrSafe Sub json_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (json_MemoryDestination As Any, json_MemorySource As Any, ByVal json_ByteLength As Long)

#Else

Private Declare Sub json_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (json_MemoryDestination As Any, json_MemorySource As Any, ByVal json_ByteLength As Long)

#End If

Private Type json_Options
    ' VBA only stores 15 significant digits, so any numbers larger than that are truncated
    ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
    ' See: http://support.microsoft.com/kb/269370
    '
    ' By default, VBA-JSON will use String for numbers longer than 15 characters that contain only digits
    ' to override set `JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True`
    UseDoubleForLargeNumbers As Boolean

    ' The JSON standard requires object keys to be quoted (" or '), use this option to allow unquoted keys
    AllowUnquotedKeys As Boolean

    ' The solidus (/) is not required to be escaped, use this option to escape them as \/ in ConvertToJson
    EscapeSolidus As Boolean
End Type
Public JsonOptions As json_Options

' ============================================= '
' Public Methods
' ============================================= '

''
' Convert JSON string to object (Dictionary/Collection)
'
' @method ParseJson
' @param {String} json_String
' @return {Object} (Dictionary or Collection)
' @throws 10001 - JSON parse error
''
Public Function ParseJson(ByVal JsonString As String) As Object
          Dim json_Index As Long
1         json_Index = 1

          ' Remove vbCr, vbLf, and vbTab from json_String
2         JsonString = VBA.Replace(VBA.Replace(VBA.Replace(JsonString, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")

3         json_SkipSpaces JsonString, json_Index
4         Select Case VBA.Mid$(JsonString, json_Index, 1)
          Case "{"
5             Set ParseJson = json_ParseObject(JsonString, json_Index)
6         Case "["
7             Set ParseJson = json_ParseArray(JsonString, json_Index)
8         Case Else
              ' Error: Invalid JSON string
9             Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(JsonString, json_Index, "Expecting '{' or '['")
10        End Select
End Function

''
' Convert object (Dictionary/Collection/Array) to JSON
'
' @method ConvertToJson
' @param {Variant} JsonValue (Dictionary, Collection, or Array)
' @param {Integer|String} Whitespace "Pretty" print json with given number of spaces per indentation (Integer) or given string
' @return {String}
''
Public Function ConvertToJson(ByVal JsonValue As Variant, Optional ByVal whitespace As Variant, Optional ByVal json_CurrentIndentation As Long = 0) As String
          Dim json_buffer As String
          Dim json_BufferPosition As Long
          Dim json_BufferLength As Long
          Dim json_Index As Long
          Dim json_LBound As Long
          Dim json_UBound As Long
          Dim json_IsFirstItem As Boolean
          Dim json_Index2D As Long
          Dim json_LBound2D As Long
          Dim json_UBound2D As Long
          Dim json_IsFirstItem2D As Boolean
          Dim json_Key As Variant
          Dim json_Value As Variant
          Dim json_DateStr As String
          Dim json_Converted As String
          Dim json_SkipItem As Boolean
          Dim json_PrettyPrint As Boolean
          Dim json_Indentation As String
          Dim json_InnerIndentation As String

1         json_LBound = -1
2         json_UBound = -1
3         json_IsFirstItem = True
4         json_LBound2D = -1
5         json_UBound2D = -1
6         json_IsFirstItem2D = True
7         json_PrettyPrint = Not IsMissing(whitespace)

8         Select Case VBA.VarType(JsonValue)
          Case VBA.vbNull
9             ConvertToJson = "null"
10        Case VBA.vbDate
              ' Date
11            json_DateStr = ConvertToIso(VBA.CDate(JsonValue))

12            ConvertToJson = """" & json_DateStr & """"
13        Case VBA.vbString
              ' String (or large number encoded as string)
14            If Not JsonOptions.UseDoubleForLargeNumbers And json_StringIsLargeNumber(JsonValue) Then
15                ConvertToJson = JsonValue
16            Else
17                ConvertToJson = """" & json_Encode(JsonValue) & """"
18            End If
19        Case VBA.vbBoolean
20            If JsonValue Then
21                ConvertToJson = "true"
22            Else
23                ConvertToJson = "false"
24            End If
25        Case VBA.vbArray To VBA.vbArray + VBA.vbByte
26            If json_PrettyPrint Then
27                If VBA.VarType(whitespace) = VBA.vbString Then
28                    json_Indentation = VBA.String$(json_CurrentIndentation + 1, whitespace)
29                    json_InnerIndentation = VBA.String$(json_CurrentIndentation + 2, whitespace)
30                Else
31                    json_Indentation = VBA.Space$((json_CurrentIndentation + 1) * whitespace)
32                    json_InnerIndentation = VBA.Space$((json_CurrentIndentation + 2) * whitespace)
33                End If
34            End If

              ' Array
35            json_BufferAppend json_buffer, "[", json_BufferPosition, json_BufferLength

36            On Error Resume Next

37            json_LBound = LBound(JsonValue, 1)
38            json_UBound = UBound(JsonValue, 1)
39            json_LBound2D = LBound(JsonValue, 2)
40            json_UBound2D = UBound(JsonValue, 2)

41            If json_LBound >= 0 And json_UBound >= 0 Then
42                For json_Index = json_LBound To json_UBound
43                    If json_IsFirstItem Then
44                        json_IsFirstItem = False
45                    Else
                          ' Append comma to previous line
46                        json_BufferAppend json_buffer, ",", json_BufferPosition, json_BufferLength
47                    End If

48                    If json_LBound2D >= 0 And json_UBound2D >= 0 Then
                          ' 2D Array
49                        If json_PrettyPrint Then
50                            json_BufferAppend json_buffer, vbNewLine, json_BufferPosition, json_BufferLength
51                        End If
52                        json_BufferAppend json_buffer, json_Indentation & "[", json_BufferPosition, json_BufferLength

53                        For json_Index2D = json_LBound2D To json_UBound2D
54                            If json_IsFirstItem2D Then
55                                json_IsFirstItem2D = False
56                            Else
57                                json_BufferAppend json_buffer, ",", json_BufferPosition, json_BufferLength
58                            End If

59                            json_Converted = ConvertToJson(JsonValue(json_Index, json_Index2D), whitespace, json_CurrentIndentation + 2)

                              ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
60                            If json_Converted = "" Then
                                  ' (nest to only check if converted = "")
61                                If json_IsUndefined(JsonValue(json_Index, json_Index2D)) Then
62                                    json_Converted = "null"
63                                End If
64                            End If

65                            If json_PrettyPrint Then
66                                json_Converted = vbNewLine & json_InnerIndentation & json_Converted
67                            End If

68                            json_BufferAppend json_buffer, json_Converted, json_BufferPosition, json_BufferLength
69                        Next json_Index2D

70                        If json_PrettyPrint Then
71                            json_BufferAppend json_buffer, vbNewLine, json_BufferPosition, json_BufferLength
72                        End If

73                        json_BufferAppend json_buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength
74                        json_IsFirstItem2D = True
75                    Else
                          ' 1D Array
76                        json_Converted = ConvertToJson(JsonValue(json_Index), whitespace, json_CurrentIndentation + 1)

                          ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
77                        If json_Converted = "" Then
                              ' (nest to only check if converted = "")
78                            If json_IsUndefined(JsonValue(json_Index)) Then
79                                json_Converted = "null"
80                            End If
81                        End If

82                        If json_PrettyPrint Then
83                            json_Converted = vbNewLine & json_Indentation & json_Converted
84                        End If

85                        json_BufferAppend json_buffer, json_Converted, json_BufferPosition, json_BufferLength
86                    End If
87                Next json_Index
88            End If

89            On Error GoTo 0

90            If json_PrettyPrint Then
91                json_BufferAppend json_buffer, vbNewLine, json_BufferPosition, json_BufferLength

92                If VBA.VarType(whitespace) = VBA.vbString Then
93                    json_Indentation = VBA.String$(json_CurrentIndentation, whitespace)
94                Else
95                    json_Indentation = VBA.Space$(json_CurrentIndentation * whitespace)
96                End If
97            End If

98            json_BufferAppend json_buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength

99            ConvertToJson = json_BufferToString(json_buffer, json_BufferPosition, json_BufferLength)

          ' Dictionary or Collection
100       Case VBA.vbObject
101           If json_PrettyPrint Then
102               If VBA.VarType(whitespace) = VBA.vbString Then
103                   json_Indentation = VBA.String$(json_CurrentIndentation + 1, whitespace)
104               Else
105                   json_Indentation = VBA.Space$((json_CurrentIndentation + 1) * whitespace)
106               End If
107           End If

              ' Dictionary
108           If VBA.TypeName(JsonValue) = "Dictionary" Then
109               json_BufferAppend json_buffer, "{", json_BufferPosition, json_BufferLength
110               For Each json_Key In JsonValue.Keys
                      ' For Objects, undefined (Empty/Nothing) is not added to object
111                   json_Converted = ConvertToJson(JsonValue(json_Key), whitespace, json_CurrentIndentation + 1)
112                   If json_Converted = "" Then
113                       json_SkipItem = json_IsUndefined(JsonValue(json_Key))
114                   Else
115                       json_SkipItem = False
116                   End If

117                   If Not json_SkipItem Then
118                       If json_IsFirstItem Then
119                           json_IsFirstItem = False
120                       Else
121                           json_BufferAppend json_buffer, ",", json_BufferPosition, json_BufferLength
122                       End If

123                       If json_PrettyPrint Then
124                           json_Converted = vbNewLine & json_Indentation & """" & json_Key & """: " & json_Converted
125                       Else
126                           json_Converted = """" & json_Key & """:" & json_Converted
127                       End If

128                       json_BufferAppend json_buffer, json_Converted, json_BufferPosition, json_BufferLength
129                   End If
130               Next json_Key

131               If json_PrettyPrint Then
132                   json_BufferAppend json_buffer, vbNewLine, json_BufferPosition, json_BufferLength

133                   If VBA.VarType(whitespace) = VBA.vbString Then
134                       json_Indentation = VBA.String$(json_CurrentIndentation, whitespace)
135                   Else
136                       json_Indentation = VBA.Space$(json_CurrentIndentation * whitespace)
137                   End If
138               End If

139               json_BufferAppend json_buffer, json_Indentation & "}", json_BufferPosition, json_BufferLength

              ' Collection
140           ElseIf VBA.TypeName(JsonValue) = "Collection" Then
141               json_BufferAppend json_buffer, "[", json_BufferPosition, json_BufferLength
142               For Each json_Value In JsonValue
143                   If json_IsFirstItem Then
144                       json_IsFirstItem = False
145                   Else
146                       json_BufferAppend json_buffer, ",", json_BufferPosition, json_BufferLength
147                   End If

148                   json_Converted = ConvertToJson(json_Value, whitespace, json_CurrentIndentation + 1)

                      ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
149                   If json_Converted = "" Then
                          ' (nest to only check if converted = "")
150                       If json_IsUndefined(json_Value) Then
151                           json_Converted = "null"
152                       End If
153                   End If

154                   If json_PrettyPrint Then
155                       json_Converted = vbNewLine & json_Indentation & json_Converted
156                   End If

157                   json_BufferAppend json_buffer, json_Converted, json_BufferPosition, json_BufferLength
158               Next json_Value

159               If json_PrettyPrint Then
160                   json_BufferAppend json_buffer, vbNewLine, json_BufferPosition, json_BufferLength

161                   If VBA.VarType(whitespace) = VBA.vbString Then
162                       json_Indentation = VBA.String$(json_CurrentIndentation, whitespace)
163                   Else
164                       json_Indentation = VBA.Space$(json_CurrentIndentation * whitespace)
165                   End If
166               End If

167               json_BufferAppend json_buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength
168           End If

169           ConvertToJson = json_BufferToString(json_buffer, json_BufferPosition, json_BufferLength)
170       Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
              ' Number (use decimals for numbers)
171           ConvertToJson = VBA.Replace(JsonValue, ",", ".")
172       Case Else
              ' vbEmpty, vbError, vbDataObject, vbByte, vbUserDefinedType
              ' Use VBA's built-in to-string
173           On Error Resume Next
174           ConvertToJson = JsonValue
175           On Error GoTo 0
176       End Select
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Function json_ParseObject(json_String As String, ByRef json_Index As Long) As Dictionary
          Dim json_Key As String
          Dim json_NextChar As String

1         Set json_ParseObject = New Dictionary
2         json_SkipSpaces json_String, json_Index
3         If VBA.Mid$(json_String, json_Index, 1) <> "{" Then
4             Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '{'")
5         Else
6             json_Index = json_Index + 1

7             Do
8                 json_SkipSpaces json_String, json_Index
9                 If VBA.Mid$(json_String, json_Index, 1) = "}" Then
10                    json_Index = json_Index + 1
11                    Exit Function
12                ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
13                    json_Index = json_Index + 1
14                    json_SkipSpaces json_String, json_Index
15                End If

16                json_Key = json_ParseKey(json_String, json_Index)
17                json_NextChar = json_Peek(json_String, json_Index)
18                If json_NextChar = "[" Or json_NextChar = "{" Then
19                    Set json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index)
20                Else
21                    json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index)
22                End If
23            Loop
24        End If
End Function

Private Function json_ParseArray(json_String As String, ByRef json_Index As Long) As Collection
1         Set json_ParseArray = New Collection

2         json_SkipSpaces json_String, json_Index
3         If VBA.Mid$(json_String, json_Index, 1) <> "[" Then
4             Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '['")
5         Else
6             json_Index = json_Index + 1

7             Do
8                 json_SkipSpaces json_String, json_Index
9                 If VBA.Mid$(json_String, json_Index, 1) = "]" Then
10                    json_Index = json_Index + 1
11                    Exit Function
12                ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
13                    json_Index = json_Index + 1
14                    json_SkipSpaces json_String, json_Index
15                End If

16                json_ParseArray.Add json_ParseValue(json_String, json_Index)
17            Loop
18        End If
End Function

Private Function json_ParseValue(json_String As String, ByRef json_Index As Long) As Variant
1         json_SkipSpaces json_String, json_Index
2         Select Case VBA.Mid$(json_String, json_Index, 1)
          Case "{"
3             Set json_ParseValue = json_ParseObject(json_String, json_Index)
4         Case "["
5             Set json_ParseValue = json_ParseArray(json_String, json_Index)
6         Case """", "'"
7             json_ParseValue = json_ParseString(json_String, json_Index)
8         Case Else
9             If VBA.Mid$(json_String, json_Index, 4) = "true" Then
10                json_ParseValue = True
11                json_Index = json_Index + 4
12            ElseIf VBA.Mid$(json_String, json_Index, 5) = "false" Then
13                json_ParseValue = False
14                json_Index = json_Index + 5
15            ElseIf VBA.Mid$(json_String, json_Index, 4) = "null" Then
16                json_ParseValue = Null
17                json_Index = json_Index + 4
18            ElseIf VBA.InStr("+-0123456789", VBA.Mid$(json_String, json_Index, 1)) Then
19                json_ParseValue = json_ParseNumber(json_String, json_Index)
20            Else
21                Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
22            End If
23        End Select
End Function

Private Function json_ParseString(json_String As String, ByRef json_Index As Long) As String
          Dim json_Quote As String
          Dim json_Char As String
          Dim json_Code As String
          Dim json_buffer As String
          Dim json_BufferPosition As Long
          Dim json_BufferLength As Long

1         json_SkipSpaces json_String, json_Index

          ' Store opening quote to look for matching closing quote
2         json_Quote = VBA.Mid$(json_String, json_Index, 1)
3         json_Index = json_Index + 1

4         Do While json_Index > 0 And json_Index <= Len(json_String)
5             json_Char = VBA.Mid$(json_String, json_Index, 1)

6             Select Case json_Char
              Case "\"
                  ' Escaped string, \\, or \/
7                 json_Index = json_Index + 1
8                 json_Char = VBA.Mid$(json_String, json_Index, 1)

9                 Select Case json_Char
                  Case """", "\", "/", "'"
10                    json_BufferAppend json_buffer, json_Char, json_BufferPosition, json_BufferLength
11                    json_Index = json_Index + 1
12                Case "b"
13                    json_BufferAppend json_buffer, vbBack, json_BufferPosition, json_BufferLength
14                    json_Index = json_Index + 1
15                Case "f"
16                    json_BufferAppend json_buffer, vbFormFeed, json_BufferPosition, json_BufferLength
17                    json_Index = json_Index + 1
18                Case "n"
19                    json_BufferAppend json_buffer, vbCrLf, json_BufferPosition, json_BufferLength
20                    json_Index = json_Index + 1
21                Case "r"
22                    json_BufferAppend json_buffer, vbCr, json_BufferPosition, json_BufferLength
23                    json_Index = json_Index + 1
24                Case "t"
25                    json_BufferAppend json_buffer, vbTab, json_BufferPosition, json_BufferLength
26                    json_Index = json_Index + 1
27                Case "u"
                      ' Unicode character escape (e.g. \u00a9 = Copyright)
28                    json_Index = json_Index + 1
29                    json_Code = VBA.Mid$(json_String, json_Index, 4)
30                    json_BufferAppend json_buffer, VBA.ChrW(VBA.Val("&h" + json_Code)), json_BufferPosition, json_BufferLength
31                    json_Index = json_Index + 4
32                End Select
33            Case json_Quote
34                json_ParseString = json_BufferToString(json_buffer, json_BufferPosition, json_BufferLength)
35                json_Index = json_Index + 1
36                Exit Function
37            Case Else
38                json_BufferAppend json_buffer, json_Char, json_BufferPosition, json_BufferLength
39                json_Index = json_Index + 1
40            End Select
41        Loop
End Function

Private Function json_ParseNumber(json_String As String, ByRef json_Index As Long) As Variant
          Dim json_Char As String
          Dim json_Value As String
          Dim json_IsLargeNumber As Boolean

1         json_SkipSpaces json_String, json_Index

2         Do While json_Index > 0 And json_Index <= Len(json_String)
3             json_Char = VBA.Mid$(json_String, json_Index, 1)

4             If VBA.InStr("+-0123456789.eE", json_Char) Then
                  ' Unlikely to have massive number, so use simple append rather than buffer here
5                 json_Value = json_Value & json_Char
6                 json_Index = json_Index + 1
7             Else
                  ' Excel only stores 15 significant digits, so any numbers larger than that are truncated
                  ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
                  ' See: http://support.microsoft.com/kb/269370
                  '
                  ' Fix: Parse -> String, Convert -> String longer than 15/16 characters containing only numbers and decimal points -> Number
                  ' (decimal doesn't factor into significant digit count, so if present check for 15 digits + decimal = 16)
8                 json_IsLargeNumber = IIf(InStr(json_Value, "."), Len(json_Value) >= 17, Len(json_Value) >= 16)
9                 If Not JsonOptions.UseDoubleForLargeNumbers And json_IsLargeNumber Then
10                    json_ParseNumber = json_Value
11                Else
                      ' VBA.Val does not use regional settings, so guard for comma is not needed
12                    json_ParseNumber = VBA.Val(json_Value)
13                End If
14                Exit Function
15            End If
16        Loop
End Function

Private Function json_ParseKey(json_String As String, ByRef json_Index As Long) As String
          ' Parse key with single or double quotes
1         If VBA.Mid$(json_String, json_Index, 1) = """" Or VBA.Mid$(json_String, json_Index, 1) = "'" Then
2             json_ParseKey = json_ParseString(json_String, json_Index)
3         ElseIf JsonOptions.AllowUnquotedKeys Then
              Dim json_Char As String
4             Do While json_Index > 0 And json_Index <= Len(json_String)
5                 json_Char = VBA.Mid$(json_String, json_Index, 1)
6                 If (json_Char <> " ") And (json_Char <> ":") Then
7                     json_ParseKey = json_ParseKey & json_Char
8                     json_Index = json_Index + 1
9                 Else
10                    Exit Do
11                End If
12            Loop
13        Else
14            Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '""' or '''")
15        End If

          ' Check for colon and skip if present or throw if not present
16        json_SkipSpaces json_String, json_Index
17        If VBA.Mid$(json_String, json_Index, 1) <> ":" Then
18            Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting ':'")
19        Else
20            json_Index = json_Index + 1
21        End If
End Function

Private Function json_IsUndefined(ByVal json_Value As Variant) As Boolean
          ' Empty / Nothing -> undefined
1         Select Case VBA.VarType(json_Value)
          Case VBA.vbEmpty
2             json_IsUndefined = True
3         Case VBA.vbObject
4             Select Case VBA.TypeName(json_Value)
              Case "Empty", "Nothing"
5                 json_IsUndefined = True
6             End Select
7         End Select
End Function

Private Function json_Encode(ByVal json_Text As Variant) As String
          ' Reference: http://www.ietf.org/rfc/rfc4627.txt
          ' Escape: ", \, /, backspace, form feed, line feed, carriage return, tab
          Dim json_Index As Long
          Dim json_Char As String
          Dim json_AscCode As Long
          Dim json_buffer As String
          Dim json_BufferPosition As Long
          Dim json_BufferLength As Long

1         For json_Index = 1 To VBA.Len(json_Text)
2             json_Char = VBA.Mid$(json_Text, json_Index, 1)
3             json_AscCode = VBA.AscW(json_Char)

              ' When AscW returns a negative number, it returns the twos complement form of that number.
              ' To convert the twos complement notation into normal binary notation, add 0xFFF to the return result.
              ' https://support.microsoft.com/en-us/kb/272138
4             If json_AscCode < 0 Then
5                 json_AscCode = json_AscCode + 65536
6             End If

              ' From spec, ", \, and control characters must be escaped (solidus is optional)

7             Select Case json_AscCode
              Case 34
                  ' " -> 34 -> \"
8                 json_Char = "\"""
9             Case 92
                  ' \ -> 92 -> \\
10                json_Char = "\\"
11            Case 47
                  ' / -> 47 -> \/ (optional)
12                If JsonOptions.EscapeSolidus Then
13                    json_Char = "\/"
14                End If
15            Case 8
                  ' backspace -> 8 -> \b
16                json_Char = "\b"
17            Case 12
                  ' form feed -> 12 -> \f
18                json_Char = "\f"
19            Case 10
                  ' line feed -> 10 -> \n
20                json_Char = "\n"
21            Case 13
                  ' carriage return -> 13 -> \r
22                json_Char = "\r"
23            Case 9
                  ' tab -> 9 -> \t
24                json_Char = "\t"
25            Case 0 To 31, 127 To 65535
                  ' Non-ascii characters -> convert to 4-digit hex
26                json_Char = "\u" & VBA.Right$("0000" & VBA.Hex$(json_AscCode), 4)
27            End Select

28            json_BufferAppend json_buffer, json_Char, json_BufferPosition, json_BufferLength
29        Next json_Index

30        json_Encode = json_BufferToString(json_buffer, json_BufferPosition, json_BufferLength)
End Function

Private Function json_Peek(json_String As String, ByVal json_Index As Long, Optional json_NumberOfCharacters As Long = 1) As String
          ' "Peek" at the next number of characters without incrementing json_Index (ByVal instead of ByRef)
1         json_SkipSpaces json_String, json_Index
2         json_Peek = VBA.Mid$(json_String, json_Index, json_NumberOfCharacters)
End Function

Private Sub json_SkipSpaces(json_String As String, ByRef json_Index As Long)
          ' Increment index to skip over spaces
1         Do While json_Index > 0 And json_Index <= VBA.Len(json_String) And VBA.Mid$(json_String, json_Index, 1) = " "
2             json_Index = json_Index + 1
3         Loop
End Sub

Private Function json_StringIsLargeNumber(json_String As Variant) As Boolean
          ' Check if the given string is considered a "large number"
          ' (See json_ParseNumber)

          Dim json_Length As Long
          Dim json_CharIndex As Long
1         json_Length = VBA.Len(json_String)

          ' Length with be at least 16 characters and assume will be less than 100 characters
2         If json_Length >= 16 And json_Length <= 100 Then
              Dim json_CharCode As String
              Dim json_Index As Long

3             json_StringIsLargeNumber = True

4             For json_CharIndex = 1 To json_Length
5                 json_CharCode = VBA.Asc(VBA.Mid$(json_String, json_CharIndex, 1))
6                 Select Case json_CharCode
                  ' Look for .|0-9|E|e
                  Case 46, 48 To 57, 69, 101
                      ' Continue through characters
7                 Case Else
8                     json_StringIsLargeNumber = False
9                     Exit Function
10                End Select
11            Next json_CharIndex
12        End If
End Function

Private Function json_ParseErrorMessage(json_String As String, ByRef json_Index As Long, ErrorMessage As String)
          ' Provide detailed parse error message, including details of where and what occurred
          '
          ' Example:
          ' Error parsing JSON:
          ' {"abcde":True}
          '          ^
          ' Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['

          Dim json_StartIndex As Long
          Dim json_StopIndex As Long

          ' Include 10 characters before and after error (if possible)
1         json_StartIndex = json_Index - 10
2         json_StopIndex = json_Index + 10
3         If json_StartIndex <= 0 Then
4             json_StartIndex = 1
5         End If
6         If json_StopIndex > VBA.Len(json_String) Then
7             json_StopIndex = VBA.Len(json_String)
8         End If

9         json_ParseErrorMessage = "Error parsing JSON:" & VBA.vbNewLine & _
                                   VBA.Mid$(json_String, json_StartIndex, json_StopIndex - json_StartIndex + 1) & VBA.vbNewLine & _
                                   VBA.Space$(json_Index - json_StartIndex) & "^" & VBA.vbNewLine & _
                                   ErrorMessage
End Function

Private Sub json_BufferAppend(ByRef json_buffer As String, _
                              ByRef json_Append As Variant, _
                              ByRef json_BufferPosition As Long, _
                              ByRef json_BufferLength As Long)
#If Mac Then
1         json_buffer = json_buffer & json_Append
#Else
          ' VBA can be slow to append strings due to allocating a new string for each append
          ' Instead of using the traditional append, allocate a large empty string and then copy string at append position
          '
          ' Example:
          ' Buffer: "abc  "
          ' Append: "def"
          ' Buffer Position: 3
          ' Buffer Length: 5
          '
          ' Buffer position + Append length > Buffer length -> Append chunk of blank space to buffer
          ' Buffer: "abc       "
          ' Buffer Length: 10
          '
          ' Copy memory for "def" into buffer at position 3 (0-based)
          ' Buffer: "abcdef    "
          '
          ' Approach based on cStringBuilder from vbAccelerator
          ' http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp

          Dim json_AppendLength As Long
          Dim json_LengthPlusPosition As Long

2         json_AppendLength = VBA.LenB(json_Append)
3         json_LengthPlusPosition = json_AppendLength + json_BufferPosition

4         If json_LengthPlusPosition > json_BufferLength Then
              ' Appending would overflow buffer, add chunks until buffer is long enough
              Dim json_TemporaryLength As Long

5             json_TemporaryLength = json_BufferLength
6             Do While json_TemporaryLength < json_LengthPlusPosition
                  ' Initially, initialize string with 255 characters,
                  ' then add large chunks (8192) after that
                  '
                  ' Size: # Characters x 2 bytes / character
7                 If json_TemporaryLength = 0 Then
8                     json_TemporaryLength = json_TemporaryLength + 510
9                 Else
10                    json_TemporaryLength = json_TemporaryLength + 16384
11                End If
12            Loop

13            json_buffer = json_buffer & VBA.Space$((json_TemporaryLength - json_BufferLength) \ 2)
14            json_BufferLength = json_TemporaryLength
15        End If

          ' Copy memory from append to buffer at buffer position
16        json_CopyMemory ByVal json_UnsignedAdd(StrPtr(json_buffer), _
                          json_BufferPosition), _
                          ByVal StrPtr(json_Append), _
                          json_AppendLength

17        json_BufferPosition = json_BufferPosition + json_AppendLength
#End If
End Sub

Private Function json_BufferToString(ByRef json_buffer As String, ByVal json_BufferPosition As Long, ByVal json_BufferLength As Long) As String
#If Mac Then
1         json_BufferToString = json_buffer
#Else
2         If json_BufferPosition > 0 Then
3             json_BufferToString = VBA.Left$(json_buffer, json_BufferPosition \ 2)
4         End If
#End If
End Function

#If VBA7 Then
Private Function json_UnsignedAdd(json_Start As LongPtr, json_Increment As Long) As LongPtr
#Else
Private Function json_UnsignedAdd(json_Start As Long, json_Increment As Long) As Long
#End If

1         If json_Start And &H80000000 Then
2             json_UnsignedAdd = json_Start + json_Increment
3         ElseIf (json_Start Or &H80000000) < -json_Increment Then
4             json_UnsignedAdd = json_Start + json_Increment
5         Else
6             json_UnsignedAdd = (json_Start + &H80000000) + (json_Increment + &H80000000)
7         End If
End Function

''
' VBA-UTC v1.0.3
' (c) Tim Hall - https://github.com/VBA-tools/VBA-UtcConverter
'
' UTC/ISO 8601 Converter for VBA
'
' Errors:
' 10011 - UTC parsing error
' 10012 - UTC conversion error
' 10013 - ISO 8601 parsing error
' 10014 - ISO 8601 conversion error
'
' @module UtcConverter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' (Declarations moved to top)

' ============================================= '
' Public Methods
' ============================================= '

''
' Parse UTC date to local date
'
' @method ParseUtc
' @param {Date} UtcDate
' @return {Date} Local date
' @throws 10011 - UTC parsing error
''
Public Function ParseUtc(utc_UtcDate As Date) As Date
1         On Error GoTo utc_ErrorHandling

#If Mac Then
2         ParseUtc = utc_ConvertDate(utc_UtcDate)
#Else
          Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
          Dim utc_LocalDate As utc_SYSTEMTIME

3         utc_GetTimeZoneInformation utc_TimeZoneInfo
4         utc_SystemTimeToTzSpecificLocalTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_UtcDate), utc_LocalDate

5         ParseUtc = utc_SystemTimeToDate(utc_LocalDate)
#End If

6         Exit Function

utc_ErrorHandling:
7         Err.Raise 10011, "UtcConverter.ParseUtc", "UTC parsing error: " & Err.Number & " - " & Err.Description
End Function

''
' Convert local date to UTC date
'
' @method ConvertToUrc
' @param {Date} utc_LocalDate
' @return {Date} UTC date
' @throws 10012 - UTC conversion error
''
Public Function ConvertToUtc(utc_LocalDate As Date) As Date
1         On Error GoTo utc_ErrorHandling

#If Mac Then
2         ConvertToUtc = utc_ConvertDate(utc_LocalDate, utc_ConvertToUtc:=True)
#Else
          Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
          Dim utc_UtcDate As utc_SYSTEMTIME

3         utc_GetTimeZoneInformation utc_TimeZoneInfo
4         utc_TzSpecificLocalTimeToSystemTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_LocalDate), utc_UtcDate

5         ConvertToUtc = utc_SystemTimeToDate(utc_UtcDate)
#End If

6         Exit Function

utc_ErrorHandling:
7         Err.Raise 10012, "UtcConverter.ConvertToUtc", "UTC conversion error: " & Err.Number & " - " & Err.Description
End Function

''
' Parse ISO 8601 date string to local date
'
' @method ParseIso
' @param {Date} utc_IsoString
' @return {Date} Local date
' @throws 10013 - ISO 8601 parsing error
''
Public Function ParseIso(utc_IsoString As String) As Date
1         On Error GoTo utc_ErrorHandling

          Dim utc_Parts() As String
          Dim utc_DateParts() As String
          Dim utc_TimeParts() As String
          Dim utc_OffsetIndex As Long
          Dim utc_HasOffset As Boolean
          Dim utc_NegativeOffset As Boolean
          Dim utc_OffsetParts() As String
          Dim utc_Offset As Date

2         utc_Parts = VBA.Split(utc_IsoString, "T")
3         utc_DateParts = VBA.Split(utc_Parts(0), "-")
4         ParseIso = VBA.DateSerial(VBA.CInt(utc_DateParts(0)), VBA.CInt(utc_DateParts(1)), VBA.CInt(utc_DateParts(2)))

5         If UBound(utc_Parts) > 0 Then
6             If VBA.InStr(utc_Parts(1), "Z") Then
7                 utc_TimeParts = VBA.Split(VBA.Replace(utc_Parts(1), "Z", ""), ":")
8             Else
9                 utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "+")
10                If utc_OffsetIndex = 0 Then
11                    utc_NegativeOffset = True
12                    utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "-")
13                End If

14                If utc_OffsetIndex > 0 Then
15                    utc_HasOffset = True
16                    utc_TimeParts = VBA.Split(VBA.Left$(utc_Parts(1), utc_OffsetIndex - 1), ":")
17                    utc_OffsetParts = VBA.Split(VBA.Right$(utc_Parts(1), Len(utc_Parts(1)) - utc_OffsetIndex), ":")

18                    Select Case UBound(utc_OffsetParts)
                      Case 0
19                        utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), 0, 0)
20                    Case 1
21                        utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), 0)
22                    Case 2
                          ' VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
23                        utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), Int(VBA.Val(utc_OffsetParts(2))))
24                    End Select

25                    If utc_NegativeOffset Then: utc_Offset = -utc_Offset
26                Else
27                    utc_TimeParts = VBA.Split(utc_Parts(1), ":")
28                End If
29            End If

30            Select Case UBound(utc_TimeParts)
              Case 0
31                ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), 0, 0)
32            Case 1
33                ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), 0)
34            Case 2
                  ' VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
35                ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), Int(VBA.Val(utc_TimeParts(2))))
36            End Select

37            ParseIso = ParseUtc(ParseIso)

38            If utc_HasOffset Then
39                ParseIso = ParseIso + utc_Offset
40            End If
41        End If

42        Exit Function

utc_ErrorHandling:
43        Err.Raise 10013, "UtcConverter.ParseIso", "ISO 8601 parsing error for " & utc_IsoString & ": " & Err.Number & " - " & Err.Description
End Function

''
' Convert local date to ISO 8601 string
'
' @method ConvertToIso
' @param {Date} utc_LocalDate
' @return {Date} ISO 8601 string
' @throws 10014 - ISO 8601 conversion error
''
Public Function ConvertToIso(utc_LocalDate As Date) As String
1         On Error GoTo utc_ErrorHandling

2         ConvertToIso = VBA.Format$(ConvertToUtc(utc_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")

3         Exit Function

utc_ErrorHandling:
4         Err.Raise 10014, "UtcConverter.ConvertToIso", "ISO 8601 conversion error: " & Err.Number & " - " & Err.Description
End Function

' ============================================= '
' Private Functions
' ============================================= '

#If Mac Then

Private Function utc_ConvertDate(utc_Value As Date, Optional utc_ConvertToUtc As Boolean = False) As Date
          Dim utc_ShellCommand As String
          Dim utc_Result As utc_ShellResult
          Dim utc_Parts() As String
          Dim utc_DateParts() As String
          Dim utc_TimeParts() As String

1         If utc_ConvertToUtc Then
2             utc_ShellCommand = "date -ur `date -jf '%Y-%m-%d %H:%M:%S' " & _
                  "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & "' " & _
                  " +'%s'` +'%Y-%m-%d %H:%M:%S'"
3         Else
4             utc_ShellCommand = "date -jf '%Y-%m-%d %H:%M:%S %z' " & _
                  "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & " +0000' " & _
                  "+'%Y-%m-%d %H:%M:%S'"
5         End If

6         utc_Result = utc_ExecuteInShell(utc_ShellCommand)

7         If utc_Result.utc_Output = "" Then
8             Err.Raise 10015, "UtcConverter.utc_ConvertDate", "'date' command failed"
9         Else
10            utc_Parts = Split(utc_Result.utc_Output, " ")
11            utc_DateParts = Split(utc_Parts(0), "-")
12            utc_TimeParts = Split(utc_Parts(1), ":")

13            utc_ConvertDate = DateSerial(utc_DateParts(0), utc_DateParts(1), utc_DateParts(2)) + _
                  TimeSerial(utc_TimeParts(0), utc_TimeParts(1), utc_TimeParts(2))
14        End If
End Function

Private Function utc_ExecuteInShell(utc_ShellCommand As String) As utc_ShellResult
#If VBA7 Then
          Dim utc_File As LongPtr
          Dim utc_Read As LongPtr
#Else
          Dim utc_File As Long
          Dim utc_Read As Long
#End If

          Dim utc_Chunk As String

1         On Error GoTo utc_ErrorHandling
2         utc_File = utc_popen(utc_ShellCommand, "r")

3         If utc_File = 0 Then: Exit Function

4         Do While utc_feof(utc_File) = 0
5             utc_Chunk = VBA.Space$(50)
6             utc_Read = utc_fread(utc_Chunk, 1, Len(utc_Chunk) - 1, utc_File)
7             If utc_Read > 0 Then
8                 utc_Chunk = VBA.Left$(utc_Chunk, utc_Read)
9                 utc_ExecuteInShell.utc_Output = utc_ExecuteInShell.utc_Output & utc_Chunk
10            End If
11        Loop

utc_ErrorHandling:
12        utc_ExecuteInShell.utc_ExitCode = utc_pclose(utc_File)
End Function

#Else

Private Function utc_DateToSystemTime(utc_Value As Date) As utc_SYSTEMTIME
1         utc_DateToSystemTime.utc_wYear = VBA.Year(utc_Value)
2         utc_DateToSystemTime.utc_wMonth = VBA.Month(utc_Value)
3         utc_DateToSystemTime.utc_wDay = VBA.Day(utc_Value)
4         utc_DateToSystemTime.utc_wHour = VBA.Hour(utc_Value)
5         utc_DateToSystemTime.utc_wMinute = VBA.Minute(utc_Value)
6         utc_DateToSystemTime.utc_wSecond = VBA.Second(utc_Value)
7         utc_DateToSystemTime.utc_wMilliseconds = 0
End Function

Private Function utc_SystemTimeToDate(utc_Value As utc_SYSTEMTIME) As Date
1         utc_SystemTimeToDate = DateSerial(utc_Value.utc_wYear, utc_Value.utc_wMonth, utc_Value.utc_wDay) + _
              TimeSerial(utc_Value.utc_wHour, utc_Value.utc_wMinute, utc_Value.utc_wSecond)
End Function

#End If
