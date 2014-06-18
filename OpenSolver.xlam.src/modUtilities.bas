Attribute VB_Name = "modUtilities"
'==============================================================================
' modUtilities
' General purpose functions
'==============================================================================
Option Explicit


'==============================================================================
Function StrEx(d As Double) As String
    ' Convert a double to a string, always with a + or -. Also ensure we have "0.", not just "." for values between -1 and 1
    Dim s As String, prependedZero As String, sign As String
    s = Mid(str(d), 2)  ' remove the initial space (reserved by VB for the sign)
    prependedZero = IIf(left(s, 1) = ".", "0", "")  ' ensure we have "0.", not just "."
    sign = IIf(d >= 0, "+", "-")
    StrEx = sign + prependedZero + s
End Function

'==============================================================================
Function TestIntersect(ByRef R1 As Range, ByRef R2 As Range) As Boolean
    Dim r3 As Range
    Set r3 = Intersect(R1, R2)
    TestIntersect = Not (r3 Is Nothing)
    ' Below: a test to see if I could do it faster - I couldn't
    'Dim R1_X1 As Long, R1_Y1 As Long, R1_X2 As Long, R1_Y2 As Long
    'Dim R2_X1 As Long, R2_Y1 As Long, R2_X2 As Long, R2_Y2 As Long
    'R1_X1 = R1.Column
    'R1_X2 = R1_X1 + R1.Columns.Count - 1
    'R2_X1 = R2.Column
    'R2_X2 = R2_X1 + R2.Columns.Count - 1
    'R1_Y1 = R1.Row
    'R1_Y2 = R1_Y2 + R1.Rows.Height - 1
    'R2_Y1 = R2.Row
    'R2_Y2 = R2_Y2 + R2.Rows.Height - 1
    'TestIntersect = _
    '    R1_X1 <= R2_X2 And _  ' Cond A
    '    R1_X2 >= R2_X1 And _  ' Cond B
    '    R1_Y1 <= R2_Y2 And _  ' Cond C
    '    R1_Y2 >= R2_Y1        ' Cond D
End Function

