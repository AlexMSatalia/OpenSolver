VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private curStart As Currency
Private curEnd As Currency
Private curFreq As Currency

Public Sub StartTimer()
    curStart = modCTimer.GetTime()
End Sub

Public Function StopTimer() As Double
    curEnd = modCTimer.GetTime
    curFreq = modCTimer.GetFreq
    StopTimer = (curEnd - curStart) / curFreq
End Function

Public Function Time() As Double
    Time = (curEnd - curStart) / curFreq
End Function
