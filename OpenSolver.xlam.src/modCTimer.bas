Attribute VB_Name = "modCTimer"
'==============================================================================
' OpenSolver
' Copyright Andrew Mason, Iain Dunning, 2011
' http://www.opensolver.org
'==============================================================================
' modCTimer
' High-performance timer API calls are defined here. Each instance of CTimer
' will use this to get start and end times.
'==============================================================================
Option Explicit

#If VBA7 Then
    ' No point writing it if it can't be tested.
    ' Not a critical feature!
    Public Function GetTime() As Currency
        GetTime = CCur(Timer)
    End Function

    Public Function GetFreq() As Currency
        GetFreq = 1
    End Function
#Else

' Returns the current high performance tick count
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
' Get the ticks per second
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
' Required for converting
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

    
' Each of the Performance Counter functions returns a 64bit count.
'   This result must be taken in it's entirety, and only the Currency
'   data type will allow 64bit values for calculation.
Public Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type



Public Function GetTime() As Currency
    Dim lgiTemp As LARGE_INTEGER
    Call QueryPerformanceCounter(lgiTemp)
    GetTime = LargeIntToCurrency(lgiTemp)
End Function

Public Function GetFreq() As Currency
    Dim lgiTemp As LARGE_INTEGER
    Call QueryPerformanceFrequency(lgiTemp)
    GetFreq = LargeIntToCurrency(lgiTemp)
End Function

Private Function LargeIntToCurrency(liInput As LARGE_INTEGER) As Currency
    ' Copy 8 bytes (64 bits) from the large integer to an empty currency
    CopyMemory LargeIntToCurrency, liInput, LenB(liInput)
    
    ' Adjust it to remove the 4 decimal positions of the Currency type
    LargeIntToCurrency = LargeIntToCurrency * 10000
End Function

#End If




