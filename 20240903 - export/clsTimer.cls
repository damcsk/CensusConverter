VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Quick timing test boilerplate
'Dim i As Integer
'Dim j As Integer
'Dim result As Double
'Dim stopwatch As clsTimer
'Set stopwatch = New clsTimer
'
'result = 0
'j = 100
'For i = 1 To j
'    stopwatch.StartTimer
'
'    result = result + stopwatch.ElapsedTime
'Next i
'result = result / j
'MsgBox result

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long

Private m_CounterStart As LARGE_INTEGER
Private m_CounterEnd As LARGE_INTEGER
Private m_crFrequency As Double

Private Const TWO_32 = 4294967296# ' = 256# * 256# * 256# * 256#

Private Function LI2Double(LI As LARGE_INTEGER) As Double
Dim Low As Double
    Low = LI.lowpart
    If Low < 0 Then
        Low = Low + TWO_32
    End If
    LI2Double = LI.highpart * TWO_32 + Low
End Function

Private Sub Class_Initialize()
Dim PerfFrequency As LARGE_INTEGER
    QueryPerformanceFrequency PerfFrequency
    m_crFrequency = LI2Double(PerfFrequency)
End Sub

Public Sub StartTimer()
    QueryPerformanceCounter m_CounterStart
End Sub

Property Get ElapsedTime() As Double
Dim crStart As Double
Dim crStop As Double
    QueryPerformanceCounter m_CounterEnd
    crStart = LI2Double(m_CounterStart)
    crStop = LI2Double(m_CounterEnd)
    ElapsedTime = 1000# * (crStop - crStart) / m_crFrequency
End Property







