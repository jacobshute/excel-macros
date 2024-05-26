VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''
'
' From https://stackoverflow.com/questions/939230/how-to-get-a-datediff-value-in-milliseconds-in-vba-excel


Private mlngStart As Long
#If Win64 Then
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As LongPtr
Public Sub StartTimer()
    mlngStart = GetTickCount
End Sub

Public Function EndTimer() As Long
    EndTimer = (GetTickCount - mlngStart)
End Function
#Else
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public Sub StartTimer()
    mlngStart = GetTickCount
End Sub

Public Function EndTimer() As Long
    EndTimer = (GetTickCount - mlngStart)
End Function
#End If



