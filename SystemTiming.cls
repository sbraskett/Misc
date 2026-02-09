VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SystemTiming"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' © Copyright/Author, Mark Kubiszyn. All Rights Reserved.  SystemTiming Class.  License MIT License (MIT) upon purchase!
Option Explicit

#If VBA7 Then
   Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
   Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

Public Sub WaitForMilliseconds(ByVal TimeInMilliseconds As Long)
   '@Ignore VariableNotUsed
   Dim WaitInMilliseconds As Long
   Dim Ticker As Long
   Ticker = GetTickCount()
   Do
      WaitInMilliseconds = DoEvents
   Loop Until Ticker + TimeInMilliseconds < GetTickCount
End Sub

