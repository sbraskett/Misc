VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If VBA7 And Win64 Then
   Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
   Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
   Private Declare PtrSafe Function GetCursorInfo Lib "user32" (ByRef pci As CursorInfo) As Boolean
   Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
   Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
   Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
   Private Declare Function GetCursorInfo Lib "user32" (ByRef pci As CursorInfo) As Boolean
   Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Enum CursorTypes
   IDC_ARROW = 32512
   IDC_IBEAM = 32513
   IDC_WAIT = 32514
   IDC_CROSS = 32515
   IDC_UPARROW = 32516
   IDC_SIZE = 32640
   IDC_ICON = 32641
   IDC_SIZENWSE = 32642
   IDC_SIZENESW = 32643
   IDC_SIZEWE = 32644
   IDC_SIZENS = 32645
   IDC_SIZEALL = 32646
   IDC_NO = 32648
   IDC_HAND = 32649
   IDC_APPSTARTING = 32650
End Enum

Private Type POINT
   X As Long
   Y As Long
End Type

Private Type CursorInfo
   cbSize As Long
   flags As Long
   hCursor As Long
   ptScreenPos As POINT
End Type

Public Sub AddCursor(ByRef CursorType As CursorTypes)
   If Not IsCursorType(CursorType) Then
      SetCursor LoadCursor(0, CursorType)
      ' you can leave this slight pause if you prefer ;)
      'Sleep 50
   End If
End Sub

Public Function IsCursorType(ByRef CursorType As CursorTypes) As Boolean
   Dim CursorHandle As Long: CursorHandle = LoadCursor(ByVal 0&, CursorType)
   Dim Cursor As CursorInfo: Cursor.cbSize = Len(Cursor)
   Dim CursorInfo As Boolean: CursorInfo = GetCursorInfo(Cursor)
   If Not CursorInfo Then
      IsCursorType = False
      Exit Function
   End If
   IsCursorType = (Cursor.hCursor = CursorHandle)
End Function

