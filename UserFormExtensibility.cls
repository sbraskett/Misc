VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserFormExtensibility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' © Copyright/Author, Mark Kubiszyn. All Rights Reserved.  UserFormExtensibility Class.  License MIT License (MIT) upon purchase!
' uses a redacted / modified version of my UserForm Extensibility Class and its respective component Classes
' https://www.kubiszyn.co.uk/documentation/userform-extensibility.html
Option Explicit

#If VBA7 Then
   Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
   Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal lngWindow As Long) As Long
   Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal lngWindow As Long, ByVal dwNewLong As Long) As Long
   Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "User32.dll" (ByVal hwnd As LongPtr, ByVal crKey As LongPtr, ByVal bAlpha As Byte, ByVal dwFlags As LongPtr) As LongPtr
   Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
   Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
   Private Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long
   Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
   Private Declare PtrSafe Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
   Private Declare PtrSafe Function SetWindowRgn Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndRect As Long, ByVal bRedraw As Long) As Long
   Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
   Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, ByRef lpRect As RECT) As Long
   Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hwnd As LongPtr) As Long
   Private Declare PtrSafe Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
   Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
   Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As Long
   Private Declare PtrSafe Function FrameRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
   Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
#Else
   Private Declare Function GetActiveWindow Lib "user32" () As Long
   Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal lngWindow As Long) As Long
   Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal lngWindow As Long, ByVal dwNewLong As Long) As Long
   Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Integer, ByVal bAlpha As Integer, ByVal dwFlags As Long) As Long
   Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
   Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
   Private Declare Function ReleaseCapture Lib "user32" () As Long
   Private Declare Function SendMessage Lib "user32" Alias "SendMessageA"(ByVal hWnd As Long,ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
   Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
   Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hwndRect As Long, ByVal bRedraw As Long) As Long
   Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
   Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
   Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
   Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
   Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
   Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
   Private Declare Function FrameRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
   Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
#End If

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents OnDeactivate As Excel.Worksheet
Attribute OnDeactivate.VB_VarHelpID = -1

Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = &HFFEC
Private Const HTCAPTION As Long = 2
Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_EX_DLGMODALFRAME As Long = &H1
Private Const LWA_COLORKEY As Long = &H1

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type TInterface
   PScreen As SystemMetric
   PTimer As SystemTiming
   PEasing As Easing
   PCursor As Cursor
   PUserForm As Object
   #If VBA7 Then
   PhWnd As LongPtr
   #Else
   PhWnd As Long
   #End If
   PUnloadOnDeactivate As Boolean
   PSprite As Object
End Type

Private this As TInterface

Public Property Set Assign(ByVal value As Object): Set this.PUserForm = value: this.PhWnd = FindWindow(vbNullString, this.PUserForm.Caption): End Property

Public Property Get Assign() As Object: Set Assign = this.PUserForm: End Property

Private Sub Class_Initialize()
   Set this.PScreen = New SystemMetric
   Set this.PTimer = New SystemTiming
   Set this.PEasing = New Easing
   Set this.PCursor = New Cursor
   
   Set OnDeactivate = ThisWorkbook.ActiveSheet
End Sub

Public Sub RemoveCaption()
   GetWindowLong this.PhWnd, GWL_STYLE
   SetWindowLong this.PhWnd, GWL_STYLE, 0
   DrawMenuBar this.PhWnd
End Sub

Public Sub RemoveCaptionAndBorders()
   Dim hStyle As Long
   hStyle = GetWindowLong(this.PhWnd, GWL_STYLE)
   hStyle = hStyle And Not WS_CAPTION
   SetWindowLong this.PhWnd, GWL_STYLE, hStyle
   hStyle = hStyle And Not WS_EX_DLGMODALFRAME
   SetWindowLong this.PhWnd, GWL_EXSTYLE, hStyle
   DrawMenuBar this.PhWnd
End Sub

Public Sub RemoveCaptionBordersAndSetTransparency(Optional ByRef TransparencyColour As Long = 255, Optional ByRef Opacity As Long = 255)
   Dim hStyle As Long
   hStyle = GetWindowLong(this.PhWnd, GWL_STYLE)
   hStyle = hStyle And Not WS_CAPTION
   SetWindowLong this.PhWnd, GWL_STYLE, hStyle
   hStyle = hStyle And Not WS_EX_DLGMODALFRAME
   SetWindowLong this.PhWnd, GWL_EXSTYLE, hStyle
   SetLayeredWindowAttributes this.PhWnd, TransparencyColour, Opacity, LWA_COLORKEY
   DrawMenuBar this.PhWnd
End Sub

Public Sub RoundCorners(Optional ByVal CornerSize As Long = 12)
   SetWindowRgn this.PhWnd, CreateRoundRectRgn(0, 0, this.PUserForm.Width / this.PScreen.PointsPerPixel, this.PUserForm.Height / this.PScreen.PointsPerPixel, CornerSize, CornerSize), True
End Sub

Public Sub RoundBorderCorners(Optional ByVal CornerSize As Long = 12)
   this.PTimer.WaitForMilliseconds 50                                     ' Important!
   Dim BrushValue As Long
   Dim DCValue As Long
   Dim RectStructure As RECT
   Dim hWndRect As Long
   BrushValue = CreateSolidBrush(CLng(this.PUserForm.BorderColor))
   GetWindowRect this.PhWnd, RectStructure
   hWndRect = CreateRoundRectRgn(0, 0, this.PUserForm.Width / this.PScreen.PointsPerPixel + 1, this.PUserForm.Height / this.PScreen.PointsPerPixel + 1, CornerSize, CornerSize)
   DCValue = GetWindowDC(this.PhWnd)
   FrameRgn DCValue, hWndRect, BrushValue, 1, 1
   ReleaseDC hWndRect, DCValue
   SetWindowRgn this.PhWnd, hWndRect, True
   DeleteObject hWndRect
   DeleteObject BrushValue
End Sub

Public Sub AllowDrag()
   ReleaseCapture
   SendMessage this.PhWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
End Sub

Public Sub MoveOffScreen()
   this.PUserForm.StartUpPosition = 0
   this.PUserForm.Move -this.PScreen.ScreenWidth, -this.PScreen.ScreenHeight
End Sub

Public Sub Pause(ByVal Milliseconds As Long)
   this.PTimer.WaitForMilliseconds Milliseconds
End Sub

Public Sub ChangeCursor(ByVal CursorName As CursorTypes)
   this.PCursor.AddCursor CursorName
End Sub

