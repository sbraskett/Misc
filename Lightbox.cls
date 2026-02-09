VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lightbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' © Copyright/Author, Mark Kubiszyn. All Rights Reserved.  Lightbox Class.  License MIT License (MIT) upon purchase!
Option Explicit

#If VBA7 Then
   Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
   Private Declare PtrSafe Function MoveWindow Lib "User32.dll" (ByVal hwnd As LongPtr, ByVal X As LongPtr, ByVal Y As LongPtr, ByVal nWidth As LongPtr, ByVal nHeight As LongPtr, ByVal bRepaint As LongPtr) As LongPtr
   Private Declare PtrSafe Function GetWindowLong Lib "User32.dll" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As LongPtr) As Long
   Private Declare PtrSafe Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As LongPtr, ByVal dwNewLongPtr As LongPtr) As LongPtr
   Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "User32.dll" (ByVal hwnd As LongPtr, ByVal crKey As LongPtr, ByVal bAlpha As Byte, ByVal dwFlags As LongPtr) As LongPtr
   Private Declare PtrSafe Function DrawMenuBar Lib "User32.dll" (ByVal hwnd As LongPtr) As LongPtr
   Private Declare PtrSafe Function SetWindowRgn Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndRect As Long, ByVal bRedraw As Long) As Long
   Private Declare PtrSafe Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
#Else
   Private Declare Function FindWindow Lib "User32.dll" Alias "FindWindowA" (ByVal lpszClass As String, ByVal lpszWindow As String) As Long
   Private Declare Function MoveWindow Lib "User32.dll" (ByVal Hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
   Private Declare Function GetWindowLong Lib "User32.dll" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
   Private Declare Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
   Private Declare Function SetLayeredWindowAttributes Lib "User32.dll" (ByVal Hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
   Private Declare Function DrawMenuBar Lib "User32.dll" (ByVal Hwnd As Long) As Long
   Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hwndRect As Long, ByVal bRedraw As Long) As Long
   Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
#End If

Public Enum TLightboxPlacement
   FitToScreen = 1
   FitToExcel = 2
End Enum

Private Type TInterface
   TScreen As SystemMetric
   TEasing As Easing
   TLightbox As Object
   TLightboxHwnd As Long
   TLightboxUserFormName As String
End Type

Private this As TInterface

Private Sub Class_Initialize()
   Set this.TLightbox = ThisWorkbook.VBProject.VBComponents.Add(3)
   this.TLightboxUserFormName = this.TLightbox.Name
   VBA.UserForms.Add (this.TLightboxUserFormName)
   Set this.TLightbox = VBA.UserForms.Item(UserForms.Count - 1)
   With this.TLightbox
      .Width = 400
      .Height = 300
      .StartUpPosition = 0
      .Enabled = False
   End With
   this.TLightboxHwnd = FindWindow("ThunderDFrame", this.TLightboxUserFormName)
   Set this.TScreen = New SystemMetric
   Set this.TEasing = New Easing
End Sub

Public Sub TransitionIn(Optional ByVal Placement As TLightboxPlacement = FitToExcel, Optional ByVal Easing As TEasing = easeBackInQuartic, Optional ByVal Opacity As Double = 120, Optional ByVal Colour As Variant = vbWhite, Optional ByVal NumberOfTweens As Double = 12, Optional ByVal TweensPerSecond As Double = 50)
   #If VBA7 Then
      Dim Style As LongPtr
   #Else
      Dim Style As Long
   #End If
   MoveWindow this.TLightboxHwnd, this.TScreen.ScreenWidth, this.TScreen.ScreenHeight, this.TScreen.ScreenWidth, this.TScreen.ScreenHeight, 1
   this.TLightbox.BackColor = Colour
   Style = GetWindowLong(this.TLightboxHwnd, -16)
   Style = Style And Not &HC00000
   SetWindowLong this.TLightboxHwnd, -16, Style
   SetWindowLong this.TLightboxHwnd, -20, &H40000
   SetWindowLong this.TLightboxHwnd, -20, &H80000
   SetLayeredWindowAttributes this.TLightboxHwnd, Colour, 0, &H2
   DrawMenuBar this.TLightboxHwnd
   If NumberOfTweens > 0 Then
      this.TLightbox.Show 0
      If Placement = FitToExcel Then
         MoveWindow this.TLightboxHwnd, (Application.Left / this.TScreen.PointsPerPixel) - 1, (Application.Top / this.TScreen.PointsPerPixel), (Application.Width / this.TScreen.PointsPerPixel), (Application.Height / this.TScreen.PointsPerPixel), 1
         ' for any fit to Excel, round windows for Windows 11.  we need to round the corners of the Window and reduce the Window dimensions slightly to take account of the latest Window styling
         If WIN11 Then SetWindowRgn this.TLightboxHwnd, CreateRoundRectRgn(7, 0, (this.TLightbox.Width - 4.5) / this.TScreen.PointsPerPixel, (this.TLightbox.Height - 4.5) / this.TScreen.PointsPerPixel, 11, 11), True
      Else
         MoveWindow this.TLightboxHwnd, 0, 0, this.TScreen.ScreenWidth, this.TScreen.ScreenHeight, 1
      End If
      Dim Tweening As Double
      Dim TweenTime As Double
      Dim LayeredOpacity As Double
      TweenTime = 0
      Do
         DoEvents
         LayeredOpacity = this.TEasing.EasingFunction(Easing, TweenTime, 0, Opacity, 1)
         If LayeredOpacity >= 0 Then SetLayeredWindowAttributes this.TLightboxHwnd, 0, LayeredOpacity, &H2
         TweenTime = TweenTime + 1 / NumberOfTweens
         Tweening = Timer
         Do
            DoEvents
         Loop While Timer - Tweening < 1 / TweensPerSecond
      Loop Until TweenTime > 1
   End If
End Sub

Public Sub TransitionOut(ByVal Easing As TEasing, Optional ByVal Opacity As Double = 120, Optional ByVal Colour As Variant = vbWhite, Optional ByVal NumberOfTweens As Double = 12, Optional ByVal TweensPerSecond As Double = 50)
   Dim Tweening As Double
   Dim TweenTime As Double
   Dim LayeredOpacity As Double
   TweenTime = 0
   Do
      DoEvents
      LayeredOpacity = this.TEasing.EasingFunction(Easing, TweenTime, 0, Opacity, 1)
      If LayeredOpacity >= 0 And LayeredOpacity <= Opacity Then SetLayeredWindowAttributes this.TLightboxHwnd, 0, (Opacity - LayeredOpacity), &H2
      TweenTime = TweenTime + 1 / NumberOfTweens
      Tweening = Timer
      Do
         DoEvents
      Loop While Timer - Tweening < 1 / TweensPerSecond
   Loop Until TweenTime > 1
End Sub

Private Sub Class_Terminate()
   Dim VBComp As Object
   If this.TLightboxUserFormName <> "" Then
      Set VBComp = ThisWorkbook.VBProject.VBComponents(this.TLightboxUserFormName)
      ThisWorkbook.VBProject.VBComponents.Remove VBComp
   End If
End Sub

Private Function WIN11() As Boolean
   On Error GoTo Catch
   Dim WMI As Object
   Dim WMIQuery As String
   Dim BEM As Object
   Set WMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
   WMIQuery = "SELECT BuildNumber FROM Win32_OperatingSystem"
   Set BEM = WMI.ExecQuery(WMIQuery)
   If BEM.ItemIndex(0).Properties_("BuildNumber") >= 22000 Then WIN11 = True Else WIN11 = False
   Exit Function
Catch:
   WIN11 = False
End Function

