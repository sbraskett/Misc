VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SystemMetric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' © Copyright/Author, Mark Kubiszyn. All Rights Reserved.  SystemMetric Class.  License MIT License (MIT) upon purchase!
Option Explicit

#If VBA7 Then
   Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
   Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
   Private Declare PtrSafe Function ReleaseDC Lib "User32.dll" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
   Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
   Private Declare Function GetSystemMetrics Lib "User32.dll" (ByVal nIndex As Long) As Long
   Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
   Private Declare Function ReleaseDC Lib "User32.dll" (ByVal Hwnd As Long, ByVal hDC As Long) As Long
   Private Declare Function GetDC Lib "User32.dll" (ByVal Hwnd As Long) As Long
#End If

Private Type TScreen
   TScreenWidth As Long
   TScreenHeight As Long
   TPointsPerPixel As Double
End Type

Private this As TScreen

Public Property Get Self() As SystemMetric
   Set Self = Me
End Property

Public Property Get ScreenWidth() As Long
   this.TScreenWidth = GetSystemMetrics(0)
   ScreenWidth = this.TScreenWidth
End Property

Public Property Let ScreenWidth(ByVal value As Long)
   this.TScreenWidth = value
End Property

Public Property Get ScreenHeight() As Long
   this.TScreenHeight = GetSystemMetrics(1)
   ScreenHeight = this.TScreenHeight
End Property

Public Property Let ScreenHeight(ByVal value As Long)
   this.TScreenHeight = value
End Property

Public Property Get PointsPerPixel() As Double
   Dim DCHwnd As Long
   Dim DotsPerInch As Long
   DCHwnd = GetDC(0)
   DotsPerInch = GetDeviceCaps(DCHwnd, 88)
   this.TPointsPerPixel = 72 / DotsPerInch
   PointsPerPixel = this.TPointsPerPixel
   ReleaseDC 0, DCHwnd
End Property

Public Property Let PointsPerPixel(ByVal value As Double)
   this.TPointsPerPixel = value
End Property

