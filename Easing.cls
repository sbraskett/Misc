VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Easing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' © Copyright/Author, Mark Kubiszyn. All Rights Reserved.  Easing Class.  License MIT License (MIT) upon purchase!
Option Explicit

Public Enum TEasing
   easeBackInCubic = 1
   easeBackInQuartic = 2
   easeInBack = 3
   easeInCirc = 4
   easeInCubic = 5
   easeInExpo = 6
   easeInOutCubic = 7
   easeInOutQuintic = 8
   easeInQuad = 9
   easeInQuadratic = 10
   easeInQuart = 11
   easeInQuartic = 12
   easeInQuintic = 13
   easeInSine = 14
   easeLinear = 15
   easeOutBack = 16
   easeOutBounce = 17
   easeOutElastic = 18
   easeOutInCubic = 19
   easeOutInQuartic = 20
   easeOutQuad = 21
   easeOutQuartic = 22
   easeOutQuintic = 23
   easeOutSine = 24
   easeInOutSine = 25
   easeOutQuart = 26
   easeOutCirc = 27
   easeInOutQuad = 28
   easeOutCubic = 29
   easeInOutQuartic = 30
   easeOutExpo = 31
   'easeInOutExpo = 32
   easeInOutCirc = 33
   easeOutElasticSmall = 34
   easeOutElasticLarge = 35
   easeInElasticSmall = 36
   easeInElasticBig = 37
End Enum

'@Ignore AssignedByValParameter
Public Function EasingFunction(ByVal Easing As TEasing, ByVal t As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Variant
   Dim ts As Double
   Dim tc As Double
   Select Case Easing
   Case easeLinear
      EasingFunction = c * t / d + b
   Case easeInQuad
      ts = t / d
      EasingFunction = c * ts * t + b
   Case easeOutQuad
      ts = t / d
      EasingFunction = -c * ts * (t - 2) + b
   Case easeInBack
      ts = t / d
      EasingFunction = c * ts * t * ((1.70158 + 1) * t - 1.70158) + b
   Case easeInQuart
      EasingFunction = c * (t / d) * t * t * t + b
   Case easeInExpo
      If t = 0 Then
         EasingFunction = b
      Else
         EasingFunction = c * Application.WorksheetFunction.Power(2, 10 * (t / d - 1)) + b
      End If
   Case easeInCirc
      EasingFunction = -c * (VBA.Sqr(1 - (t / d) * t) - 1) + b
   Case easeInQuintic
      EasingFunction = b + c * (((t / d) * t * t) * ((t / d) * t))
   Case easeOutBounce
      If (t / d < (1 / 2.75)) Then
         EasingFunction = c * (7.5625 * t * t) + b
      ElseIf (t < (2 / 2.75)) Then
         t = t - (1.5 / 2.75)
         EasingFunction = c * (7.5625 * (t) * t + 0.75) + b
      ElseIf (t < (2.5 / 2.75)) Then
         t = t - (2.25 / 2.75)
         EasingFunction = c * (7.5625 * (t) * t + 0.9375) + b
      Else
         t = t - (2.625 / 2.75)
         EasingFunction = c * (7.5625 * (t) * t + 0.984375) + b
      End If
   Case easeOutQuartic
      ts = (t / d) * t
      tc = ts * t
      EasingFunction = b + c * (-1 * ts * ts + 4 * tc + -6 * ts + 4 * t)
   Case easeOutQuintic
      ts = (t / d) * t
      tc = ts * t
      EasingFunction = b + c * (tc * ts + -5 * ts * ts + 10 * tc + -10 * ts + 5 * t)
   Case easeOutElastic
      Dim s As Double
      Dim p As Double
      Dim a As Double
      s = 1.70158
      p = 0
      a = c
      If t = 0 Then
         EasingFunction = b
         Exit Function
      End If
      t = t / d
      If t = 1 Then
         EasingFunction = b + c
         Exit Function
      End If
      p = d * 0.3
      If a < Math.Abs(c) Then
         a = c
         s = p / 4
      Else
         s = p / (2 * Application.WorksheetFunction.Pi()) * Application.WorksheetFunction.Asin(c / a)
      End If
      EasingFunction = (a * (2 ^ (-10 * t))) * Math.Sin(((t * d - s) * (2 * Application.WorksheetFunction.Pi()) / p)) + c + b
   Case easeBackInCubic
      ts = (t / d) * t
      tc = ts * t
      EasingFunction = b + c * (4 * tc + -3 * ts)
   Case easeInCubic
      ts = (t / d) * t * t
      EasingFunction = b + c * (ts)
   Case easeBackInQuartic
      ts = (t / d) * t
      tc = ts * t
      EasingFunction = b + c * (2 * ts * ts + 2 * tc + -3 * ts)
   Case easeOutInQuartic
      ts = (t / d) * t
      tc = ts * t
      EasingFunction = b + c * (6 * tc + -9 * ts + 4 * t)
   Case easeInQuartic
      ts = (t / d) * t
      EasingFunction = b + c * (ts * ts)
   Case easeOutInCubic
      ts = (t / d) * t
      tc = ts * t
      EasingFunction = b + c * (4 * tc + -6 * ts + 3 * t)
   Case easeInQuadratic
      ts = (t / d) * t
      EasingFunction = b + c * (ts)
   Case easeInOutQuintic
      ts = (t / d) * t
      tc = ts * t
      EasingFunction = b + c * (6 * tc * ts + -15 * ts * ts + 10 * tc)
   Case easeInOutCubic
      ts = (t / d) * t
      tc = ts * t
      EasingFunction = b + c * (-2 * tc + 3 * ts)
   Case easeInSine
      EasingFunction = -c * Math.Cos(t / d * (Application.WorksheetFunction.Pi() / 2)) + c + b
   Case easeOutSine
      EasingFunction = c * Math.Sin(t / d * (Application.WorksheetFunction.Pi() / 2)) + b
   Case easeOutBack
      ts = t / d - 1
      EasingFunction = c * ((ts) * t * ((1.70158 + 1) * t + 1.70158) + 1) + b
   Case easeInOutSine
      EasingFunction = -c / 2 * (Math.Cos(Application.WorksheetFunction.Pi() * t / d) - 1) + b
   Case easeOutQuart
      t = t / d
      t = t - 1
      EasingFunction = -c * (t * t * t * t - 1) + b
   Case easeOutCirc
      t = t / d
      t = t - 1
      EasingFunction = c * VBA.Sqr(1 - t * t) + b
   Case easeInOutQuad
      t = t / (d / 2)
      If t < 1 Then
         EasingFunction = c / 2 * t * t + b
      Else
         t = t - 1
         EasingFunction = -c / 2 * (t * (t - 2) - 1) + b
      End If
   Case easeOutCubic
      t = t / d
      t = t - 1
      EasingFunction = c * (t * t * t + 1) + b
   Case easeInOutQuartic
      t = t / (d / 2)
      If t < 1 Then
         EasingFunction = c / 2 * t * t * t * t + b
      Else
         t = t - 2
         EasingFunction = -c / 2 * (t * t * t * t - 2) + b
      End If
   Case easeOutExpo
      t = t / d
      EasingFunction = c * (-Application.WorksheetFunction.Power(2, -10 * t) + 1) + b
      '   Case easeInOutExpo
      '      t = t / (d / 2)
      '      If t < 1 Then
      '         EasingFunction = c / 2 * Application.WorksheetFunction.Power(2, 10 * (t - 1)) + b
      '      Else
      '         t = t - 1
      '         EasingFunction = c / 2 * (-Application.WorksheetFunction.Power(2, -10 * t) + 2) + b
      '      End If
   Case easeInOutCirc
      t = t / (d / 2)
      If t < 1 Then
         EasingFunction = -c / 2 * (VBA.Sqr(1 - t * t) - 1) + b
      Else
         t = t - 2
         EasingFunction = c / 2 * (VBA.Sqr(1 - t * t) + 1) + b
      End If
   Case easeOutElasticSmall
      ts = (t / d) * t
      tc = ts * t
      EasingFunction = b + c * (33 * tc * ts + -106 * ts * ts + 126 * tc + -67 * ts + 15 * t)
   Case easeOutElasticLarge
      ts = (t / d) * t
      tc = ts * t
      EasingFunction = b + c * (56 * tc * ts + -175 * ts * ts + 200 * tc + -100 * ts + 20 * t)
   Case easeInElasticSmall
      ts = (t / d) * t
      tc = ts * t
      EasingFunction = b + c * (33 * tc * ts + -59 * ts * ts + 32 * tc + -5 * ts)
   Case easeInElasticBig
      ts = (t / d) * t
      tc = ts * t
      EasingFunction = b + c * (56 * tc * ts + -105 * ts * ts + 60 * tc + -10 * ts)
   End Select
End Function

