VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "Box.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule IntegerDataType, SheetAccessedUsingString

Option Explicit

' add UserForm extensibility Classes
Private Self As New UserFormExtensibility

' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' -----¬ UserForm
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub UserForm_Initialize()
   
   ' UserForm extensibility
   Set Self.Assign = Me
   Me.CloseCrossCircle.Visible = False

   ' Radiobuttons.  You will need to initialise each Group.  a Group is "A", "B" etc. if more than one set are used
   Dim ID As Long
   ID = [RadiobuttonASelected].Value2
   Dim ControlName As String
   ControlName = "A" & ID & "Radiobutton"
   Me.Controls(ControlName).ZOrder msoSendToBack
   Me.Controls(ControlName & "Checked").ZOrder msoBringToFront
   
   ' Checkboxes
   If [Checkbox1Value].Value2 Then Tick Me.Checkbox1 Else Untick Me.Checkbox1
   If [Checkbox2Value].Value2 Then Tick Me.Checkbox2 Else Untick Me.Checkbox2
   If [Checkbox3Value].Value2 Then Tick Me.Checkbox3 Else Untick Me.Checkbox3
   If [Checkbox4Value].Value2 Then Tick Me.Checkbox4 Else Untick Me.Checkbox4
   
   ' Toggle
   Toggle Me.Toggle1, True
   
   ' Sliders
   InitialiseSlider Me.Slider1, ThisWorkbook.Sheets("Sheet1").Range("B17").Value2, 1, 10
   InitialiseSlider Me.Slider2, ThisWorkbook.Sheets("Sheet1").Range("B19").Value2, 1, 5
 
End Sub

Private Sub UserForm_Activate()
   
   ' remove UserForm caption and borders
   Self.RemoveCaptionAndBorders
   
   ' round the UserForm corners
   Self.RoundBorderCorners

End Sub

' RepaintCorners, exposed to repaint the UserForm corners should a Modeless UserForm be Minimised and then restored with its Excel Application Window
Public Sub RepaintCorners()
   Self.RoundBorderCorners
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   
   ' allow the UserForm to be dragged
   Self.AllowDrag

End Sub

'@Ignore EmptyMethod
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
   ' Optional UserForm Close
   'Unload Me
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   Me.CloseCrossCircle.Visible = False
End Sub

' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' -----¬ End of UserForm
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'
'
'
'
'
'
'
'
'
'
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' -----¬ Sliders, requires the external mdUserFormSlider Code Module
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub Slider1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   DragSlider SliderControl:=Me.Slider1, UpdateRange:=ThisWorkbook.Sheets("Sheet1").Range("B17"), Minimum:=1, Maximum:=10
End Sub

Private Sub Slider1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   ' //  instantiate our Cursor Class into an object
   '     we will avoid auto-instancing
   Dim MouseCursor As Cursor
   Set MouseCursor = New Cursor
   MouseCursor.AddCursor IDC_HAND
End Sub

Private Sub Slider2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   DragSlider SliderControl:=Me.Slider2, UpdateRange:=ThisWorkbook.Sheets("Sheet1").Range("B19"), Minimum:=1, Maximum:=5
End Sub

Private Sub Slider2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   ' //  instantiate our Cursor Class into an object
   '     we will avoid auto-instancing
   Dim MouseCursor As Cursor
   Set MouseCursor = New Cursor
   MouseCursor.AddCursor IDC_HAND
End Sub

' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' -----¬ End of Slider
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'
'
'
'
'
'
'
'
'
'
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' -----¬ Close Cross
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub CloseCross_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   If Button = 1 Then Running = False: Unload Me                          ' Me.Hide
End Sub

Private Sub CloseCross_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   Me.CloseCrossCircle.Visible = True
   Self.ChangeCursor IDC_HAND
End Sub

' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' -----¬ End of Close Cross
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'
'
'
'
'
'
'
'
'
'
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' -----¬ Checkboxes
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub Checkbox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   Self.ChangeCursor IDC_HAND
End Sub

Private Sub Checkbox1Ticked_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   Self.ChangeCursor IDC_HAND
End Sub

Private Sub Checkbox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   If Button = 1 Then
      Tick Me.Checkbox1
      [Checkbox1Value].Value2 = True
   End If
End Sub

Private Sub Checkbox1Ticked_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   If Button = 1 Then
      Untick Me.Checkbox1
      [Checkbox1Value].Value2 = False
   End If
End Sub

Private Sub Checkbox2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   Self.ChangeCursor IDC_HAND
End Sub

Private Sub Checkbox2Ticked_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   Self.ChangeCursor IDC_HAND
End Sub

Private Sub Checkbox2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   If Button = 1 Then
      Tick Me.Checkbox2
      [Checkbox2Value].Value2 = True
   End If
End Sub

Private Sub Checkbox2Ticked_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   If Button = 1 Then
      Untick Me.Checkbox2
      [Checkbox2Value].Value2 = False
   End If
End Sub

Private Sub Checkbox3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   Self.ChangeCursor IDC_HAND
End Sub

Private Sub Checkbox3Ticked_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   Self.ChangeCursor IDC_HAND
End Sub

Private Sub Checkbox3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   If Button = 1 Then
      Tick Me.Checkbox3
      [Checkbox3Value].Value2 = True
   End If
End Sub

Private Sub Checkbox3Ticked_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   If Button = 1 Then
      Untick Me.Checkbox3
      [Checkbox3Value].Value2 = False
   End If
End Sub

Private Sub Checkbox4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   Self.ChangeCursor IDC_HAND
End Sub

Private Sub Checkbox4Ticked_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   Self.ChangeCursor IDC_HAND
End Sub

Private Sub Checkbox4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   If Button = 1 Then
      Tick Me.Checkbox4
      [Checkbox4Value].Value2 = True
   End If
End Sub

Private Sub Checkbox4Ticked_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   If Button = 1 Then
      Untick Me.Checkbox4
      [Checkbox4Value].Value2 = False
   End If
End Sub

Private Sub Tick(ByVal UserFormControl As Control)
   Me.Controls(UserFormControl.Name & "Ticked").ZOrder msoBringToFront
   Me.Controls(UserFormControl.Name).ZOrder msoSendToBack
End Sub

Private Sub Untick(ByVal UserFormControl As Control)
   Me.Controls(UserFormControl.Name & "Ticked").ZOrder msoSendToBack
   Me.Controls(UserFormControl.Name).ZOrder msoBringToFront
End Sub

' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' -----¬ End of Checkboxes
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'
'
'
'
'
'
'
'
'
'

Private Sub Toggle1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   Self.ChangeCursor IDC_HAND
End Sub

Private Sub Toggle1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   If Button = 1 Then
      Dim Direction As Boolean
      Direction = Toggle(Me.Toggle1)
      [Toggle1Value].Value2 = Not Direction
   End If
End Sub

Private Function Toggle(ByVal UserFormControl As Control, Optional ByRef StartPosition As Boolean = False) As Boolean
   On Error Resume Next
   Dim TweenTime As Double
   TweenTime = 0
   Dim Tweening As Double
   Dim ButtonMoveFrom As Long
   Dim ButtonMoveTo As Long
   Dim MoveAmount As Long
   MoveAmount = 14
   Dim Direction As Boolean
   ' assign the direction status which will be the last know toggle ie. On / Off, Left or Right
   Direction = [Toggle1Value].Value2
   
   ' only ran if set to True.  this is used to position the Toggle Button on initialisation and set the Label
   If StartPosition Then
      If Direction Then
         Me.Controls(UserFormControl.Name).Left = Me.Controls(UserFormControl.Name).Left + MoveAmount
         Me.Controls(UserFormControl.Name & "Label") = "On"
      Else
         Me.Controls(UserFormControl.Name & "Label") = "Off"
      End If
      Exit Function
   End If
   
   ButtonMoveFrom = Me.Controls(UserFormControl.Name).Left
   If Direction Then
      ButtonMoveTo = Me.Controls(UserFormControl.Name).Left - MoveAmount
   Else
      ButtonMoveTo = Me.Controls(UserFormControl.Name).Left + MoveAmount
   End If
   
   Do
      DoEvents
      
      ' easing, .easeOutQuintic
      Me.Controls(UserFormControl.Name).Left = easeOutQuintic(TweenTime, ButtonMoveFrom, ButtonMoveTo - ButtonMoveFrom, 1)
   
      ' tween refresh rate ie. 24 per second, 1 / 24 = 0.0416666666666667
      TweenTime = TweenTime + 1 / 12
      Tweening = Timer
      Do
         DoEvents
         ' timer refresh rate ie. 40 per second, 1 / 40 = 0.025
      Loop While Timer - Tweening < 1 / 40
   Loop Until TweenTime >= 1
   
   ' output the opposite of our eased direction
   If Not Direction Then Me.Controls(UserFormControl.Name & "Label") = "On" Else Me.Controls(UserFormControl.Name & "Label") = "Off"
   Toggle = Direction
   On Error GoTo 0
End Function

Private Function easeOutQuintic(ByVal t As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Variant
   Dim ts As Double
   ts = (t / d) * t
   Dim tc As Double
   tc = ts * t
   easeOutQuintic = b + c * (tc * ts + -5 * ts * ts + 10 * tc + -10 * ts + 5 * t)
End Function

'
'
'
'
'
'
'
'
'
'
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' -----¬ Radiobuttons
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub A1Radiobutton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   Self.ChangeCursor IDC_HAND
End Sub

Private Sub A1RadiobuttonChecked_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   Self.ChangeCursor IDC_HAND
End Sub

Private Sub A1Radiobutton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   If Button = 1 Then
      Check Me.A1Radiobutton
   End If
End Sub

Private Sub A1RadiobuttonChecked_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   If Button = 1 Then
      Check Me.A1Radiobutton
   End If
End Sub

Private Sub A2Radiobutton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   Self.ChangeCursor IDC_HAND
End Sub

Private Sub A2RadiobuttonChecked_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   Self.ChangeCursor IDC_HAND
End Sub

Private Sub A2Radiobutton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   If Button = 1 Then
      Check Me.A2Radiobutton
   End If
End Sub

Private Sub A2RadiobuttonChecked_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   If Button = 1 Then
      Check Me.A2Radiobutton
   End If
End Sub

Private Sub Check(ByVal UserFormControl As Control)
   Me.Controls(UserFormControl.Name & "Checked").ZOrder msoBringToFront
   Me.Controls(UserFormControl.Name).ZOrder msoSendToBack
   
   Dim Group As String
   Group = Left$(UserFormControl.Name, 1)
   Dim ID As Long
   Dim TempID As Long
   ID = Int(Right$(Left$(UserFormControl.Name, 2), 1))
   Application.Evaluate("Radiobutton" & ID & "Value").Value2 = True
   Application.Evaluate("Radiobutton" & Group & "Selected").Value2 = ID
   
   Dim Ctrl As Control
   For Each Ctrl In Me.Controls
      If Ctrl.Name <> UserFormControl.Name Then
         If Left$(Ctrl.Name, 1) = Group Then
            If Int(Right$(Left$(Ctrl.Name, 2), 1)) <> ID Then
               TempID = Int(Right$(Left$(Ctrl.Name, 2), 1))
               Me.Controls(Group & TempID & "RadiobuttonChecked").ZOrder msoSendToBack
               Me.Controls(Group & TempID & "Radiobutton").ZOrder msoBringToFront
               Application.Evaluate("Radiobutton" & TempID & "Value").Value2 = False
            End If
         End If
      End If
   Next Ctrl
End Sub

