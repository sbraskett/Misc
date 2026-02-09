Attribute VB_Name = "mdUserFormSlider"
'@IgnoreModule MoveFieldCloserToUsage

Option Explicit
Option Private Module

#If VBA7 Then
   Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
   Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
   Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As Long
   Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
   Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
   Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
   Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
   Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Type POINTAPI
   X As Long
   Y As Long
End Type

#Else
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetAsyncKeystate Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As Point) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Type POINT
   X As Long
   Y As Long
End Type

#End If

Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const TWIPSPERINCH As Long = 1440

' the one and only Interface, used to hold all local variables
Private Type TInterface
   PRunning As Boolean
   PXPixelsPerInch As Long
   PYPixelsPerInch As Long
   PUserForm As UserForm
   PSliderButton As Control
   PSliderBar As Control
   PPlaceholder As Control
   PValue As Control
   PCursor As Cursor
   PSliderButtonLeft As Single
   PSliderButtonTop As Single
   PSliderButtonMinimum As Single
   PSliderButtonMaximum As Single
   PRangeMinimum As Single
   PRangeMaximum As Single
   PPosition As Single
   PUpdateRange As Range
   #If VBA7 Then
   PInitialMousePosition As POINTAPI
   PTimerId As LongPtr
   #Else
   PInitialMousePosition As POINT
   PTimerId As Long
   #End If
End Type

Private this As TInterface
Private that As TInterface

'Private Self As New Cursor

' Pause, used to stop time for a small interval ;)
Private Sub Pause(ByVal Milliseconds As Long)
   Dim TickCount As Long: TickCount = GetTickCount
   Do: DoEvents: Loop Until TickCount + Milliseconds < GetTickCount
End Sub

' ScaleBetweenRange, scales the Slider Button between the Slider Placeholder Left and the Slider Placeholder Width within a Range specified
Public Function ScaleBetweenRange(ByRef SliderPosition As Variant, ByRef RangeMin As Variant, ByRef RangeMax As Variant, ByRef PlaceholderMin As Variant, ByRef PlaceholderMax As Variant) As Variant
   ScaleBetweenRange = (RangeMax - RangeMin) * (SliderPosition - PlaceholderMin) / (PlaceholderMax - PlaceholderMin) + RangeMin
End Function

' InitialiseSlider, set up for each Slider
' in order to maintain this correctly we also need to store the current values, see the UserForm Code for how we retrieve and Initialiase these using this Subroutine
Public Sub InitialiseSlider(ByRef SliderControl As Control, ByVal SliderPosition As Variant, Optional ByRef Minimum As Long = 1, Optional ByRef Maximum As Long = 100)
   'Debug.Print SliderControl.Name
   'Debug.Print SliderControl.Parent.Name
   
   ' the use of On Error Resume Next here means that we don't neccesarily have to maintain a Label for our Sliders!
   On Error Resume Next
   With SliderControl.Parent
   
      Dim Step As Variant
      Step = .Controls(SliderControl.Name & "Placeholder").Width / Maximum
   
      If SliderPosition <= Minimum Then
         .Controls(SliderControl.Name).Left = .Controls(SliderControl.Name & "Placeholder").Left - (.Controls(SliderControl.Name).Width / 2)
         .Controls(SliderControl.Name & "Bar").Left = .Controls(SliderControl.Name & "Placeholder").Left
         .Controls(SliderControl.Name & "Bar").Width = 1
      ElseIf SliderPosition >= Maximum Then
         .Controls(SliderControl.Name).Left = .Controls(SliderControl.Name & "Placeholder").Left + .Controls(SliderControl.Name & "Placeholder").Width - (.Controls(SliderControl.Name).Width / 2)
         .Controls(SliderControl.Name & "Bar").Left = .Controls(SliderControl.Name & "Placeholder").Left
         .Controls(SliderControl.Name & "Bar").Width = .Controls(SliderControl.Name & "Placeholder").Width
      Else
         .Controls(SliderControl.Name).Left = .Controls(SliderControl.Name & "Placeholder").Left + (SliderPosition * Step) - (.Controls(SliderControl.Name).Width / 2)
         .Controls(SliderControl.Name & "Bar").Left = .Controls(SliderControl.Name & "Placeholder").Left
         .Controls(SliderControl.Name & "Bar").Width = .Controls(SliderControl.Name).Left - (.Controls(SliderControl.Name & "Placeholder").Left - (.Controls(SliderControl.Name).Width / 2))
      End If
      .Controls(SliderControl.Name & "Value").Caption = SliderPosition
   End With
   On Error GoTo 0
End Sub

' DragSlider, fires once the main Slider Button Object is dragged by the Mouse. pass in the specific UserForm SliderButton Control that is moved
Public Sub DragSlider(ByRef SliderControl As Control, Optional ByRef UpdateRange As Range, Optional ByRef Minimum As Long = 1, Optional ByRef Maximum As Long = 100)
   'Debug.Print SliderControl.Name
   'Debug.Print SliderControl.Parent.Name
   
   On Error Resume Next
   If this.PRunning Then
      ReleaseSliderButton
   Else
   
      Dim hDCHwnd As Long
      hDCHwnd = GetDC(0)
      this.PXPixelsPerInch = GetDeviceCaps(hDCHwnd, LOGPIXELSX)
      this.PYPixelsPerInch = GetDeviceCaps(hDCHwnd, LOGPIXELSY)
      ReleaseDC 0, hDCHwnd
      
      this.PRangeMinimum = Minimum
      this.PRangeMaximum = Maximum
      Set this.PUpdateRange = UpdateRange                                 ' will be error skipped if not used
      
      Set this.PUserForm = SliderControl.Parent                           ' future dev
      With this.PUserForm
         Set this.PSliderButton = .Controls(SliderControl.Name)
         Set this.PPlaceholder = .Controls(SliderControl.Name & "Placeholder")
         Set this.PSliderBar = .Controls(SliderControl.Name & "Bar")
         Set this.PValue = .Controls(SliderControl.Name & "Value")
      
         this.PSliderButtonLeft = this.PSliderButton.Left
         this.PSliderButtonTop = this.PSliderButton.Top
         this.PSliderButtonMinimum = this.PPlaceholder.Left - (this.PSliderButton.Width / 2)
         this.PSliderButtonMaximum = this.PPlaceholder.Left + this.PPlaceholder.Width - (this.PSliderButton.Width / 2)
      End With
       
      GetCursorPos this.PInitialMousePosition
      Pause 20
      StartTimer
      this.PRunning = True
   End If
   On Error GoTo 0
End Sub

' ReleaseSliderButton, fired upon releasing the Slider Button and terminates the Interface
Public Sub ReleaseSliderButton()
   On Error Resume Next
   this.PRunning = False
   StopTimer
   Pause 20
   On Error GoTo 0
End Sub

' StartTime, starts the timer
Private Sub StartTimer()
   On Error Resume Next
   this.PTimerId = SetTimer(0, 0, 20, AddressOf TimerProc)
   On Error GoTo 0
End Sub

' StopTimer, stops the timer
Private Sub StopTimer()
   On Error Resume Next
   KillTimer 0, this.PTimerId
   ' this method will reset a Type structure
   this = that
   On Error GoTo 0
End Sub

' TimerProc, the one and only timer callback ;)
'@Ignore HungarianNotation
Private Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
   #If VBA7 Then
      Dim MousePosition As POINTAPI
   #Else
      Dim MousePosition As POINT
   #End If
   Dim MouseXAdjustment As Single
   '@Ignore VariableNotUsed
   Dim MouseYAdjustment As Single
   On Error Resume Next
   If GetAsyncKeyState(1) = 0 Then
      ReleaseSliderButton
   End If
   If this.PRunning Then
      ' derive the new Cursor position for the Slider Button
      GetCursorPos MousePosition
      MouseXAdjustment = (MousePosition.X - this.PInitialMousePosition.X) * TWIPSPERINCH / 20 / this.PXPixelsPerInch
      MouseYAdjustment = (MousePosition.Y - this.PInitialMousePosition.Y) * TWIPSPERINCH / 20 / this.PYPixelsPerInch
      
      this.PSliderButton.Top = this.PSliderButtonTop
      this.PSliderButton.Left = this.PSliderButtonLeft + MouseXAdjustment
      
      If this.PSliderButton.Left < this.PSliderButtonMinimum Then
         this.PSliderButton.Left = this.PSliderButtonMinimum
         this.PSliderBar.Left = this.PPlaceholder.Left
         this.PSliderBar.Width = 1
      End If
      
      If this.PSliderButton.Left > this.PSliderButtonMaximum Then
         this.PSliderButton.Left = this.PSliderButtonMaximum
         this.PSliderBar.Left = this.PPlaceholder.Left
         this.PSliderBar.Width = this.PPlaceholder.Width
      End If
      
      this.PSliderBar.Width = this.PSliderButton.Left - this.PSliderButtonMinimum
      this.PPosition = ScaleBetweenRange(SliderPosition:=this.PSliderBar.Width, RangeMin:=this.PRangeMinimum, RangeMax:=this.PRangeMaximum, PlaceholderMin:=1, PlaceholderMax:=this.PPlaceholder.Width)
      this.PValue.Caption = Round(this.PPosition, 0)
      this.PUpdateRange.Value2 = Round(this.PPosition, 0)                 ' will be error skipped if not used
   End If
   On Error GoTo 0
End Sub

