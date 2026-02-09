Attribute VB_Name = "BoX"

' Software Name:
' BoX
' © Copyright/Author:
' Mark Kubiszyn. All Rights Reserved
' Website/Follow/Help:
' https://www.kubiszyn.co.uk/
' https://www.facebook.com/Kubiszyn.co.uk/
' https://www.kubiszyn.co.uk/documentation/box.html
'
' License:
' This Software is released under an MIT License (MIT)
' https://www.kubiszyn.co.uk/license.html
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files
' (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge,
' publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so,
' subject to the following conditions:
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
' MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR
' ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
' SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE

Option Explicit
Option Private Module

' used for the Demo so that you can't launch multiple dialogs!
'@Ignore EncapsulatePublicField
Public Running As Boolean

' Example1, an example of a Modeless BoX Dialog without a Lightbox
Public Sub Example1()

   If Not Running Then
      Running = True
      
      With UserForm1
         ' you can adjust the UserForm common members here like the .Height or .Width
         .Height = 200
         .Show 0
      End With
   
   End If
   
End Sub

' Example2, an example of a Modal BoX Dialog with a Lightbox
Public Sub Example2()

   If Not Running Then
      Running = True
      
      Dim Effect As Lightbox
      Set Effect = New Lightbox
      Effect.TransitionIn FitToExcel, easeOutSine, 160

      With UserForm1
         ' you can adjust the UserForm common members here like the .Height or .Width
         .Height = 200
         .Show 1                                                          ' 1 = Modal
      End With
      Effect.TransitionOut easeOutSine, 160
   
   End If

End Sub

