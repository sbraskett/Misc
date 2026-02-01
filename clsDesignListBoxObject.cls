VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDesignListBoxObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents pLabl As MSForms.Label
Attribute pLabl.VB_VarHelpID = -1
Private pParent As clsDesignListBox
Private pSuppressClick As Boolean
Private pLastX As Single
Private pWasHeaderHover As Boolean
Private pIsHeaderPressed As Boolean
Private Const RESIZE_HIT As Long = 6  'pixels from right edg


Private Function IsNearRightEdge(ByVal X As Single) As Boolean
    IsNearRightEdge = (pLabl.Width - X) <= RESIZE_HIT
End Function

Private Function HeaderColIndex() As Long
    Dim parts() As String
    parts = Split(pLabl.Name, ";")
    If UBound(parts) >= 1 Then
        If IsNumeric(parts(1)) Then
            HeaderColIndex = CLng(parts(1))  '0-based
            Exit Function
        End If
    End If
    HeaderColIndex = -1
End Function

Public Sub Create(InpLabel As MSForms.Label, InpParent As clsDesignListBox)
    Set pLabl = InpLabel
    Set pParent = InpParent
End Sub
Private Sub pLabl_MouseDown3(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    'Prevent the subsequent Click event from re-processing
    'pSuppressClick = True

    Call pParent.RaiseEventBeforeClick
    If pLabl.Tag = "GridLine" Then Exit Sub

    If pLabl.Tag = "Header" Then

        Dim parts() As String
        parts = Split(pLabl.Name, ";")
        If UBound(parts) >= 1 And IsNumeric(parts(1)) Then

            Dim colIdx As Long
            colIdx = CLng(parts(1))   '0-based column

            '--- resize if near right edge ---
            If IsNearRightEdge(X) Then
                'startX should be frame-relative: label.Left + X
                pParent.BeginResize colIdx, (pLabl.Left + X)
                Exit Sub
            End If

            '--- otherwise normal header click sort ---
            Call pParent.HeaderSort(colIdx)
        End If

        Call pParent.RaiseEventClick
        Exit Sub
    End If

    '--- NORMAL CELL CLICK = SELECTION ---
    Dim RowNumber As Long
    RowNumber = CLng(Split(pLabl.Name, ";")(0))

    'Shift bitmask: 1=Shift, 2=Ctrl, 4=Alt
    Dim isShift As Boolean, isCtrl As Boolean
    isShift = ((Shift And 1) <> 0)
    isCtrl = ((Shift And 2) <> 0)

    Call pParent.SelectRow(RowNumber, isCtrl, isShift)
    Call pParent.RaiseEventClick

End Sub


Private Sub pLabl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Call pParent.RaiseEventBeforeClick

    If pLabl.Tag = "GridLine" Then Exit Sub

    '--- HEADER: RESIZE has priority ---
    If pLabl.Tag = "Header" Then

        Dim colIdx As Long
        colIdx = HeaderColIndex()
        If colIdx < 0 Then Exit Sub

        If IsNearRightEdge(X) Then
            'Begin resize instead of sorting
            pLabl.MousePointer = fmMousePointerSizeWE
            pParent.BeginResize colIdx, pLabl.Left + X   '<<< you must have BeginResize in parent
            Exit Sub
        End If

        'Not resizing -> treat as sort click
        pLabl.BackColor = pParent.HeaderPressedBackColor
        pParent.HeaderSort colIdx
        pLabl.BackColor = pParent.HeaderHoverBackColor

        Call pParent.RaiseEventClick
        'pParent.FocusKeySink
        pParent.TakeKeyboardFocus
        Exit Sub
    End If

    '--- NORMAL CELL CLICK = SELECTION ---
    Dim RowNumber As Long
    RowNumber = CLng(Split(pLabl.Name, ";")(0))

    'Shift bitmask: 1=Shift, 2=Ctrl, 4=Alt
    Dim isShift As Boolean, isCtrl As Boolean
    isShift = ((Shift And 1) <> 0)
    isCtrl = ((Shift And 2) <> 0)

    Call pParent.SelectRow(RowNumber, isCtrl, isShift)
    Call pParent.RaiseEventClick
    'pParent.FocusKeySink
    pParent.TakeKeyboardFocus
End Sub


Private Sub pLabl_Click()
    'MouseDown already did the work; ignore the follow-on Click
    If pSuppressClick Then
        pSuppressClick = False
        Exit Sub
    End If
End Sub
Private Sub pLabl_MouseMove_null(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    pLastX = X
    'pParent.NotifyMouseOverLabel
    
    '--- HEADER: resize cursor + drag sizing ---
    If pLabl.Tag = "Header" Then

        'Cursor feedback
        If pParent.IsResizing Or IsNearRightEdge(X) Then
            pLabl.MousePointer = fmMousePointerSizeWE
        Else
            pLabl.MousePointer = fmMousePointerUpArrow
        End If

        'Drag resize
        If pParent.IsResizing Then
            pParent.DragResize (pLabl.Left + X)
        End If

        Exit Sub
    End If

    'Ignore gridlines
    If pLabl.Tag = "GridLine" Then Exit Sub

    'If resizing is active, don't do hover
    If pParent.IsResizing Then Exit Sub

    '--- ROW HOVER for normal cells ---
    Dim parts() As String
    parts = Split(pLabl.Name, ";")
    If UBound(parts) < 1 Then Exit Sub
    If Not IsNumeric(parts(0)) Then Exit Sub

    pParent.SetHoverRow CLng(parts(0))

End Sub
Private Sub pLabl_MouseMove3(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    pLastX = X
    pParent.NotifyMouseOverLabel
    
    '--- HEADER: resize cursor + drag sizing ---
    If pLabl.Tag = "Header" Then

        'Cursor feedback
        If pParent.IsResizing Or IsNearRightEdge(X) Then
            pLabl.MousePointer = fmMousePointerSizeWE
        Else
            pLabl.MousePointer = fmMousePointerUpArrow
        End If

        'Drag resize
        If pParent.IsResizing Then
            pParent.DragResize (pLabl.Left + X)
        End If

        Exit Sub
    End If

    'Ignore gridlines
    If pLabl.Tag = "GridLine" Then Exit Sub

    'If resizing is active, don't do hover
    If pParent.IsResizing Then Exit Sub

    '--- ROW HOVER for normal cells ---
    Dim parts() As String
    parts = Split(pLabl.Name, ";")
    If UBound(parts) < 1 Then Exit Sub
    If Not IsNumeric(parts(0)) Then Exit Sub

    pParent.SetHoverRow CLng(parts(0))

End Sub
Private Sub pLabl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pLastX = X
    '--- HEADER: resize cursor + drag sizing ---

    'If we're resizing, keep resizing no matter what
    If pParent.IsResizing Then
        pParent.DragResize (pLabl.Left + X)
        Exit Sub
    End If

    If pLabl.Tag = "Header" Then
        'Cursor feedback has priority over hover visuals
        If IsNearRightEdge(X) Then
            pLabl.MousePointer = fmMousePointerSizeWE
            'optional: don't apply hover while on the boundary
            pLabl.BackColor = pParent.HeaderBackColor
        Else
            pLabl.MousePointer = fmMousePointerUpArrow
            'optional hover
            pLabl.BackColor = pParent.HeaderHoverBackColor
        End If
        
        If pLabl.Tag = "Header" And Not pParent.IsResizing Then
            pLabl.BackColor = RGB(225, 230, 240)
        End If

    Else
    
        If pLabl.Tag <> "Header" And pLabl.Tag <> "GridLine" Then
            pParent.ResetHeaderHover
        End If
    
        If pParent.IsResizing Then
            pParent.DragResize (pLabl.Left + X)
        End If
    End If
    
    ' added for hover
    'Ignore gridlines
    If pLabl.Tag = "GridLine" Then Exit Sub

    'If resizing is active, don't do hover
    If pParent.IsResizing Then Exit Sub

    '--- ROW HOVER for normal cells ---
    Dim parts() As String
    parts = Split(pLabl.Name, ";")
    If UBound(parts) < 1 Then Exit Sub
    If Not IsNumeric(parts(0)) Then Exit Sub

    pParent.SetHoverRow CLng(parts(0))



End Sub

Private Sub pLabl_MouseUp3(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If pParent.IsResizing Then
        pParent.EndResize
    End If
End Sub
Private Sub pLabl_MouseUp4(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If pLabl.Tag = "Header" Then
        pIsHeaderPressed = False
        pLabl.BackColor = pParent.HeaderHoverBackColor
    End If
End Sub
Private Sub pLabl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If pParent.IsResizing Then
        pParent.EndResize
    End If
    If pLabl.Tag = "Header" Then
        pLabl.BackColor = RGB(240, 240, 240)
    End If
End Sub

Private Function NextDescendingFromCaption(ByVal cap As String) As Boolean
    cap = Trim$(cap)

    'Your Sort code:
    'Descending=True  -> adds ? (2191)
    'Descending=False -> adds ? (2193)

    If Len(cap) = 0 Then
        NextDescendingFromCaption = False
        Exit Function
    End If

    Select Case Right$(cap, 1)
        Case ChrW(&H2191) '? currently descending -> next ascending
            NextDescendingFromCaption = False
        Case ChrW(&H2193) '? currently ascending -> next descending
            NextDescendingFromCaption = True
        Case Else
            'No arrow yet -> first click ascending (Descending=False)
            NextDescendingFromCaption = False
    End Select
End Function

Private Sub pLabl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If pLabl.Tag <> "Header" Then Exit Sub
    If (pLabl.Width - pLastX) > 6 Then Exit Sub   '6px boundary hit

    Dim parts() As String
    parts = Split(pLabl.Name, ";")
    If UBound(parts) >= 1 And IsNumeric(parts(1)) Then
        pParent.AutoFitColumn CLng(parts(1))
    End If

End Sub
