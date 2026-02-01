VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDesignListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pHeadersCol As Collection
Private pDesignListBoxObjectsCol As New Collection
Private pFreezeColumnsCol As Collection
Private pFreezeRowsCol As Collection
Private pInpArr() As Variant
Private pParent As UserForm
Private WithEvents pFrame As MSForms.Frame
Attribute pFrame.VB_VarHelpID = -1
Private WithEvents pKeySink As MSForms.TextBox
Attribute pKeySink.VB_VarHelpID = -1
Private pKeySinkName As String
Private pIsEmpty As Boolean
Private pGridLineColor As Long
Private pSelectedRows As Object
Private pParamDict As Object
Private pAnchorRow As Long
Private pLastSortCol As Long
Private pLastSortDescending As Boolean
Private pIsResizing As Boolean
Private pResizeCol As Long          '0-based column index being resized
Private pResizeStartX As Single
Private pResizeStartW As Long
Private pMinColWidth As Long
Private Const AUTOFIT_PAD As Long = 10
Private pActiveRow As Long

Private pHoverRow As Long
Private pHoverColor As Long
Private pHoverEnabled As Boolean
Private pHoverSteps As Long
Private pHoverDelayMs As Long
Private pHoverToken As Long
Private pLastLabelMove As Single   'Timer() when label last moved

'--- Filtering ---
Private pBaseArr() As Variant          'original unfiltered data (2D)
Private pRowMap() As Long              'display row -> source row (0-based, excluding header row)
Private pFilters As Collection         'each item is a filter rule array

'=== clsDesignListBox ===
Private pHeaderBackColor As Long
Private pHeaderForeColor As Long
Private pHeaderBorderColor As Long
Private pHeaderHoverBackColor As Long
Private pHeaderPressedBackColor As Long
Private pHeaderSortMarkColor As Long
Private pHeaderMask As MSForms.Label

Private pRowHeights() As Single
Private pMeasureLbl As MSForms.Label

'Public Events
Public Event BeforeClick()
Public Event Change()
Public Event Click()

'--- Multi-select support ---
Public Enum eDLBMultiSelectMode
    dlbSingle = 0          'existing behavior
    dlbMulti = 1           'click toggles
    dlbExtended = 2        'Ctrl toggles; Shift ranges
End Enum
Private pInFill As Boolean
Private pInLayout As Boolean
Private pInColumnWidthsLet As Boolean
Private pMultiSelectMode As eDLBMultiSelectMode
Private pFormatRules As Collection


Public Property Get IsDLBEmpty() As Boolean
    IsDLBEmpty = pIsEmpty
End Property
Public Property Get IsListEmpty() As Boolean
    IsListEmpty = pIsEmpty
End Property
Private Sub Class_Initialize()

    '--- core state ---
    pIsEmpty = True
    pAnchorRow = -1
    pLastSortCol = -1
    pLastSortDescending = False
    pMultiSelectMode = dlbExtended
    pHoverRow = -1
    pHoverColor = RGB(242, 247, 255)   'very light blue (Excel-ish)
    pHoverEnabled = True
    pHoverSteps = 6
    pHoverDelayMs = 0
    pHoverToken = 0
    pLastLabelMove = 0
    pActiveRow = -1
    
    '--- collections ---
    Set pFreezeColumnsCol = New Collection
    Set pFreezeRowsCol = New Collection
    Set pHeadersCol = New Collection
    Set pDesignListBoxObjectsCol = New Collection
    Set pFormatRules = New Collection
    Set pFilters = New Collection
    
    '--- dictionaries (LATE BINDING) ---
    Set pSelectedRows = CreateObject("Scripting.Dictionary")
    pSelectedRows.CompareMode = vbTextCompare

    Set pParamDict = CreateObject("Scripting.Dictionary")
    pParamDict.CompareMode = vbTextCompare

    '--- defaults stored in param dict (avoid Key not found later) ---
    pParamDict("SelectionColor") = RGB(221, 235, 247)
    pParamDict("ColumnSource") = 0
    pParamDict("Headers") = False
    pParamDict("FreezeRows") = 0
    pParamDict("FreezeColumns") = 0
    pParamDict("RowHeight") = 12
    pParamDict("ColumnWidths") = vbNullString
    pParamDict("ColumnsCount") = 0
    pParamDict("RowsCount") = 0
    pParamDict("DeviationRows") = 0
    pParamDict("DeviationColumns") = 0
    InitHeaderThemeDefaults
End Sub
Public Property Get HoverEnabled() As Boolean
    HoverEnabled = pHoverEnabled
End Property
Public Property Let HoverEnabled(ByVal v As Boolean)
    pHoverEnabled = v
    If Not v Then ClearHover
End Property

Public Property Get HeaderBackColor() As Long: HeaderBackColor = pHeaderBackColor: End Property
Public Property Let HeaderBackColor(ByVal v As Long): pHeaderBackColor = v: End Property

Public Property Get HeaderHoverBackColor() As Long: HeaderHoverBackColor = pHeaderHoverBackColor: End Property
Public Property Let HeaderHoverBackColor(ByVal v As Long): pHeaderHoverBackColor = v: End Property

Public Property Get HeaderPressedBackColor() As Long: HeaderPressedBackColor = pHeaderPressedBackColor: End Property
Public Property Let HeaderPressedBackColor(ByVal v As Long): pHeaderPressedBackColor = v: End Property

Public Property Get HeaderBorderColor() As Long: HeaderBorderColor = pHeaderBorderColor: End Property
Public Property Let HeaderBorderColor(ByVal v As Long): pHeaderBorderColor = v: End Property

Public Property Get HeaderForeColor() As Long: HeaderForeColor = pHeaderForeColor: End Property
Public Property Let HeaderForeColor(ByVal v As Long): pHeaderForeColor = v: End Property

Public Property Get HeaderSortMarkColor() As Long: HeaderSortMarkColor = pHeaderSortMarkColor: End Property
Public Property Let HeaderSortMarkColor(ByVal v As Long): pHeaderSortMarkColor = v: End Property
Public Property Let HoverFadeSteps(ByVal v As Long)
    If v < 1 Then v = 1
    pHoverSteps = v
End Property

Public Property Let HoverFadeDelayMs(ByVal v As Long)
    If v < 0 Then v = 0
    pHoverDelayMs = v
End Property
Public Property Get HoverColor() As Long
    HoverColor = pHoverColor
End Property
Public Property Let HoverColor(ByVal v As Long)
    pHoverColor = v
End Property
Public Property Get MultiSelectMode() As eDLBMultiSelectMode
    MultiSelectMode = pMultiSelectMode
End Property
Public Property Get SelectedRowCount() As Long
    If pSelectedRows Is Nothing Then
        SelectedRowCount = 0
    Else
        SelectedRowCount = pSelectedRows.Count
    End If
End Property


Public Property Get SelectedRows() As Variant
    Dim k As Variant
    
    If pSelectedRows Is Nothing Or pSelectedRows.Count = 0 Then
        SelectedRows = Empty
        Exit Property
    End If

    Dim arr() As Variant, i As Long
    ReDim arr(0 To pSelectedRows.Count - 1)

    i = 0
    For Each k In pSelectedRows.Keys
        arr(i) = CLng(k)
        i = i + 1
    Next k

    SelectedRows = arr
End Property
Public Property Get Frame() As MSForms.Frame
    Set Frame = pFrame
End Property
Public Property Get SelectedValues() As Collection
    Dim result As New Collection

    If pSelectedRows Is Nothing Or pSelectedRows.Count = 0 Then
        Set SelectedValues = result
        Exit Property
    End If

    Dim rowsArr As Variant, i As Long
    rowsArr = Me.SelectedRows

    If Not IsArray(rowsArr) Then
        Set SelectedValues = result
        Exit Property
    End If

    Call SortLongArrayInPlace(rowsArr)

    For i = LBound(rowsArr) To UBound(rowsArr)
        result.Add pInpArr(CLng(rowsArr(i)) + GetDeviation("Rows"), _
                           Me.ColumnSource + GetDeviation("Columns"))
    Next i

    Set SelectedValues = result
End Property
Private Sub SortLongArrayInPlace(ByRef v As Variant)
    If Not IsArray(v) Then Exit Sub

    On Error Resume Next
    Dim lb As Long, ub As Long
    lb = LBound(v)
    ub = UBound(v)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    Dim i As Long, j As Long
    Dim tmp As Long

    For i = lb + 1 To ub
        tmp = CLng(v(i))
        j = i - 1
        Do While j >= lb And CLng(v(j)) > tmp
            v(j + 1) = v(j)
            j = j - 1
        Loop
        v(j + 1) = tmp
    Next i
End Sub

Public Property Let MultiSelectMode(ByVal v As eDLBMultiSelectMode)
    pMultiSelectMode = v

    'If switching to single, collapse selection
    If v = dlbSingle Then
        Dim idx As Long: idx = Me.ListIndex
        Call ClearSelection
        If idx >= 0 Then
            pSelectedRows(CStr(idx)) = True
            pAnchorRow = idx
            Call RepaintSelection
        End If
    End If
End Property

Public Property Get ColumnsCount() As Long

    If pParamDict.Exists("ColumnsCount") Then
        ColumnsCount = CLng(pParamDict("ColumnsCount"))
    Else
        ColumnsCount = 0
    End If

End Property
Public Property Get RowsCount() As Long

'has to return 0 if empty. No Error Handler required
RowsCount = pParamDict("RowsCount")

End Property
Public Property Get RowHeight() As Long

'has to return 0 if empty. No Error Handler required
RowHeight = pParamDict("RowHeight")

End Property
Public Property Let RowHeight(ByVal InpHeight As Long)

    If InpHeight < 1 Then InpHeight = 1
    pParamDict("RowHeight") = InpHeight

    If pFrame Is Nothing Then Exit Property
    If Me.IsListEmpty Then Exit Property

    Dim c As Control
    Dim L As MSForms.Label
    Dim parts() As String
    Dim r As Long

    '1) Resize ONLY cell/header labels (NOT gridlines)
    For Each c In pFrame.Controls
        If TypeName(c) = "Label" Then
            Set L = c

            If L.Tag <> "GridLine" Then
                L.Height = InpHeight
            Else
                'keep gridlines as 1px
                L.Height = 1
            End If
        End If
    Next c

    '2) Reposition cell/header labels by row index
    For Each c In pFrame.Controls
        If TypeName(c) = "Label" Then
            Set L = c

            If L.Tag <> "GridLine" Then
                parts = Split(L.Name, ";")
                If UBound(parts) >= 1 Then
                    If IsNumeric(parts(0)) Then
                        r = CLng(parts(0))
                        L.Top = r * InpHeight
                    End If
                End If
            End If
        End If
    Next c

    '3) Update scroll height
    pFrame.ScrollHeight = (Me.RowsCount * InpHeight)

    '4) Redraw horizontal gridlines at the new row boundaries
    ApplyHorizontalGridlines IIf(pGridLineColor = 0, RGB(230, 230, 230), pGridLineColor)

End Property

Public Property Let RowHeight_Save(ByVal InpHeight As Long)
    If InpHeight < 1 Then InpHeight = 1

    pParamDict("RowHeight") = CLng(InpHeight)

    'If we already have data, recompute row heights and rebuild
    If Not Me.IsDLBEmpty Then
        On Error Resume Next
        ComputeRowHeights pInpArr     'recalc pRowHeights()
        Err.Clear
        On Error GoTo 0

        'Easiest + safest: rebuild controls to match new heights
        Me.Fill pInpArr
    End If
End Property
Public Property Get ColumnWidths() As String
    If pParamDict.Exists("ColumnWidths") Then
        ColumnWidths = CStr(pParamDict("ColumnWidths"))
    Else
        ColumnWidths = vbNullString
    End If
End Property

Public Property Let ColumnWidths(ByVal strWidths As String)

    Dim parts() As String
    Dim outParts() As String
    Dim i As Long
    Dim n As Long
    Dim colCount As Long
    Dim w As Long

    strWidths = Trim$(strWidths)

    'Remove trailing semicolons (one or many)
    Do While Len(strWidths) > 0 And Right$(strWidths, 1) = ";"
        strWidths = Left$(strWidths, Len(strWidths) - 1)
    Loop

    'If empty string -> treat as zero widths, but don't blow up
    If Len(strWidths) = 0 Then
        'If we don't know column count yet, nothing we can do
        If (Not pParamDict.Exists("ColumnsCount")) Or CLng(pParamDict("ColumnsCount")) <= 0 Then Exit Property
        colCount = CLng(pParamDict("ColumnsCount"))
        ReDim outParts(0 To colCount - 1)
        For i = 0 To colCount - 1
            outParts(i) = CStr(pMinColWidth)
        Next i
        pParamDict("ColumnWidths") = Join(outParts, ";")
        GoTo MaybeLayout
    End If

    'Split
    parts = Split(strWidths, ";")
    n = UBound(parts) - LBound(parts) + 1

    'Infer ColumnsCount if unknown/0
    If (Not pParamDict.Exists("ColumnsCount")) Or CLng(pParamDict("ColumnsCount")) <= 0 Then
        pParamDict("ColumnsCount") = n
    End If

    colCount = CLng(pParamDict("ColumnsCount"))
    If colCount <= 0 Then Exit Property

    'Normalize to exactly ColumnsCount
    ReDim outParts(0 To colCount - 1)

    For i = 0 To colCount - 1

        If i <= UBound(parts) Then
            outParts(i) = Trim$(parts(i))
        Else
            outParts(i) = CStr(pMinColWidth)
        End If

        If Len(outParts(i)) = 0 Then outParts(i) = CStr(pMinColWidth)

        If Not IsNumeric(outParts(i)) Then
            Err.Raise 123, "clsDesignListBox", _
                      "Input for ColumnWidths is not correct. Bad value: [" & outParts(i) & "]", _
                      "Correct notation example: '30;40;60;50'"
            Exit Property
        End If

        'Clamp to min width to prevent “autofit to tiny”
        w = CLng(val(outParts(i)))
        If w < pMinColWidth Then w = pMinColWidth
        outParts(i) = CStr(w)

    Next i

    'Persist
    pParamDict("ColumnWidths") = Join(outParts, ";")

MaybeLayout:
    'Relayout only if safe (prevents recursion/crash)
    If pInFill Then Exit Property
    If pInLayout Then Exit Property
    If pFrame Is Nothing Then Exit Property
    If Me.IsDLBEmpty Then Exit Property

    On Error GoTo ErrHandler
    pInLayout = True
    LayoutFromColumnWidths
    pInLayout = False
    Exit Property

ErrHandler:
    pInLayout = False
    Debug.Print "ColumnWidths Let error: "; Err.Number; Err.Description
End Property
Public Property Let ShowGrid(ByVal v As Boolean)
    Dim Labl
    For Each Labl In pFrame.Controls
        If Labl.Tag <> "Header" Then
            If v Then
                Labl.BorderStyle = fmBorderStyleSingle
                Labl.BorderColor = RGB(200, 200, 200)
            Else
                Labl.BorderStyle = fmBorderStyleNone
            End If
        End If
    Next Labl
End Property

Public Property Get GridLineColor() As Long
    GridLineColor = pGridLineColor
End Property

Public Property Let GridLineColor(ByVal v As Long)
    pGridLineColor = v
    If Not pFrame Is Nothing Then
        If Not Me.IsDLBEmpty Then ApplyHorizontalGridlines pGridLineColor
    End If
End Property

Public Property Get Headers() As Boolean
    Headers = pParamDict("Headers")
End Property
Public Property Let Headers(isOn As Boolean)

pParamDict("Headers") = isOn

If Not Me.IsDLBEmpty Then
    Call ActivateDeactivateHeaders(isOn)
    If Me.FreezeRows = 0 Then
        Me.FreezeRows = 1
    End If
End If

End Property
Public Property Get AllLabels() As Collection

Dim ResultCol As New Collection

Dim Labl
For Each Labl In pFrame.Controls
    ResultCol.Add Labl
Next Labl

Set AllLabels = ResultCol

End Property
Public Property Get RowLabels(RowNumber As Long, Optional InludingHeaders As Boolean) As Collection

Dim ResultCol As New Collection

If Not CheckInputRowNumber(RowNumber) Then Exit Property

Dim Labl
For Each Labl In pFrame.Controls
    If Not Labl.Tag = "Header" Or InludingHeaders Then
        If Split(Labl.Name, ";")(0) = RowNumber Then
            ResultCol.Add Labl
        End If
    End If
Next Labl

Set RowLabels = ResultCol

End Property
Public Property Get ColumnLabels(ColumnNumber As Long, Optional InludingHeaders As Boolean) As Collection

If Not CheckInputColumnNumber(ColumnNumber) Then Exit Property

Dim ResultCol As New Collection

Dim Labl
For Each Labl In pFrame.Controls
    If Not Labl.Tag = "Header" Or InludingHeaders Then
        If Split(Labl.Name, ";")(1) = ColumnNumber Then
            ResultCol.Add Labl
        End If
    End If
Next Labl

Set ColumnLabels = ResultCol

End Property
Public Property Get ExactLabel(RowNumber As Long, ColumnNumber As Long) As Object

If Not CheckInputRowNumber(RowNumber) Then Exit Property
If Not CheckInputColumnNumber(ColumnNumber) Then Exit Property

Dim Labl
For Each Labl In pFrame.Controls
    If Split(Labl.Name, ";")(0) = RowNumber And Split(Labl.Name, ";")(1) = ColumnNumber Then
        Set ExactLabel = Labl
        Exit Property
    End If
Next Labl

End Property
Public Property Get HeadersLabels() As Collection

Set HeadersLabels = pHeadersCol

End Property
Public Property Get ColumnSource() As Long

Dim TestNumber As Long
TestNumber = Round(pParamDict("ColumnSource"))

If TestNumber < 0 Then TestNumber = 0
If TestNumber > Me.ColumnsCount Then TestNumber = Me.ColumnsCount

ColumnSource = TestNumber

End Property
Public Property Let ColumnSource(ColumnNumber As Long)

pParamDict("ColumnSource") = Round(ColumnNumber)

End Property
Public Property Get TrueSelectedValue() As Variant

Dim ColumnSourceNumber As Long
ColumnSourceNumber = Me.ColumnSource

Dim Labl
For Each Labl In pFrame.Controls
    If Labl.BackColor = Me.SelectionColor And Split(Labl.Name, ";")(1) = ColumnSourceNumber Then
        SelectedValue = Labl.Caption
    End If
Next Labl

End Property
Public Property Get SelectedValue() As Variant

Dim ColumnSourceNumber As Long
ColumnSourceNumber = Me.ColumnSource + GetDeviation("Columns")

Dim RowNumber As Long
RowNumber = Me.ListIndex + GetDeviation("Rows")

SelectedValue = pInpArr(RowNumber, ColumnSourceNumber)

End Property

Public Property Get SelectionColor() As Long
    SelectionColor = pParamDict("SelectionColor")
End Property
Public Property Let SelectionColor(ColorNumber As Long)
    If Not Me.IsDLBEmpty Then Call DeselectAllLabels
    pParamDict("SelectionColor") = ColorNumber
End Property
Public Property Get ListIndex() As Long
    'Compatibility: return smallest selected row, or -1
    Dim minRow As Long, k As Variant
    ListIndex = -1

    If pSelectedRows Is Nothing Or pSelectedRows.Count = 0 Then Exit Property

    minRow = 2147483647
    For Each k In pSelectedRows.Keys
        If CLng(k) < minRow Then minRow = CLng(k)
    Next k
    ListIndex = minRow
End Property
Public Property Let ListIndex(RowNumber As Long)
    If RowNumber = -1 Then
        Call ClearSelection
    Else
        SelectRow RowNumber
    End If
End Property
Public Property Get FreezeRows() As Long

FreezeRows = pParamDict("FreezeRows")

End Property
Public Property Let FreezeRows(RowsFromTop As Long)

RowsFromTop = Round(RowsFromTop)
If RowsFromTop < 0 Then RowsFromTop = 0
If RowsFromTop > Me.RowsCount Then RowsFromTop = Me.RowsCount

pParamDict("FreezeRows") = RowsFromTop
Set pFreezeRowsCol = New Collection

pFrame.Scroll fmScrollActionBegin, fmScrollActionBegin

If RowsFromTop = 0 Then Exit Property

Dim i As Long
For i = 0 To RowsFromTop - 1
    pFreezeRowsCol.Add Me.RowLabels(i, True)
Next i

Dim Col, Labl

' removed to make this feature work..
'For Each Col In pFreezeColumnsCol
'    For Each Labl In Col
'        Labl.ZOrder msoBringToFront
'    Next Labl
'Next Col

For Each Col In pFreezeRowsCol
    For Each Labl In Col
        Labl.ZOrder msoBringToFront
    Next Labl
Next Col

End Property

Public Property Get IsResizing() As Boolean
    IsResizing = pIsResizing
End Property
Public Property Get FreezeColumns() As Long

FreezeColumns = pParamDict("FreezeColumns")

End Property
Public Property Let FreezeColumns(ColumnsFromLeft As Long)

ColumnsFromLeft = Round(ColumnsFromLeft)
If ColumnsFromLeft < 0 Then ColumnsFromLeft = 0
If ColumnsFromLeft > Me.ColumnsCount Then ColumnsFromLeft = Me.ColumnsCount

pParamDict("FreezeColumns") = ColumnsFromLeft
Set pFreezeColumnsCol = New Collection

pFrame.Scroll fmScrollActionBegin, fmScrollActionBegin

If ColumnsFromLeft = 0 Then Exit Property

Dim i As Long
For i = 0 To ColumnsFromLeft - 1
    pFreezeColumnsCol.Add Me.ColumnLabels(i, True)
Next i

Dim Col, Labl

For Each Col In pFreezeColumnsCol
    For Each Labl In Col
        Labl.ZOrder msoBringToFront
    Next Labl
Next Col

For Each Col In pFreezeRowsCol
    For Each Labl In Col
        Labl.ZOrder msoBringToFront
    Next Labl
Next Col

End Property
Private Sub InitHeaderThemeDefaults()
    'Excel-ish light gray header
    pHeaderBackColor = RGB(245, 245, 245)
    pHeaderHoverBackColor = RGB(235, 235, 235)
    pHeaderPressedBackColor = RGB(225, 225, 225)
    pHeaderForeColor = RGB(40, 40, 40)
    pHeaderBorderColor = RGB(200, 200, 200)
    pHeaderSortMarkColor = RGB(90, 90, 90)
End Sub
Private Function rowTop(ByVal row0 As Long) As Single
    Dim r As Long, t As Single

    If row0 <= 0 Then
        rowTop = 0
        Exit Function
    End If

    'If row heights not allocated yet, fall back to fixed height math
    On Error Resume Next
    Dim ub As Long
    ub = UBound(pRowHeights)
    If Err.Number <> 0 Then
        Err.Clear
        rowTop = row0 * Me.RowHeight
        Exit Function
    End If
    On Error GoTo 0

    'Clamp row0 to array bounds
    If row0 > ub + 1 Then row0 = ub + 1

    t = 0
    For r = 0 To row0 - 1
        t = t + pRowHeights(r)
    Next r

    rowTop = t
End Function

Private Function RowTop_new(ByVal row0 As Long) As Single
    Dim r As Long, t As Single, ub As Long

    If row0 <= 0 Then
        rowTop = 0
        Exit Function
    End If

    'If row heights not allocated yet, fall back to fixed height math
    On Error Resume Next
    ub = UBound(pRowHeights)
    If Err.Number <> 0 Then
        Err.Clear
        rowTop = row0 * Me.RowHeight
        Exit Function
    End If
    On Error GoTo 0

    'Clamp row0 to valid range: 0..(ub+1)
    If row0 > (ub + 1) Then row0 = ub + 1

    t = 0
    For r = 0 To row0 - 1
        t = t + pRowHeights(r)
    Next r

    rowTop = t
End Function


Private Sub EnsureMeasureLabel_2()
    If Not pMeasureLbl Is Nothing Then Exit Sub
    Set pMeasureLbl = pFrame.Controls.Add("Forms.Label.1", "dlbMeasure", True)
    With pMeasureLbl
        .Visible = False
        .WordWrap = True
        .AutoSize = True
    End With
End Sub

Private Sub ComputeRowHeights(ByVal InpArr As Variant)

    Call EnsureMeasureLabel

    Dim rows0 As Long, cols0 As Long
    rows0 = UBound(InpArr, 1) - LBound(InpArr, 1) + 1
    cols0 = UBound(InpArr, 2) - LBound(InpArr, 2) + 1

    ReDim pRowHeights(0 To rows0 - 1)

    Dim baseH As Single
    baseH = CSng(Me.RowHeight)
    If baseH < 1 Then baseH = 12

    Dim row0 As Long, col0 As Long
    Dim needH As Single, maxH As Single
    Dim cap As String, w As Single

    For row0 = 0 To rows0 - 1
        maxH = baseH

        For col0 = 0 To cols0 - 1
            w = CSng(GetColWidth(col0))
            If w > 4 Then
                cap = CStr(InpArr(LBound(InpArr, 1) + row0, LBound(InpArr, 2) + col0))

                If Not pMeasureLbl Is Nothing Then
                    With pMeasureLbl
                        .Font.Name = pFrame.Font.Name
                        .Font.Size = pFrame.Font.Size
                        .Font.Bold = False
                        .Width = w - 4
                        .Caption = cap
                        .WordWrap = True
                        .AutoSize = True
                        needH = .Height + 2
                    End With

                    If needH > maxH Then maxH = needH
                End If
            End If
        Next col0

        'Keep header row fixed height (optional)
        If Me.Headers And row0 = 0 Then maxH = baseH

        'CRITICAL: never allow zero/negative heights
        If maxH < baseH Then maxH = baseH

        pRowHeights(row0) = maxH
    Next row0

End Sub
Private Sub EnsureMeasureLabel()
    If Not pMeasureLbl Is Nothing Then Exit Sub
    If pFrame Is Nothing Then Exit Sub

    'Hidden label used only for measuring wrapped text height
    Set pMeasureLbl = pFrame.Controls.Add("Forms.Label.1", "__dlbMeasure", False)
    With pMeasureLbl
        .Visible = False
        .WordWrap = True
        .AutoSize = True
        .Caption = vbNullString
    End With
End Sub
Private Function GetColWidthSafe(ByVal zeroBasedCol As Long) As Single
    On Error GoTo EH

    Dim parts() As String
    parts = Split(Me.ColumnWidths, ";")

    If zeroBasedCol < 0 Then Exit Function
    If zeroBasedCol > UBound(parts) Then Exit Function

    GetColWidthSafe = CSng(val(parts(zeroBasedCol)))
    Exit Function

EH:
    GetColWidthSafe = 0
End Function
Public Function SelectedRowArray() As Variant

    If pSelectedRows Is Nothing Or pSelectedRows.Count = 0 Then
        SelectedRowArray = Empty
        Exit Function
    End If

    Dim arr() As Long
    Dim i As Long
    Dim k As Variant

    ReDim arr(0 To pSelectedRows.Count - 1)

    i = 0
    For Each k In pSelectedRows.Keys
        arr(i) = CLng(k)
        i = i + 1
    Next k

    Call SortLongArrayInPlace(arr)

    SelectedRowArray = arr

End Function

Public Function SelectedDataArray() As Variant

    Dim rows As Variant
    rows = SelectedRowArray()

    If IsEmpty(rows) Then
        SelectedDataArray = Empty
        Exit Function
    End If

    Dim r As Long, c As Long
    Dim outArr() As Variant

    ReDim outArr(0 To UBound(rows), 0 To Me.ColumnsCount - 1)

    For r = 0 To UBound(rows)
        For c = 0 To Me.ColumnsCount - 1
            outArr(r, c) = pInpArr( _
                LBound(pInpArr, 1) + CLng(rows(r)), _
                LBound(pInpArr, 2) + c _
            )
        Next c
    Next r

    SelectedDataArray = outArr

End Function
Public Sub ApplyHeaderVisuals()
    If pFrame Is Nothing Then Exit Sub
    If pHeadersCol Is Nothing Then Exit Sub

    Dim L As MSForms.Label
    For Each L In pHeadersCol
        With L
            .BackStyle = fmBackStyleOpaque
            .BackColor = pHeaderBackColor
            .ForeColor = pHeaderForeColor
            .Font.Bold = True
            .Font.Size = 9
            .TextAlign = fmTextAlignLeft

            'Remove per-cell borders (we'll use a single divider line)
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat

            'Give the header a little left padding by adding a leading space
            If Len(.Caption) > 0 Then
                If Left$(.Caption, 1) <> " " Then .Caption = " " & .Caption
            End If
        End With
    Next L

    'Add/refresh a strong bottom divider for the header row
    AddOrUpdateHeaderDivider
End Sub
'=== clsDesignListBox ===
Public Sub ResetHeaderHover()
    If pHeadersCol Is Nothing Then Exit Sub
    Dim L As MSForms.Label
    For Each L In pHeadersCol
        L.BackColor = pHeaderBackColor
    Next L
End Sub
Private Sub AddOrUpdateHeaderDivider()
    If pFrame Is Nothing Then Exit Sub
    Dim nm As String: nm = "hline;header_div"

    On Error Resume Next
    pFrame.Controls.Remove nm
    On Error GoTo 0

    Dim hdrLine As MSForms.Label
    Set hdrLine = pFrame.Controls.Add("Forms.Label.1", nm, True)
    With hdrLine
        .Tag = "GridLine"
        .Left = 0
        .Top = Me.RowHeight - 1
        .Width = pFrame.ScrollWidth
        .Height = 1
        .BackStyle = fmBackStyleOpaque
        .BackColor = pHeaderBorderColor
        .Caption = vbNullString
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .Enabled = False
        .ZOrder 0
    End With
End Sub
Public Sub SetHoverRow(ByVal RowNumber As Long)

    If Not pHoverEnabled Then Exit Sub
    If Me.IsListEmpty Then Exit Sub
    If Me.Headers And RowNumber = 0 Then Exit Sub

    'same row? nothing to do
    If RowNumber = pHoverRow Then Exit Sub

    'clear old hover first
    If pHoverRow <> -1 Then ClearHover

    pHoverRow = RowNumber

    Dim L As MSForms.Label
    For Each L In Me.RowLabels(RowNumber)
        'don't paint gridlines/headers
        If L.Tag <> "Header" And L.Tag <> "GridLine" Then
            'don't override selected row highlight
            If Not IsRowSelected(RowNumber) Then
                L.BackColor = pHoverColor
            End If
        End If
    Next L

End Sub


Public Sub ClearHover()

    If pHoverRow = -1 Then Exit Sub
    If Me.IsListEmpty Then
        pHoverRow = -1
        Exit Sub
    End If

    Dim r As Long
    r = pHoverRow
    pHoverRow = -1

    Dim L As MSForms.Label
    For Each L In Me.RowLabels(r)
        If L.Tag <> "Header" And L.Tag <> "GridLine" Then
            'restore based on selection state
            If IsRowSelected(r) Then
                L.BackColor = Me.SelectionColor
            Else
                L.BackColor = vbWhite
            End If
        End If
    Next L

End Sub

Public Sub HoverRow(ByVal RowNumber As Long)
    If Me.IsListEmpty Then Exit Sub
    If Me.Headers And RowNumber = 0 Then Exit Sub
    If RowNumber < 0 Or RowNumber > Me.RowsCount - 1 Then Exit Sub

    'Selected rows: don't hover (optional)
    If Not pSelectedRows Is Nothing Then
        If pSelectedRows.Exists(RowNumber) Or pSelectedRows.Exists(CStr(RowNumber)) Then Exit Sub
    End If

    If pHoverRow = RowNumber Then Exit Sub

    'Cancel any in-flight animation
    pHoverToken = pHoverToken + 1
    Dim myToken As Long: myToken = pHoverToken

    'Unhover previous row (no ping-pong)
    If pHoverRow >= 0 Then
        UnhoverRow pHoverRow, myToken
    End If

    'IMPORTANT: set hover row early so repeated MouseMove doesn't restart fade
    pHoverRow = RowNumber

    Dim rowLabs As Collection
    Set rowLabs = Me.RowLabels(RowNumber, False)

    Dim fromC As Long
    fromC = vbWhite
    If rowLabs.Count > 0 Then
        fromC = rowLabs(1).BackColor   'fade from current state
    End If

    FadeRowLabels rowLabs, fromC, pHoverColor, pHoverToken, myToken, pHoverSteps, pHoverDelayMs
End Sub

Public Sub ClearHover2()
    If pHoverRow < 0 Then Exit Sub
    pHoverToken = pHoverToken + 1
    Dim myToken As Long: myToken = pHoverToken

    UnhoverRow pHoverRow, myToken
    pHoverRow = -1
End Sub
Public Sub NotifyMouseOverLabel()
    pLastLabelMove = Timer
End Sub
Private Sub UnhoverRow(ByVal RowNumber As Long, ByVal myToken As Long)
    Dim rowLabs As Collection
    Set rowLabs = Me.RowLabels(RowNumber, False)

    Dim Target As Long: Target = vbWhite
    If Not pSelectedRows Is Nothing Then
        If pSelectedRows.Exists(RowNumber) Or pSelectedRows.Exists(CStr(RowNumber)) Then
            Target = Me.SelectionColor
        End If
    End If

    Dim fromC As Long
    fromC = pHoverColor
    If rowLabs.Count > 0 Then fromC = rowLabs(1).BackColor

    FadeRowLabels rowLabs, fromC, Target, pHoverToken, myToken, pHoverSteps, pHoverDelayMs
End Sub
Private Function IsRowSelected(ByVal RowNumber As Long) As Boolean
    If pSelectedRows Is Nothing Then Exit Function
    IsRowSelected = pSelectedRows.Exists(CStr(RowNumber)) Or pSelectedRows.Exists(RowNumber)
End Function
Public Sub BeginResize(ByVal colIdx As Long, ByVal startX As Single)
    pIsResizing = True
    pResizeCol = colIdx
    pResizeStartX = startX
    pResizeStartW = GetColWidth(colIdx)
End Sub
Public Sub EndResize()
    pIsResizing = False
End Sub
Public Sub DragResize(ByVal currentX As Single)
    If Not pIsResizing Then Exit Sub

    Dim dx As Long
    dx = CLng(currentX - pResizeStartX)

    Dim newW As Long
    newW = pResizeStartW + dx
    If newW < pMinColWidth Then newW = pMinColWidth

    'Update just this column and relayout
    SetColWidth pResizeCol, newW
    LayoutFromColumnWidths
End Sub
Private Sub pFrame_MouseMove2(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Mouse is on the frame background, not a label
    ClearHover
End Sub
Private Function GetColWidthsArray() As Variant
    GetColWidthsArray = Split(Me.ColumnWidths, ";")
End Function
Public Sub LayoutFromColumnWidths()
    If pFrame Is Nothing Then Exit Sub
    If Me.IsDLBEmpty Then Exit Sub   'or Me.IsEmpty if you kept that name

    Dim ctrl As Control
    Dim L As MSForms.Label
    Dim parts() As String
    Dim r As Long, c As Long
    Dim leftX As Long, w As Long

    'Reposition all cell/header labels based on their row;col Name
    For Each ctrl In pFrame.Controls
        If TypeName(ctrl) = "Label" Then
            Set L = ctrl
            If IsCellLabel(L) Or L.Tag = "Header" Then
                parts = Split(L.Name, ";")
                If UBound(parts) >= 1 Then
                    r = CLng(parts(0))
                    c = CLng(parts(1))

                    leftX = 0
                    Dim k As Long
                    For k = 0 To c - 1
                        leftX = leftX + GetColWidth(k)
                    Next k

                    w = GetColWidth(c)

                    L.Left = leftX
                    L.Width = w
                End If
            End If
        End If
    Next ctrl

    'Recompute ScrollWidth for gridlines spanning full width
    Dim totalW As Long, i As Long
    totalW = 0
    For i = 0 To Me.ColumnsCount - 1
        totalW = totalW + GetColWidth(i)
    Next i
    pFrame.ScrollWidth = totalW

    'Reapply gridlines
    ApplyHorizontalGridlines RGB(230, 230, 230)
End Sub
Private Sub SetColWidth(ByVal colIdx As Long, ByVal newW As Long)
    Dim parts As Variant, i As Long, s As String
    parts = GetColWidthsArray()

    If colIdx < 0 Or colIdx > UBound(parts) Then Exit Sub

    If newW < pMinColWidth Then newW = pMinColWidth
    parts(colIdx) = CStr(newW)

    For i = 0 To UBound(parts)
        s = s & parts(i)
        If i < UBound(parts) Then s = s & ";"
    Next i

    'This will call your Property Let ColumnWidths
    Me.ColumnWidths = s
End Sub


Public Sub ClearFormats()
    If pFormatRules Is Nothing Then Set pFormatRules = New Collection
    Set pFormatRules = New Collection
End Sub

Public Sub AddFormat( _
    ByVal ColIndex As Long, _
    ByVal Operator As String, _
    ByVal CompareValue As Variant, _
    ByVal ForeColor As Long, _
    Optional ByVal ApplyToWholeRow As Boolean = False, _
    Optional ByVal UseCaption As Boolean = False, _
    Optional ByVal CaseSensitive As Boolean = False _
)
    'Stores one rule. Rules are applied in the order added (later rules can override earlier ones).
    If pFormatRules Is Nothing Then Set pFormatRules = New Collection

    Dim rule(0 To 6) As Variant
    rule(0) = CLng(ColIndex)
    rule(1) = LCase$(Trim$(CStr(Operator)))
    rule(2) = CompareValue
    rule(3) = CLng(ForeColor)
    rule(4) = CBool(ApplyToWholeRow)
    rule(5) = CBool(UseCaption)
    rule(6) = CBool(CaseSensitive)

    pFormatRules.Add rule
End Sub
Public Sub HeaderSort(ByVal ColIndex As Long)

    'toggle direction if clicking same column again
    If ColIndex = pLastSortCol Then
        pLastSortDescending = Not pLastSortDescending
    Else
        pLastSortCol = ColIndex
        pLastSortDescending = False   'default first click = ascending
    End If

    'Perform sort
    Me.Sort ColIndex, pLastSortDescending

    'Optional: reapply formatting rules after sort
    If Not pFormatRules Is Nothing Then
        If pFormatRules.Count > 0 Then Me.ApplyFormats
    End If

End Sub
'Public Sub SortByHeaderClick(ByVal ColIndex As Long)
'    'TEMP: prove this is the one being called
'    Debug.Print "SortByHeaderClick called. ColIndex=" & ColIndex
'End Sub


Public Sub SortByHeaderClick2(ByVal ColIndex As Long)

    If ColIndex = pLastSortCol Then
        pLastSortDescending = Not pLastSortDescending
    Else
        pLastSortCol = ColIndex
        pLastSortDescending = False
    End If

    Me.Sort ColIndex, pLastSortDescending

End Sub

Public Sub ApplyFormats( _
    Optional ByVal ResetNonHeaderToColor As Long = vbBlack, _
    Optional ByVal SkipHeaderRow As Boolean = True _
)
    If pFrame Is Nothing Then Exit Sub
    If Me.IsDLBEmpty Then Exit Sub
    If pFormatRules Is Nothing Then Exit Sub
    If pFormatRules.Count = 0 Then Exit Sub

    BeginUpdate

    'Reset existing colors (avoid old formatting sticking around)
    Dim L As MSForms.Label
    For Each L In Me.AllLabels
        If L.Tag <> "Header" And L.Tag <> "GridLine" Then
            L.ForeColor = ResetNonHeaderToColor
        End If
    Next L

    Dim r As Long
    Dim rule As Variant
    Dim colIdx As Long, op As String
    Dim cmp As Variant, color As Long
    Dim wholeRow As Boolean, useCap As Boolean, caseSens As Boolean

    For Each rule In pFormatRules
        colIdx = CLng(rule(0))
        op = CStr(rule(1))
        cmp = rule(2)
        color = CLng(rule(3))
        wholeRow = CBool(rule(4))
        useCap = CBool(rule(5))
        caseSens = CBool(rule(6))

        For r = 0 To Me.RowsCount - 1
            If SkipHeaderRow And Me.Headers And r = 0 Then GoTo nextRow

            If RuleMatchesCell(r, colIdx, op, cmp, useCap, caseSens) Then
                If wholeRow Then
                    Dim RL As MSForms.Label
                    For Each RL In Me.RowLabels(r)
                        RL.ForeColor = color
                    Next RL
                Else
                    Dim cell As MSForms.Label
                    Set cell = Me.ExactLabel(r, colIdx)
                    If Not cell Is Nothing Then cell.ForeColor = color
                End If
            End If

nextRow:
        Next r
    Next rule

    EndUpdate
End Sub

Public Sub AutoFitColumn(ByVal ColIndex As Long)

    If pFrame Is Nothing Then Exit Sub
    If Me.IsDLBEmpty Then Exit Sub
    If Not CheckInputColumnNumber(ColIndex) Then Exit Sub

    Const MIN_W As Long = 30
    Const PAD As Long = 12

    Dim r As Long, w As Long, maxW As Long
    Dim L As MSForms.Label

    maxW = MIN_W

    For r = 0 To Me.RowsCount - 1
        Set L = Me.ExactLabel(r, ColIndex)
        If Not L Is Nothing Then
            w = MeasureText(CStr(L.Caption), L) + PAD
            If w > maxW Then maxW = w
        End If
    Next r

    SetColWidth ColIndex, maxW
    LayoutFromColumnWidths
End Sub
Private Sub BeginUpdate()
    If Not pFrame Is Nothing Then pFrame.Visible = False
End Sub

Private Sub EndUpdate()
    If Not pFrame Is Nothing Then pFrame.Visible = True
End Sub
Private Function MeasureText(ByVal txt As String, ByVal sample As MSForms.Label) As Long

    If pFrame Is Nothing Then Exit Function

    If pMeasureLbl Is Nothing Then
        Set pMeasureLbl = pFrame.Controls.Add("Forms.Label.1", "__dlb_measure", True)
        With pMeasureLbl
            .AutoSize = True
            .WordWrap = False
            .Caption = vbNullString

            'VISIBLE but off-screen so AutoSize/Width actually updates
            .Visible = True
            .Left = -20000
            .Top = -20000
        End With
    End If

    'Match font to sample label
    With pMeasureLbl.Font
        .Name = sample.Font.Name
        .Size = sample.Font.Size
        .Bold = sample.Font.Bold
        .Italic = sample.Font.Italic
    End With

    pMeasureLbl.Caption = txt

    'Force layout update (usually not needed, but helps in some cases)
    DoEvents

    MeasureText = CLng(pMeasureLbl.Width)
End Function
Private Function TryCDbl(ByVal v As Variant, ByRef outVal As Double) As Boolean
    Dim s As String

    If IsNumeric(v) Then
        outVal = CDbl(v)
        TryCDbl = True
        Exit Function
    End If

    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function

    s = Replace$(s, "$", "")
    s = Replace$(s, "%", "")
    s = Replace$(s, ",", "")
    s = Trim$(s)

    If IsNumeric(s) Then
        outVal = CDbl(s)
        TryCDbl = True
    End If
End Function

Private Function RuleMatchesCell( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long, _
    ByVal op As String, _
    ByVal cmp As Variant, _
    ByVal UseCaption As Boolean, _
    ByVal CaseSensitive As Boolean _
) As Boolean

    Dim cell As MSForms.Label
    Set cell = Me.ExactLabel(RowIndex, ColIndex)
    If cell Is Nothing Then Exit Function

    Dim v As Variant
    If UseCaption Then
        v = cell.Caption
    Else
        v = pInpArr(RowIndex + GetDeviation("Rows"), ColIndex + GetDeviation("Columns"))
    End If

    'Decide whether this is a numeric or text compare based on operator + convertability
    Dim dv As Double, dcmp As Double
    Dim isNumV As Boolean, isNumC As Boolean

    isNumV = TryCDbl(v, dv)
    isNumC = TryCDbl(cmp, dcmp)

    Select Case op
        '---- Numeric operators ----
        Case "<", "<=", ">", ">=", "=", "<>"
            If isNumV And isNumC Then
                Select Case op
                    Case "<":  RuleMatchesCell = (dv < dcmp)
                    Case "<=": RuleMatchesCell = (dv <= dcmp)
                    Case ">":  RuleMatchesCell = (dv > dcmp)
                    Case ">=": RuleMatchesCell = (dv >= dcmp)
                    Case "=":  RuleMatchesCell = (dv = dcmp)
                    Case "<>": RuleMatchesCell = (dv <> dcmp)
                End Select
            Else
                'Fallback to text compare for = and <> if not numeric
                If op = "=" Or op = "<>" Then
                    RuleMatchesCell = TextCompare(CStr(v), CStr(cmp), op, CaseSensitive)
                End If
            End If

        '---- Text operators ----
        Case "contains", "startswith", "endswith", "like"
            RuleMatchesCell = TextCompare(CStr(v), CStr(cmp), op, CaseSensitive)

        Case Else
            'Unknown operator -> no match
            RuleMatchesCell = False
    End Select

End Function
Private Function TextCompare( _
    ByVal leftText As String, _
    ByVal rightText As String, _
    ByVal op As String, _
    ByVal CaseSensitive As Boolean _
) As Boolean

    Dim cm As VbCompareMethod
    cm = IIf(CaseSensitive, vbBinaryCompare, vbTextCompare)

    Select Case op
        Case "=", "<>"
            If op = "=" Then
                TextCompare = (StrComp(leftText, rightText, cm) = 0)
            Else
                TextCompare = (StrComp(leftText, rightText, cm) <> 0)
            End If

        Case "contains"
            TextCompare = (InStr(1, leftText, rightText, cm) > 0)

        Case "startswith"
            If Len(rightText) = 0 Then Exit Function
            TextCompare = (StrComp(Left$(leftText, Len(rightText)), rightText, cm) = 0)

        Case "endswith"
            If Len(rightText) = 0 Then Exit Function
            TextCompare = (StrComp(Right$(leftText, Len(rightText)), rightText, cm) = 0)

        Case "like"
            'Like is case sensitive by default; simulate case-insensitive by uppercasing both
            If CaseSensitive Then
                TextCompare = (leftText Like rightText)
            Else
                TextCompare = (UCase$(leftText) Like UCase$(rightText))
            End If
    End Select
End Function

Private Function IsArrayAllocated(ByRef v As Variant) As Boolean
    If Not IsArray(v) Then
        IsArrayAllocated = False
        Exit Function
    End If

    On Error Resume Next
    Dim ub As Long
    ub = UBound(v, 1)          'unallocated dynamic array -> error 9
    IsArrayAllocated = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function
Public Sub ColorColumnByValueFast(ByVal ColIndex As Long)

    If pFrame Is Nothing Then Exit Sub

    pFrame.Visible = False

    Call ColorColumnByValue(ColIndex)

    pFrame.Visible = True

End Sub
Public Sub ColorColumnByValue(ByVal ColIndex As Long)

    Dim r As Long
    Dim val As Double
    Dim L As MSForms.Label

    For r = 0 To Me.RowsCount - 1

        Set L = Me.ExactLabel(r, ColIndex)

        If InStr(1, L.Caption, "Red", vbTextCompare) > 0 Then
            L.ForeColor = vbRed
        End If

    Next r

End Sub
Private Sub pFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Only clear hover if we haven't had a label MouseMove very recently
    If (Timer - pLastLabelMove) > 0.05 Then  '50ms debounce
        Me.ClearHover
    End If
End Sub
Private Sub pFrame_Scroll_shaky( _
    ByVal ActionX As MSForms.fmScrollAction, _
    ByVal ActionY As MSForms.fmScrollAction, _
    ByVal RequestDx As Single, _
    ByVal RequestDy As Single, _
    ByVal ActualDx As MSForms.ReturnSingle, _
    ByVal ActualDy As MSForms.ReturnSingle)

    Dim Col As Collection
    Dim Labl As MSForms.Label

    'Freeze columns: cancel horizontal scroll
    For Each Col In pFreezeColumnsCol
        For Each Labl In Col
            Labl.Left = CLng(-pFrame.ScrollLeft)
        Next Labl
    Next Col

    'Freeze rows (header): cancel vertical scroll
    For Each Col In pFreezeRowsCol
        For Each Labl In Col
            Labl.Top = CLng(-pFrame.ScrollTop)
        Next Labl
    Next Col

    'Keep frozen stuff above gridlines/content
    For Each Col In pFreezeColumnsCol
        For Each Labl In Col
            Labl.ZOrder 0
        Next Labl
    Next Col

    For Each Col In pFreezeRowsCol
        For Each Labl In Col
            Labl.ZOrder 0
        Next Labl
    Next Col

End Sub
' handles the scrolling; use this
Private Sub pFrame_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)

Dim Col, Labl As MSForms.Label

'For Each Col In pFreezeColumnsCol
'    For Each Labl In Col
'        Labl.Left = Labl.Left + ActualDx
'        If Labl.Left < 0 Then Labl.Left = 0
'    Next Labl
'Next Col

    For Each Col In pFreezeRowsCol
        For Each Labl In Col
            Labl.Top = Labl.Top + ActualDy
            If Labl.Top < 0 Then Labl.Top = 0
        Next Labl
    Next Col
    
' added to fix header moving issue
    
'    For Each Col In pFreezeRowsCol
'        For Each Labl In Col
'            Labl.Top = 0  'pin the frozen row(s)
'        Next Labl
'    Next Col
    
    
If Not pHeaderMask Is Nothing Then
    pHeaderMask.Left = pHeaderMask.Left + ActualDx
    If pHeaderMask.Left < 0 Then pHeaderMask.Left = 0 '

    pHeaderMask.Top = pHeaderMask.Top + ActualDy
    If pHeaderMask.Top < 0 Then pHeaderMask.Top = 0 '

'    'Make sure it stays above gridlines, below header text
    pHeaderMask.ZOrder 0
    Dim h As MSForms.Label
    If Not pHeadersCol Is Nothing Then
        For Each h In pHeadersCol
            h.ZOrder 0
        Next h
    End If
End If


End Sub
'Public Sub FocusKeySink()
'    On Error Resume Next
'    If Not pKeySink Is Nothing Then pKeySink.SetFocus
'    On Error GoTo 0
'End Sub

Public Sub Create_preKeyStroke(Parent As UserForm, inpTop As Long, inpLeft As Long, InpHeight As Long, inpWidth As Long, InpArr As Variant)

    Call AddControlsToUserform(Parent, Parent, TypeName(Parent), "Frame", "DesignListBox", InpHeight, inpLeft, inpTop, inpWidth, , , , , , , , , , , , True)

    Set pFrame = Parent.DesignListBox
    pFrame.ScrollBars = fmScrollBarsVertical
    'Set pKeySink = pParent.Controls.Add("Forms.TextBox.1", "dlbKeySink_" & CStr(Timer * 1000), True)

    'With pKeySink
        .Left = -200          'off-screen
        .Top = -200
        .Width = 10
        .Height = 10
        .Text = vbNullString
        .TabStop = False
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleNone
    End With
    
    ' or put this at the bottom of ExecuteFill
    EnsureKeySink
    
    Set pParent = Parent

    Call Fill(InpArr)

End Sub

Public Sub Create(Parent As UserForm, inpTop As Long, inpLeft As Long, InpHeight As Long, inpWidth As Long, InpArr As Variant)

    'Create the frame
    Call AddControlsToUserform(Parent, Parent, TypeName(Parent), _
                               "Frame", "DesignListBox", _
                               InpHeight, inpLeft, inpTop, inpWidth, , , , , , , , , , , , True)

    'Store parent FIRST (so we can add controls to it)


    Set pFrame = Parent.DesignListBox
    pFrame.ScrollBars = fmScrollBarsVertical
    
    Set pParent = Parent
    'Create key sink (off-screen but focusable)
    EnsureKeySink

    'Fill data
    Call Fill(InpArr)

    'Give focus to sink so arrow keys start working immediately
    'FocusKeySink

End Sub

Private Sub EnsureKeySink()
    If pParent Is Nothing Then Exit Sub

    If pKeySink Is Nothing Then
        Set pKeySink = pParent.Controls.Add("Forms.TextBox.1", "dlbKeySink_" & CLng(Timer * 1000), True)
        With pKeySink
            .Left = -200
            .Top = -200
            .Width = 10
            .Height = 10
            .Text = vbNullString
            .TabStop = False
            .SpecialEffect = fmSpecialEffectFlat
            .BorderStyle = fmBorderStyleNone
        End With
    End If

    'keep it focused
    On Error Resume Next
    pKeySink.SetFocus
    On Error GoTo 0
End Sub

Public Sub EnsureKeyFocus()
    If Not pKeySink Is Nothing Then pKeySink.SetFocus
End Sub
Public Sub TakeKeyboardFocus()
    EnsureKeySink
End Sub
Public Sub FocusKeySink()
    On Error Resume Next
    If Not pKeySink Is Nothing Then pKeySink.SetFocus
    Debug.Print "ActiveControl=", TypeName(pParent.ActiveControl)
    On Error GoTo 0
End Sub
'Private Sub pKeySink_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    Debug.Print "KEY=" & KeyCode & " SHIFT=" & Shift   'prove it fires
'    Me.HandleKey CLng(KeyCode), CLng(Shift)
'End Sub


Private Function firstDataRow() As Long
    If Me.Headers Then
        firstDataRow = 1
    Else
        firstDataRow = 0
    End If
End Function

Private Function LastDataRow() As Long
    LastDataRow = Me.RowsCount - 1
    If LastDataRow < firstDataRow() Then LastDataRow = firstDataRow()
End Function

Private Sub EnsureRowVisible(ByVal rowIdx As Long)

    If pFrame Is Nothing Then Exit Sub
    If rowIdx < 0 Then Exit Sub

    Dim rowTop As Single, rowBottom As Single
    rowTop = rowIdx * Me.RowHeight
    rowBottom = rowTop + Me.RowHeight

    Dim viewTop As Single, viewBottom As Single
    viewTop = pFrame.ScrollTop
    viewBottom = viewTop + pFrame.InsideHeight  'better than .Height

    If rowBottom > viewBottom Then
        pFrame.ScrollTop = rowBottom - pFrame.InsideHeight
    ElseIf rowTop < viewTop Then
        pFrame.ScrollTop = rowTop
    End If

End Sub
Private Sub PaintSelections()

    Dim Labl As MSForms.Label
    Dim r As Long

    'Clear existing highlight
    For Each Labl In pFrame.Controls
        If Labl.Tag <> "Header" And Labl.Tag <> "GridLine" Then
            Labl.BackColor = vbWhite
        End If
    Next Labl

    'Apply highlight
    If pSelectedRows Is Nothing Then Exit Sub

    Dim key As Variant

    For Each key In pSelectedRows.Keys
        For Each Labl In Me.RowLabels(CLng(key))
            Labl.BackColor = Me.SelectionColor
        Next Labl
    Next key

End Sub
Public Sub ClearSelection_old()
    If pSelectedRows Is Nothing Then
        Set pSelectedRows = CreateObject("Scripting.Dictionary")
        pSelectedRows.CompareMode = vbTextCompare   'optional but recommended
    Else
        pSelectedRows.RemoveAll
    End If

    pAnchorRow = -1
    Call DeselectAllLabels
End Sub

Public Sub ClearSelection()
    If pSelectedRows Is Nothing Then Set pSelectedRows = CreateObject("Scripting.Dictionary")
    pSelectedRows.RemoveAll
    pAnchorRow = -1
    pActiveRow = -1
    DeselectAllLabels
End Sub

Public Sub Fill(InpArr As Variant)

If pInFill Then Exit Sub
pInFill = True
On Error GoTo CleanFail

If pParent Is Nothing Or pFrame Is Nothing Then
    Err.Raise 91, "clsDesignListBox.Fill", "DesignListBox not initialized. Call .Create before .Fill."
End If
'Debug.Print "TypeName=" & TypeName(InpArr) & " IsArray=" & IsArray(InpArr)
If Not IsArray(InpArr) Then
    Err.Raise 123, "clsDesignListBox", "Input is not an Array", "Use Arrays only."
    Exit Sub
End If

InpArr = To2DArray(InpArr)

Call Me.Clear

'Store inpArr
pInpArr = InpArr

'Store original (unfiltered) copy
pBaseArr = InpArr

'Store Deviation
pParamDict("DeviationRows") = LBound(InpArr, 1)
pParamDict("DeviationColumns") = LBound(InpArr, 2)

Dim ArrayDimensions As Integer

ArrayDimensions = NumberOfArrayDimensions(InpArr)

'ColumnsCount
If ArrayDimensions = 1 Then
    pParamDict("ColumnsCount") = 1
ElseIf ArrayDimensions = 2 Then
    pParamDict("ColumnsCount") = UBound(InpArr, 2) - LBound(InpArr, 2) + 1
Else
    Err.Raise 123, "clsDesignListBox", "Input Array has more than 2 dimensions", "Use 1 or 2 dimension arrays."
End If

'RowsCount
pParamDict("RowsCount") = UBound(InpArr, 1) - LBound(InpArr, 1) + 1

'ColumnWidths
If Me.ColumnWidths = vbNullString Then
    Dim strWidths As String
    Dim strWidth As String
    strWidth = CStr(Round((pFrame.Width - 20) / pParamDict("ColumnsCount"), 2)) 'subtract 20 because of Scroll Bar
    Dim i As Long
    For i = 1 To pParamDict("ColumnsCount")
        strWidths = strWidths + strWidth & ";"
    Next i
    'Eliminate last ;
    If Right(strWidths, 1) = ";" Then
        strWidths = Left(strWidths, Len(strWidths) - 1)
    End If
    pParamDict("ColumnWidths") = strWidths
End If

'RowHeight
pParamDict("RowHeight") = 12
If Me.Headers Then
    Set pHeadersCol = New Collection
Else
    Set pHeadersCol = Nothing
End If



Call BuildRowMapForNoFilter
'Call ComputeRowHeights(InpArr)
Call ExecuteFill(InpArr)

CleanExit:
    pInFill = False
    Exit Sub

CleanFail:
    Debug.Print "Fill error: "; Err.Number; Err.Description
    Resume CleanExit
End Sub

Public Sub CopySelectionToClipboard(Optional ByVal IncludeHeaders As Boolean = True)

    If Me.IsDLBEmpty Then Exit Sub

    Dim firstDataRow As Long
    firstDataRow = IIf(Me.Headers, 1, 0)

    Dim sb As String
    Dim r As Long, c As Long

    '--- decide which rows to copy ---
    Dim rowsToCopy As Object  'Dictionary: key=row0
    Set rowsToCopy = CreateObject("Scripting.Dictionary")

    If (Not pSelectedRows Is Nothing) And (pSelectedRows.Count > 0) Then
        Dim k As Variant
        For Each k In pSelectedRows.Keys
            rowsToCopy(CLng(k)) = True
        Next k
    Else
        'no selection -> copy all data rows
        For r = firstDataRow To Me.RowsCount - 1
            rowsToCopy(r) = True
        Next r
    End If

    '--- headers line ---
    If IncludeHeaders And Me.Headers Then
        For c = 0 To Me.ColumnsCount - 1
            sb = sb & SafeCaption(pInpArr(LBound(pInpArr, 1), LBound(pInpArr, 2) + c))
            If c < Me.ColumnsCount - 1 Then sb = sb & vbTab
        Next c
        sb = sb & vbCrLf
    End If

    '--- rows (in ascending row order) ---
    Dim rr As Variant
    rr = DictKeysToLongArray(rowsToCopy)
    Call SortLongArrayInPlace(rr)

    For r = LBound(rr) To UBound(rr)
        For c = 0 To Me.ColumnsCount - 1
            sb = sb & SafeCaption(pInpArr(LBound(pInpArr, 1) + CLng(rr(r)), LBound(pInpArr, 2) + c))
            If c < Me.ColumnsCount - 1 Then sb = sb & vbTab
        Next c
        sb = sb & vbCrLf
    Next r

    '--- clipboard (late bound MSForms.DataObject) ---
    'Dim dobj As Object
    'Set dobj = CreateObject("MSForms.DataObject")
    'dobj.SetText sb
    'dobj.PutInClipboard

    Call modClipboard.ClipboardSetText(sb)
    If Not modClipboard.ClipboardSetText(sb) Then
        Debug.Print "Clipboard copy failed."
    End If
End Sub

Private Function DictKeysToLongArray(ByVal d As Object) As Variant
    Dim arr() As Variant, i As Long, k As Variant
    ReDim arr(0 To d.Count - 1)
    i = 0
    For Each k In d.Keys
        arr(i) = CLng(k)
        i = i + 1
    Next k
    DictKeysToLongArray = arr
End Function


Public Sub ToggleSelectAll()

    If Me.IsDLBEmpty Then Exit Sub
    If pSelectedRows Is Nothing Then Set pSelectedRows = CreateObject("Scripting.Dictionary")

    Dim firstRow As Long
    firstRow = IIf(Me.Headers, 1, 0)

    Dim totalRows As Long
    totalRows = Me.RowsCount - firstRow

    'If already fully selected ? clear
    If pSelectedRows.Count >= totalRows And totalRows > 0 Then
        ClearSelection
        RaiseEvent Change
        Exit Sub
    End If

    'Else select all
    pSelectedRows.RemoveAll

    Dim r As Long
    For r = firstRow To Me.RowsCount - 1
        pSelectedRows(CStr(r)) = True
    Next r

    pAnchorRow = firstRow
    pActiveRow = firstRow

    PaintSelections
    EnsureRowVisible firstRow

    RaiseEvent Change
End Sub

Private Sub HandleArrowKey(ByVal key As Long, ByVal Shift As Integer)

    Dim idx As Long
    idx = Me.ListIndex

    If idx < 0 Then
        'no selection yet: pick first data row
        idx = IIf(Me.Headers, 1, 0)
    End If

    If key = vbKeyUp Then idx = idx - 1
    If key = vbKeyDown Then idx = idx + 1

    'Clamp
    If Me.Headers And idx < 1 Then idx = 1
    If Not Me.Headers And idx < 0 Then idx = 0
    If idx > Me.RowsCount - 1 Then idx = Me.RowsCount - 1

    Dim isShift As Boolean, isCtrl As Boolean
    isShift = ((Shift And 1) <> 0)
    isCtrl = ((Shift And 2) <> 0)

    'Arrow keys: normal move selects single row; Shift extends range
    Call Me.SelectRow(idx, isCtrl, isShift)

End Sub
Public Sub AddFilter(ByVal ColIndex As Long, ByVal Operator As String, ByVal CompareValue As Variant, _
                     Optional ByVal CaseSensitive As Boolean = False)

    If pFilters Is Nothing Then Set pFilters = New Collection

    Dim rule(0 To 3) As Variant
    rule(0) = CLng(ColIndex)
    rule(1) = LCase$(Trim$(CStr(Operator)))
    rule(2) = CompareValue
    rule(3) = CBool(CaseSensitive)

    pFilters.Add rule
End Sub
Private Sub pKeySink_KeyDown_old2(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    'Debug.Print "DLB KeyDown:", KeyCode, Shift

    Select Case KeyCode
        Case vbKeyUp, vbKeyDown
            HandleArrowKey CLng(KeyCode), Shift
            KeyCode = 0   'eat the key so the textbox doesn't do anything weird
            EnsureKeySink 'force focus back (prevents "one step then dead")
    End Select

End Sub

Private Sub pKeySink_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    Dim cur As Long
    cur = pActiveRow

    If cur < 0 Then
        cur = Me.ListIndex
        If cur < 0 Then cur = IIf(Me.Headers, 1, 0)
        pActiveRow = cur
    End If

    Dim nextRow As Long
    nextRow = cur
    
    'Ctrl+A toggle
    If (Shift And 2) <> 0 Then
        If KeyCode = vbKeyA Then
            Me.ToggleSelectAll
            KeyCode = 0
            Exit Sub
        End If
    End If
    
    'Ctrl+C copy
    'If (Shift And 2) <> 0 Then
    '    If KeyCode = vbKeyC Then
    '        Me.CopySelectionToClipboard True
    '        KeyCode = 0
    '        Exit Sub
    '    End If
    'End If
    
    Select Case KeyCode
        Case vbKeyDown: nextRow = cur + 1
        Case vbKeyUp:   nextRow = cur - 1
        Case vbKeyHome: nextRow = IIf(Me.Headers, 1, 0)
        Case vbKeyEnd:  nextRow = Me.RowsCount - 1
        Case Else
            Exit Sub
    End Select

    If Me.Headers And nextRow < 1 Then nextRow = 1
    If nextRow < 0 Then nextRow = 0
    If nextRow > Me.RowsCount - 1 Then nextRow = Me.RowsCount - 1

    Dim isShift As Boolean, isCtrl As Boolean
    isShift = ((Shift And 1) <> 0)
    isCtrl = ((Shift And 2) <> 0)

    Me.SelectRow nextRow, isCtrl, isShift

    KeyCode = 0
End Sub
Public Sub ClearFilters()
    If pFilters Is Nothing Then Set pFilters = New Collection
    Set pFilters = New Collection
End Sub

Public Sub ApplyFilters()
    If pParent Is Nothing Or pFrame Is Nothing Then Exit Sub
    If Not IsArrayAllocated(pBaseArr) Then Exit Sub

    BeginUpdate

    Dim viewArr As Variant
    viewArr = BuildFilteredView()

    'refresh display
    Call Me.Clear
    pInpArr = viewArr

    'update counts/deviation based on view
    pParamDict("DeviationRows") = LBound(pInpArr, 1)
    pParamDict("DeviationColumns") = LBound(pInpArr, 2)
    pParamDict("ColumnsCount") = UBound(pInpArr, 2) - LBound(pInpArr, 2) + 1
    pParamDict("RowsCount") = UBound(pInpArr, 1) - LBound(pInpArr, 1) + 1

    Call ExecuteFill(pInpArr)

    EndUpdate
End Sub
Public Sub SelectAll(Optional ByVal KeepAnchor As Boolean = True)
    If Me.IsDLBEmpty Then Exit Sub
    If pSelectedRows Is Nothing Then Set pSelectedRows = CreateObject("Scripting.Dictionary")

    pSelectedRows.RemoveAll

    Dim r As Long
    Dim firstRow As Long
    firstRow = 0
    If Me.Headers Then firstRow = 1   'skip header row

    For r = firstRow To Me.RowsCount - 1
        pSelectedRows(CStr(r)) = True
    Next r

    'Anchor: keep current if requested, else set to first row
    If Not KeepAnchor Or pAnchorRow < firstRow Then
        If Me.RowsCount > firstRow Then pAnchorRow = firstRow
    End If

    PaintSelections
    If Me.RowsCount > firstRow Then EnsureRowVisible pAnchorRow

    RaiseEvent Change
End Sub
Public Sub ResetFilters()
    ClearFilters
    ApplyFilters
End Sub

Private Sub BuildRowMapForNoFilter()
    Dim hasHdr As Boolean: hasHdr = Me.Headers
    Dim srcLb As Long: srcLb = LBound(pBaseArr, 1)
    Dim srcUb As Long: srcUb = UBound(pBaseArr, 1)

    Dim dataStart As Long
    dataStart = IIf(hasHdr, srcLb + 1, srcLb)

    Dim n As Long: n = srcUb - dataStart + 1
    If n < 0 Then n = 0

    ReDim pRowMap(0 To n - 1)

    Dim i As Long, d As Long
    d = 0
    For i = dataStart To srcUb
        pRowMap(d) = i - dataStart    '0..n-1 in source-data space
        d = d + 1
    Next i
End Sub
Private Sub UpdateScrollExtents()
    If pFrame Is Nothing Then Exit Sub
    If Me.IsListEmpty Then Exit Sub

    Dim maxRight As Single, maxBottom As Single
    Dim L As MSForms.Label

    maxRight = 0
    maxBottom = 0

    For Each L In pFrame.Controls
        If TypeName(L) = "Label" Then
            If L.Tag <> "GridLine" Then
                If L.Left + L.Width > maxRight Then maxRight = L.Left + L.Width
                If L.Top + L.Height > maxBottom Then maxBottom = L.Top + L.Height
            End If
        End If
    Next L

    'Add a little padding so scrollbars engage reliably
    pFrame.ScrollWidth = maxRight + 6
    pFrame.ScrollHeight = maxBottom + 6

    'Optional: force scrollbars visible (otherwise they can "disappear" when content fits)
    pFrame.KeepScrollBarsVisible = fmScrollBarsBoth
End Sub

Private Function BuildFilteredView() As Variant

    Dim src As Variant
    src = pBaseArr

    Dim rLb As Long, rUb As Long, cLb As Long, cUb As Long
    rLb = LBound(src, 1): rUb = UBound(src, 1)
    cLb = LBound(src, 2): cUb = UBound(src, 2)

    Dim hasHdr As Boolean: hasHdr = Me.Headers

    Dim dataStart As Long
    dataStart = IIf(hasHdr, rLb + 1, rLb)

    'Collect matching source rows (in a dynamic list)
    Dim matches() As Long
    Dim mCount As Long
    mCount = 0

    Dim r As Long
    For r = dataStart To rUb
        If RowPassesAllFilters(src, r, cLb) Then
            ReDim Preserve matches(0 To mCount)
            matches(mCount) = r
            mCount = mCount + 1
        End If
    Next r

    'Build view array: header + matches
    Dim outRows As Long
    outRows = mCount + IIf(hasHdr, 1, 0)

    Dim outArr() As Variant
    If outRows = 0 Then
        'Return an empty 1xN header-only array if headers, else 0 rows is messy in VBA.
        ReDim outArr(0 To IIf(hasHdr, 0, -1), cLb To cUb)
        If hasHdr Then
            Dim cc As Long
            For cc = cLb To cUb
                outArr(0, cc) = src(rLb, cc)
            Next cc
        End If
        BuildFilteredView = outArr
        Exit Function
    End If

    ReDim outArr(0 To outRows - 1, cLb To cUb)

    Dim outR As Long, c As Long

    'Header row
    outR = 0
    If hasHdr Then
        For c = cLb To cUb
            outArr(outR, c) = src(rLb, c)
        Next c
        outR = outR + 1
    End If

    'Data rows
    ReDim pRowMap(0 To mCount - 1)

    Dim i As Long
    For i = 0 To mCount - 1
        r = matches(i)
        pRowMap(i) = r  'store SOURCE row index (absolute in pBaseArr)

        For c = cLb To cUb
            outArr(outR, c) = src(r, c)
        Next c

        outR = outR + 1
    Next i

    BuildFilteredView = outArr
End Function


Private Function RowPassesAllFilters(ByRef src As Variant, ByVal srcRow As Long, ByVal cLb As Long) As Boolean
    If pFilters Is Nothing Then
        RowPassesAllFilters = True
        Exit Function
    End If
    If pFilters.Count = 0 Then
        RowPassesAllFilters = True
        Exit Function
    End If

    Dim rule As Variant
    For Each rule In pFilters
        If Not RuleMatches(src, srcRow, CLng(rule(0)), CStr(rule(1)), rule(2), CBool(rule(3)), cLb) Then
            RowPassesAllFilters = False
            Exit Function
        End If
    Next rule

    RowPassesAllFilters = True
End Function



Private Function RuleMatches(ByRef src As Variant, ByVal srcRow As Long, ByVal ColIndex As Long, _
                             ByVal op As String, ByVal cmp As Variant, ByVal caseSens As Boolean, _
                             ByVal cLb As Long) As Boolean

    Dim v As Variant
    v = src(srcRow, cLb + ColIndex)

    Dim dv As Double, dc As Double
    Dim isNumV As Boolean, isNumC As Boolean

    isNumV = TryCDbl(v, dv)
    isNumC = TryCDbl(cmp, dc)

    Select Case op
        Case ">", ">=", "<", "<=", "=", "<>"
            If isNumV And isNumC Then
                Select Case op
                    Case ">":  RuleMatches = (dv > dc)
                    Case ">=": RuleMatches = (dv >= dc)
                    Case "<":  RuleMatches = (dv < dc)
                    Case "<=": RuleMatches = (dv <= dc)
                    Case "=":  RuleMatches = (dv = dc)
                    Case "<>": RuleMatches = (dv <> dc)
                End Select
            Else
                'fallback text compare only for "=" and "<>"
                If op = "=" Or op = "<>" Then
                    RuleMatches = TextCompare(CStr(v), CStr(cmp), op, caseSens)
                Else
                    RuleMatches = False
                End If
            End If

        Case "contains", "startswith", "endswith", "like"
            RuleMatches = TextCompare(CStr(v), CStr(cmp), op, caseSens)

        Case Else
            RuleMatches = False
    End Select
End Function

Public Sub Sort(ByVal ColumnNumber As Long, Optional ByVal Descending As Boolean = False)

    If Not CheckInputColumnNumber(ColumnNumber) Then Exit Sub

    '1) Sort the backing array
    If Me.Headers Then
        SortArray pInpArr, LBound(pInpArr, 1) + 1, , ColumnNumber + GetDeviation("Columns"), Descending
    Else
        SortArray pInpArr, , , ColumnNumber + GetDeviation("Columns"), Descending
    End If

    '2) Update header arrow indicators
    UpdateHeaderArrows ColumnNumber, Descending

    '3) Rewrite captions in-place (NO renaming, NO moving)
    RewriteGridCaptions

    '4) Repaint selection + gridlines + formats
    PaintSelections
    ApplyHorizontalGridlines RGB(230, 230, 230)

    If Not pFormatRules Is Nothing Then
        If pFormatRules.Count > 0 Then ApplyFormats
    End If

End Sub

Private Sub RewriteGridCaptions()

    Dim ctrl As Control
    Dim L As MSForms.Label
    Dim parts() As String
    Dim r As Long, c As Long

    For Each ctrl In pFrame.Controls
        If TypeName(ctrl) = "Label" Then
            Set L = ctrl

            If IsCellLabel(L) Then
                parts = Split(L.Name, ";")
                r = CLng(parts(0))
                c = CLng(parts(1))

                'pull from sorted array
                L.Caption = CStr(pInpArr(r + GetDeviation("Rows"), c + GetDeviation("Columns")))
            End If
        End If
    Next ctrl

End Sub
Private Function CompareForSort(ByVal a As Variant, ByVal b As Variant) As Long
    'Returns: -1 if a<b, 0 if a=b, +1 if a>b

    '--- push “bad” values to the bottom ---
    If IsError(a) Then
        If IsError(b) Then CompareForSort = 0 Else CompareForSort = 1
        Exit Function
    End If
    If IsError(b) Then CompareForSort = -1: Exit Function

    If IsNull(a) Or VarType(a) = vbEmpty Or a = "" Then
        If IsNull(b) Or VarType(b) = vbEmpty Or b = "" Then
            CompareForSort = 0
        Else
            CompareForSort = 1
        End If
        Exit Function
    End If

    If IsNull(b) Or VarType(b) = vbEmpty Or b = "" Then
        CompareForSort = -1
        Exit Function
    End If

    '--- numeric compare if possible for both ---
    Dim da As Double, db As Double
    Dim na As Boolean, nb As Boolean
    na = TryCDbl(a, da)
    nb = TryCDbl(b, db)

    If na And nb Then
        If da < db Then
            CompareForSort = -1
        ElseIf da > db Then
            CompareForSort = 1
        Else
            CompareForSort = 0
        End If
        Exit Function
    End If

    '--- fallback: text compare ---
    CompareForSort = Sgn(StrComp(CStr(a), CStr(b), vbTextCompare))
End Function
Private Sub UpdateHeaderArrows(ByVal ColumnNumber As Long, ByVal Descending As Boolean)

    Dim ctrl As Control
    Dim L As MSForms.Label
    Dim parts() As String

    For Each ctrl In pFrame.Controls
        If TypeName(ctrl) = "Label" Then
            Set L = ctrl

            If L.Tag = "Header" Then
                'strip existing arrow if present
                If Right$(L.Caption, 1) = ChrW(&H2191) Or Right$(L.Caption, 1) = ChrW(&H2193) Then
                    L.Caption = Left$(L.Caption, Len(L.Caption) - 2)
                End If

                parts = Split(L.Name, ";")
                If UBound(parts) >= 1 Then
                    If CLng(parts(1)) = ColumnNumber Then
                        If Descending Then
                            L.Caption = L.Caption & " " & ChrW(&H2191)
                        Else
                            L.Caption = L.Caption & " " & ChrW(&H2193)
                        End If
                    End If
                End If
            End If
        End If
    Next ctrl

End Sub

Public Sub SelectRow_old(RowNumber As Long, Optional isCtrl As Boolean = False, Optional isShift As Boolean = False)

    If Not CheckInputRowNumber(RowNumber) Then Exit Sub
    If Me.Headers And RowNumber = 0 Then Exit Sub

    If pSelectedRows Is Nothing Then
        Set pSelectedRows = CreateObject("Scripting.Dictionary")
        pSelectedRows.CompareMode = vbTextCompare   'optional but recommended
    End If
    'No modifiers ? single select
    If Not isCtrl And Not isShift Then
        Call ClearSelection
        pSelectedRows(RowNumber) = True
        pAnchorRow = RowNumber

    'CTRL ? toggle row
    ElseIf isCtrl Then
        If pSelectedRows.Exists(RowNumber) Then
            pSelectedRows.Remove RowNumber
        Else
            pSelectedRows(RowNumber) = True
        End If
        pAnchorRow = RowNumber

    'SHIFT ? range select
    ElseIf isShift Then

        If pAnchorRow < 0 Then pAnchorRow = RowNumber

        Dim startRow As Long, endRow As Long, r As Long
        startRow = IIf(pAnchorRow < RowNumber, pAnchorRow, RowNumber)
        endRow = IIf(pAnchorRow > RowNumber, pAnchorRow, RowNumber)

        Call ClearSelection

        For r = startRow To endRow
            If Not (Me.Headers And r = 0) Then
                pSelectedRows(r) = True
            End If
        Next r
    End If

    Call PaintSelections
    RaiseEvent Change

End Sub


Public Sub SelectRow(RowNumber As Long, Optional isCtrl As Boolean = False, Optional isShift As Boolean = False)

    If Not CheckInputRowNumber(RowNumber) Then Exit Sub
    If Me.Headers And RowNumber = 0 Then Exit Sub

    If pSelectedRows Is Nothing Then Set pSelectedRows = CreateObject("Scripting.Dictionary")

    'No modifiers -> single select
    If Not isCtrl And Not isShift Then
        pSelectedRows.RemoveAll
        pSelectedRows(CStr(RowNumber)) = True
        pAnchorRow = RowNumber

    'CTRL -> toggle
    ElseIf isCtrl Then
        If pSelectedRows.Exists(CStr(RowNumber)) Then
            pSelectedRows.Remove CStr(RowNumber)
        Else
            pSelectedRows(CStr(RowNumber)) = True
        End If
        pAnchorRow = RowNumber

    'SHIFT -> range select (KEEP anchor!)
    ElseIf isShift Then

        If pAnchorRow < 0 Then pAnchorRow = RowNumber

        Dim startRow As Long, endRow As Long, r As Long
        If pAnchorRow < RowNumber Then
            startRow = pAnchorRow
            endRow = RowNumber
        Else
            startRow = RowNumber
            endRow = pAnchorRow
        End If

        'IMPORTANT: do NOT ClearSelection here (it resets anchor)
        pSelectedRows.RemoveAll

        For r = startRow To endRow
            If Not (Me.Headers And r = 0) Then
                pSelectedRows(CStr(r)) = True
            End If
        Next r

        'anchor stays as original
    End If
    Debug.Print "SelectRow row=" & RowNumber & " shift=" & isShift & " anchor=" & pAnchorRow
    
    pActiveRow = RowNumber
    
    PaintSelections
    EnsureRowVisible RowNumber

    RaiseEvent Change
End Sub



Public Sub SelectRow_old2(RowNumber As Long, Optional isCtrl As Boolean = False, Optional isShift As Boolean = False)

    If Not CheckInputRowNumber(RowNumber) Then Exit Sub
    If Me.Headers And RowNumber = 0 Then Exit Sub

    If pSelectedRows Is Nothing Then Set pSelectedRows = CreateObject("Scripting.Dictionary")
    pSelectedRows.CompareMode = vbTextCompare

    Dim k As String
    k = CStr(RowNumber)

    'No modifiers -> single select + update anchor
    If Not isCtrl And Not isShift Then
        pSelectedRows.RemoveAll
        pSelectedRows(k) = True
        pAnchorRow = RowNumber

    'CTRL -> toggle + update anchor
    ElseIf isCtrl Then
        If pSelectedRows.Exists(k) Then
            pSelectedRows.Remove k
        Else
            pSelectedRows(k) = True
        End If
        pAnchorRow = RowNumber

    'SHIFT -> range from anchor (do NOT clear anchor)
    ElseIf isShift Then

        If pAnchorRow < 0 Then pAnchorRow = RowNumber

        Dim startRow As Long, endRow As Long, r As Long
        If pAnchorRow < RowNumber Then
            startRow = pAnchorRow
            endRow = RowNumber
        Else
            startRow = RowNumber
            endRow = pAnchorRow
        End If

        'Clear selection WITHOUT touching the anchor
        pSelectedRows.RemoveAll

        For r = startRow To endRow
            If Not (Me.Headers And r = 0) Then
                pSelectedRows(CStr(r)) = True
            End If
        Next r

    End If

    PaintSelections
    RaiseEvent Change

End Sub
Private Function GetColWidth(ByVal zeroBasedCol As Long) As Long
    Dim parts() As String

    parts = Split(Me.ColumnWidths, ";")

    If zeroBasedCol < 0 Or zeroBasedCol > UBound(parts) Then
        GetColWidth = 0
    Else
        GetColWidth = CLng(val(parts(zeroBasedCol)))
    End If
End Function

Private Sub RepaintSelection()
    Dim Labl As MSForms.Label
    Dim rowIdx As Long

    For Each Labl In pFrame.Controls
        If Labl.Tag <> "Header" Then
            rowIdx = CLng(Split(Labl.Name, ";")(0))
            If pSelectedRows.Exists(CStr(rowIdx)) Then
                Labl.BackColor = Me.SelectionColor
                'Labl.BorderStyle = fmBorderStyleSingle
            Else
                Labl.BackColor = vbWhite
                'Labl.BorderStyle = fmBorderStyleSingle
            End If
        End If
    Next Labl
End Sub
Private Sub AddRangeToSelection(ByVal a As Long, ByVal b As Long)
    Dim lo As Long, hi As Long, r As Long
    If a <= b Then
        lo = a: hi = b
    Else
        lo = b: hi = a
    End If
    For r = lo To hi
        If r >= 0 And r <= Me.RowsCount - 1 Then
            pSelectedRows(CStr(r)) = True
        End If
    Next r
End Sub
Public Sub Clear()
    If pFrame Is Nothing Then Exit Sub

    Dim ctrl As Control
    For Each ctrl In pFrame.Controls
        pFrame.Controls.Remove ctrl.Name
    Next ctrl

    pIsEmpty = True
End Sub
Public Sub RaiseEventBeforeClick()

RaiseEvent BeforeClick

End Sub
Public Sub RaiseEventClick()

RaiseEvent Click

End Sub

Private Sub ExecuteFill_FixedRowHeight(ByVal InpArr As Variant)

    Dim i As Long, j As Long
    Dim setTop As Single, setLeft As Single
    Dim setHeight As Single, setWidth As Single
    Dim CreatedLabel As MSForms.Label
    Dim InpBorderStyle As Long, InpIsBold As Boolean
    Dim DLBObject As clsDesignListBoxObject
    Dim zCol As Long
    Dim totalW As Long, c As Long

    If Me.Headers Then
        Set pHeadersCol = New Collection
    Else
        Set pHeadersCol = Nothing
    End If

    Set pDesignListBoxObjectsCol = New Collection

    For i = LBound(InpArr, 1) To UBound(InpArr, 1)

        setTop = (i - LBound(InpArr, 1)) * Me.RowHeight
        setHeight = Me.RowHeight

        For j = LBound(InpArr, 2) To UBound(InpArr, 2)

            zCol = (j - LBound(InpArr, 2))

            setLeft = 0
            For c = 0 To zCol - 1
                setLeft = setLeft + GetColWidth(c)
            Next c

            setWidth = GetColWidth(zCol)
            If setWidth <= 0 Then GoTo NextIteration

            If Me.Headers And i = LBound(InpArr, 1) Then
                InpBorderStyle = fmBorderStyleNone
                InpIsBold = True
            Else
                InpBorderStyle = fmBorderStyleNone
                InpIsBold = False
            End If

            Set CreatedLabel = AddControlsToUserform( _
                pParent, pFrame, "DesignListBox", "Label", _
                (i - LBound(InpArr, 1)) & ";" & zCol, _
                setHeight, setLeft, setTop, setWidth, _
                SafeCaption(InpArr(i, j)), _
                SafeCaption(InpArr(i, j)), _
                , , , , , , , InpBorderStyle, , , InpIsBold _
            )

            CreatedLabel.WordWrap = False
            CreatedLabel.AutoSize = False

            If Me.Headers And i = LBound(InpArr, 1) Then
                pHeadersCol.Add CreatedLabel
                CreatedLabel.Tag = "Header"
            End If

            Set DLBObject = New clsDesignListBoxObject
            DLBObject.Create CreatedLabel, Me
            pDesignListBoxObjectsCol.Add DLBObject

NextIteration:
        Next j
    Next i

    pIsEmpty = False

    totalW = 0
    For c = 0 To Me.ColumnsCount - 1
        totalW = totalW + GetColWidth(c)
    Next c

    pFrame.ScrollWidth = totalW
    pFrame.ScrollHeight = (Me.RowsCount * Me.RowHeight)

    ApplyHorizontalGridlines RGB(230, 230, 230)

    If Not pFormatRules Is Nothing Then
        If pFormatRules.Count > 0 Then ApplyFormats
    End If
    
    If Me.Headers Then Me.FreezeRows = 1
End Sub



Private Sub ExecuteFill(ByVal InpArr As Variant)

    Dim i As Long, j As Long, k As Long
    Dim row0 As Long, zCol As Long
    Dim setTop As Single, setLeft As Single, setHeight As Single, setWidth As Single
    Dim CreatedLabel As MSForms.Label
    Dim InpBorderStyle As Long, InpIsBold As Boolean
    Dim DLBObject As clsDesignListBoxObject
    Dim cap As String

    Dim totalW As Long
    Dim colLefts() As Long

    '--- sanity ---
    If pFrame Is Nothing Then Exit Sub
    If Not IsArray(InpArr) Then Exit Sub

    '--- headers collection safety ---
    If Me.Headers Then
        Set pHeadersCol = New Collection
    Else
        Set pHeadersCol = Nothing
    End If

    '--- reset event wrappers so they don't pile up ---
    Set pDesignListBoxObjectsCol = New Collection

    '--- precompute column lefts + total width (fast + avoids drift) ---
    colLefts = BuildColLefts(totalW)   'ByRef totalW

    '--- build cells ---
    For i = LBound(InpArr, 1) To UBound(InpArr, 1)

        row0 = (i - LBound(InpArr, 1))

        'Top based on variable heights if available
        'setTop = RowTop(row0)
        setTop = row0 * Me.RowHeight
        
        'Height: use pRowHeights(row0) if available, else fixed RowHeight
 
        On Error Resume Next
        If row0 >= 0 Then
            setHeight = pRowHeights(row0)
            'setHeight = Me.RowHeight
            If Err.Number <> 0 Then
                Err.Clear
                setHeight = Me.RowHeight
            End If
        End If
        On Error GoTo 0
        If setHeight < 1 Then setHeight = Me.RowHeight
        If setHeight < 1 Then setHeight = 12

        For j = LBound(InpArr, 2) To UBound(InpArr, 2)

            zCol = (j - LBound(InpArr, 2))  '0..cols-1
            If zCol < 0 Then GoTo NextCell
            If zCol > UBound(colLefts) Then GoTo NextCell

            setLeft = colLefts(zCol)

            setWidth = GetColWidth(zCol)
            If setWidth <= 0 Then GoTo NextCell

            'Header formatting
            If Me.Headers And i = LBound(InpArr, 1) Then
                InpBorderStyle = fmBorderStyleNone
                InpIsBold = True
            Else
                InpBorderStyle = fmBorderStyleNone
                InpIsBold = False
            End If

            cap = SafeCaption(InpArr(i, j))

            '--- create label ---
            On Error GoTo EH_Create

            Set CreatedLabel = AddControlsToUserform( _
                pParent, pFrame, "DesignListBox", "Label", _
                CStr(row0) & ";" & CStr(zCol), _
                setHeight, setLeft, setTop, setWidth, _
                cap, cap, , , , , , , , InpBorderStyle, , , InpIsBold _
            )

            On Error GoTo 0

            'Force multiline layout behavior
            With CreatedLabel
                .WordWrap = True
                .AutoSize = False
                .Height = setHeight
                .TextAlign = fmTextAlignLeft
            End With

            'Tag + styling for header cells
            If Me.Headers And i = LBound(InpArr, 1) Then
                If Not pHeadersCol Is Nothing Then pHeadersCol.Add CreatedLabel
                CreatedLabel.Tag = "Header"
                CreatedLabel.MousePointer = fmMousePointerUpArrow

                'Optional: keep your header visuals consistent
                'If you already have ApplyHeaderVisuals, you can omit these lines
                CreatedLabel.BackColor = RGB(240, 240, 240)
                CreatedLabel.ForeColor = RGB(30, 30, 30)
                CreatedLabel.BorderStyle = fmBorderStyleSingle
            End If

            '--- wire events ---
            Set DLBObject = New clsDesignListBoxObject
            DLBObject.Create CreatedLabel, Me
            pDesignListBoxObjectsCol.Add DLBObject

NextCell:
        Next j
    Next i

    '--- mark non-empty ---
    pIsEmpty = False



    '--- scrolling extents ---
    pFrame.ScrollWidth = totalW

    'Total height: prefer RowTop(rows) (variable height), fallback fixed math
    Dim totalH As Single
    totalH = 0

    On Error Resume Next
    totalH = rowTop(Me.RowsCount)   'RowTop sums row heights
    If Err.Number <> 0 Or totalH <= 0 Then
        Err.Clear
        totalH = CSng(Me.RowsCount) * CSng(Me.RowHeight)
    End If
    On Error GoTo 0

    If totalH < 1 Then totalH = CSng(Me.RowsCount) * 12!

    pFrame.ScrollHeight = totalH

    '--- gridlines + header visuals + formats ---
    ApplyHorizontalGridlines RGB(230, 230, 230)

    If Me.Headers Then
        On Error Resume Next
        ApplyHeaderVisuals
        On Error GoTo 0
    End If
    
    If Me.Headers Then EnsureHeaderMask
    
    If Not pFormatRules Is Nothing Then
        If pFormatRules.Count > 0 Then ApplyFormats
    End If


    
    '--- freeze header row if enabled ---
    If Me.Headers Then Me.FreezeRows = 1
   'UpdateScrollExtents
   ' Debug.Print "InsideH=" & pFrame.InsideHeight & " ScrollH=" & pFrame.ScrollHeight
   ' Debug.Print "InsideW=" & pFrame.InsideWidth & " ScrollW=" & pFrame.ScrollWidth
    Exit Sub
EH_Create:
    Debug.Print "ExecuteFill create error: i=" & i & " j=" & j & _
                " row0=" & row0 & " col0=" & zCol & _
                " cap=[" & cap & "] Err=" & Err.Number & " " & Err.Description
    Err.Clear
    On Error GoTo 0
    Resume Next

End Sub
'Private Procedures
Private Sub ExecuteFill_save(ByVal InpArr As Variant)

    Dim i As Long, j As Long, k As Long
    Dim setTop As Single, setLeft As Single, setHeight As Single, setWidth As Single
    Dim CreatedLabel As MSForms.Label
    Dim InpBorderStyle As Long, InpIsBold As Boolean
    Dim DLBObject As clsDesignListBoxObject
    Dim zCol As Long
    Dim cap As String

    Dim totalW As Long, c As Long

    'Headers collection safety
    If Me.Headers Then
        Set pHeadersCol = New Collection
    Else
        Set pHeadersCol = Nothing
    End If

    'Reset wrappers so events don’t pile up
    Set pDesignListBoxObjectsCol = New Collection

    'Precompute total width for ScrollWidth and gridlines
    totalW = 0
    For c = 0 To Me.ColumnsCount - 1
        totalW = totalW + GetColWidth(c)
    Next c

    'Build labels
    For i = LBound(InpArr, 1) To UBound(InpArr, 1)

        setTop = (i - LBound(InpArr, 1)) * Me.RowHeight
        setHeight = Me.RowHeight

        For j = LBound(InpArr, 2) To UBound(InpArr, 2)

            zCol = (j - LBound(InpArr, 2)) '0..cols-1

            'Compute Left by summing widths of prior columns
            setLeft = 0
            For k = 0 To zCol - 1
                setLeft = setLeft + GetColWidth(k)
            Next k

            setWidth = GetColWidth(zCol)
            If setWidth <= 0 Then GoTo NextIteration

            'Header formatting
            If Me.Headers And i = LBound(InpArr, 1) Then
                InpBorderStyle = fmBorderStyleNone
                InpIsBold = True
            Else
                InpBorderStyle = fmBorderStyleNone
                InpIsBold = False
            End If

            cap = SafeCaption(InpArr(i, j))

            Set CreatedLabel = AddControlsToUserform( _
                pParent, pFrame, "DesignListBox", "Label", _
                (i - LBound(InpArr, 1)) & ";" & zCol, _
                setHeight, setLeft, setTop, setWidth, _
                cap, cap, , , , , , , , InpBorderStyle, , , InpIsBold _
            )

            With CreatedLabel
                .WordWrap = False
                .AutoSize = False
                .Height = setHeight
                .TextAlign = fmTextAlignLeft
            End With

            'Header tag + pointer
            If Me.Headers And i = LBound(InpArr, 1) Then
                pHeadersCol.Add CreatedLabel
                CreatedLabel.Tag = "Header"
                CreatedLabel.MousePointer = fmMousePointerUpArrow
            End If

            'Wire events (selection + header click)
            Set DLBObject = New clsDesignListBoxObject
            DLBObject.Create CreatedLabel, Me
            pDesignListBoxObjectsCol.Add DLBObject

NextIteration:
        Next j
    Next i

    pIsEmpty = False

    If Me.Headers Then Me.FreezeRows = 1

    pFrame.ScrollWidth = totalW
    pFrame.ScrollHeight = (Me.RowsCount) * Me.RowHeight

    ApplyHorizontalGridlines RGB(230, 230, 230)

    If Not pFormatRules Is Nothing Then
        If pFormatRules.Count > 0 Then ApplyFormats
    End If

End Sub
Private Sub ExecuteFill_rollback(ByVal InpArr As Variant)

    Dim i As Long, j As Long
    Dim setTop As Single, setLeft As Single, setHeight As Single, setWidth As Single
    Dim CreatedLabel As MSForms.Label
    Dim InpBorderStyle As Long, InpIsBold As Boolean
    Dim DLBObject As clsDesignListBoxObject
    Dim zCol As Long
    Dim row0 As Long
    Dim cap As String
    Dim totalW As Long

    '---- headers collection ----
    If Me.Headers Then
        Set pHeadersCol = New Collection
    Else
        Set pHeadersCol = Nothing
    End If

    '---- reset event wrappers so they don't pile up ----
    Set pDesignListBoxObjectsCol = New Collection

    '---- precompute column lefts + total width ----
    Dim colLefts() As Long
    colLefts = BuildColLefts(totalW)   'expects BuildColLefts(ByRef totalW As Long)

    '---- build labels ----
    For i = LBound(InpArr, 1) To UBound(InpArr, 1)

        row0 = (i - LBound(InpArr, 1))

        setTop = rowTop(row0)
        setHeight = GetRowHeightSafe(row0)   '<<< SAFE (prevents subscript out of range)

        For j = LBound(InpArr, 2) To UBound(InpArr, 2)

            zCol = (j - LBound(InpArr, 2))   '0..cols-1

            'guard col array
            If zCol < LBound(colLefts) Or zCol > UBound(colLefts) Then GoTo NextIteration

            setLeft = colLefts(zCol)
            setWidth = GetColWidth(zCol)
            If setWidth <= 0 Then GoTo NextIteration

            'Header formatting
            If Me.Headers And i = LBound(InpArr, 1) Then
                InpBorderStyle = fmBorderStyleNone
                InpIsBold = True
            Else
                InpBorderStyle = fmBorderStyleNone
                InpIsBold = False
            End If

            cap = SafeCaption(InpArr(i, j))

            On Error GoTo EH_CreateLabel

            Set CreatedLabel = AddControlsToUserform( _
                pParent, pFrame, "DesignListBox", "Label", _
                row0 & ";" & zCol, _
                setHeight, setLeft, setTop, setWidth, _
                cap, cap, , , , , , , , InpBorderStyle, , , InpIsBold _
            )

            'multi-line cell settings
            With CreatedLabel
                .WordWrap = True
                .AutoSize = False
                .Height = setHeight
                .TextAlign = fmTextAlignLeft
            End With

            On Error GoTo 0

            'Header tag + visuals
            If Me.Headers And i = LBound(InpArr, 1) Then
                pHeadersCol.Add CreatedLabel
                CreatedLabel.Tag = "Header"
                CreatedLabel.MousePointer = fmMousePointerUpArrow

                CreatedLabel.BackColor = RGB(240, 240, 240)
                CreatedLabel.ForeColor = RGB(30, 30, 30)
                CreatedLabel.BorderStyle = fmBorderStyleSingle
            End If

            'Wire events
            Set DLBObject = New clsDesignListBoxObject
            DLBObject.Create CreatedLabel, Me
            pDesignListBoxObjectsCol.Add DLBObject

NextIteration:
        Next j
    Next i

    pIsEmpty = False
    If Me.Headers Then Me.FreezeRows = 1

    'Scroll extents
    pFrame.ScrollWidth = totalW

    Dim totalH As Single
    totalH = rowTop(Me.RowsCount)   'RowTop should accept row0 = RowsCount to mean "end"
    pFrame.ScrollHeight = totalH    '<<< DO NOT overwrite this later

    'Gridlines + formats
    ApplyHorizontalGridlines RGB(230, 230, 230)

    If Me.Headers Then
        ApplyHeaderVisuals
    End If

    If Not pFormatRules Is Nothing Then
        If pFormatRules.Count > 0 Then ApplyFormats
    End If

    Exit Sub

EH_CreateLabel:
    Debug.Print "ERROR creating label at i=" & i & " j=" & j & _
                " row0=" & row0 & " col0=" & zCol & _
                " value=[" & SafeCaption(InpArr(i, j)) & "]" & _
                " TypeName=" & TypeName(InpArr(i, j)) & _
                " VarType=" & VarType(InpArr(i, j)) & _
                " Err=" & Err.Number & " " & Err.Description
    Err.Clear
    On Error GoTo 0
    Resume Next

End Sub

Private Function GetRowHeightSafe(ByVal row0 As Long) As Single
    'Returns pRowHeights(row0) if allocated and in range; otherwise Me.RowHeight

    GetRowHeightSafe = Me.RowHeight

    If row0 < 0 Then Exit Function
    If Not IsArray(pRowHeights) Then Exit Function

    On Error Resume Next
    Dim lb As Long, ub As Long
    lb = LBound(pRowHeights)
    ub = UBound(pRowHeights)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    If row0 < lb Or row0 > ub Then Exit Function

    GetRowHeightSafe = CSng(pRowHeights(row0))
End Function

Private Function BuildColLefts(ByRef totalW As Long) As Long()
    Dim lefts() As Long
    Dim c As Long, running As Long
    Dim cc As Long

    totalW = 0
    cc = Me.ColumnsCount

    'Guard: no columns yet
    If cc <= 0 Then
        ReDim lefts(0 To 0)
        lefts(0) = 0
        BuildColLefts = lefts
        Exit Function
    End If

    ReDim lefts(0 To cc - 1)

    running = 0
    For c = 0 To cc - 1
        lefts(c) = running
        running = running + GetColWidth(c)
    Next c

    totalW = running
    BuildColLefts = lefts
End Function
Private Function SafeCaption(ByVal v As Variant) As String
    On Error GoTo EH

    If IsError(v) Then
        SafeCaption = "#ERR"
    ElseIf IsNull(v) Then
        SafeCaption = vbNullString
    ElseIf IsEmpty(v) Then
        SafeCaption = vbNullString
    ElseIf VarType(v) = vbError Then
        SafeCaption = "#ERR"
    Else
        SafeCaption = CStr(v)
    End If

    Exit Function
EH:
    SafeCaption = vbNullString
End Function
Private Sub EnsureHeaderMask()
    If Not Me.Headers Then Exit Sub
    If pFrame Is Nothing Then Exit Sub

    On Error Resume Next
    If pHeaderMask Is Nothing Then
        Set pHeaderMask = pFrame.Controls.Add("Forms.Label.1", "hdrmask", True)
    End If
    On Error GoTo 0

    With pHeaderMask
        .Tag = "HeaderMask"
        .Caption = vbNullString
        .BackStyle = fmBackStyleOpaque
        .BackColor = RGB(240, 240, 240)   'match your header background
        .BorderStyle = fmBorderStyleNone
        .Left = 0
        .Top = -2
        .Width = pFrame.ScrollWidth
        .Height = Me.RowHeight + 4        'or your header row height
        .Enabled = False                  'don’t intercept clicks
    End With

    'Bring mask forward, then headers forward on top of it
    pHeaderMask.ZOrder 0
    Dim h As MSForms.Label
    If Not pHeadersCol Is Nothing Then
        For Each h In pHeadersCol
            h.ZOrder 0
        Next h
    End If
End Sub
Private Function SafeDbg(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Then
        SafeDbg = "#ERR(" & CStr(v) & ")"
    Else
        SafeDbg = CStr(v)
    End If
End Function
Private Function SafeText(ByVal v As Variant) As String
    If IsError(v) Then
        SafeText = ""                 'or "N/A"
    ElseIf IsNull(v) Then
        SafeText = ""
    ElseIf LenB(v) = 0 Then           'handles Empty safely
        SafeText = ""
    Else
        SafeText = CStr(v)
    End If
End Function

Private Function To2DArray(ByVal v As Variant) As Variant
    'If v is 1-D, convert to 2-D (n x 1)
    Dim d As Integer
    d = NumberOfArrayDimensions(v)

    If d = 2 Then
        To2DArray = v
        Exit Function
    End If

    If d = 1 Then
        Dim i As Long, lb As Long, ub As Long
        lb = LBound(v): ub = UBound(v)
        Dim outArr() As Variant
        ReDim outArr(lb To ub, 0 To 0)
        For i = lb To ub
            outArr(i, 0) = v(i)
        Next i
        To2DArray = outArr
        Exit Function
    End If

    Err.Raise 123, "clsDesignListBox", "Input Array has 0 or >2 dimensions"
End Function


Public Sub ApplyHorizontalGridlines(Optional ByVal LineColor As Long = -1)

    If pFrame Is Nothing Then Exit Sub
    If Me.IsDLBEmpty Then Exit Sub

    If LineColor = -1 Then LineColor = RGB(230, 230, 230)

    Dim i As Long
    Dim nm As String

    '---- SAFE REMOVE: iterate backwards by index ----
    'MSForms.Controls is 0-based indexable (Controls(i))
    For i = pFrame.Controls.Count - 1 To 0 Step -1
        If TypeName(pFrame.Controls(i)) = "Label" Then
            nm = pFrame.Controls(i).Name
            If Left$(nm, 6) = "hline;" Then
                pFrame.Controls.Remove nm
            End If
        End If
    Next i

    'Turn off borders on all cell labels (prevents vertical lines)
    Dim L As MSForms.Label
    For i = 0 To pFrame.Controls.Count - 1
        If TypeName(pFrame.Controls(i)) = "Label" Then
            Set L = pFrame.Controls(i)
            If L.Tag <> "Header" And L.Tag <> "GridLine" Then
                L.BorderStyle = fmBorderStyleNone
            End If
        End If
    Next i

    'Create a 1px line at the bottom of each row
    Dim r As Long
    Dim topY As Single, lineName As String
    Dim lineLbl As MSForms.Label

    For r = 0 To Me.RowsCount - 1

        'Skip header row line (optional)
        If Me.Headers And r = 0 Then GoTo nextRow

        topY = rowTop(r + 1) - 1
        lineName = "hline;" & CStr(r)

        Set lineLbl = pFrame.Controls.Add("Forms.Label.1", lineName, True)
        With lineLbl
            .Tag = "GridLine"
            .Left = 0
            .Top = topY
            .Width = pFrame.ScrollWidth
            .Height = 1
            .BackStyle = fmBackStyleOpaque
            .BackColor = LineColor
            .Caption = vbNullString
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat

            .Enabled = False    '<<< key: don’t intercept mouse
            .ZOrder 0           '<<< key: bring to front so it’s visible
        End With

nextRow:
    Next r

    'Header divider (optional)
    If Me.Headers Then
        Dim hdrLine As MSForms.Label
        Set hdrLine = pFrame.Controls.Add("Forms.Label.1", "hline;header", True)
        With hdrLine
            .Tag = "GridLine"   ' <-- FIXED (was lineLbl.Tag)
            .Left = 0
            .Top = Me.RowHeight - 1
            .Width = pFrame.ScrollWidth
            .Height = 2
            .BackStyle = fmBackStyleOpaque
            .BackColor = RGB(180, 180, 180)
            .Caption = vbNullString
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .Enabled = False
            .ZOrder 0
        End With
    End If

End Sub
Private Function IsCellLabel(ByVal L As MSForms.Label) As Boolean
    Dim parts() As String

    If L Is Nothing Then Exit Function
    If L.Tag = "Header" Or L.Tag = "GridLine" Then Exit Function

    parts = Split(L.Name, ";")
    If UBound(parts) < 1 Then Exit Function

    If Not IsNumeric(parts(0)) Then Exit Function
    If Not IsNumeric(parts(1)) Then Exit Function

    IsCellLabel = True
End Function
Public Sub ApplyExcelGridlines()

    Dim Labl As MSForms.Label
    Dim rowIdx As Long

    For Each Labl In pFrame.Controls

        If Labl.Tag <> "Header" Then

            rowIdx = CLng(Split(Labl.Name, ";")(0))

            ' Apply light horizontal separator
            Labl.BorderStyle = fmBorderStyleSingle
            'Labl.BorderColor = pGridLineColor
            Labl.BorderColor = RGB(150, 150, 150)

        Else
            ' Optional: Stronger header separator
            Labl.BorderStyle = fmBorderStyleSingle
            Labl.BorderColor = RGB(180, 180, 180)
        End If

    Next Labl

End Sub


Private Function GetDeviation(Mode As String) As Long
'Mode can be "Rows" or "Columns"

If Mode = "Rows" Then
    GetDeviation = pParamDict("DeviationRows")
ElseIf Mode = "Columns" Then
    GetDeviation = pParamDict("DeviationColumns")
End If



End Function
Private Sub ActivateDeactivateHeaders(isOn As Boolean)

Dim Labl

If isOn Then
    Set pHeadersCol = New Collection
    Set pHeadersCol = Me.RowLabels(0)
    For Each Labl In pHeadersCol
        Labl.BorderStyle = 1
        Labl.Font.Bold = True
        Labl.Tag = "Header"
    Next Labl
Else
For Each Labl In pHeadersCol
    Labl.BorderStyle = 1
    Labl.Font.Bold = True
    Labl.Font.Size = 9
    Labl.Font.Name = "Segoe UI"
    Labl.Tag = "Header"
Next Labl
    Set pHeadersCol = New Collection
End If

End Sub
Private Sub SelectCorrectLabels(RowNumber As Long)

Dim Labl
For Each Labl In pParent.Controls
    If Split(Labl.Name, ";")(0) = RowNumber Then
        Labl.BackColor = Me.SelectionColor
    End If
Next Labl

End Sub
Private Sub DeselectAllLabels()

    Dim Labl As MSForms.Label

    For Each Labl In pFrame.Controls
        If Labl.Tag <> "Header" Then
            If Labl.BackColor = Me.SelectionColor Then
                Labl.BackColor = vbWhite
            End If
        End If
    Next Labl

End Sub
Private Function CheckInputRowNumber(ByRef RowNumber As Long) As Boolean

Err.Clear

RowNumber = Round(RowNumber)

If Me.IsDLBEmpty Then
    Call IsEmptyErr
ElseIf RowNumber < 0 Then
    Err.Raise 123, "clsDesignListBox", "RowNumber must be bigger than 0"
ElseIf RowNumber > Me.RowsCount - 1 Then
    Err.Raise 123, "clsDesignListBox", "RowNumber must be smaller than " & Me.RowsCount - 1, "Rows are zero based"
End If

If Err.Number = 0 Then
    CheckInputRowNumber = True
End If

End Function

Private Function MaxLong(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then
        MaxLong = a
    Else
        MaxLong = b
    End If
End Function

Private Function ClampLong(ByVal v As Long, ByVal lo As Long, ByVal hi As Long) As Long
    If v < lo Then v = lo
    If v > hi Then v = hi
    ClampLong = v
End Function


Private Function CheckInputColumnNumber(ByRef ColumnNumber As Long) As Boolean

Err.Clear

ColumnNumber = Round(ColumnNumber)

If Me.IsDLBEmpty Then
    Call IsEmptyErr
ElseIf ColumnNumber < 0 Then
    Err.Raise 123, "clsDesignListBox", "ColumnNumber must be bigger than 0"
ElseIf ColumnNumber > Me.ColumnsCount - 1 Then
    Err.Raise 123, "clsDesignListBox", "ColumnNumber must be smaller than " & Me.ColumnsCount - 1, "Columns are zero based"
End If

If Err.Number = 0 Then
    CheckInputColumnNumber = True
End If


End Function
Private Sub IsEmptyErr()

Err.Raise 123, "clsDesignListBox", "DesignListBox is empty", "Use .create or .fill to instantiate DesignListBox"

End Sub


Private Function NumberOfArrayDimensions(ByVal v As Variant) As Long
    'Only needs to distinguish 1D vs 2D
    If Not IsArray(v) Then
        NumberOfArrayDimensions = 0
        Exit Function
    End If

    On Error Resume Next

    Dim ub1 As Long
    ub1 = UBound(v, 1)
    If Err.Number <> 0 Then
        Err.Clear
        NumberOfArrayDimensions = 0   'unallocated / not really an array
        Exit Function
    End If

    Dim ub2 As Long
    ub2 = UBound(v, 2)
    If Err.Number = 0 Then
        NumberOfArrayDimensions = 2
    Else
        Err.Clear
        NumberOfArrayDimensions = 1
    End If

    On Error GoTo 0
End Function
Private Function AddControlsToUserform(UForm As UserForm, OutpLocation As Object, strOutputLocation As String, InpType, InpName, InpHeight, inpLeft, inpTop, _
inpWidth, Optional InpCaption, Optional InpValue, Optional InpTextAlign, Optional InpTextSize, _
Optional InpForeColor, Optional InpItems As Collection, Optional InpListIndex, Optional InpTipText, _
Optional InpSpecialEffect, Optional InpBorderStyle, Optional InpVisibility As Boolean = True, Optional InpParent As Boolean, Optional InpIsBold) As MSForms.Control
'InpType can be: CheckBox, ComboBox, CommandButton,Frame,Image,Label,ListBox,MultiPage,OptionButton,ScrollBar
'SpinButton,TabStrip,TextBox,ToggleButton
'''''''''''''''''''''''
'If a problem with adding objects to a frame arises (automation error), then its probably because the frame is not the last
'created frame on the userform. The frame to fill has to be the last created (on runtime or on designtime)
'''''''''''''''''''''''
'ControlClass defines the WithEvents Class assigned to the control. 0 means, that no class is assigned

Dim ProgId As String
ProgId = "Forms." & InpType & ".1"

Dim NewForm, ControlToEdit

If Not InpParent Then
    Set NewForm = OutpLocation.Controls.Add(ProgId, InpName)
Else
    Set NewForm = OutpLocation.Add(ProgId, InpName)
End If
Set ControlToEdit = NewForm

With ControlToEdit
    On Error Resume Next
    .Height = InpHeight
    .Left = inpLeft
    .Top = inpTop
    .Width = inpWidth
    If Not IsMissing(InpCaption) Then
        .Caption = InpCaption
    End If
    If Not IsMissing(InpValue) Then
        .Value = InpValue
    End If
    If Not IsMissing(InpTextAlign) Then
        .TextAlign = InpTextAlign
    End If
    If Not IsMissing(InpTextSize) Then
        .Font.Size = InpTextSize
    End If
    If Not IsMissing(InpForeColor) Then
        .ForeColor = InpForeColor
    End If
    'items: safe, only for list controls
    If Not InpItems Is Nothing Then
        Select Case TypeName(ControlToEdit)
            Case "ListBox", "ComboBox"
                Dim Item As Variant
                For Each Item In InpItems
                    .AddItem Item
                Next Item
        End Select
    End If
    If Not IsMissing(InpListIndex) Then
        .ListIndex = InpListIndex
    End If
    If Not IsMissing(InpTipText) Then
        .ControlTipText = InpTipText
        .MousePointer = fmMousePointerHelp
    End If
    If Not IsMissing(InpSpecialEffect) Then
        .SpecialEffect = InpSpecialEffect
    End If
    If Not IsMissing(InpBorderStyle) Then
        .BorderStyle = InpBorderStyle
    End If
    
    If Not IsMissing(InpIsBold) Then .Font.Bold = InpIsBold
    .Visible = InpVisibility
End With

OutpLocation.Repaint

Set AddControlsToUserform = ControlToEdit

End Function
Private Sub SortArray(ByRef InpArr As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional lngColumn As Long = 0, Optional Descending As Boolean)


If NumberOfArrayDimensions(InpArr) = 1 Then
    Call QuickSortSingleDimArray(InpArr, lngMin, lngMax)
Else
    Call QuickSortMultiDimArray(InpArr, lngMin, lngMax, lngColumn)
End If

If Descending Then
    Call ReverseArrayInPlace(InpArr, lngMin, lngMax)
End If

End Sub
Private Sub QuickSortSingleDimArray(ByRef InpArr As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1)
    On Error Resume Next

    'Sort a 1-Dimensional array

    ' SampleUsage: sort arrData
    '
    '   QuickSortSingleDimArray arrData

    '
    ' Originally posted by Jim Rech 10/20/98 Excel.Programming


    ' Modifications, Nigel Heffernan:
    '       ' Escape failed comparison with an empty variant in the array
    '       ' Defensive coding: check inputs

    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim varX As Variant

    If TypeName(InpArr) = "Empty" Then
        Exit Sub
    End If
    If InStr(TypeName(InpArr), "()") < 1 Then  'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(InpArr)
    End If
    If lngMax = -1 Then
        lngMax = UBound(InpArr)
    End If
    If lngMin >= lngMax Then    ' no sorting required
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = InpArr((lngMin + lngMax) \ 2)

    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(InpArr(n)) - varMid *might* pick up a default member or property
        i = lngMax
        j = lngMin
    ElseIf TypeName(varMid) = "Empty" Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    While i <= j
        
        While (CompareForSort(InpArr(i), varMid) < 0) And (i < lngMax)
            i = i + 1
        Wend
        While (CompareForSort(varMid, InpArr(j)) < 0) And (j > lngMin)
            j = j - 1
        Wend

        If i <= j Then
            ' Swap the item
            varX = InpArr(i)
            InpArr(i) = InpArr(j)
            InpArr(j) = varX

            i = i + 1
            j = j - 1
        End If

    Wend

    If (lngMin < j) Then Call QuickSortSingleDimArray(InpArr, lngMin, j)
    If (i < lngMax) Then Call QuickSortSingleDimArray(InpArr, i, lngMax)

End Sub
Private Sub QuickSortMultiDimArray(ByRef InpArr As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional lngColumn As Long = 0)
    On Error Resume Next

    'Sort a 2-Dimensional array

    ' SampleUsage: sort arrData by the contents of column 3
    '
    '   QuickSortArray arrData, , , 3

    '
    'Posted by Jim Rech 10/20/98 Excel.Programming

    'Modifications, Nigel Heffernan:

    '       ' Escape failed comparison with empty variant
    '       ' Defensive coding: check inputs

    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim arrRowTemp As Variant
    Dim lngColTemp As Long

    If TypeName(InpArr) = "Empty" Then
        Exit Sub
    End If
    If InStr(TypeName(InpArr), "()") < 1 Then  'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(InpArr, 1)
    End If
    If lngMax = -1 Then
        lngMax = UBound(InpArr, 1)
    End If
    If lngMin >= lngMax Then    ' no sorting required
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = InpArr((lngMin + lngMax) \ 2, lngColumn)

    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(InpArr(n)) - varMid *might* pick up a valid default member or property
        i = lngMax
        j = lngMin
    ElseIf TypeName(varMid) = "Empty" Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    While i <= j
        While (CompareForSort(InpArr(i, lngColumn), varMid) < 0) And (i < lngMax)
            i = i + 1
        Wend
        While (CompareForSort(varMid, InpArr(j, lngColumn)) < 0) And (j > lngMin)
            j = j - 1
        Wend

        If i <= j Then
            ' Swap the rows
            ReDim arrRowTemp(LBound(InpArr, 2) To UBound(InpArr, 2))
            For lngColTemp = LBound(InpArr, 2) To UBound(InpArr, 2)
                arrRowTemp(lngColTemp) = InpArr(i, lngColTemp)
                InpArr(i, lngColTemp) = InpArr(j, lngColTemp)
                InpArr(j, lngColTemp) = arrRowTemp(lngColTemp)
            Next lngColTemp
            Erase arrRowTemp

            i = i + 1
            j = j - 1
        End If
    Wend

    If (lngMin < j) Then Call QuickSortMultiDimArray(InpArr, lngMin, j, lngColumn)
    If (i < lngMax) Then Call QuickSortMultiDimArray(InpArr, i, lngMax, lngColumn)
    
End Sub
Private Function ReverseArrayInPlace(InputArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, _
    Optional Reversed As Boolean = False, Optional DoubleReversed As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ReverseArrayInPlace
' This procedure reverses the order of an array in place -- this is, the array variable
' If Reversed is true then it does reorder it in the other dimension
' in the calling procedure is reversed. This works only on arrays
' of simple data types (String, Single, Double, Integer, Long). It will not work
' on arrays of objects. Use ReverseArrayOfObjectsInPlace to reverse an array of objects.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Temp As Variant
Dim Ndx As Long
Dim Ndx2 As Long
Dim i

If lngMin = -1 Then lngMin = LBound(InputArray, 1)
If lngMax = -1 Then lngMax = UBound(InputArray, 1)

'''''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''''
ReverseArrayInPlace = False

Ndx2 = lngMax
''''''''''''''''''''''''''''''''''''''
' loop from the LBound of InputArray to
' the midpoint of InputArray
''''''''''''''''''''''''''''''''''''''
If Reversed = False Then
    For Ndx = lngMin To ((lngMax - lngMin + 1) \ 2)
        For i = LBound(InputArray, 2) To UBound(InputArray, 2)
            'swap the elements
            Temp = InputArray(Ndx, i)
            InputArray(Ndx, i) = InputArray(Ndx2, i)
            InputArray(Ndx2, i) = Temp
        Next i
        ' decrement the upper index
        Ndx2 = Ndx2 - 1
    Next Ndx
    If DoubleReversed = True Then GoTo Rev
ElseIf Reversed = True Or DoubleReversed = True Then
Rev:
    Ndx2 = UBound(InputArray, 2)
    For Ndx = LBound(InputArray, 2) To Int(((UBound(InputArray, 2) - LBound(InputArray, 2) + 1) \ 2))
        For i = LBound(InputArray, 1) To UBound(InputArray, 1)
            'swap the elements
            Temp = InputArray(i, Ndx)
            InputArray(i, Ndx) = InputArray(i, Ndx2)
            InputArray(i, Ndx2) = Temp
        Next i
        ' decrement the upper index
        Ndx2 = Ndx2 - 1
    Next Ndx
End If

''''''''''''''''''''''''''''''''''''''
' OK - Return True
''''''''''''''''''''''''''''''''''''''
ReverseArrayInPlace = True

End Function

Private Function GetPositionInArr(Content As Variant, ColumnNumber, LastRowInput As Variant) As String

Dim PivotCol As Collection

Dim i, Element
For i = LBound(pInpArr, 1) To UBound(pInpArr, 1)
    If CStr(Content) = CStr(pInpArr(i, ColumnNumber + GetDeviation("Columns"))) Then
        If TypeName(LastRowInput) = "Collection" Then
            Set PivotCol = LastRowInput
            If Not IsInCollection(PivotCol, , i - GetDeviation("Rows") & ";" & ColumnNumber) Then
                GetPositionInArr = i - GetDeviation("Rows") & ";" & ColumnNumber
                'Debug.Assert i - GetDeviation("Rows") <> 0
                Exit Function
            End If
        Else
            GetPositionInArr = i - GetDeviation("Rows") & ";" & ColumnNumber
            'Debug.Assert i - GetDeviation("Rows") <> 0
            Exit Function
        End If
    End If
Next i

Debug.Print "Result: None"

End Function
Public Function IsInCollection(Kollection As Collection, Optional key As Variant, Optional Item As Variant) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INPUT       : Kollection, the collection we would like to examine
'            : (Optional) Key, the Key we want to find in the collection
'            : (Optional) Item, the Item we want to find in the collection
'OUTPUT      : True if Key or Item is found, False if not
'SPECIAL CASE: If both Key and Item are missing, return False
    
    Dim strKey As String
    Dim var As Variant

    'First, investigate assuming a Key was provided
    If Not IsMissing(key) Then
    
        strKey = CStr(key)
        
        'Handling errors is the strategy here
        On Error Resume Next
            IsInCollection = True
            var = Kollection(strKey) '<~ this is where our (potential) error will occur
            If Err.Number = 91 Then GoTo CheckForObject
            If Err.Number = 5 Then GoTo NotFound
        On Error GoTo 0
        Exit Function

CheckForObject:
        If IsObject(Kollection(strKey)) Then
            IsInCollection = True
            On Error GoTo 0
            Exit Function
        End If

NotFound:
        IsInCollection = False
        On Error GoTo 0
        Exit Function
        
    'If the Item was provided but the Key was not, then...
    ElseIf Not IsMissing(Item) Then
    
        IsInCollection = False '<~ assume that we will not find the item
    
        'We have to loop through the collection and check each item against the passed-in Item
        For Each var In Kollection
            If TypeName(var) = "clsSymbol" Then var = var.ID
            If TypeName(var) = "clsDividedSymbol" Then var = var.CombinedId
            If var = Item Then
                IsInCollection = True
                Exit Function
            End If
        Next var
    
    'Otherwise, no Key OR Item was provided, so we default to False
    Else
        IsInCollection = False
    End If
    
End Function
'Private Sub pKeySink_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    Debug.Print "DLB KeyDown: "; CLng(KeyCode), Shift
'    HandleKey CLng(KeyCode), CLng(Shift)
'End Sub

Public Sub HandleKey(ByVal KeyCode As Long, ByVal Shift As Long)
    'Shift bit flags:
    ' 1 = Shift, 2 = Ctrl, 4 = Alt
    Dim isShift As Boolean, isCtrl As Boolean
    isShift = ((Shift And 1) <> 0)
    isCtrl = ((Shift And 2) <> 0)

    Dim cur As Long, nxt As Long
    cur = Me.ListIndex

    'If nothing selected yet, pick first data row
    If cur < 0 Then cur = firstDataRow()

    Select Case KeyCode
        Case vbKeyUp
            nxt = cur - 1
        Case vbKeyDown
            nxt = cur + 1
        Case vbKeyHome
            nxt = firstDataRow()
        Case vbKeyEnd
            nxt = LastDataRow()
        Case Else
            Exit Sub
    End Select

    'Clamp + skip header row
    If nxt < firstDataRow() Then nxt = firstDataRow()
    If nxt > LastDataRow() Then nxt = LastDataRow()

    'If we have headers, never land on row 0
    If Me.Headers And nxt = 0 Then
        If KeyCode = vbKeyUp Then nxt = 1 Else nxt = 1
    End If

    'Apply selection logic
    If isShift And (pMultiSelectMode = dlbExtended) Then
        'extend range using existing anchor behavior in SelectRow
        Me.SelectRow nxt, False, True
    Else
        'single select move (even if ctrl is held; keep it simple/standard)
        Me.SelectRow nxt, False, False
    End If

    'Keep selected row visible
    EnsureRowVisible nxt

    'Keep focus on sink so repeated arrow presses work
    Me.FocusKeySink
End Sub
Private Sub pKeySink_KeyDown_old(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If Me.IsEmpty Then Exit Sub
    If Me.RowsCount <= 0 Then Exit Sub

    Dim isShift As Boolean, isCtrl As Boolean
    isShift = ((Shift And 1) <> 0) 'Shift key
    isCtrl = ((Shift And 2) <> 0)  'Ctrl key

    Dim cur As Long
    cur = Me.ListIndex
    If cur < 0 Then cur = firstDataRow()

    Dim nextRow As Long
    nextRow = cur

    Dim page As Long
    page = CLng(pFrame.InsideHeight \ Me.RowHeight)
    If page < 1 Then page = 1

    Select Case KeyCode
        Case vbKeyUp
            nextRow = cur - 1

        Case vbKeyDown
            nextRow = cur + 1

        Case vbKeyHome
            nextRow = firstDataRow()

        Case vbKeyEnd
            nextRow = LastDataRow()

        Case vbKeyPageUp
            nextRow = cur - page

        Case vbKeyPageDown
            nextRow = cur + page

        Case vbKeySpace, vbKeyReturn
            'Optional: Space toggles current row in multi modes
            If Me.MultiSelectMode <> dlbSingle Then
                Call Me.SelectRow(cur, True, False) 'toggle via ctrl behavior
            Else
                Call Me.SelectRow(cur, False, False)
            End If
            EnsureRowVisible cur
            KeyCode = 0
            Exit Sub

        Case Else
            Exit Sub
    End Select

    nextRow = ClampLong(nextRow, firstDataRow(), LastDataRow())

    Call Me.SelectRow(nextRow, isCtrl, isShift)
    EnsureRowVisible nextRow

    KeyCode = 0
End Sub

