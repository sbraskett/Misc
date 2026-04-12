Public Enum eListType
    ltNone = 0
    ltUL = 1
    ltOL = 2
End Enum

Public Enum eTableMode
    tmDisabled = 0
    tmFlattenText = 1
    tmPptTableShape = 2
End Enum

Public Type HtmlParseOptions
    TableMode As eTableMode
End Type

opts.TableMode = tmFlattenText

Public Sub WriteHtmlToPowerPointShape_HighPerf(ByVal html As String, ByVal shp As Shape)
    Dim segments() As HtmlTextSegment
    Dim segCount As Long
    
    Dim fullText As String
    Dim paraBullet() As Boolean
    Dim paraLevel() As Long
    Dim paraCount As Long
    Dim paraListType() As Long
    
    ParseSimpleHtmlToSegments html, segments, segCount, paraBullet, paraLevel, paraListType, paraCount
    fullText = BuildFullTextAndMapPositions(segments, segCount)
    
    With shp.TextFrame2
        .WordWrap = msoTrue
        .TextRange.Text = fullText
    End With
    
    ApplyParagraphFormatting shp.TextFrame2.TextRange, paraBullet, paraLevel, paraListType, paraCount
    ApplyCharacterFormattingRuns shp.TextFrame2.TextRange, segments, segCount
End Sub

Public Sub ParseSimpleHtmlToSegments( _
    ByVal html As String, _
    ByRef segments() As HtmlTextSegment, _
    ByRef segCount As Long, _
    ByRef paraBullet() As Boolean, _
    ByRef paraLevel() As Long, _
    ByRef paraListType() As Long, _
    ByRef paraCount As Long, _
    ByRef opts As HtmlParseOptions)

    Dim doc As Object
    Set doc = CreateObject("HTMLFILE")
    
    doc.body.innerHTML = html
    
    ReDim segments(1 To 1)
    ReDim paraBullet(1 To 1)
    ReDim paraLevel(1 To 1)
    ReDim paraListType(1 To 1)
    
    paraCount = 1
    paraBullet(1) = False
    paraLevel(1) = 0
    paraListType(1) = ltNone
    
    WalkHtmlNodes doc.body, segments, segCount, _
                  False, False, False, False, 0, _
                  paraBullet, paraLevel, paraListType, paraCount, _
                  ltNone, 0, opts
End Sub

Private Sub WalkHtmlNodes( _
    ByVal node As Object, _
    ByRef segments() As HtmlTextSegment, _
    ByRef segCount As Long, _
    ByVal curBold As Boolean, _
    ByVal curItalic As Boolean, _
    ByVal curUnderline As Boolean, _
    ByVal curHasColor As Boolean, _
    ByVal curColor As Long, _
    ByRef paraBullet() As Boolean, _
    ByRef paraLevel() As Long, _
    ByRef paraListType() As Long, _
    ByRef paraCount As Long, _
    ByVal currentListType As Long, _
    ByVal currentListLevel As Long)

    Dim child As Object
    Dim tagName As String
    
    Dim nextBold As Boolean
    Dim nextItalic As Boolean
    Dim nextUnderline As Boolean
    Dim nextHasColor As Boolean
    Dim nextColor As Long
    
    Dim nextListType As Long
    Dim nextListLevel As Long
    
    If node Is Nothing Then Exit Sub
    
    If node.nodeType = 3 Then
        If Len(node.nodeValue) > 0 Then
            AddSegment segments, segCount, node.nodeValue, _
                       curBold, curItalic, curUnderline, _
                       False, 0, paraCount, curHasColor, curColor
        End If
        Exit Sub
    End If
    
    On Error Resume Next
    tagName = LCase$(node.nodeName)
    On Error GoTo 0
    
    nextBold = curBold
    nextItalic = curItalic
    nextUnderline = curUnderline
    nextHasColor = curHasColor
    nextColor = curColor
    
    nextListType = currentListType
    nextListLevel = currentListLevel
    
    Select Case tagName
        Case "b", "strong"
            nextBold = True
            
        Case "i", "em"
            nextItalic = True
            
        Case "u"
            nextUnderline = True
            
        Case "span", "font"
            If TryGetNodeColor(node, nextColor) Then
                nextHasColor = True
            End If
            
        Case "ul"
            nextListType = ltUL
            nextListLevel = currentListLevel + 1
            
        Case "ol"
            nextListType = ltOL
            nextListLevel = currentListLevel + 1
            
        Case "br"
            AddSegment segments, segCount, vbCr, False, False, False, False, 0, paraCount, False, 0
            paraCount = paraCount + 1
            GrowParagraphArrays paraBullet, paraLevel, paraListType, paraCount
            
        Case "p", "div"
            If segCount > 0 Then
                AddSegment segments, segCount, vbCr, False, False, False, False, 0, paraCount, False, 0
                paraCount = paraCount + 1
                GrowParagraphArrays paraBullet, paraLevel, paraListType, paraCount
            End If
            
        Case "li"
            If segCount > 0 Then
                AddSegment segments, segCount, vbCr, False, False, False, False, 0, paraCount, False, 0
                paraCount = paraCount + 1
                GrowParagraphArrays paraBullet, paraLevel, paraListType, paraCount
            End If
            
            paraBullet(paraCount) = (nextListType <> ltNone)
            paraLevel(paraCount) = IIf(nextListLevel > 0, nextListLevel, 1)
            paraListType(paraCount) = nextListType
		Case "table"
			Select Case opts.TableMode
				Case tmDisabled
					' do nothing special; either ignore or walk children normally

				Case tmFlattenText
					AddFlattenedTable node, segments, segCount, paraCount
					Exit Sub

				Case tmPptTableShape
					' future: extract structured table data instead
					Exit Sub
			End Select
	End Select
    
    For Each child In node.childNodes
        WalkHtmlNodes child, segments, segCount, _
                      nextBold, nextItalic, nextUnderline, _
                      nextHasColor, nextColor, _
                      paraBullet, paraLevel, paraListType, paraCount, _
                      nextListType, nextListLevel
    Next child
End Sub

Private Sub GrowParagraphArrays( _
    ByRef paraBullet() As Boolean, _
    ByRef paraLevel() As Long, _
    ByRef paraListType() As Long, _
    ByVal newSize As Long)

    If newSize > UBound(paraBullet) Then
        ReDim Preserve paraBullet(1 To newSize * 2)
        ReDim Preserve paraLevel(1 To newSize * 2)
        ReDim Preserve paraListType(1 To newSize * 2)
    End If
End Sub

ApplyParagraphFormatting2 shp.TextFrame2.TextRange, paraBullet, paraLevel, paraListType, paraCount
ApplyCharacterFormattingRuns2 shp.TextFrame2.TextRange, segments, segCount

Private Sub ApplyParagraphFormatting( _
    ByVal tr As Object, _
    ByRef paraBullet() As Boolean, _
    ByRef paraLevel() As Long, _
    ByRef paraListType() As Long, _
    ByVal paraCount As Long)

    Dim i As Long
    Dim paraTotal As Long
    
    paraTotal = tr.Paragraphs.Count
    If paraTotal < paraCount Then paraCount = paraTotal
    
    For i = 1 To paraCount
        With tr.Paragraphs(i).ParagraphFormat
            Select Case paraListType(i)
                Case ltUL
                    On Error Resume Next
                    .Bullet.Visible = msoFalse
                    .Bullet.Visible = msoTrue
                    .Bullet.Type = msoBulletUnnumbered
                    .Bullet.Character = 8226
                    If paraLevel(i) > 0 Then .IndentLevel = paraLevel(i)
                    On Error GoTo 0
                    
                Case ltOL
                    On Error Resume Next
                    .Bullet.Visible = msoFalse
                    .Bullet.Visible = msoTrue
                    .Bullet.Type = msoBulletNumbered
                    If paraLevel(i) > 0 Then .IndentLevel = paraLevel(i)
                    On Error GoTo 0
                    
                Case Else
                    On Error Resume Next
                    .Bullet.Visible = msoFalse
                    On Error GoTo 0
            End Select
        End With
    Next i
End Sub

Private Sub ApplyCharacterFormattingRuns(ByVal tr As Object, ByRef segments() As HtmlTextSegment, ByVal segCount As Long)
    Dim i As Long
    Dim runStart As Long
    Dim runLen As Long
    Dim j As Long
    
    i = 1
    Do While i <= segCount
        If segments(i).TextLength > 0 And segments(i).Text <> vbCr Then
            runStart = segments(i).StartPos
            runLen = segments(i).TextLength
            
            j = i + 1
            Do While j <= segCount
                If CanMergeSegments(segments(i), segments(j)) Then
                    runLen = runLen + segments(j).TextLength
                    j = j + 1
                Else
                    Exit Do
                End If
            Loop
            
            ApplyFormatToRange2 tr.Characters(runStart, runLen), segments(i)
            i = j
        Else
            i = i + 1
        End If
    Loop
End Sub

Private Sub ApplyFormatToRange(ByVal tr As Object, ByRef seg As HtmlTextSegment)
    With tr.Font
        .Bold = IIf(seg.Bold, msoTrue, msoFalse)
        .Italic = IIf(seg.Italic, msoTrue, msoFalse)
        .UnderlineStyle = IIf(seg.Underline, msoUnderlineSingleLine, msoNoUnderline)
        
        If seg.HasColor Then
            .Fill.ForeColor.RGB = seg.ColorValue
		Else
			.Fill.ForeColor.RGB = RGB(0,0,0)
        End If
    End With
End Sub

Private Sub AddFlattenedTable( _
    ByVal tableNode As Object, _
    ByRef segments() As HtmlTextSegment, _
    ByRef segCount As Long, _
    ByVal paraCount As Long)

    Dim rowNode As Object
    Dim cellNode As Object
    Dim rowText As String
    Dim cellText As String
    Dim firstCell As Boolean
    
    For Each rowNode In tableNode.getElementsByTagName("tr")
        rowText = vbNullString
        firstCell = True
        
        For Each cellNode In rowNode.childNodes
            Select Case LCase$(cellNode.nodeName)
                Case "th", "td"
                    cellText = CleanNodeInnerText(cellNode)
                    If Not firstCell Then rowText = rowText & " | "
                    rowText = rowText & cellText
                    firstCell = False
            End Select
        Next cellNode
        
        If Len(rowText) > 0 Then
            AddSegment segments, segCount, rowText, False, False, False, False, 0, paraCount, False, 0
            AddSegment segments, segCount, vbCr, False, False, False, False, 0, paraCount, False, 0
        End If
    Next rowNode
End Sub

Private Function CleanNodeInnerText(ByVal node As Object) As String
    Dim s As String
    
    On Error Resume Next
    s = node.innerText
    On Error GoTo 0
    
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    
    CleanNodeInnerText = Trim$(s)
End Function