Public Enum eListType
    ltNone = 0
    ltUL = 1
    ltOL = 2
End Enum

Public Enum eTableMode
    tmDisabled = 0
    tmFlattenText = 1
End Enum

'table mode off: 
'WriteHtmlToPowerPointShape_HighPerf html, shp

'table mode on
'WriteHtmlToPowerPointShape_HighPerf html, shp, True


Public Type HtmlParseOptions
    TableMode As eTableMode
End Type

Public Sub WriteHtmlToPowerPointShape_HighPerf( _
    ByVal html As String, _
    ByVal shp As Shape, _
    Optional ByVal EnableTables As Boolean = False)

    Dim segments() As HtmlTextSegment
    Dim segCount As Long
    
    Dim fullText As String
    Dim paraBullet() As Boolean
    Dim paraLevel() As Long
    Dim paraCount As Long
    Dim paraListType() As Long
    Dim opts As HtmlParseOptions
    
    If EnableTables Then
        opts.TableMode = tmFlattenText
    Else
        opts.TableMode = tmDisabled
    End If
    
    ParseSimpleHtmlToSegments html, segments, segCount, paraBullet, paraLevel, paraListType, paraCount, opts
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
    
    segCount = 0
    paraCount = 1
    paraBullet(1) = False
    paraLevel(1) = 0
    paraListType(1) = ltNone
    
    WalkHtmlNodes doc.body, segments, segCount, _
                  False, False, False, False, 0, _
                  paraBullet, paraLevel, paraListType, paraCount, _
                  ltNone, 0, opts
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
    
    segCount = 0
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
    ByVal currentListLevel As Long, _
    ByRef opts As HtmlParseOptions)

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
            
        Case "table"
            If opts.TableMode = tmFlattenText Then
                AddFlattenedTable node, segments, segCount, paraCount
                Exit Sub
            End If
            
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
    End Select
    
    For Each child In node.childNodes
        WalkHtmlNodes child, segments, segCount, _
                      nextBold, nextItalic, nextUnderline, _
                      nextHasColor, nextColor, _
                      paraBullet, paraLevel, paraListType, paraCount, _
                      nextListType, nextListLevel, opts
    Next child
End Sub

Private Sub AddFlattenedTable( _
    ByVal tableNode As Object, _
    ByRef segments() As HtmlTextSegment, _
    ByRef segCount As Long, _
    ByRef paraCount As Long)

    Dim rowNodes As Object
    Dim rowNode As Object
    Dim cellNode As Object
    Dim rowText As String
    Dim cellText As String
    Dim firstCell As Boolean
    
    On Error Resume Next
    Set rowNodes = tableNode.getElementsByTagName("tr")
    On Error GoTo 0
    
    If rowNodes Is Nothing Then Exit Sub
    
    If segCount > 0 Then
        AddSegment segments, segCount, vbCr, False, False, False, False, 0, paraCount, False, 0
        paraCount = paraCount + 1
    End If
    
    For Each rowNode In rowNodes
        rowText = vbNullString
        firstCell = True
        
        For Each cellNode In rowNode.childNodes
            Select Case LCase$(cellNode.nodeName)
                Case "th", "td"
                    cellText = CleanNodeInnerText(cellNode)
                    If Len(cellText) > 0 Then
                        If Not firstCell Then rowText = rowText & " | "
                        rowText = rowText & cellText
                        firstCell = False
                    End If
            End Select
        Next cellNode
        
        If Len(rowText) > 0 Then
            AddSegment segments, segCount, rowText, False, False, False, False, 0, paraCount, False, 0
            AddSegment segments, segCount, vbCr, False, False, False, False, 0, paraCount, False, 0
            paraCount = paraCount + 1
        End If
    Next rowNode
End Sub

Private Function CleanNodeInnerText(ByVal node As Object) As String
    Dim s As String
    
    On Error Resume Next
    s = node.innerText
    On Error GoTo 0
    
    s = DecodeHtmlEntities(s)
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    s = Replace$(s, vbTab, " ")
    
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    
    CleanNodeInnerText = Trim$(s)
End Function
