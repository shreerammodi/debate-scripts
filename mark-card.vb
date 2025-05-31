Sub MarkCard()
    Dim p As Paragraph
    Dim rngCard As Range
    Dim markerText As String
    Dim markerLen As Long

    markerText = vbCrLf & "<<MARKED>>" & vbCrLf
    markerLen = Len(markerText)

    ' Insert marker text before cursor
    Selection.InsertAfter markerText

    ' Remove previous paragraph from selection
    Selection.SetRange Start:=Selection.Start + 1, End:=Selection.End

    ' Make the marker normal text
    With Selection.Range
        .Style = ActiveDocument.Styles("Normal")
        .Font.Underline = wdUnderlineNone
        .HighlightColorIndex = wdNoHighlight
    End With

    Selection.Collapse(wdCollapseEnd)

    ' Get paragraph immediately after the marker
    Set p = Selection.Paragraphs(1)

    ' Get rest of the card
    Set rngCard = Paperless.SelectCardTextRange(p)

    ' Exclude anything before marker
    rngCard.Start = Selection.Start

    rngCard.Font.Color = wdColorRed
End Sub
