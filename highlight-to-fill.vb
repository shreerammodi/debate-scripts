Sub ConvertHighlightsToFills()
    Dim rng As Range
    Dim ch As Range
    Dim fillCol As Long

    ' Make sure something is selected
    If Selection Is Nothing Or Len(Selection.Range.Text) = 0 Then
        MsgBox "Please select some text before running this macro.", vbExclamation
        Exit Sub
    End If

    ' Turn off screen updating for speed
    Application.ScreenUpdating = False

    ' Work on the selected range
    Set rng = Selection.Range

    ' Loop through each character in the selection
    For Each ch In rng.Characters
        ' Only process highlighted text (skip paragraph marks)
        If ch.HighlightColorIndex <> wdNoHighlight And ch.Text <> vbCr Then
            ' Map the current highlight color to an RGB fill
            fillCol = MapHighlightToRGB(ch.HighlightColorIndex)
            ' Apply it as a shading background
            ch.Shading.BackgroundPatternColor = fillCol
            ' Clear the original highlight
            ch.HighlightColorIndex = wdNoHighlight
        End If
    Next ch

    ' Restore updating
    Application.ScreenUpdating = True

    ' Clean up
    Set ch = Nothing
    Set rng = Nothing
End Sub

Private Function MapHighlightToRGB(highlightIndex As WdColorIndex) As Long
    Select Case highlightIndex
        Case wdYellow:      MapHighlightToRGB = RGB(255, 255, 0)
        Case wdBrightGreen: MapHighlightToRGB = RGB(0, 255, 0)
        Case wdTurquoise:   MapHighlightToRGB = RGB(0, 255, 255)
        Case wdPink:        MapHighlightToRGB = RGB(255, 0, 255)
        Case wdBlue:        MapHighlightToRGB = RGB(0, 0, 255)
        Case wdRed:         MapHighlightToRGB = RGB(255, 0, 0)
        Case wdDarkBlue:    MapHighlightToRGB = RGB(0, 0, 139)
        Case wdTeal:        MapHighlightToRGB = RGB(0, 128, 128)
        Case wdGreen:       MapHighlightToRGB = RGB(0, 128, 0)
        Case wdViolet:      MapHighlightToRGB = RGB(238, 130, 238)
        Case wdDarkRed:     MapHighlightToRGB = RGB(139, 0, 0)
        Case wdDarkYellow:  MapHighlightToRGB = RGB(128, 128, 0)
        Case wdGray25:      MapHighlightToRGB = RGB(210, 210, 210)
        Case wdGray50:      MapHighlightToRGB = RGB(128, 128, 128)
        Case Else:          MapHighlightToRGB = RGB(255, 255, 255)  ' Default to white
    End Select
End Function
