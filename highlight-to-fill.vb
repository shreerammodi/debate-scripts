Sub ConvertHighlightsToFills()
    Dim rng As Range
    Dim char As Range
    Dim highlightColor As Long

    ' Check if there is a selection
    If Selection.Type = wdSelectionIP Then
        MsgBox "No text selected. Please select the text you want to modify.", vbExclamation
        Exit Sub
    End If

    ' Set the range to the current selection
    Set rng = Selection.Range

    ' Loop through each character in the selection
    For Each char In rng.Characters
        ' Check if the character is highlighted and is not a paragraph mark
        If char.HighlightColorIndex <> wdNoHighlight And char.Text <> vbCr Then
            ' Get the RGB color based on the highlight index
            highlightColor = MapHighlightToRGB(char.HighlightColorIndex)

            ' Set the background color to the highlight color
            char.Shading.BackgroundPatternColor = highlightColor

            ' Remove the highlight
            char.HighlightColorIndex = wdNoHighlight
        End If
    Next char

    ' Clean up
    Set rng = Nothing
    Set char = Nothing
End Sub

Function MapHighlightToRGB(highlightIndex As WdColorIndex) As Long
    Select Case highlightIndex
        Case wdYellow: MapHighlightToRGB = RGB(255, 255, 0)
        Case wdBrightGreen: MapHighlightToRGB = RGB(0, 255, 0)
        Case wdTurquoise: MapHighlightToRGB = RGB(0, 255, 255)
        Case wdPink: MapHighlightToRGB = RGB(255, 0, 255)
        Case wdBlue: MapHighlightToRGB = RGB(0, 0, 255)
        Case wdRed: MapHighlightToRGB = RGB(255, 0, 0)
        Case wdDarkBlue: MapHighlightToRGB = RGB(0, 0, 139)
        Case wdTeal: MapHighlightToRGB = RGB(0, 128, 128)
        Case wdGreen: MapHighlightToRGB = RGB(0, 128, 0)
        Case wdViolet: MapHighlightToRGB = RGB(238, 130, 238)
        Case wdDarkRed: MapHighlightToRGB = RGB(139, 0, 0)
        Case wdDarkYellow: MapHighlightToRGB = RGB(128, 128, 0)
        Case wdGray25: MapHighlightToRGB = RGB(210, 210, 210)
        Case wdGray50: MapHighlightToRGB = RGB(128, 128, 128)
        Case Else: MapHighlightToRGB = RGB(255, 255, 255) ' Default to white for unknown colors
    End Select
End Function
