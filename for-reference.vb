Sub ForReference()
    Dim rng As Range
    Dim char As Range

    ' Exit if there is no selection or if the selection is empty
    If Selection Is Nothing Or Selection.Range.Text = "" Then
        MsgBox "Please select text to apply the highlight change.", vbExclamation
        Exit Sub
    End If

    ' Speed up the macro by disabling screen updates and other processes
    Application.ScreenUpdating = False

    ' Set the range to the current selection
    Set rng = Selection.Range

    ' Loop through each word (or character range) within the selection
    For Each char In rng.Characters
        If char.HighlightColorIndex <> wdNoHighlight Then
            char.HighlightColorIndex = wdGray25 ' Change to gray highlight
        End If
    Next char

    ' Restore screen updates
    Application.ScreenUpdating = True
End Sub
