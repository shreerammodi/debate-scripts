Public Sub ForReferenceSlow()
    Dim selRange As Range
    Dim i As Long
    Dim charRange As Range

    If Selection.Type = wdNoSelection Or Len(Selection.Range.Text) = 0 Then
        MsgBox "Please select some text before running this macro.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Set selRange = Selection.Range

    For i = 0 To selRange.Characters.Count - 1
        Set charRange = selRange.Duplicate
        charRange.SetRange Start:=selRange.Start + i, End:=selRange.Start + i + 1

        If charRange.HighlightColorIndex <> wdNoHighlight Then
            charRange.HighlightColorIndex = wdNoHighlight
            charRange.Shading.BackgroundPatternColor = wdColorGray20
        End If
    Next i

    Application.ScreenUpdating = True
End Sub

Public Sub ForReferenceFast()
    Dim selRange As Range
    Dim procRange As Range
    Dim selEnd   As Long

    If Selection.Type = wdNoSelection Or Len(Selection.Range.Text) = 0 Then
        MsgBox "Please select some text before running this macro.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Grab your selection and remember its end
    Set selRange = Selection.Range
    selEnd = selRange.End

    Set procRange = selRange.Duplicate

    With procRange.Find
        .ClearFormatting
        .Text = ""
        .Highlight = True
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While procRange.Find.Execute
        If procRange.Start >= selEnd Then Exit Do
            procRange.HighlightColorIndex = wdNoHighlight
            procRange.Shading.BackgroundPatternColor = wdColorGray20
            procRange.Collapse wdCollapseEnd
        End If
    Loop

    Application.ScreenUpdating = True
End Sub
