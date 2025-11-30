Sub AcronymHighlight()
    Dim i As Integer

    For i = 1 To Selection.Words.Count
        Dim word As String
        word = Selection.Words(i)
        If word <> "-" And word <> "," And word <> "." Then
            Selection.Words(i).Characters(1).HighlightColorIndex = Options.DefaultHighlightColorIndex
        End If
    Next
End Sub

Sub AcronymEmphasize()
    Dim i As Integer

    For i = 1 To Selection.Words.Count
        Dim word As String
        word = Selection.Words(i)
        If word <> "-" And word <> "," And word <> "." Then
            Selection.Words(i).Characters(1).Style = "Emphasis"
        End If
    Next
End Sub

Sub AcronymUnderline()
    Dim i As Integer

    For i = 1 To Selection.Words.Count
        Dim word As String
        word = Selection.Words(i)
        If word <> "-" And word <> "," And word <> "." Then
            Selection.Words(i).Characters(1).Style = "Underline"
        End If
    Next
End Sub
