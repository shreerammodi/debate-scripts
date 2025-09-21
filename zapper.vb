Private Sub Zap()
    ZapRange ActiveDocument.Content
End Sub

Private Sub ZapRange(ByVal targetRange As Range)
    Dim styles As Variant
    Dim s As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Styles that should be preserved in the read doc
    styles = Array("Pocket", "Hat", "Block", "Tag", "Cite", "Analytic", "Undertag")
    If Options.DefaultHighlightColorIndex = wdNoHighlight Then
        Options.DefaultHighlightColorIndex = wdBrightGreen
    End If

    ' First pass: turn on highlights for styles
    For Each s In styles
        With targetRange.Find
            .ClearFormatting
            .Style = s
            .Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchWildcards = True

            With .Replacement
                .ClearFormatting
                .Text = "^&"
                .Highlight = True
            End With

            .Execute Replace:=wdReplaceAll
        End With
    Next s


    ' Second pass: Delete anything that isn't highlighted
    With targetRange.Find
        .ClearFormatting
        .Highlight = False
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True

        With .Replacement
            .ClearFormatting
            .Text = "^p"
        End With
        .Execute Replace:=wdReplaceAll
    End With

    ' Third pass: Remove highlighting from styles
    For Each s In styles
        With targetRange.Find
            .ClearFormatting
            .Style = s
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchWildcards = True

            With .Replacement
                .Text = "^&"
                .ClearFormatting
                .Highlight = False
            End With
            .Execute Replace:=wdReplaceAll
        End With
    Next s

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Private Sub CondenseCards()
    Dim p As Paragraph

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each p In ActiveDocument.Paragraphs
        ' Only process Tags
        if p.OutlineLevel = wdOutlineLevel4 Then
            Set CondenseRange = SelectCardTextRange(p)

            ' Drop trailing paragraph mark if present
            If CondenseRange.Characters.Last.Text = Chr(13) Then
                CondenseRange.MoveEnd wdCharacter, -1
            End If

            ' Skip cards with nothing but the Tag itself
            If CondenseRange.Paragraphs.Count > 1 Then
                ' Replace all paragraph marks in the card with a space
                With CondenseRange.Find
                    .ClearFormatting
                    .Text = "^p"
                    With .Replacement
                        .ClearFormatting
                        .Text = " "
                        .Highlight = False
                        .Style = ActiveDocument.Styles("Normal")
                    End With
                    .Wrap = wdFindStop
                    .Format = True
                    .Execute Replace:=wdReplaceAll
                End With
            End If
        End If
    Next p

    ' Remove duplicate spaces
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Text = " {2,}"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = True
    Application.DisplayAlerts  = True
End Sub

Public Sub CreateZappedDoc()
    Dim originalDoc As Document
    Dim originalFilePath As String
    Dim originalFolderPath As String
    Dim savePath As String
    Dim zappedDoc As Document

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ActiveDocument.Save
    Set originalDoc = ActiveDocument

    Set ZappedDoc = Documents.Add(ActiveDocument.FullName)

    Call Zap
    Call CondenseCards

    downloadsDirPath = GetDownloadsDir()
    savePath = downloadsDirPath & "[R] " & originalDoc.Name
    ActiveDocument.SaveAs2 Filename:=savePath, FileFormat:=wdFormatDocumentDefault

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Public Sub ZapSelection()
    ' Make sure something is selected
    If Selection.Type = wdNoSelection Or Len(Selection.Range.Text) = 0 Then
        MsgBox "Please select some text before running this macro.", vbExclamation
        Exit Sub
    End If

    ' Apply zapping to only the selected range
    ZapRange Selection.Range
End Sub
