Public Sub Zap()
    ZapRange ActiveDocument.Content
End Sub

Private Sub CondenseCards()
    CondenseCardsRange ActiveDocument.Content
End Sub

Private Sub ZapRange(ByVal targetRange As Range)
    Dim styles As Variant
    Dim s As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Styles that should be preserved in the read doc
    styles = Array("Pocket", "Hat", "Block", "Tag", "Cite", "Analytic", "Analytics")

    If Options.DefaultHighlightColorIndex = wdNoHighlight Then
        Options.DefaultHighlightColorIndex = wdTurquoise
    End If

    ' First pass: turn on highlights for styles
    For Each s In styles
        If StyleExists(CStr(s)) Then
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
        End If
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
        If StyleExists(CStr(s)) Then
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
        End If
    Next s

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Private Sub CondenseCardsRange(ByVal targetRange As Range)
    Dim p As Paragraph

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each p In targetRange.Paragraphs
        ' Only process Tags
        if p.OutlineLevel = wdOutlineLevel4 Then
            Set CondenseRange = SelectCardTextRange(p)

            ' Drop beginning whitespace if present
            If CondenseRange.Characters.First.Text = Chr(13) Then
                CondenseRange.Characters(1).Delete
                CondenseRange.MoveStart wdParagraph, +1
            End If

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
    With targetRange.Find
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
    Dim savePath As String
    Dim zappedDoc As Document
    Dim baseFileName As String
    Dim counter As Integer
    Dim fileExist As Boolean
    Dim docPath As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ActiveDocument.Save
    Set originalDoc = ActiveDocument

    Set zappedDoc = Documents.Add(ActiveDocument.FullName)

    Call Zap
    Call CondenseCards

    docPath = DocDir()
    baseFileName = "[R] " & originalDoc.Name
    savePath = docPath & baseFileName

    counter = 1

    fileExist = (Dir(savePath) <> "")

    Do While fileExist
        Dim baseName As String
        Dim extension As String
        Dim dotPos As Integer

        dotPos = InStrRev(baseFileName, ".")

        baseName = Left(baseFileName, dotPos - 1)
        extension = Mid(baseFileName, dotPos)

        savePath = docPath & baseName & " (" & counter & ")" & extension
        counter = counter + 1

        fileExist = (Dir(savePath) <> "")
    Loop

    ActiveDocument.SaveAs2 Filename:=savePath, FileFormat:=wdFormatDocumentDefault

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Public Sub ZapCard()
    Dim p As Paragraph
    Dim cardRange As Range

    ' Get the paragraph containing the cursor
    Set p = Selection.Paragraphs(1)

    ' Check if cursor is in a heading 4 (Tag)
    If p.OutlineLevel = wdOutlineLevel4 Then
        ' Get the range for this card
        Set cardRange = SelectHeadingAndContentRange(p)

        ' Apply zapping and condensing to just this card
        ZapRange cardRange
        CondenseCardsRange cardRange
    Else
        MsgBox "Please place your cursor in a Tag before running this macro.", vbExclamation
    End If
End Sub
