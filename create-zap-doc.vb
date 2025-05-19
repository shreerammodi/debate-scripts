Sub Zap()
    Dim styles As Variant, s As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Styles that should be preserved in the read doc
    styles = Array("Pocket", "Hat", "Block", "Tag", "Cite", "Analytic", "Undertag")
    Options.DefaultHighlightColorIndex = wdTurquoise

    ' First pass: turn on highlights for styles
    For Each s In styles
        With ActiveDocument.Content.Find
            .ClearFormatting
            .Style = s
            .Text = ""
            .Forward = True
            .Wrap = wdFindContinue
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
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Highlight = False
        .Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True

        With .Replacement
            .ClearFormatting
            .Text = "^p"
        End With
        .Execute Replace:=wdReplaceAll
    End With

    ' Third pass: Remove highlighting from styles
    For Each s In styles
        With ActiveDocument.Content.Find
            .ClearFormatting
            .Style = s
            .Forward = True
            .Wrap = wdFindContinue
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

Sub CondenseZap()
    Dim rngTemp As Range
    Dim rngStart As Range, rngEnd As Range

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Set the range to the entire document
    Set rngTemp = ActiveDocument.Content

    With rngTemp.Find
        .ClearFormatting
        .Text = Chr(13) ' Paragraph break
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False

        Do While .Execute()
            If Not .Found Then Exit Do

                Set rngStart = rngTemp.Duplicate
                Set rngEnd = rngTemp.Duplicate
                rngStart.Collapse wdCollapseStart
                rngEnd.Collapse wdCollapseEnd

                ' If both ranges are highlighted, replace paragraph break with a space
                If rngStart.HighlightColorIndex <> wdNoHighlight And rngEnd.HighlightColorIndex <> wdNoHighlight Then
                    rngStart.End = rngEnd.Start
                    rngStart.Text = " "
                End If

                ' Reset the range
                rngTemp.SetRange rngEnd.End, rngTemp.Document.Range.End
            Loop
    End With

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub CreateZappedDoc()
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

    Call Zap()
    Call CondenseZap()

    ' Get the Downloads folder path
    downloadsDirPath = GetDownloadsDir()
    savePath = downloadsDirPath & "[R] " & originalDoc.Name
    ActiveDocument.SaveAs2 Filename:=savePath, FileFormat:=wdFormatDocumentDefault

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
