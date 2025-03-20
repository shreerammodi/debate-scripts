Sub Zap(Optional targetDoc As Document = Nothing)
    ' If targetDoc is not passed, use ActiveDocument
    If targetDoc Is Nothing Then
        Set targetDoc = ActiveDocument
    End If

    Application.ScreenUpdating = False
    Options.DefaultHighlightColorIndex = wdTurquoise
    With targetDoc.Content.Find
        .ClearFormatting
        .Style = "Tag"
        With .Replacement
            .Text = "^&"
            .ClearFormatting
            .Highlight = True
        End With
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = False
    Options.DefaultHighlightColorIndex = wdTurquoise
    With targetDoc.Content.Find
        .ClearFormatting
        .Style = "Cite"
        With .Replacement
            .Text = "^&"
            .ClearFormatting
            .Highlight = True
        End With
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = False
    Options.DefaultHighlightColorIndex = wdTurquoise
    With targetDoc.Content.Find
        .ClearFormatting
        .Style = "Pocket"
        With .Replacement
            .Text = "^&"
            .ClearFormatting
            .Highlight = True
        End With
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = False
    Options.DefaultHighlightColorIndex = wdTurquoise
    With targetDoc.Content.Find
        .ClearFormatting
        .Style = "Hat"
        With .Replacement
            .Text = "^&"
            .ClearFormatting
            .Highlight = True
        End With
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = False
    Options.DefaultHighlightColorIndex = wdTurquoise
    With targetDoc.Content.Find
        .ClearFormatting
        .Style = "Block"
        With .Replacement
            .Text = "^&"
            .ClearFormatting
            .Highlight = True
        End With
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = False
    Options.DefaultHighlightColorIndex = wdTurquoise
    With targetDoc.Content.Find
        .ClearFormatting
        .Style = "Analytic"
        With .Replacement
            .Text = "^&"
            .ClearFormatting
            .Highlight = True
        End With
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = False
    Options.DefaultHighlightColorIndex = wdTurquoise
    With targetDoc.Content.Find
        .ClearFormatting
        .Style = "Undertag"
        With .Replacement
            .Text = "^&"
            .ClearFormatting
            .Highlight = True
        End With
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    Selection.Find.ClearFormatting
    Selection.Find.Highlight = False
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Application.ScreenUpdating = False
    Options.DefaultHighlightColorIndex = wdTurquoise
    With targetDoc.Content.Find
        .ClearFormatting    ' If targetDoc is not passed, use ActiveDocument
        If targetDoc Is Nothing Then
            Set targetDoc = ActiveDocument
        End If
        .Style = "Tag"
        With .Replacement
            .Text = "^&"
            .ClearFormatting
            .Highlight = False
        End With
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = False
    Options.DefaultHighlightColorIndex = wdTurquoise
    With targetDoc.Content.Find
        .ClearFormatting
        .Style = "Cite"
        With .Replacement
            .Text = "^&"
            .ClearFormatting
            .Highlight = False
        End With
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = False
    Options.DefaultHighlightColorIndex = wdTurquoise
    With targetDoc.Content.Find
        .ClearFormatting
        .Style = "Block"
        With .Replacement
            .Text = "^&"
            .ClearFormatting
            .Highlight = False
        End With
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = False
    Options.DefaultHighlightColorIndex = wdTurquoise
    With targetDoc.Content.Find
        .ClearFormatting
        .Style = "Pocket"
        With .Replacement
            .Text = "^&"
            .ClearFormatting
            .Highlight = False
        End With
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = False
    Options.DefaultHighlightColorIndex = wdTurquoise
    With targetDoc.Content.Find
        .ClearFormatting
        .Style = "Hat"
        With .Replacement
            .Text = "^&"
            .ClearFormatting
            .Highlight = False
        End With
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = False
    Options.DefaultHighlightColorIndex = wdTurquoise
    With targetDoc.Content.Find
        .ClearFormatting
        .Style = "Analytic"
        With .Replacement
            .Text = "^&"
            .ClearFormatting
            .Highlight = False
        End With
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = False
    Options.DefaultHighlightColorIndex = wdTurquoise
    With targetDoc.Content.Find
        .ClearFormatting
        .Style = "Undertag"
        With .Replacement
            .Text = "^&"
            .ClearFormatting
            .Highlight = False
        End With
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " ^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWholeWord = False
        .MatchWildcards = False
    End With
    Selection.Find.Execute
    While Selection.Find.Found
        Selection.HomeKey Unit:=wdStory
        Selection.Find.Execute Replace:=wdReplaceAll
        Selection.Find.Execute
    Wend

    Selection.Find.Text = " ^l"
    Selection.Find.Replacement.Text = "^l"
    Selection.Find.Execute
    While Selection.Find.Found
        Selection.HomeKey Unit:=wdStory
        Selection.Find.Execute Replace:=wdReplaceAll
        Selection.Find.Execute
    Wend

End Sub

Sub CondenseZap(Optional targetDoc As Document = Nothing)
    ' If targetDoc is not passed, use ActiveDocument
    If targetDoc Is Nothing Then
        Set targetDoc = ActiveDocument
    End If

    Dim rngTemp As Range
    Dim rngStart As Range, rngEnd As Range

    Application.ScreenUpdating = False

    ' Set the range to the entire document
    Set rngTemp = targetDoc.Content

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
    End Sub

Sub CreateZappedDoc()
    Dim originalDoc As Document
    Dim newDoc As Document
    Dim savePath As String
    Dim originalFolderPath As String
    Dim originalFilePath As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Save the original document
    ActiveDocument.Save

    ' Assign the original document to a variable
    Set originalDoc = ActiveDocument

    ' Get the folder path and file path
    originalFolderPath = Left(originalDoc.FullName, InStrRev(originalDoc.FullName, Application.PathSeparator))
    originalFilePath = originalDoc.FullName

    ' Define save path for the modified document
    savePath = originalFolderPath & "[R] " & originalDoc.Name

    ' Save a copy of the document
    originalDoc.SaveAs2 Filename:=savePath, FileFormat:=wdFormatXMLDocument
    Set newDoc = Documents.Open(savePath)

    ' Call Zap and CondenseZap on the new document
    Call Zap(newDoc)
    Call CondenseZap(newDoc)

    ' Save and close the modified document
    newDoc.Save
    newDoc.Close

    ' Reopen the original document
    Documents.Open originalFilePath

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Read version created and saved as " & savePath
End Sub
