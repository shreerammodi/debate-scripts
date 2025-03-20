Sub Zap()
    Application.ScreenUpdating = False
    Options.DefaultHighlightColorIndex = wdTurquoise
    With ActiveDocument.Content.Find
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
    With ActiveDocument.Content.Find
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
    With ActiveDocument.Content.Find
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
    With ActiveDocument.Content.Find
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
    With ActiveDocument.Content.Find
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
    With ActiveDocument.Content.Find
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
    With ActiveDocument.Content.Find
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
    With ActiveDocument.Content.Find
        .ClearFormatting    ' If ActiveDocument is not passed, use ActiveDocument
        If ActiveDocument Is Nothing Then
            Set ActiveDocument = ActiveDocument
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
    With ActiveDocument.Content.Find
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
    With ActiveDocument.Content.Find
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
    With ActiveDocument.Content.Find
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
    With ActiveDocument.Content.Find
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
    With ActiveDocument.Content.Find
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
    With ActiveDocument.Content.Find
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

Sub CondenseZap()
    Dim rngTemp As Range
    Dim rngStart As Range, rngEnd As Range

    Application.ScreenUpdating = False

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
    End Sub

Sub CreateZappedDoc()
    Dim newDoc As Document
    Dim originalDoc As Document
    Dim originalFilePath As String
    Dim originalFolderPath As String
    Dim savePath As String

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
End Sub
