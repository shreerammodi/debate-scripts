Private Sub DeleteAnalytics()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    If StyleExists("Analytic")
        With ActiveDocument.Content.Find
            .ClearFormatting
            .Style = "Analytic"
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End If

    If StyleExists("Undertag")
        With ActiveDocument.Content.Find
            .ClearFormatting
            .Style = "Undertag"
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End If

    If StyleExists("Analytics")
        With ActiveDocument.Content.Find
            .ClearFormatting
            .Style = "Analytics"
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End If

End Sub

Public Sub CreateSendDoc()
    Dim downloadsDirPath As String
    Dim originalDoc As Document
    Dim savePath As String
    Dim sendDoc As Document

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ActiveDocument.Save
    set originalDoc = ActiveDocument

    Set sendDoc = Documents.Add(ActiveDocument.FullName)

    ' Process the document to remove analytics content
    Call DeleteAnalytics

    ' Get the Downloads folder path
    downloadsDirPath = GetDownloadsDir()
    savePath = downloadsDirPath & "[S] " & originalDoc.Name
    ActiveDocument.SaveAs2 Filename:=savePath, FileFormat:=wdFormatDocumentDefault
    sendDoc.Close

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
