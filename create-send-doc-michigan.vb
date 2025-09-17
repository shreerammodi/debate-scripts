Sub DeleteAnalyticsMichigan()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    With ActiveDocument.Content.Find
        .ClearFormatting
        .Style = "Analytic"
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Content.Find
        .ClearFormatting
        .Style = "Undertag"
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Content.Find
        .ClearFormatting
        .Style = "Analytics"
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With

End Sub

Sub CreateSendDocMichigan()
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
    Call DeleteAnalyticsMichigan

    ' Get the Downloads folder path
    downloadsDirPath = GetDownloadsDir()
    savePath = downloadsDirPath & "[S] " & originalDoc.Name
    ActiveDocument.SaveAs2 Filename:=savePath, FileFormat:=wdFormatDocumentDefault
    sendDoc.Close

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
