Private Sub DeleteStyles(styles As Variant)
    For Each s in styles
        If StyleExists(CStr(s)) Then
            With ActiveDocument.Content.Find
                .ClearFormatting
                .Style = s
                .Text = ""
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
        End If
    Next s
End Sub

Public Sub CreateSendDoc()
    Dim styles As Variant
    styles = Array("Analytic", "Analytics", "Undertag")
    Call SendDoc(styles)
End Sub

Public Sub CreateSendDocNoHeaders()
    Dim styles As Variant
    styles = Array("Analytic", "Analytics", "Undertag", "Block", "Hat", "Pocket")
    Call SendDoc(styles)
End Sub

Private Sub SendDoc(styles as Variant)
    Dim docPath As String
    Dim originalDoc As Document
    Dim savePath As String
    Dim sendDoc As Document
    Dim baseFileName As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ActiveDocument.Save
    set originalDoc = ActiveDocument

    docPath = DocDir()
    baseFileName = "[S] " & originalDoc.Name
    savePath = docPath & baseFileName

    Call CloseDocumentIfOpen(baseFileName)

    Set sendDoc = Documents.Add(ActiveDocument.FullName)

    ' Process the document to remove analytics content
    Call DeleteStyles(styles)

    ActiveDocument.SaveAs2 Filename:=savePath, FileFormat:=wdFormatDocumentDefault
    sendDoc.Close

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
