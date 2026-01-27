' Send Doc - Debate Scripts - Version 3.2.0
' Copyright (C) 2025 Shreeram Modi
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program. If not, see <https://www.gnu.org/licenses/>.

Private Sub DeleteStyles(styles As Variant)
    Dim s As Variant
    Dim targetStyle As Style
    Dim shouldDelete As Boolean

    For Each s In styles
        If StyleExists(CStr(s)) Then
            Set targetStyle = ActiveDocument.styles(CStr(s))

            shouldDelete = True

            ' Don't delete any tags
            If InStr(1, CStr(targetStyle), "Tag,", vbTextCompare) > 0 Then
                shouldDelete = False
            End If

            If shouldDelete Then
                With ActiveDocument.Content.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Style = targetStyle
                    .Text = ""
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                End With
            End If
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

    docPath = SendDocDir()
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
