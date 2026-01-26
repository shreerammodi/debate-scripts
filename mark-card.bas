' Mark Card - Debate Scripts - Version 3.1.1
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

Public Sub MarkCard()
    Dim p As Paragraph
    Dim rngCard As Range
    Dim markerText As String

    markerText = vbCrLf & "<<MARKED>>" & vbCrLf

    ' Insert marker text before cursor
    Selection.InsertAfter markerText

    ' Remove previous paragraph from selection
    Selection.SetRange Start:=Selection.Start + 1, End:=Selection.End

    ' Make the marker normal text
    With Selection
        .ClearFormatting
        .Style = ActiveDocument.styles("Normal")
        .Font.Underline = wdUnderlineNone
        .Range.HighlightColorIndex = wdNoHighlight
    End With

    Selection.Collapse(wdCollapseEnd)

    ' Get paragraph immediately after the marker
    Set p = Selection.Paragraphs(1)

    ' Get rest of the card
    Set rngCard = Paperless.SelectCardTextRange(p)

    ' Exclude anything before marker
    rngCard.Start = Selection.Start

    rngCard.Font.Color = wdColorRed
End Sub

Public Sub CompileMarkedCards()
    Dim p As Paragraph
    Dim cardRng as Range
    Dim hdrRng As Range
    Dim insertRng As Range
    Dim searchRng As Range
    Dim foundRng As Range

    Set markedCards = New Collection

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Create range to paste marked cards
    Set searchRng = ActiveDocument.Content

    With searchRng.Find
        .ClearFormatting
        .Text = "<<MARKED>>"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
    End With

    ' Loop over every marker
    Do While searchRng.Find.Execute
        Set foundRng = searchRng.Duplicate
        foundRng.Collapse wdCollapseStart

        Set p = foundRng.Paragraphs(1)

        ' Add marked card to marked cards collection
        Set cardRng = SelectHeadingAndContentRange(p)
        markedCards.Add cardRng.Duplicate

        ' Move search start past current card
        searchRng.Start = foundRng.End
    Loop

    ' Get end of document
    Set hdrRng = ActiveDocument.Content
    hdrRng.Collapse wdCollapseEnd

    ' Insert new paragraph
    hdrRng.InsertParagraphAfter
    hdrRng.Collapse wdCollapseEnd

    ' Add "Marked Cards" pocket
    hdrRng.Text = "Marked Cards"
    hdrRng.Style = ActiveDocument.Styles("Pocket")
    hdrRng.InsertParagraphAfter

    Set insertRng = ActiveDocument.Range(hdrRng.End, hdrRng.End)
    insertRng.Collapse wdCollapseEnd

    ' Paste all marked cards at bottom
    For Each Item In markedCards
        With insertRng
            .FormattedText = Item.FormattedText
            .Collapse wdCollapseEnd
            .InsertParagraphAfter
            .Collapse wdCollapseEnd
        End With
    Next item

    ' Move cursor to beginning of Marked Cards section
    hdrRng.Select

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
