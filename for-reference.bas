' Debate Scripts - Version 3.1.0
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

Public Sub ForReferenceSlow()
    Dim selRange As Range
    Dim i As Long
    Dim charRange As Range
    Dim sfr As Boolean
    sfr = ShrinkForReference()
    Dim shrinkByAmount As Integer
    shrinkByAmount = ShrinkBy()

    If Selection.Type = wdNoSelection Or Len(Selection.Range.Text) = 0 Then
        MsgBox "Please select some text before running this macro.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Set selRange = Selection.Range

    If sfr = True Then
        For i = 0 To shrinkByAmount
            selRange.Font.Shrink
        Next i
    End If

    For i = 0 To selRange.Characters.Count - 1
        Set charRange = selRange.Duplicate
        charRange.SetRange Start:=selRange.Start + i, End:=selRange.Start + i + 1

        If charRange.HighlightColorIndex <> wdNoHighlight Then
            charRange.HighlightColorIndex = wdNoHighlight
            charRange.Shading.BackgroundPatternColor = wdColorGray20
        End If
    Next i

    Application.ScreenUpdating = True
End Sub

Public Sub ForReferenceFast()
    Dim selRange As Range
    Dim procRange As Range
    Dim selEnd   As Long
    Dim i As Long
    Dim sfr As Boolean
    sfr = ShrinkForReference()
    Dim shrinkByAmount As Integer
    shrinkByAmount = ShrinkBy()

    If Selection.Type = wdNoSelection Or Len(Selection.Range.Text) = 0 Then
        MsgBox "Please select some text before running this macro.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Grab your selection and remember its end
    Set selRange = Selection.Range
    selEnd = selRange.End

    If sfr = True Then
        For i = 0 To shrinkByAmount
            selRange.Font.Shrink
        Next i
    End If

    Set procRange = selRange.Duplicate

    With procRange.Find
        .ClearFormatting
        .Text = ""
        .Highlight = True
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While procRange.Find.Execute
        If procRange.Start >= selEnd Then Exit Do
        procRange.HighlightColorIndex = wdNoHighlight
        procRange.Shading.BackgroundPatternColor = wdColorGray20
        procRange.Collapse wdCollapseEnd
    Loop

    Application.ScreenUpdating = True
End Sub
