' Acronyms - Debate Scripts - Version 3.2.0
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

Sub AcronymHighlight()
    Dim i As Integer

    For i = 1 To Selection.Words.Count
        Dim word As String
        word = Selection.Words(i)
        If word <> "-" And word <> "," And word <> "." Then
            Selection.Words(i).Characters(1).HighlightColorIndex = Options.DefaultHighlightColorIndex
        End If
    Next
End Sub

Sub AcronymEmphasize()
    Dim i As Integer

    For i = 1 To Selection.Words.Count
        Dim word As String
        word = Selection.Words(i)
        If word <> "-" And word <> "," And word <> "." Then
            Selection.Words(i).Characters(1).Style = "Emphasis"
        End If
    Next
End Sub

Sub AcronymUnderline()
    Dim i As Integer

    For i = 1 To Selection.Words.Count
        Dim word As String
        word = Selection.Words(i)
        If word <> "-" And word <> "," And word <> "." Then
            Selection.Words(i).Characters(1).Style = "Underline"
        End If
    Next
End Sub
