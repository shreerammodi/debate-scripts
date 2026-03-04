' Acronyms - Debate Scripts - Version 3.3.0
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

' =============================================================
' Custom Acronym Table
'
' To add an entry, add a Case line in LookupCustom below.
'
' Pattern format: comma-separated "wordIdx:charSpec" pairs
'   wordIdx  = which word (1-based)
'   charSpec = one of:
'     "N"    -> character N
'     "N-M"  -> characters N through M
'     "last" -> last character of that word
'
' Examples:
'   "nuclear weapons"            -> "nucs": "1:1-3,2:last"
'   "economy"                    -> "econ": "1:1-4"
'   "weapon of mass destruction" -> "WMD":  "1:1,3:1,4:1"
' =============================================================

Private Function LookupCustom(phrase As String) As String
    Dim normalized As String
    normalized = LCase(Trim(phrase))

    ' Strip trailing punctuation or CR that Word may append to Selection.Text
    Dim lc As String
    Do While Len(normalized) > 0
        lc = Right(normalized, 1)
        If lc = Chr(13) Or lc = Chr(10) Or lc = "." Or lc = "," Then
            normalized = Left(normalized, Len(normalized) - 1)
        Else
            Exit Do
        End If
    Loop

    Select Case normalized
        Case "nuclear weapons"                     : LookupCustom = "1:1-3,2:last"           ' nucs
        Case "nuclear warheads"                    : LookupCustom = "1:1-3,2:last"           ' nucs
        Case "economy"                             : LookupCustom = "1:1-4"                  ' econ
        Case "communications"                      : LookupCustom = "1:1-5"                  ' comms
        Case "weapon of mass destruction"          : LookupCustom = "1:1,3  :1,4:1"          ' wmd
        Case "existential"                         : LookupCustom = "1:2"                    ' x
        Case "extinction"                          : LookupCustom = "1:2"                    ' x
        Case "utilitarianism"                      : LookupCustom = "1:1-4"                  ' util
        Case "evidence"                            : LookupCustom = "1:1-2"                  ' ev
        Case "become"                              : LookupCustom = "1:1-2"                  ' be
        Case "submarines"                          : LookupCustom = "1:1-3,1:last"           ' subs
        Case "technology"                          : LookupCustom = "1:1-4"                  ' tech
        Case "intercontinental ballistic missile"  : LookupCustom = "1:1,1:6,2:1,3:1"        ' icbm
        Case "intercontinental ballistic missiles" : LookupCustom = "1:1,1:6,2:1,3:1,3:last" ' icbms
        Case "miscalculation"                      : LookupCustom = "1:1-7"                  ' miscalc
        Case "cooperation"                         : LookupCustom = "1:1-4"                  ' coop
        Case Else                                  : LookupCustom = ""
    End Select
End Function

Private Sub ApplyCustomPattern(sel As Object, pattern As String, action As String)
    Dim specs() As String
    Dim parts() As String
    Dim rangeParts() As String
    Dim i As Integer
    Dim c As Integer
    Dim wordIdx As Integer
    Dim charSpec As String
    Dim wordLen As Integer
    Dim startC As Integer
    Dim endC As Integer

    specs = Split(pattern, ",")

    For i = 0 To UBound(specs)
        parts = Split(Trim(specs(i)), ":")

        If UBound(parts) >= 1 Then
            wordIdx = CInt(Trim(parts(0)))

            If wordIdx >= 1 And wordIdx <= sel.Words.Count Then
                charSpec = Trim(parts(1))
                wordLen = Len(Trim(sel.Words(wordIdx)))

                If charSpec = "last" Then
                    startC = wordLen
                    endC = wordLen
                ElseIf InStr(charSpec, "-") > 0 Then
                    rangeParts = Split(charSpec, "-")
                    startC = CInt(Trim(rangeParts(0)))
                    endC = CInt(Trim(rangeParts(1)))
                Else
                    startC = CInt(charSpec)
                    endC = startC
                End If

                For c = startC To endC
                    If c >= 1 And c <= wordLen Then
                        Select Case action
                            Case "highlight" : sel.Words(wordIdx).Characters(c).HighlightColorIndex = Options.DefaultHighlightColorIndex
                            Case "emphasize" : sel.Words(wordIdx).Characters(c).Style = "Emphasis"
                            Case "underline" : sel.Words(wordIdx).Characters(c).Style = "Underline"
                        End Select
                    End If
                Next c
            End If
        End If
    Next i
End Sub

Sub AcronymHighlight()
    Dim pattern As String
    pattern = LookupCustom(Selection.Text)

    If pattern <> "" Then
        ApplyCustomPattern Selection, pattern, "highlight"
        Exit Sub
    End If

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
    Dim pattern As String
    pattern = LookupCustom(Selection.Text)

    If pattern <> "" Then
        ApplyCustomPattern Selection, pattern, "emphasize"
        Exit Sub
    End If

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
    Dim pattern As String
    pattern = LookupCustom(Selection.Text)

    If pattern <> "" Then
        ApplyCustomPattern Selection, pattern, "underline"
        Exit Sub
    End If

    Dim i As Integer
    For i = 1 To Selection.Words.Count
        Dim word As String
        word = Selection.Words(i)
        If word <> "-" And word <> "," And word <> "." Then
            Selection.Words(i).Characters(1).Style = "Underline"
        End If
    Next
End Sub
