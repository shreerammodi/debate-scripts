' Word Count - Debate Scripts - Version 4.0.0
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

' Zapper - Debate Scripts - Version 3.2.0
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

Public Sub WordCount()
    Dim originalDoc As Document
    Dim newDoc As Document
    Dim styles As Variant
    Dim wordCount As Long
    Dim wpm As Long
    Dim totalSeconds As Long
    Dim mins As Long
    Dim secs As Long

    ' ENTER YOUR WPM BELOW
    wpm = 250

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ActiveDocument.Save
    Set originalDoc = ActiveDocument

    Set newDoc = Documents.Add(ActiveDocument.FullName)

    Call Zap()

    styles = Array("Undertag", "Block", "Hat", "Pocket")

    Call DeleteStyles(styles)

    wordCount = ActiveDocument.Range.ComputeStatistics(wdStatisticWords)

    totalSeconds = CLng((wordCount / wpm) * 60)
    mins = totalSeconds \ 60
    secs = totalSeconds Mod 60

    MsgBox wordCount & " words." & vbNewLine & vbNewLine & mins & "m " & secs & "s at " & wpm & " wpm."

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ActiveDocument.Close(wdDoNotSaveChanges)
End Sub
