Public Sub ForReference()
    Dim selRange As Range
    Dim procRange As Range
    Dim selEnd   As Long

    ' Make sure something is selected
    If Selection.Type = wdNoSelection Or Len(Selection.Range.Text) = 0 Then
        MsgBox "Please select some text before running this macro.", vbExclamation
        Exit Sub
    End If

    ' Turn off screen updating for speed
    Application.ScreenUpdating = False

    ' Grab your selection and remember its end
    Set selRange = Selection.Range
    selEnd = selRange.End

    ' Work on a copy so we don't lose the original boundaries
    Set procRange = selRange.Duplicate

    With procRange.Find
        .ClearFormatting
        .Text = ""                ' find formatting only
        .Highlight = True         ' only existing highlights
        .Format = True
        .Forward = True
        .Wrap = wdFindStop        ' do not go past procRange
    End With

    ' Loop through each found highlighted run
    Do While procRange.Find.Execute
        ' If we've passed the original selection, exit
        If procRange.Start >= selEnd Then Exit Do

            ' Recolor this run
            procRange.HighlightColorIndex = wdGray25

            ' Move past it
            procRange.Collapse wdCollapseEnd
        Loop

    ' Restore updating
    Application.ScreenUpdating = True
End Sub
