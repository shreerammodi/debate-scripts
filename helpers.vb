Private Function GetDownloadsDir() As String
    Dim downloadsPath As String

    #If Mac Then
        username = Environ("USER")
        downloadsPath = "/Users/" & username & "/Downloads/"
    #Else
        Dim WshShell As Object
        Set WshShell = CreateObject("WScript.Shell")
        downloadsPath = WshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Downloads\"
    #End If

    GetDownloadsDir = downloadsPath
End Function

Private Function StyleExists(styleName As String) As Boolean
    On Error Resume Next
    StyleExists = Not ActiveDocument.Styles(styleName) Is Nothing
    On Error GoTo 0
End Function
