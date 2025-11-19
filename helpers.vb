Public Function GetDownloadsDir() As String
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

Public Function StyleExists(styleName As String) As Boolean
    On Error Resume Next
    StyleExists = Not ActiveDocument.Styles(styleName) Is Nothing
    On Error GoTo 0
End Function

' ==================================================
' SETTINGS - Edit these values to configure behavior
' ==================================================

' Set to True to shrink text when running ForReference
' Set to False to keep text at original size

Public Function ShrinkForReference() As Boolean
    ShrinkForReference = True
End Function

' Set to the font size you want to shrink text in ForReference by
Public Function ShrinkBy() As Integer
    ShrinkBy = 1
End Function
