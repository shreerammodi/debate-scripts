Public Function GetDownloadsPath() As String
    Dim downloadsPath As String

    #If Mac Then
        username = Environ("USER")
        downloadsPath = "/Users/" & username & "/Downloads/"
    #Else
        Dim WshShell As Object
        Set WshShell = CreateObject("WScript.Shell")
        downloadsPath = WshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Downloads\"
    #End If

    GetDownloadsPath = downloadsPath
End Function

Public Function GetDesktopPath() As String
    Dim downloadsPath As String

    #If Mac Then
        username = Environ("USER")
        desktopPath = "/Users/" & username & "/Desktop/"
    #Else
        Dim WshShell As Object
        Set WshShell = CreateObject("WScript.Shell")
        desktopPath = WshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Desktop\"
    #End If

    GetDesktopPath = desktopPath
End Function

Public Function StyleExists(styleName As String) As Boolean
    On Error Resume Next
    StyleExists = Not ActiveDocument.Styles(styleName) Is Nothing
    On Error GoTo 0
End Function

Public Sub CloseDocumentIfOpen(docName)
    On Error Resume Next
    Documents(docName).Close SaveChanges:=wdDoNotSaveChanges
    On Error GoTo 0
End Sub

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

' Set to True if you want Send and Read docs to appear in your
' Downloads directory
Public Function DocsInDownloads As Boolean
    DocsInDownloads = True
End Function

' Set to True if you want Send and Read docs to appear in your
' Desktop directory
Public Function DocsInDesktop As Boolean
    DocsInDesktop = False
End Function

Public Function DocDir As String
    Dim inDownloads As Boolean
    inDownloads = DocsInDownloads()
    Dim inDesktop As Boolean
    inDesktop = DocsInDesktop()

    Dim downloadsPath As String
    downloadsPath = GetDownloadsPath()
    Dim desktopPath As String
    desktopPath = GetDesktopPath()

    If inDownloads Then
        DocDir = downloadsPath
    End If

    If inDesktop Then
        DocDir = desktopPath
    End If

    ' Alternatively, set the above two functions as False and replace
    ' <<Custom Path Here>> with the path of the folderyou want your Send and
    ' Read docs to appear in.
    '
    ' Ensure that there is a trailing /
    If Not inDownloads And Not inDesktop Then
        DocDir =  "<<Custom Path Here>>"
    End If
End Function
