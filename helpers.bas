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

Public Function GetDownloadsPath() As String
    Dim downloadsPath As String
    Dim username As String

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
    Dim desktopPath As String
    Dim username As String

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
    ' Ensure that there is a trailing slash
    If Not inDownloads And Not inDesktop Then
        DocDir =  "<<Custom Path Here>>"
    End If
End Function
