Attribute VB_Name = "Filesystem"
' Katip - Add-in to support spellchecking in Microsoft Word desktop application using Hunspell library and dictionaries
'
' Copyright (C) 2012-2025 Nazar Mammedov
' https://github.com/berkesas/katip4
'
' This software uses Hunspell Copyright (C) 2002-2022 Németh László
' https://github.com/hunspell/hunspell
'
' Hunspell dictionaries are copyright by respective developers
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

Option Explicit
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
#Else
    Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
#End If

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Const DICTIONARY_DIR As String = "dictionaries"
Function IsWindowsXP() As Boolean
    Dim osInfo As OSVERSIONINFO
    osInfo.dwOSVersionInfoSize = Len(osInfo)
    
    If GetVersionEx(osInfo) Then
        If osInfo.dwMajorVersion = 5 And osInfo.dwMinorVersion = 1 Then
            IsWindowsXP = True
        Else
            IsWindowsXP = False
        End If
    Else
        IsWindowsXP = False
    End If
End Function
Function GetAppDataFolder() As String
    If IsWindowsXP() Then
        GetAppDataFolder = Environ("ALLUSERSPROFILE") & "\Application Data\Katip" & "\"
    Else
        GetAppDataFolder = Environ("ProgramData") & "\Katip" & "\"
    End If
End Function
Function GetDictionaryFolder() As String
    GetDictionaryFolder = GetAppDataFolder & DICTIONARY_DIR
End Function
Function FileExists(filePath As String) As Boolean
    FileExists = (Dir(filePath) <> "")
End Function

Function ReadUTF8File(filePath As String) As String
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    Dim adoStream As Object
    Dim fileContent As String
    
    If Not FileExists(filePath) Then
        MsgBox "File not found: " & filePath, vbExclamation
        Exit Function
    End If
    
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Charset = "UTF-8"
    
    adoStream.Open
    adoStream.LoadFromFile filePath
    fileContent = adoStream.ReadText
    adoStream.Close
    Set adoStream = Nothing
    ReadUTF8File = fileContent
#If DebugMode = 0 Then
PROC_EXIT:
    Exit Function
PROC_ERROR:
    Helper.ErrorHandler "Filesystem.ReadUTF8File: " & filePath
    GoTo PROC_EXIT
#End If
End Function

Function ParseFileContent(ByVal filePath As String) As Object
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    Dim elements() As String
    Dim fileContent As String
    Dim result As Long
    fileContent = ReadUTF8File(filePath)
    Dim returnVal As Object
    
    Set returnVal = New Scripting.Dictionary
    
    If Len(fileContent) > 0 Then
        Dim lines() As String
        Dim i As Long
        lines = Split(fileContent, vbCrLf)
        For i = 0 To UBound(lines)
            elements = Split(lines(i), "=")
            If UBound(elements) = 1 Then
                returnVal(elements(0)) = elements(1)
            End If
        Next
        result = UBound(lines)
    End If
    
    Set ParseFileContent = returnVal
#If DebugMode = 0 Then
PROC_EXIT:
    Exit Function
PROC_ERROR:
    Helper.ErrorHandler "Filesystem.ParseFileContent"
    GoTo PROC_EXIT
#End If
End Function
Public Function GetFilesCollection(folderPath As String) As Collection
    Dim fileName As String
    Dim fileList As Collection
    Dim fileItem As Variant

    Set fileList = New Collection
    
    fileName = Dir(folderPath)
    
    Do While fileName <> ""
        fileList.Add fileName
        fileName = Dir
    Loop
    
    Set GetFilesCollection = fileList
End Function
Function RemoveFileExtension(fileName As String) As String
    Dim pos As Integer
    pos = InStrRev(fileName, ".")
    If pos = 0 Then
        RemoveFileExtension = fileName
    Else
        RemoveFileExtension = Left(fileName, pos - 1)
    End If
End Function
Sub AppendToFile(filePath As String, lineToAdd As String)
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    Dim fileStream As Object
    Dim tempStream As Object
    Dim fileContent As String
    Dim lines() As String
    Dim lineCount As Long
    Dim newContent As String
    Dim i As Long
    Dim boolAlreadyExists As Boolean
    
    If Not FileExists(filePath) Then
       UI.ShowMessage Localization.GetLocalizedString("msgFileDoesNotExist", "File does not exist:") & " " & filePath
       Exit Sub
    End If
    
    boolAlreadyExists = False
    
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = 2 ' Specify text stream type
    fileStream.Charset = "UTF-8"
    fileStream.Open
    fileStream.LoadFromFile filePath
    fileContent = fileStream.ReadText
    
    lines = Split(fileContent, vbCrLf)
    lineCount = UBound(lines) + 1
    
    For i = 1 To UBound(lines)
        If lines(i) = lineToAdd Then
            boolAlreadyExists = True
            Exit For
        End If
        newContent = newContent & lines(i) & vbCrLf
    Next i
    
    If Not boolAlreadyExists Then
        Set tempStream = CreateObject("ADODB.Stream")
        tempStream.Type = 2 ' Specify text stream type
        tempStream.Charset = "UTF-8"
        tempStream.Open
        tempStream.WriteText lineCount & vbCrLf & newContent & lineToAdd
        tempStream.SaveToFile filePath, 2 ' Overwrite the file
        tempStream.Close
    End If
    
    fileStream.Close
    Set fileStream = Nothing
    Set tempStream = Nothing
#If DebugMode = 0 Then
PROC_EXIT:
    Exit Sub
PROC_ERROR:
    Helper.ErrorHandler "Filesystem.AppendToFile"
    GoTo PROC_EXIT
#End If
End Sub
Sub CreateFile(filePath As String, initialContent As String)
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    Dim fileStream As Object
    
    If Not FileExists(filePath) Then
        Set fileStream = CreateObject("ADODB.Stream")
        fileStream.Type = 2 ' Specify text stream type
        fileStream.Charset = "UTF-8"
        fileStream.Open
        fileStream.WriteText initialContent
        
        fileStream.SaveToFile filePath, 2 ' Create the file (overwrite mode)
        fileStream.Close
        Set fileStream = Nothing
    End If
#If DebugMode = 0 Then
PROC_EXIT:
    Exit Sub
PROC_ERROR:
    Helper.ErrorHandler "Filesystem.CreateFile"
    GoTo PROC_EXIT
#End If
End Sub
Sub OpenTextFile(filePath As String)
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    If Dir(filePath) <> "" Then
        shell "cmd.exe /c start " & filePath, vbHide
    Else
        UI.ShowMessage Localization.GetLocalizedString("msgFileDoesNotExist", "File does not exist:") & " " & filePath
    End If
#If DebugMode = 0 Then
PROC_EXIT:
    Exit Sub
PROC_ERROR:
    Helper.ErrorHandler "Filesystem.OpenTextFile:" & filePath
    GoTo PROC_EXIT
#End If
End Sub
Sub OpenFolderInExplorer(filePath As String)
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    If Dir(filePath) <> "" Then
        shell "explorer.exe """ & filePath & """", vbNormalFocus
    Else
        UI.ShowMessage Localization.GetLocalizedString("msgFolderDoesNotExist", "Folder does not exist:") & " " & filePath
    End If
#If DebugMode = 0 Then
PROC_EXIT:
    Exit Sub
PROC_ERROR:
    Helper.ErrorHandler "Filesystem.OpenFolderInExplorer: " & filePath
    GoTo PROC_EXIT
#End If
End Sub
Function LogError(logText As String)
    Dim filePath As String
    filePath = Filesystem.GetAppDataFolder & "error.log"
    If FileExists(filePath) Then
        Dim stream
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2 ' adTypeText
        stream.Charset = "utf-8"
        stream.Open
        stream.LoadFromFile filePath
        stream.Position = stream.Size
        stream.WriteText logText, 1 ' adWriteLine
        stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
        stream.Close
        Set stream = Nothing
    End If
End Function

