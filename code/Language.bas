Attribute VB_Name = "Language"
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
Private Const LANGUAGES_FILE As String = "languages.txt"
Private languages() As Language.LanguageInfo
Public Type LanguageInfo
    Index As Long
    LanguageID As Long
    Name As String
    Locale As String
End Type
Public Sub Initialize()
    LoadLanguages Filesystem.GetAppDataFolder & LANGUAGES_FILE
End Sub
Public Function GetLanguage(Locale As String) As String
    Dim result As String
    Dim i As Long
    result = "English (U.S.)"
    For i = 0 To UBound(languages)
        If languages(i).Locale = Locale Then
            result = languages(i).Name
            Exit For
        End If
    Next
    GetLanguage = result
End Function

Public Function GetLanguageID(Locale As String) As Long
    Dim result As Long
    Dim i As Long
    result = 1033
    For i = 0 To UBound(languages)
        If languages(i).Locale = Locale Then
            result = languages(i).LanguageID
            Exit For
        End If
    Next
    GetLanguageID = result
End Function
Sub LoadLanguages(filePath As String)
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    Dim dict As Object
    Dim fso As Object
    Dim fileStream As Object
    Dim fileContent As String
    Dim lines() As String
    Dim line As Variant
    Dim fields() As String
    Dim i As Long
    Dim intIndex As Long
    Dim lineCount As Long
    
    If Not Filesystem.FileExists(filePath) Then
        UI.ShowMessage Localization.GetLocalizedString("msgFileDoesNotExist", "File does not exist:") & " " & filePath
        Exit Sub
    End If

    Set fileStream = CreateObject("ADODB.Stream")
    With fileStream
        .Type = 2 ' Specify text stream type
        .Charset = "UTF-8"
        .Open
        .LoadFromFile filePath
        fileContent = .ReadText
        .Close
    End With
    
    'Debug.Print fileContent
    lines = Split(fileContent, vbCrLf)
    lineCount = UBound(lines)
    'Debug.Print lineCount
    ReDim languages(0 To lineCount)
    
    For i = 1 To lineCount
        If Trim(lines(i)) <> "" Then
            fields = Split(lines(i), ",")
            intIndex = i - 1
            languages(intIndex).Index = intIndex
            languages(intIndex).LanguageID = fields(0)
            languages(intIndex).Locale = fields(1)
            languages(intIndex).Name = fields(2)
        End If
    Next i
    
    Set fileStream = Nothing
#If DebugMode = 0 Then
PROC_EXIT:
    Exit Sub
PROC_ERROR:
    Helper.ErrorHandler "Language.LoadLanguages"
    GoTo PROC_EXIT
#End If
End Sub
