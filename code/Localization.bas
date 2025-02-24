Attribute VB_Name = "Localization"
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
Private controlLabels As Object
Public Sub Initialize()
    InitializeLocalizedStrings
End Sub
Private Sub InitializeLocalizedStrings()
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    Dim result As Long
    Dim filePath As String
    filePath = Filesystem.GetAppDataFolder & "locale\" & Settings.GetDisplayLocale & ".txt"
    If Filesystem.FileExists(filePath) Then
        Set controlLabels = Filesystem.ParseFileContent(filePath)
    End If
#If DebugMode = 0 Then
PROC_EXIT:
    Exit Sub
PROC_ERROR:
    Helper.ErrorHandler "Localization.InitializeLocalizedStrings"
    GoTo PROC_EXIT
#End If
End Sub
Public Function GetLocalizedString(key As String, Optional defaultText As String = "Translation not available") As String
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    If controlLabels.Exists(key) Then
        GetLocalizedString = controlLabels(key)
    Else
        GetLocalizedString = defaultText
    End If
#If DebugMode = 0 Then
PROC_EXIT:
    Exit Function
PROC_ERROR:
    Helper.ErrorHandler "Localization.GetLocalizedString"
    GoTo PROC_EXIT
#End If
End Function
