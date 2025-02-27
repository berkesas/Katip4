Attribute VB_Name = "Main"
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
Public isLoaded As Boolean
Dim eventHandler As AppEventClass
Public CompatibilityVersion As Long
Public Version As Long
Public Const MACRO_VERSION As String = "v4.0.1.x"
Public Sub Initialize()
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    If Not isLoaded Then
        If Not ActiveDocument Is Nothing Then
            Version = CLng(Application.Version)
            If Version >= 14 Then
                CompatibilityVersion = ActiveDocument.CompatibilityMode
            Else
                CompatibilityVersion = 12
            End If
            frmHelp.lblVersion.Caption = Main.MACRO_VERSION & ":" & System.GetWindowsBitVersion & ":" & System.GetWordBitVersion
            Settings.Initialize
            Language.Initialize
            Localization.Initialize
            Spelling.Initialize
            isLoaded = True
        End If
    End If
#If DebugMode = 0 Then
PROC_EXIT:
    Exit Sub
PROC_ERROR:
    Helper.ErrorHandler "Main.Initialize"
    GoTo PROC_EXIT
#End If
End Sub
Sub InitializeAppEventHandler()
    Set eventHandler = New AppEventClass
    Set eventHandler.App = word.Application
End Sub

