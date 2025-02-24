VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GlobalEventManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Public Event DisplayLanguageChanged(ByVal newLanguage As String)
Private currentLanguage As String
Public Sub RaiseDisplayLanguageChanged(ByVal newLanguage As String)
    Localization.Initialize
    UI.UpdateControlLabels
    RaiseEvent DisplayLanguageChanged(newLanguage)
End Sub
Public Property Get Language() As String
    Language = currentLanguage
End Property
Public Property Let Language(ByVal newLanguage As String)
    If currentLanguage <> newLanguage Then
        currentLanguage = newLanguage
        RaiseEvent DisplayLanguageChanged(newLanguage)
    End If
End Property

