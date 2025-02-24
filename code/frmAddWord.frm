VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddWord 
   Caption         =   "Add word"
   ClientHeight    =   2760
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   5604
   OleObjectBlob   =   "frmAddWord.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Private WithEvents EventManager As GlobalEventManager
Attribute EventManager.VB_VarHelpID = -1
Private Sub btnAdd_Click()
    Spelling.AddToDictionary
End Sub
Private Sub UserForm_Initialize()
    Set EventManager = Settings.GlobalEvent
    'If Not EventManager Is Nothing Then
    'End If
End Sub
Private Sub EventManager_DisplayLanguageChanged(ByVal newLanguage As String)
    LoadLabels
End Sub
Sub LoadLabels()
    Me.lblWord.Caption = Localization.GetLocalizedString("lblWord", "Word")
    Me.lblAffix.Caption = Localization.GetLocalizedString("lblAffix", "Affixes")
    Me.lblWordNote.Caption = Localization.GetLocalizedString("lblWordNote", "Input word as plain text")
    Me.lblAffixNote.Caption = Localization.GetLocalizedString("lblAffixNote", "Comma separated affix flags from .aff file valid for this word")
    Me.btnAdd.Caption = Localization.GetLocalizedString("btnAdd", "Add")
    Me.btnCancel.Caption = Localization.GetLocalizedString("btnCancel", "Cancel")
    Me.lblDictionaryFile.Caption = Localization.GetLocalizedString("lblOpenDictionary", "Open dictionary")
End Sub
Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub lblDictionaryFile_Click()
    Settings.OpenDictionary
End Sub
