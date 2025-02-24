VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSpelling 
   Caption         =   "Spelling"
   ClientHeight    =   4428
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   6684
   OleObjectBlob   =   "frmSpelling.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSpelling"
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
Private Sub UserForm_Initialize()
    Set EventManager = Settings.GlobalEvent
    'If Not EventManager Is Nothing Then
    'End If
End Sub
Private Sub EventManager_DisplayLanguageChanged(ByVal newLanguage As String)
    LoadLabels
End Sub
Sub LoadLabels()
    frmSpelling.Caption = Localization.GetLocalizedString("frmSpellingCaption") & ": " & Settings.GetLanguageName
    frmSpelling.lblNotInDictionary.Caption = Localization.GetLocalizedString("lblNotInDictionary", "Not in Dictionary:")
    frmSpelling.lblSuggestions.Caption = Localization.GetLocalizedString("lblSuggestions", "Suggestions:")
    frmSpelling.btnIgnoreOnce.Caption = Localization.GetLocalizedString("btnIgnoreOnce", "Ignore Once")
    frmSpelling.btnIgnoreAll.Caption = Localization.GetLocalizedString("btnIgnoreAll", "Ignore All")
    frmSpelling.btnAddToDictionary.Caption = Localization.GetLocalizedString("btnAddToDictionary", "Add to Dictionary")
    frmSpelling.btnChange.Caption = Localization.GetLocalizedString("btnChange", "Change")
    frmSpelling.btnChangeAll.Caption = Localization.GetLocalizedString("btnChangeAll", "Change All")
    frmSpelling.btnClear.Caption = Localization.GetLocalizedString("btnClear", "Clear")
    frmSpelling.btnCheck.Caption = Localization.GetLocalizedString("btnCheck", "Check")
    frmSpelling.chkAutoCheck.Caption = Localization.GetLocalizedString("chkAutocheckShort", "Automatic check")
    frmSpelling.chkAutoClear.Caption = Localization.GetLocalizedString("chkAutoclearShort", "Automatic clear")
    frmSpelling.btnCancel.Caption = Localization.GetLocalizedString("btnCancel", "Cancel")
End Sub
Private Sub btnAddToDictionary_Click()
    Spelling.ShowAddWord
End Sub

Private Sub btnCancel_Click()
    Spelling.HideSpelling
End Sub

Private Sub btnChange_Click()
    Spelling.Change
End Sub

Private Sub btnChangeAll_Click()
    Spelling.ChangeAll
End Sub

Private Sub btnCheck_Click()
    Spelling.CheckDocument
End Sub

Private Sub btnClear_Click()
    Spelling.ResetCheck
End Sub

Private Sub btnIgnoreAll_Click()
    Spelling.IgnoreAll
End Sub

Private Sub btnIgnoreOnce_Click()
    Spelling.IgnoreOnce
End Sub
Private Sub btnSpin_SpinDown()
    Spelling.PreviousError
End Sub

Private Sub btnSpin_SpinUp()
    Spelling.NextError
End Sub

Private Sub lbxSuggestions_Change()
    Spelling.SuggestionChanged
End Sub

Public Sub DisableControls()
    Me.btnIgnoreOnce.Enabled = False
    Me.btnIgnoreAll.Enabled = False
    Me.btnAddToDictionary.Enabled = False
    Me.btnChange.Enabled = False
    Me.btnChangeAll.Enabled = False
    Me.btnClear.Enabled = False
    Me.btnSpin.Enabled = False
    Me.lbxSuggestions.Clear
End Sub
Public Sub EnableControls()
    Me.btnIgnoreOnce.Enabled = True
    Me.btnIgnoreAll.Enabled = True
    Me.btnAddToDictionary.Enabled = True
    Me.btnChange.Enabled = True
    Me.btnChangeAll.Enabled = True
    Me.btnClear.Enabled = True
    Me.btnSpin.Enabled = True
End Sub

