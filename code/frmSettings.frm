VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Settings"
   ClientHeight    =   4752
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   5556
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
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

Option Explicit
Private isLoaded As Boolean
Private WithEvents EventManager As GlobalEventManager
Attribute EventManager.VB_VarHelpID = -1

Private Sub lblOpenDictionaries_Click()
    Settings.OpenDictionariesFolder
End Sub
Private Sub UserForm_Initialize()
    Set EventManager = Settings.GlobalEvent
    'If Not EventManager Is Nothing Then
    'End If
End Sub
Private Sub EventManager_DisplayLanguageChanged(ByVal newLanguage As String)
    LoadLabels
End Sub
Public Sub LoadForm()
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    Main.Initialize
    If Not isLoaded Then
        LoadValues
        isLoaded = True
    End If
    LoadLabels
    Me.Show
PROC_EXIT:
    Exit Sub
PROC_ERROR:
    Helper.ErrorHandler "frmSettings.LoadForm"
    GoTo PROC_EXIT
End Sub
Sub LoadLabels()
        Me.Caption = Localization.GetLocalizedString("rbtnSettingsCaption", "Settings")
        Me.btnCancel.Caption = Localization.GetLocalizedString("btnCancel", "Cancel")
        Me.btnSave.Caption = Localization.GetLocalizedString("btnSave", "Save")
        Me.btnChoose.Caption = Localization.GetLocalizedString("btnChoose", "Choose")
        Me.lblDisplayLanguage.Caption = Localization.GetLocalizedString("frmSettingsDisplayLanguage", "Display Language")
        Me.lblSpellingLanguage.Caption = Localization.GetLocalizedString("frmSettingsSpellingLanguage", "Spelling Language")
        Me.tabSettings.Pages(0).Caption = Localization.GetLocalizedString("frmSettingsDisplay", "Display")
        Me.tabSettings.Pages(1).Caption = Localization.GetLocalizedString("frmSettingsSpelling", "Spelling")
        Me.lblMisspelling.Caption = Localization.GetLocalizedString("frmSettingsMisspellingColor", "Misspelling Color")
        Me.lblSplitCharacters.Caption = Localization.GetLocalizedString("lblSplitCharacters", "Split Characters")
        Me.btnReload.Caption = Localization.GetLocalizedString("btnReload", "Reload")
        Me.lblOpenDictionaries.Caption = Localization.GetLocalizedString("lblOpenDictionariesFolder", "Open dictionaries folder")
        Me.chkAutoCheck.Caption = Localization.GetLocalizedString("chkAutocheck", "Check automatically on window open")
        Me.chkAutoClear.Caption = Localization.GetLocalizedString("chkAutoclear", "Clear automatically on window close")
        Me.chkWindowless.Caption = Localization.GetLocalizedString("chkWindowless", "Use windowless mode")
        Me.txtColor.value = Localization.GetLocalizedString("txtSampleText", "Sample text")
End Sub
Private Sub btnCancel_Click()
    Me.Hide
End Sub
Private Sub LoadValues()
'On Error GoTo PROC_ERROR
    Dim displayLanguages() As Language.LanguageInfo
    Dim spellingLanguages() As Language.LanguageInfo

    Dim i As Long
    displayLanguages = Settings.GetDisplayLanguages
    For i = 0 To UBound(displayLanguages)
        cbxDisplayLanguages.AddItem displayLanguages(i).Name
    Next
    
    cbxDisplayLanguages.ListIndex = Settings.GetDisplayLocaleIndex
    
    spellingLanguages = Settings.GetSpellingLanguages
    For i = 0 To UBound(spellingLanguages)
        cbxSpellingLanguages.AddItem spellingLanguages(i).Name
    Next
    
    cbxSpellingLanguages.ListIndex = Settings.GetSpellingLocaleIndex
    
    txtSplitCharacters.value = Settings.GetSplitCharacters
    chkAutoCheck.value = Settings.GetAutoCheck
    chkAutoClear.value = Settings.GetAutoClear
    txtColor.ForeColor = Settings.GetErrorColorIndex
    chkWindowless.value = Settings.GetUseWindowless
'PROC_EXIT:
'    Exit Sub
'PROC_ERROR:
'    Helper.ErrorHandler "frmSettings.LoadValues"
'    GoTo PROC_EXIT
End Sub

Private Sub btnChoose_Click()
    Settings.OpenColorPicker
End Sub

Private Sub btnReload_Click()
    Spelling.ReloadHunspell
    UI.ShowMessage Localization.GetLocalizedString("msgDictionaryReloaded", "Dictionaries are reloaded.")
End Sub

Private Sub btnSave_Click()
    Settings.SaveSettings
    Me.Hide
End Sub

Private Sub lblDictionaryFile_Click()
    Settings.OpenDictionariesFolder
End Sub
