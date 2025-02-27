Attribute VB_Name = "Settings"
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
Private languages(2) As String
Private locales(2) As String
Private displayLanguageInfos() As Language.LanguageInfo
Private spellingLanguageInfos() As Language.LanguageInfo
Private displayLocale As String
Private spellingLocale As String
Private strSplitCharacters As String
Private intLanguageID As Long
Private strLanguageName As String
Private autoCheck As Boolean
Private autoClear As Boolean
Private errorColorIndex As Long
Private underlineStyle As Long
Private underlineColor As Long
Private showErrors As Boolean
Private useWindowless As Boolean
Public GlobalEvent As GlobalEventManager
Public Sub Initialize()
    LoadDisplayLanguages
    LoadSpellingLanguages
    LoadSettings
    Set GlobalEvent = New GlobalEventManager
End Sub
Sub DisplayLanguageEvent(newLanguage As String)
    If Not GlobalEvent Is Nothing Then
        GlobalEvent.RaiseDisplayLanguageChanged newLanguage
    End If
End Sub
Public Sub LoadSettings()
    SetDisplayLocale AppSettings.ReadSetting("General", "displayLocale", "en-US")
    SetSpellingLocale AppSettings.ReadSetting("General", "spellingLocale", "en-US")
    SetSplitCharacters AppSettings.ReadSetting("General", "splitCharacters", """',.!?:;{}()\[\]/\\=+±\^\$\*<>|¦#@%&~…©›·`?×«»—°¤¬¨‘’™")
    intLanguageID = Language.GetLanguageID(GetSpellingLocale)
    strLanguageName = spellingLanguageInfos(GetLocaleIndex(GetSpellingLocale, spellingLanguageInfos)).Name
    SetAutoCheck CBool(AppSettings.ReadSetting("General", "autoCheck", False))
    SetAutoClear CBool(AppSettings.ReadSetting("General", "autoClear", False))
    SetShowErrors CBool(AppSettings.ReadSetting("General", "showErrors", False))
    SetUseWindowless CBool(AppSettings.ReadSetting("General", "useWindowless", False))
    SetErrorColorIndex AppSettings.ReadSetting("General", "errorColorIndex", RGB(255, 0, 0))
    SetUnderlineStyle AppSettings.ReadSetting("General", "underlineStyle", WdUnderline.wdUnderlineWavy)
    SetUnderlineColor AppSettings.ReadSetting("General", "underlineColor", WdColor.wdColorRed)
    DisplayLanguageEvent (GetDisplayLocale)
End Sub
Sub LoadDisplayLanguages()
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    Dim filePath As String
    Dim fileName As Variant
    Dim iFileCount As Long
    Dim fileElements As Object
    
    filePath = Filesystem.GetAppDataFolder & "locale"
    Dim fileList As Collection
    Set fileList = Filesystem.GetFilesCollection(filePath & "\*.*")
    ReDim displayLanguageInfos(0 To fileList.count - 1)
    iFileCount = 0
    For Each fileName In fileList
        Set fileElements = Filesystem.ParseFileContent(filePath & "\" & CStr(fileName))
        
        displayLanguageInfos(iFileCount).Index = iFileCount
        displayLanguageInfos(iFileCount).Name = fileElements("language")
        displayLanguageInfos(iFileCount).Locale = fileElements("locale")
        
        iFileCount = iFileCount + 1
        
        Set fileElements = Nothing
    Next fileName
#If DebugMode = 0 Then
PROC_EXIT:
    Exit Sub
PROC_ERROR:
    Helper.ErrorHandler "Settings.LoadDisplayLanguages"
    GoTo PROC_EXIT
#End If
End Sub
Public Function GetDisplayLanguages() As Language.LanguageInfo()
    GetDisplayLanguages = displayLanguageInfos
End Function
Sub LoadSpellingLanguages()
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    Language.Initialize
    Dim filePath As String
    Dim fileName As Variant
    Dim iFileCount As Long
    Dim fileElements As Object
    Dim Locale As String
    
    filePath = Filesystem.GetAppDataFolder & "dictionaries"
    Dim fileList As Collection
    Set fileList = Filesystem.GetFilesCollection(filePath & "\*.aff")
    ReDim spellingLanguageInfos(0 To fileList.count - 1)
    iFileCount = 0
    For Each fileName In fileList
        Locale = Filesystem.RemoveFileExtension(CStr(fileName))
        spellingLanguageInfos(iFileCount).Index = iFileCount
        spellingLanguageInfos(iFileCount).Name = Language.GetLanguage(Locale)
        spellingLanguageInfos(iFileCount).Locale = Locale
        
        iFileCount = iFileCount + 1
    Next fileName
#If DebugMode = 0 Then
PROC_EXIT:
    Exit Sub
PROC_ERROR:
    Helper.ErrorHandler "Settings.LoadSpellingLanguages"
    GoTo PROC_EXIT
#End If
End Sub
Public Function GetSpellingLanguages() As Language.LanguageInfo()
    GetSpellingLanguages = spellingLanguageInfos
End Function
Sub SaveSettings()
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    AppSettings.WriteSetting "General", "displayLocale", displayLanguageInfos(frmSettings.cbxDisplayLanguages.ListIndex).Locale
    AppSettings.WriteSetting "General", "spellingLocale", spellingLanguageInfos(frmSettings.cbxSpellingLanguages.ListIndex).Locale
    AppSettings.WriteSetting "General", "splitCharacters", frmSettings.txtSplitCharacters.value
    AppSettings.WriteSetting "General", "autoCheck", frmSettings.chkAutoCheck.value
    AppSettings.WriteSetting "General", "autoClear", frmSettings.chkAutoClear.value
    AppSettings.WriteSetting "General", "errorColorIndex", frmSettings.txtColor.ForeColor
    AppSettings.WriteSetting "General", "underlineColor", GetUnderlineColor
    AppSettings.WriteSetting "General", "underlineStyle", GetUnderlineStyle
    AppSettings.WriteSetting "General", "useWindowless", frmSettings.chkWindowless.value
    LoadSettings
    Spelling.ReloadHunspell
    UI.ShowMessage Localization.GetLocalizedString("frmSettingsSettingsSaved")
#If DebugMode = 0 Then
PROC_EXIT:
    Exit Sub
PROC_ERROR:
    Helper.ErrorHandler "Settings.SaveSettings"
    GoTo PROC_EXIT
#End If
End Sub
Public Function GetLocaleIndex(Locale As String, languageInfos() As Language.LanguageInfo) As Long
Dim i As Long
Dim result As Long
    result = 0
    For i = 0 To UBound(languageInfos)
        If languageInfos(i).Locale = Locale Then
            result = i
            Exit For
        End If
    Next
    GetLocaleIndex = result
End Function
Sub GetLocaleIndexTest()
    Initialize
End Sub
Public Function GetSplitCharacters() As String
    GetSplitCharacters = strSplitCharacters
End Function
Public Sub SetSplitCharacters(value As String)
    strSplitCharacters = value
End Sub
Public Sub SetDisplayLocale(value As String)
    displayLocale = value
End Sub
Public Function GetDisplayLocale() As String
    GetDisplayLocale = displayLocale
End Function
Public Sub SetSpellingLocale(value As String)
    spellingLocale = value
End Sub
Public Function GetSpellingLocale() As String
    GetSpellingLocale = spellingLocale
End Function
Public Function GetLanguageID() As Long
    GetLanguageID = intLanguageID
End Function
Public Function GetLanguageName() As String
    GetLanguageName = strLanguageName
End Function
Public Function GetDisplayLocaleIndex()
    GetDisplayLocaleIndex = GetLocaleIndex(GetDisplayLocale, displayLanguageInfos)
End Function
Public Function GetSpellingLocaleIndex()
    GetSpellingLocaleIndex = GetLocaleIndex(GetSpellingLocale, spellingLanguageInfos)
End Function
Public Sub SetAutoCheck(value As Boolean)
    autoCheck = value
End Sub
Public Function GetAutoCheck() As Boolean
    GetAutoCheck = autoCheck
End Function
Public Sub SetAutoClear(value As Boolean)
    autoClear = value
End Sub
Public Function GetAutoClear() As Boolean
    GetAutoClear = autoClear
End Function
Public Sub SetErrorColorIndex(value As Long)
    errorColorIndex = value
End Sub
Public Function GetErrorColorIndex() As Long
    GetErrorColorIndex = errorColorIndex
End Function
Public Sub SetShowErrors(value As Boolean)
    showErrors = value
End Sub
Public Function GetShowErrors() As Boolean
    GetShowErrors = showErrors
End Function
Sub OpenColorPicker()
    On Error Resume Next
    Dim OriginalColor As Long
    Dim selectedColor As Long
    Dim selectedUnderline As Long
    Dim selectedUnderlineColor As Long
    
    OriginalColor = Spelling.GetRangeColor(Selection.Range)
    If Application.Dialogs(wdDialogFormatFont).Show = -1 Then
        selectedColor = Spelling.GetRangeColor(Selection.Range)
        selectedUnderline = Selection.Range.Font.underline
        selectedUnderlineColor = Selection.Range.Font.underlineColor
        ActiveDocument.Undo
        frmSettings.txtColor.ForeColor = selectedColor
        If selectedUnderline > 2 Then
            frmSettings.txtColor.Font.underline = selectedUnderline
            SetUnderlineStyle selectedUnderline
        Else
            UI.ShowMessage Localization.GetLocalizedString("msgInvalidUnderline", "Choosing no underline is not allowed.")
        End If
        SetUnderlineColor selectedUnderlineColor
        'frmSettings.txtColor.Font.UnderlineColor FOR FUTURE IMPLEMENTATION
        If Err.Number = 380 Then
            UI.ShowMessage Localization.GetLocalizedString("msgStandardColorsOnly", "Only standard colors can be chosen")
        End If
    End If
End Sub
Sub OpenDictionary()
    Dim filePath As String
    filePath = Filesystem.GetDictionaryFolder & "\" & Settings.GetSpellingLocale & ".dic"
    Filesystem.OpenTextFile filePath
End Sub
Sub OpenDictionariesFolder()
    Filesystem.OpenFolderInExplorer Filesystem.GetDictionaryFolder & "\"
End Sub
Public Sub SetUseWindowless(value As Boolean)
    useWindowless = value
End Sub
Public Function GetUseWindowless() As Boolean
    GetUseWindowless = useWindowless
End Function
Public Sub SetUnderlineStyle(value As Long)
    underlineStyle = value
End Sub
Public Function GetUnderlineStyle() As Long
    GetUnderlineStyle = underlineStyle
End Function
Public Sub SetUnderlineColor(value As Long)
    underlineColor = value
End Sub
Public Function GetUnderlineColor() As Long
    GetUnderlineColor = underlineColor
End Function
