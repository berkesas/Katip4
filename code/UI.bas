Attribute VB_Name = "UI"
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
Dim KatipRibbon As IRibbonUI
Dim controlLabels As Object
Public isLoaded As Boolean
Sub RibbonOnLoad(ribbon As IRibbonUI)
'On Error Resume Next
    Set KatipRibbon = ribbon
    Initialize
End Sub
Sub Initialize()
    If Not isLoaded Then
        Main.Initialize
        LoadLabels
        frmAddWord.LoadLabels
        frmHelp.LoadLabels
        frmSettings.LoadLabels
        frmSpelling.LoadLabels
        isLoaded = True
    End If
End Sub
Sub LoadLabels()
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    If controlLabels Is Nothing Then
        Set controlLabels = CreateObject("Scripting.Dictionary")
    End If
    
    controlLabels("tabKatip") = Localization.GetLocalizedString("rtabKatipLabel", "Katip")
    controlLabels("grpSpelling") = Localization.GetLocalizedString("rgrpSpellingLabel", "Spelling")
    controlLabels("grpSettings") = Localization.GetLocalizedString("rgrpSettingsLabel", "Settings")
    controlLabels("grpHelp") = Localization.GetLocalizedString("rgrpHelpLabel", "Help")
    controlLabels("btnCheck") = Localization.GetLocalizedString("rbtnCheckCaption", "Check")
    controlLabels("btnSettings") = Localization.GetLocalizedString("rbtnSettingsCaption", "Settings")
    controlLabels("btnHelp") = Localization.GetLocalizedString("rbtnHelpCaption", "Help")
PROC_EXIT:
    Exit Sub
PROC_ERROR:
    Helper.ErrorHandler "UI.LoadLabels"
    GoTo PROC_EXIT
End Sub
Sub GetLabel(control As IRibbonControl, ByRef returnedVal)
    If Not controlLabels Is Nothing Then
        returnedVal = controlLabels(control.ID)
    Else
        returnedVal = "Not loaded"
    End If
End Sub
Sub UpdateControlLabels()
    If Not KatipRibbon Is Nothing Then
        Dim key As Variant
        LoadLabels
        For Each key In controlLabels.Keys
            KatipRibbon.InvalidateControl key
        Next key
    End If
End Sub
Public Sub btnCheck_Click(control As IRibbonControl)
    Spelling.ShowSpelling
End Sub
Public Sub btnSettings_Click(control As IRibbonControl)
    frmSettings.LoadForm
End Sub
Public Sub btnHelp_Click(control As IRibbonControl)
    frmHelp.LoadForm
End Sub
Public Sub ShowMessage(message As String)
    MsgBox message, vbOKOnly, Localization.GetLocalizedString("appName", "Katip")
End Sub
