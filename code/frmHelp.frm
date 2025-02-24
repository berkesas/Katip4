VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHelp 
   Caption         =   "Katip"
   ClientHeight    =   2340
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4632
   OleObjectBlob   =   "frmHelp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHelp"
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
    frmHelp.lblAbout.Caption = Localization.GetLocalizedString("lblAbout", "Open-source spellchecker add-in")
    frmHelp.btnOK.Caption = Localization.GetLocalizedString("btnOK", "OK")
End Sub
Private Sub btnOK_Click()
    frmHelp.Hide
End Sub
Public Sub LoadForm()
    Me.Show
End Sub
Private Sub lblLink_Click()
    Browser.OpenLinkInBrowser "https://berkesas.github.io/katip/"
End Sub
