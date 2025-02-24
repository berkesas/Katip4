Attribute VB_Name = "AppSettings"
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

#If VBA7 And Win64 Then
    Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" _
        Alias "WritePrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, _
        ByVal lpString As String, _
        ByVal lpFileName As String) As Long
    
    Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" _
        Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long, _
        ByVal lpFileName As String) As Long
#Else
    Declare Function WritePrivateProfileString Lib "kernel32" _
        Alias "WritePrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, _
        ByVal lpString As String, _
        ByVal lpFileName As String) As Long
    
    Declare Function GetPrivateProfileString Lib "kernel32" _
        Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long, _
        ByVal lpFileName As String) As Long
#End If
Option Explicit
Const SETTINGS_FILE As String = "settings.ini"
Sub WriteSetting(section As String, key As String, value As String)
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    Dim result As Long
    result = WritePrivateProfileString(section, key, value, Filesystem.GetAppDataFolder & "\" & SETTINGS_FILE)
    If result = 0 Then
        UI.ShowMessage Localization.GetLocalizedString("txtFailIniFile", "Failed to write to INI file")
    End If
#If DebugMode = 0 Then
PROC_EXIT:
    Exit Sub
PROC_ERROR:
    Helper.ErrorHandler "AppSettings.WriteSetting"
    GoTo PROC_EXIT
#End If
End Sub
Function ReadSetting(section As String, key As String, defaultValue As String) As String
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    Dim result As String
    Dim buffer As String * 255
    Dim bufferSize As Long
    
    bufferSize = GetPrivateProfileString(section, key, defaultValue, buffer, Len(buffer), Filesystem.GetAppDataFolder & "\" & SETTINGS_FILE)
    If bufferSize > 0 Then
        result = Left(buffer, bufferSize)
    Else
        result = defaultValue
    End If
    
    ReadSetting = result
#If DebugMode = 0 Then
PROC_EXIT:
    Exit Function
PROC_ERROR:
    Helper.ErrorHandler "AppSettings.ReadSetting"
    GoTo PROC_EXIT
#End If
End Function


