Attribute VB_Name = "Test"
' Katip - Add-in to support spellchecking in Microsoft Word desktop application using Hunspell library and dictionaries
'
' Copyright (C) 2012-2025 Nazar Mammedov
' https://github.com/berkesas/katip4
'
' This software uses Hunspell Copyright (C) 2002-2022 N�meth L�szl�
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

Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As LongPtr) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long

Option Explicit
Dim oPopup As CommandBarPopup
Dim oCtr As CommandBarControl
Sub LoadAndCallDLL()
    'On Error Resume Next
    Dim hModule As LongPtr
    Dim funcPtr As LongPtr
    
    Dim strDll As String
    strDll = "hunspellvba.dll"
    hModule = LoadLibrary(StrPtr(strDll))
    
    If hModule = 0 Then
        Exit Sub
    End If

    funcPtr = GetProcAddress(hModule, "HunspellInit")
    If funcPtr = 0 Then
        Debug.Print "Failed to find function in DLL!"
        FreeLibrary hModule
        Exit Sub
    End If

    FreeLibrary hModule

    Debug.Print "Function loaded successfully, but needs Declare to call."
End Sub
