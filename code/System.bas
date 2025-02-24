Attribute VB_Name = "System"
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
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetSystemWow64Directory Lib "kernel32" Alias "GetSystemWow64DirectoryW" (ByVal lpBuffer As String, ByVal uSize As Long) As Long
#Else
    Private Declare Function GetSystemWow64Directory Lib "kernel32" Alias "GetSystemWow64DirectoryW" (ByVal lpBuffer As String, ByVal uSize As Long) As Long
#End If
Function Is64BitWindows() As Boolean
    Dim buffer As String
    Dim result As Long
    buffer = String$(260, vbNullChar)
    
    On Error Resume Next
    result = GetSystemWow64Directory(buffer, 260)
    On Error GoTo 0
    
    Is64BitWindows = (result > 0)
End Function

Function GetWindowsBitVersion() As String
    If Is64BitWindows() Then
        GetWindowsBitVersion = "64-bit"
    Else
        GetWindowsBitVersion = "32-bit"
    End If
End Function

Function GetWordBitVersion() As String
    #If VBA7 Then
        #If Win64 Then
            GetWordBitVersion = "64-bit"
        #Else
            GetWordBitVersion = "32-bit"
        #End If
    #Else
        GetWordBitVersion = "32-bit"
    #End If
End Function
