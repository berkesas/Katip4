Attribute VB_Name = "Developer"
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
Private Const EXPORT_PATH As String = "C:\www\katip\Katip4\code"
Sub ExportAllComponents()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim exportPath As String
    Dim fileName As String

    exportPath = EXPORT_PATH

    If Right(exportPath, 1) <> "\" Then
        exportPath = exportPath & "\"
    End If

    Set vbProj = ThisDocument.VBProject

    For Each vbComp In vbProj.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule
                fileName = exportPath & vbComp.Name & ".bas"
            Case vbext_ct_Document
                fileName = exportPath & vbComp.Name & ".cls"
            Case vbext_ct_MSForm
                fileName = exportPath & vbComp.Name & ".frm"
        End Select
        
        ' Export the component to a text file if the filename is set
        If fileName <> "" Then
            vbComp.Export fileName
        End If
    Next vbComp
    
    ' Cleanup
    Set vbComp = Nothing
    Set vbProj = Nothing

    MsgBox "Modules, class modules, forms, and ThisDocument have been exported successfully.", vbInformation
End Sub
