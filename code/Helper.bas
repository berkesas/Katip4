Attribute VB_Name = "Helper"
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
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Any, ByVal length As LongPtr)
Private Declare PtrSafe Function lstrlenA Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Any, _
    ByVal cbMultiByte As Long, ByVal lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Any, ByVal length As Long)
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Any, _
    ByVal cbMultiByte As Long, ByVal lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
#End If
Option Explicit
' Constants
Const CP_UTF8 As Long = 65001
Function IsArrayReady(items() As String) As Boolean
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    Dim result As Boolean
    If IsArray(items) Then
        On Error Resume Next
        Dim arrayAssigned As Boolean
        result = (LBound(items) <= UBound(items))
        On Error GoTo 0
    Else
        result = False
    End If
    IsArrayReady = result
#If DebugMode = 0 Then
PROC_EXIT:
    Exit Function
PROC_ERROR:
    Helper.ErrorHandler "Helper.IsArrayReady"
    GoTo PROC_EXIT
#End If
End Function
Public Function GetRGB(wdColorValue As Long) As Long
    Dim redValue As Long
    Dim greenValue As Long
    Dim blueValue As Long
    
    ' Extract the red, green, and blue components
    redValue = wdColorValue Mod 256
    greenValue = (wdColorValue \ 256) Mod 256
    blueValue = (wdColorValue \ 65536) Mod 256
    
    GetRGB = RGB(redValue, greenValue, blueValue)
End Function
Function GetRGBValue(colorValue As Long) As Long
    Dim red As Integer
    Dim green As Integer
    Dim blue As Integer

    colorValue = -654245889
    red = (colorValue And &HFF)
    green = (colorValue \ &H100 And &HFF)
    blue = (colorValue \ &H10000 And &HFF)

    GetRGBValue = RGB(red, green, blue)
End Function
Sub ErrorHandler(Optional caller As String = "")
    If Err Then
        Filesystem.LogError CStr(Now()) & " " & Main.MACRO_VERSION & " Error code: " & Err.Number & " Description: " & Err.Description & " Caller:" & caller
        If Settings.GetShowErrors Then
            UI.ShowMessage " Error code: " & Err.Number & " Description: " & Err.Description & " Caller:" & caller
        End If
    End If
End Sub
#If VBA7 And Win64 Then
Public Function ConvertPointerToArray(itemsPtr As LongPtr, count As Long) As String()
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    Dim resultArray() As String
    Dim itemPtr As LongPtr
    Dim item As String
    Dim i As Long

    If itemsPtr <> 0 Then
        ReDim resultArray(0 To count - 1)
        For i = 0 To count - 1
            CopyMemory itemPtr, ByVal (itemsPtr + i * LenB(itemPtr)), LenB(itemPtr)
            resultArray(i) = PtrToStringUTF8(itemPtr)
        Next i

        FreeItems itemsPtr, count
    End If
    ConvertPointerToArray = resultArray
#If DebugMode = 0 Then
PROC_EXIT:
    Exit Function
PROC_ERROR:
    Helper.ErrorHandler "Helper.ConvertPointerToArray"
    GoTo PROC_EXIT
#End If
End Function
Public Function PtrToStr(ptr As LongPtr) As String
    Dim strLen As Long
    Dim strBuffer As String

    strLen = lstrlenA(ptr)
    strBuffer = String$(strLen, vbNullChar)
    CopyMemory ByVal strBuffer, ByVal ptr, strLen

    PtrToStr = strBuffer
End Function
Public Function UTF8ToString(itemsPtr As LongPtr, count As Long) As String()
    Dim resultArray() As String
    Dim itemPtr As LongPtr
    Dim item As String
    Dim i As Long

    If itemsPtr <> 0 Then
        ReDim resultArray(0 To count)
        For i = 0 To count - 1
            CopyMemory itemPtr, ByVal (itemsPtr + i * LenB(itemPtr)), LenB(itemPtr)
            resultArray(i) = PtrToStringUTF8(itemPtr)
        Next i

        FreeItems itemsPtr, count
    End If
    UTF8ToString = resultArray
End Function
Private Function PtrToStringUTF8(ByVal ptr As LongPtr) As String
    Dim byteArray() As Byte
    Dim strLen As Long
    Dim wcharArray() As LongPtr
    Dim wcharLen As Long
    Dim resultStr As String

    strLen = lstrlenA(ptr)
    If strLen > 0 Then
        ReDim byteArray(0 To strLen - 1)
        CopyMemory byteArray(0), ByVal ptr, strLen

        wcharLen = MultiByteToWideChar(CP_UTF8, 0, VarPtr(byteArray(0)), strLen, ByVal 0&, 0)
        If wcharLen > 0 Then
            ReDim wcharArray(0 To (wcharLen * 2) - 1)
            MultiByteToWideChar CP_UTF8, 0, VarPtr(byteArray(0)), strLen, VarPtr(wcharArray(0)), wcharLen

            resultStr = String$(wcharLen, vbNullChar)
            CopyMemory ByVal StrPtr(resultStr), VarPtr(wcharArray(0)), wcharLen * 2

            PtrToStringUTF8 = resultStr
        Else
            PtrToStringUTF8 = ""
        End If
    Else
        PtrToStringUTF8 = ""
    End If
End Function
#Else
Public Function ConvertPointerToArray(itemsPtr As Long, count As Long) As String()
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    Dim resultArray() As String
    Dim itemPtr As Long
    Dim item As String
    Dim i As Long

    If itemsPtr <> 0 Then
        ReDim resultArray(0 To count)
        For i = 0 To count
            CopyMemory itemPtr, ByVal (itemsPtr + i * LenB(itemPtr)), LenB(itemPtr)
            resultArray(i) = PtrToStringUTF8(itemPtr)
        Next i

        FreeItems itemsPtr, count
    End If
    ConvertPointerToArray = resultArray
#If DebugMode = 0 Then
PROC_EXIT:
    Exit Function
PROC_ERROR:
    Helper.ErrorHandler "Helper.ConvertPointerToArray 32bit"
    GoTo PROC_EXIT
#End If
End Function
Public Function PtrToStr(ptr As Long) As String
    Dim strLen As Long
    Dim strBuffer As String

    strLen = lstrlenA(ptr)
    strBuffer = String$(strLen, vbNullChar)
    CopyMemory ByVal strBuffer, ByVal ptr, strLen

    PtrToStr = strBuffer
End Function
Public Function UTF8ToString(itemsPtr As Long, count As Long) As String()
    Dim resultArray() As String
    Dim itemPtr As Long
    Dim item As String
    Dim i As Long

    If itemsPtr <> 0 Then
        ReDim resultArray(0 To count)
        For i = 0 To count - 1
            CopyMemory itemPtr, ByVal (itemsPtr + i * LenB(itemPtr)), LenB(itemPtr)
            resultArray(i) = PtrToStringUTF8(itemPtr)
        Next i

        FreeItems itemsPtr, count
    End If
    UTF8ToString = resultArray
End Function
Private Function PtrToStringUTF8(ByVal ptr As Long) As String
    Dim byteArray() As Byte
    Dim strLen As Long
    Dim wcharArray() As Long
    Dim wcharLen As Long
    Dim resultStr As String

    strLen = lstrlenA(ptr)
    If strLen > 0 Then
        ReDim byteArray(0 To strLen - 1)
        CopyMemory byteArray(0), ByVal ptr, strLen

        wcharLen = MultiByteToWideChar(CP_UTF8, 0, VarPtr(byteArray(0)), strLen, ByVal 0&, 0)
        If wcharLen > 0 Then
            ReDim wcharArray(0 To (wcharLen * 2) - 1)
            MultiByteToWideChar CP_UTF8, 0, VarPtr(byteArray(0)), strLen, VarPtr(wcharArray(0)), wcharLen

            resultStr = String$(wcharLen, vbNullChar)
            CopyMemory ByVal StrPtr(resultStr), VarPtr(wcharArray(0)), wcharLen * 2

            PtrToStringUTF8 = resultStr
        Else
            PtrToStringUTF8 = ""
        End If
    Else
        PtrToStringUTF8 = ""
    End If
End Function
#End If
