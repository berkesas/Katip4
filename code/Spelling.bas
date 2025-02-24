Attribute VB_Name = "Spelling"
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
    Private Declare PtrSafe Sub HunspellInit Lib "hunspellvba.dll" _
        (ByRef hunspell As LongPtr, ByVal affPath As String, ByVal dicPath As String)
    Private Declare PtrSafe Function CheckSpelling Lib "hunspellvba.dll" _
        (ByVal hunspell As LongPtr, ByVal word As LongPtr) As Boolean
    Private Declare PtrSafe Sub HunspellFree Lib "hunspellvba.dll" _
        (ByVal hunspell As LongPtr)
    Private Declare PtrSafe Function GetSuggestions Lib "hunspellvba.dll" _
        (ByVal hunspell As LongPtr, ByVal word As LongPtr, ByRef count As Long) As LongPtr
    Private Declare PtrSafe Function GetMisspellings Lib "hunspellvba.dll" _
        (ByVal hunspell As LongPtr, ByVal word As LongPtr, ByRef count As Long) As LongPtr
    'Not private because called in Helper module
    Declare PtrSafe Sub FreeItems Lib "hunspellvba.dll" _
        (ByVal lpItems As LongPtr, ByVal count As Long)
    Private Declare PtrSafe Function AddDictionary Lib "hunspellvba.dll" _
        (ByVal hunspell As LongPtr, ByVal dicPath As String) As Long
    Private Declare PtrSafe Function AddWord Lib "hunspellvba.dll" _
        (ByVal hunspell As LongPtr, ByVal word As LongPtr) As Long
#Else
    Private Declare Sub HunspellInit Lib "hunspellvba.dll" _
        (ByRef hunspell As Long, ByVal affPath As String, ByVal dicPath As String)
    Private Declare Function CheckSpelling Lib "hunspellvba.dll" _
        (ByVal hunspell As Long, ByVal word As Long) As Boolean
    Private Declare Sub HunspellFree Lib "hunspellvba.dll" _
        (ByVal hunspell As Long)
    Private Declare Function GetSuggestions Lib "hunspellvba.dll" _
        (ByVal hunspell As Long, ByVal word As Long, ByRef count As Long) As Long
    Private Declare Function GetMisspellings Lib "hunspellvba.dll" _
        (ByVal hunspell As Long, ByVal word As Long, ByRef count As Long) As Long
    'Not private because called in Helper module
    Declare Sub FreeItems Lib "hunspellvba.dll" _
        (ByVal lpItems As Long, ByVal count As Long)
    Private Declare Function AddDictionary Lib "hunspellvba.dll" _
        (ByVal hunspell As Long, ByVal dicPath As String) As Long
    Private Declare Function AddWord Lib "hunspellvba.dll" _
        (ByVal hunspell As Long, ByVal word As Long) As Long
#End If
Option Explicit
Private Const LENGTH_BEFORE_MISSPELLING As Long = 30
Private Const PROGRESS_WIDTH As Long = 246
Public Enum MisspellingStatus
    Error = -1
    Ignored = 0
    Fixed = 1
    NotFound = 2
End Enum
Public Type MisspellingRange
    Text As String
    Start As Long
    End As Long
    Status As Long
    OriginalColor As Long
End Type
Private spellingLocale As String
#If VBA7 And Win64 Then
Private hunspell As LongPtr
#Else
Private hunspell As Long
#End If
Private misspellings() As MisspellingRange
Dim intMisspellingCount As Long
Private regex As VBScript_RegExp_55.RegExp
Private intCurrentError As Long
Private errorColorIndex As Long
Sub Initialize()
    LoadHunspell
End Sub
Public Sub ShowSpelling()
    UI.Initialize
    frmSpelling.chkAutoCheck.value = Settings.GetAutoCheck
    frmSpelling.chkAutoClear.value = Settings.GetAutoClear
    FirstError
    'frmSpelling.btnChange.Enabled = False
    'frmSpelling.btnChangeAll.Enabled = False
    If Settings.GetAutoCheck = True Then
        CheckDocument
    End If
    frmSpelling.Show (0)
End Sub
Public Sub HideSpelling()
    If Settings.GetAutoClear = True Then
        ResetCheck
    End If
    frmSpelling.Hide
End Sub
Sub LoadHunspell()
#If DebugMode = 0 Then
On Error GoTo PROC_ERROR
#End If
    Dim dictionaryFolder As String
    Dim spellingLocale As String
    Dim affixPath As String
    Dim dictionaryPath As String
    Dim userDictionaryPath As String

    dictionaryFolder = Filesystem.GetDictionaryFolder
    spellingLocale = Settings.GetSpellingLocale

    affixPath = dictionaryFolder & "\" & spellingLocale & ".aff"
    dictionaryPath = dictionaryFolder & "\" & spellingLocale & ".dic"
    userDictionaryPath = dictionaryFolder & "\" & spellingLocale & "_user.dic"

    ' Debug output for file paths
    'Debug.Print affixPath
    'Debug.Print dictionaryPath
    'Debug.Print userDictionaryPath

    If Filesystem.FileExists(affixPath) And Filesystem.FileExists(dictionaryPath) Then
        Call HunspellInit(hunspell, affixPath, dictionaryPath)
        If Filesystem.FileExists(userDictionaryPath) Then
            Call AddDictionary(hunspell, userDictionaryPath)
        End If
    End If
    
    'Debug.Print "Dictionary loaded: " & hunspell
PROC_EXIT:
    Exit Sub
PROC_ERROR:
    Helper.ErrorHandler "Spelling.LoadHunspell"
    GoTo PROC_EXIT
End Sub

Sub UnloadHunspell()
    On Error Resume Next
    If hunspell <> 0 Then
        'Debug.Print "Dictionary unloading: " & hunspell
        HunspellFree hunspell
    End If
End Sub
Sub ReloadHunspell()
    UnloadHunspell
    LoadHunspell
End Sub
Sub CheckDocument()
#If DebugMode = 0 Then
'On Error GoTo PROC_ERROR
#End If
    Dim rng As Range
    'Dim spellingRange As Range
    'Dim result As Boolean
    'Dim checkableWord As String
    'Dim utf8Bytes() As Byte
    'Dim b As Variant
    'Dim wordPtr As Long
    Dim misRange As MisspellingRange
    Dim intTotalWords As Long
    Dim intCurrentWord As Long
    
    
    ResetCheck
    
    Set regex = New VBScript_RegExp_55.RegExp
    With regex
        .Pattern = "[" & Settings.GetSplitCharacters & _
        Chr(13) & Chr(10) & Chr(11) & ChrW(160) & Chr(7) & "]"
        'ChrW(&HA4) & Chr(7) & ChrW(&H200C) & ChrW(&H200B) & ChrW(&H200D) & Chr(109) &
        '.Pattern = "[" & "',.!?:;{}()\[\]/\\=\+±\^\$\*<>|¦#@%&~…©›·`?×«»—°¨‘’™" & "]"
        '"',.!?:;{}()[]/\=+±^$*<>|¦#@%&~…©›·`?×«»—°¨‘’™"
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
    End With
    
    errorColorIndex = Settings.GetErrorColorIndex
    
    If hunspell = 0 Then
        'Debug.Print "Dictionary is not loaded. Re-loading."
        ReloadHunspell
    End If
    
    'check spelling of selection
    'If Selection.Start <> Selection.End Then
    '    Set spellingRange = Selection.Range
    'Else
    '    Set spellingRange = ActiveDocument.Range
    'End If
    
    ActiveDocument.Range.LanguageID = Settings.GetLanguageID
    ActiveDocument.Range.NoProofing = True
    
    Application.ScreenUpdating = True
    frmSpelling.lblProgress.Width = 0
    frmSpelling.lblProgress.Visible = True
    
    intTotalWords = ActiveDocument.Words.count
    intCurrentWord = 1
    
    intMisspellingCount = 0
    intCurrentError = -1
    
    Dim para As Paragraph
    
    If hunspell <> 0 Then
        For Each rng In ActiveDocument.StoryRanges
            If rng.StoryType = wdMainTextStory Then
                CheckRangeSpelling rng
            End If
            DoEvents
        Next
    Else
        'Debug.Print "Dictionary is not loaded: " & hunspell
    End If
    
    If intMisspellingCount > 0 Then
        intCurrentError = 0
    End If
    NavigateErrors
    
    frmSpelling.lblProgress.Visible = False
    Set regex = Nothing
PROC_EXIT:
    Exit Sub
PROC_ERROR:
    Helper.ErrorHandler "Spelling.CheckDocument"
    GoTo PROC_EXIT
End Sub
Function CheckRangeSpelling(rngStory As Range)
    Dim arrMisspellings() As String
    Dim i As Long
    Dim strText As String
    Dim rng As Range
    Dim lastPosition As Long
    
    lastPosition = 0
    strText = CleanWord(rngStory.Text)
    Do
        arrMisspellings = GetMisspellingsList(hunspell, strText)
        If Helper.IsArrayReady(arrMisspellings) = True Then
            For i = LBound(arrMisspellings) To UBound(arrMisspellings)
                'Debug.Print arrMisspellings(i)
                Set rng = FindTextRange(rngStory, arrMisspellings(i), lastPosition)
                If Not rng Is Nothing Then
                    'Debug.Print rng.Text & ":" & rng.Start & ":" & rng.End
                    ReDim Preserve misspellings(0 To intMisspellingCount)
                    misspellings(intMisspellingCount).Text = rng.Text
                    misspellings(intMisspellingCount).Start = rng.Start
                    misspellings(intMisspellingCount).End = rng.End
                    misspellings(intMisspellingCount).Status = MisspellingStatus.Error
                    misspellings(intMisspellingCount).OriginalColor = GetRangeColor(rng)
                    SetRangeColor rng, intMisspellingCount, errorColorIndex
                    'rng.Font.TextColor.RGB = errorColorIndex
                    'rng.NoProofing = False
                    intMisspellingCount = intMisspellingCount + 1
                    lastPosition = rng.End
                End If
            Next
        End If
        Set rngStory = rngStory.NextStoryRange
    Loop Until rngStory Is Nothing
End Function
Function FindTextRange(searchRange As Range, strSearch As String, startPosition As Long) As Range
    Dim foundRange As Range
    Dim startPos As Long
    Dim endPos As Long
    Dim i As Long
    
    Set foundRange = searchRange.Duplicate
    searchRange.Start = startPosition
        
    ' Initialize the search range
    Set foundRange = searchRange.Duplicate
    With foundRange.Find
        .Text = strSearch
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = False
        .MatchWholeWord = False
        .Execute
    End With

    ' If the string is found, return the range
    If foundRange.Find.found Then
        Set FindTextRange = foundRange
    Else
        Set FindTextRange = Nothing
    End If
End Function
#If VBA7 And Win64 Then
Function GetMisspellingsList(hunspell As LongPtr, strText As String) As String()
    Dim itemsPtr As LongPtr
    Dim items() As String
    Dim count As Long

    If hunspell <> 0 Then
        itemsPtr = GetMisspellings(hunspell, StrPtr(strText), count)
        If count > 0 Then
            items = Helper.UTF8ToString(itemsPtr, count)
        End If
    End If
    GetMisspellingsList = items
End Function
#Else
Function GetMisspellingsList(hunspell As Long, strText As String) As String()
    Dim itemsPtr As Long
    Dim items() As String
    Dim count As Long

    If hunspell <> 0 Then
        itemsPtr = GetMisspellings(hunspell, StrPtr(strText), count)
        If count > 0 Then
            items = Helper.ConvertPointerToArray(itemsPtr, count)
        End If
    End If
    GetMisspellingsList = items
End Function
#End If
Function CleanWord(strInput As String) As String
'On Error GoTo PROC_ERROR
'On Error Resume Next
    Dim result As String
    If Len(strInput) > 0 Then
        result = regex.Replace(strInput, " ")
        'If Err.no = 5018 Then
        '    MsgBox "Split characters pattern is not correct. Check Settings > Spelling section."
        'End If
    Else
        result = ""
    End If
    CleanWord = result
PROC_EXIT:
    Exit Function
PROC_ERROR:
    Helper.ErrorHandler "Spelling.CleanWord:" & strInput
    GoTo PROC_EXIT
End Function
Sub NavigateErrors()
    'On Error Resume Next
    Dim key As Long
    Dim currentSentence As String
    Dim rng As Range
    Dim startPos As Long
    Dim misspelling As MisspellingRange
            
    If GetErrorCount() = 0 Then
        frmSpelling.txtMisspelling.Text = Localization.GetLocalizedString("txtNoErrors", "No misspellings found")
        frmSpelling.DisableControls
        intCurrentError = -1
    Else
        frmSpelling.EnableControls
        'Debug.Print "Current Error:" & intCurrentError & ":" & misspellings(intCurrentError).Text
        misspelling = misspellings(intCurrentError)
        If misspelling.Start >= ActiveDocument.Range.Start And misspelling.End <= ActiveDocument.Range.End Then
            Set rng = ActiveDocument.Range(Start:=misspelling.Start, End:=misspelling.End)
            currentSentence = GetSentenceContainingRange(rng)
            frmSpelling.txtMisspelling.Text = currentSentence
            
            startPos = InStr(currentSentence, misspelling.Text)
            If startPos > 0 Then
                SetTextColor frmSpelling.txtMisspelling, 0, Len(currentSentence), WdColor.wdColorAutomatic
                SetTextColor frmSpelling.txtMisspelling, startPos - 1, Len(misspelling.Text), errorColorIndex
            End If
            GetSuggestionsList hunspell, misspelling.Text
            ActivateRange misspelling
        End If
    End If
End Sub
Function GetErrorCount() As Long
    Dim result As Long
    Dim i As Long
    result = 0
    If intMisspellingCount > 0 Then
        For i = 0 To UBound(misspellings)
            If misspellings(i).Status = MisspellingStatus.Error Then
                result = result + 1
            End If
        Next
    Else
        result = 0
    End If
    GetErrorCount = result
End Function
Function GetSentenceContainingRange(rng As Range) As String
    Dim foundPos As Long
    Dim strResult As String
    If rng.Information(wdWithInTable) Then
        Dim tempSentence As String
        tempSentence = rng.Cells(1).Range.Text
        strResult = Left(tempSentence, Len(tempSentence) - 2)
    Else
        strResult = rng.Sentences(1).Text
    End If
    
    'Debug.Print "The sentence containing the range is: " & rng.Text & ":" & sentenceRng
    'If Len(strResult) > Len(rng.Text) Then
        foundPos = InStr(strResult, rng.Text)
        If foundPos > LENGTH_BEFORE_MISSPELLING Then
            strResult = Mid(strResult, foundPos - LENGTH_BEFORE_MISSPELLING)
        End If
    'Else
    '    strResult = rng.Text
    'End If
    GetSentenceContainingRange = strResult
    ' Display the sentence in a message box
End Function
Private Sub SetTextColor(ByRef inkEditControl As Object, ByVal startPos As Long, ByVal length As Long, ByVal color As Long)
    With inkEditControl
        .SelStart = startPos
        .SelLength = length
        .SelColor = Helper.GetRGB(color)
    End With
End Sub
Sub FirstError()
    If intCurrentError <> -1 Then
        intCurrentError = 0
    End If
    NavigateErrors
End Sub
Sub PreviousError()
    intCurrentError = GetErrorIndex(intCurrentError, -1)
    NavigateErrors
End Sub
Sub NextError()
    intCurrentError = GetErrorIndex(intCurrentError, 1)
    NavigateErrors
End Sub
Function GetErrorIndex(intErrorIndex As Long, intDirection As Long) As Long
    Dim i As Long
    Dim intStart As Long
    Dim intEnd As Long
    Dim result As Long
    
    result = intErrorIndex
    If intDirection = 1 Then
        intStart = intErrorIndex + 1
        intEnd = UBound(misspellings)
    Else
        intStart = intErrorIndex - 1
        intEnd = LBound(misspellings)
    End If
    
    'Debug.Print intErrorIndex & ":" & "0" & ":" & UBound(misspellings)
        
    For i = intStart To intEnd Step intDirection
        If misspellings(i).Status = MisspellingStatus.Error Then
            result = i
            Exit For
        End If
    Next
    
    GetErrorIndex = result
End Function
#If VBA7 And Win64 Then
Sub GetSuggestionsList(hunspell As LongPtr, word As String)
    Dim suggestionsPtr As LongPtr
    Dim suggestions() As String
    Dim count As Long
    Dim wordPtr As LongPtr
    Dim i As Long
    
    frmSpelling.lbxSuggestions.Clear
    
    If hunspell <> 0 Then
        'Debug.Print "Suggestions from hunspell: " & hunspell
        wordPtr = StrPtr(word)
        suggestionsPtr = GetSuggestions(hunspell, wordPtr, count)
        suggestions = Helper.UTF8ToString(suggestionsPtr, count)
        If UBound(suggestions) > -1 Then
            For i = 0 To UBound(suggestions) - 1
                frmSpelling.lbxSuggestions.AddItem suggestions(i)
                'Debug.Print suggestions(i)
            Next i
            If frmSpelling.lbxSuggestions.ListCount > 0 Then
                frmSpelling.lbxSuggestions.ListIndex = 0
            End If
        End If
    End If
End Sub
#Else
Sub GetSuggestionsList(hunspell As Long, word As String)
    Dim suggestionsPtr As Long
    Dim suggestions() As String
    Dim count As Long
    Dim wordPtr As Long
    Dim i As Long
    
    frmSpelling.lbxSuggestions.Clear
    
    If hunspell <> 0 Then
        'Debug.Print "Suggestions from hunspell: " & hunspell
        wordPtr = StrPtr(word)
        suggestionsPtr = GetSuggestions(hunspell, wordPtr, count)
        suggestions = Helper.UTF8ToString(suggestionsPtr, count)
        If UBound(suggestions) > -1 Then
            For i = 0 To UBound(suggestions) - 1
                frmSpelling.lbxSuggestions.AddItem suggestions(i)
                'Debug.Print suggestions(i)
            Next i
            If frmSpelling.lbxSuggestions.ListCount > 0 Then
                frmSpelling.lbxSuggestions.ListIndex = 0
            End If
        End If
    End If
End Sub
#End If
Sub ActivateRange(misspelling As MisspellingRange)
    Dim rng As Range
    Set rng = ActiveDocument.Range(Start:=misspelling.Start, End:=misspelling.End)
    rng.Select
    ActiveWindow.ScrollIntoView Selection.Range, True
    'ShowDialogAwayFromSelection
End Sub
Sub IgnoreOnce()
    Dim strSource As String
    
    strSource = misspellings(intCurrentError).Text
    
    IgnoreRange intCurrentError, strSource
    NextError
End Sub
Sub IgnoreAll()
    Dim strSource As String
    Dim i As Long
    
    strSource = misspellings(intCurrentError).Text
    If intMisspellingCount > 0 Then
        For i = 0 To UBound(misspellings)
            If misspellings(i).Text = strSource Then
                IgnoreRange i, strSource
            End If
        Next
    End If
    NextError
End Sub
Sub IgnoreRange(intErrorIndex As Long, strSource As String)
    Dim rng As Range
    Dim i As Long
    
    Set rng = ActiveDocument.Range(Start:=misspellings(intErrorIndex).Start, End:=misspellings(intErrorIndex).End)
    #If VBA7 Then
    rng.Font.TextColor.RGB = misspellings(intErrorIndex).OriginalColor
    #Else
    rng.Font.color = misspellings(intErrorIndex).OriginalColor
    #End If
    misspellings(intErrorIndex).Status = MisspellingStatus.Ignored
    i = AddWord(hunspell, StrPtr(strSource))
End Sub
Sub Change()
    Dim strSource As String
    Dim strTarget As String
    
    strSource = misspellings(intCurrentError).Text
    strTarget = frmSpelling.lbxSuggestions.value
    
    ChangeRange intCurrentError, strSource, strTarget
    NextError
End Sub
Sub ChangeAll()
    Dim strSource As String
    Dim strTarget As String
    Dim i As Long
    strSource = misspellings(intCurrentError).Text
    strTarget = frmSpelling.lbxSuggestions.value
    
    For i = 0 To UBound(misspellings)
        If misspellings(i).Text = strSource Then
            'intCurrentError = i
            ChangeRange i, strSource, strTarget
        End If
    Next
    NextError
End Sub
Function GetRangeColor(rng As Range) As Long
    'On Error Resume Next
    #If VBA7 Then
    If Main.CompatibilityVersion >= 14 Then
        GetRangeColor = rng.Font.TextColor.RGB
    Else
        GetRangeColor = rng.Font.color
    End If
    #Else
        GetRangeColor = rng.Font.color
    #End If
End Function
Sub SetRangeColor(rng As Range, intErrorIndex As Long, color As Long)
    #If VBA7 Then
    If Main.CompatibilityVersion >= 14 Then
        rng.Font.TextColor.RGB = color
    Else
        rng.Font.color = color
    End If
    #Else
        rng.Font.color = color
    #End If
End Sub
Sub ChangeRange(intErrorIndex As Long, strSource As String, strTarget As String)
    Dim rng As Range
    Set rng = ActiveDocument.Range(Start:=misspellings(intErrorIndex).Start, End:=misspellings(intErrorIndex).End)
    SetRangeColor rng, intErrorIndex, misspellings(intErrorIndex).OriginalColor
    rng.Text = strTarget
    misspellings(intErrorIndex).Status = MisspellingStatus.Fixed
    UpdateRanges intErrorIndex, Len(strTarget) - Len(strSource)
End Sub
Sub UpdateRanges(intErrorIndex As Long, intDifference As Long)
    Dim i As Long
    For i = intErrorIndex + 1 To UBound(misspellings)
        misspellings(i).Start = misspellings(i).Start + intDifference
        misspellings(i).End = misspellings(i).End + intDifference
    Next
End Sub
Sub SuggestionChanged()
    If frmSpelling.lbxSuggestions.ListIndex = -1 Then
        frmSpelling.btnChange.Enabled = False
        frmSpelling.btnChangeAll.Enabled = False
    Else
        frmSpelling.btnChange.Enabled = True
        frmSpelling.btnChangeAll.Enabled = True
    End If
End Sub
Sub ResetCheck()
    Dim i As Long
    Dim rng As Range
    Dim misspelling As MisspellingRange
    If intMisspellingCount > 0 Then
        For i = 0 To UBound(misspellings)
            misspelling = misspellings(i)
            If misspelling.Status = MisspellingStatus.Error Then
                If misspelling.Start >= ActiveDocument.Range.Start And misspelling.End <= ActiveDocument.Range.End Then
                    Set rng = ActiveDocument.Range(Start:=misspelling.Start, End:=misspelling.End)
                    'Debug.Print rng.Text & ":" & misspelling.OriginalColor
                    SetRangeColor rng, i, misspelling.OriginalColor
                End If
            End If
        Next
    End If
    intCurrentError = -1
    intMisspellingCount = 0
    ReDim misspellings(0 To 0)
    NavigateErrors
End Sub
Sub ShowDialogAwayFromSelection()
    Dim selectionRange As Range
    Dim leftPos As Long, topPos As Long
    Dim dialogWidth As Long, dialogHeight As Long
    Dim windowWidth As Long, windowHeight As Long
    
    Set selectionRange = Selection.Range
    
    leftPos = selectionRange.Information(wdHorizontalPositionRelativeToTextBoundary)
    topPos = selectionRange.Information(wdVerticalPositionRelativeToTextBoundary)
    
    windowWidth = ActiveWindow.Width
    windowHeight = ActiveWindow.Height
    
    dialogWidth = frmSpelling.Width
    dialogHeight = frmSpelling.Height
    
    If topPos + dialogHeight > windowHeight Then
        frmSpelling.Top = topPos - dialogHeight - 20
    Else
        frmSpelling.Top = topPos + 20
    End If
End Sub
Sub AddToDictionary()
    Dim strWord As String
    Dim strAffixes As String
    Dim strNewLine As String
    Dim i As Long
    Dim userDictionaryPath As String
    
    userDictionaryPath = Filesystem.GetDictionaryFolder & "\" & Settings.GetSpellingLocale & "_user.dic"
    
    strWord = frmAddWord.txtWord.value
    strAffixes = frmAddWord.txtAffixes.value
    
    If Not FileExists(userDictionaryPath) Then
        Filesystem.CreateFile userDictionaryPath, "0"
    End If
    If Len(strWord) > 0 Then
        If Len(strAffixes) > 0 Then
            strNewLine = strWord & "/" & strAffixes
        Else
            strNewLine = strWord
        End If
        Filesystem.AppendToFile userDictionaryPath, strNewLine
        UI.ShowMessage Localization.GetLocalizedString("msgWordAdded", "Word added.")
    Else
        UI.ShowMessage Localization.GetLocalizedString("msgWordBlank", "Word cannot be empty.")
    End If
    frmAddWord.Hide
End Sub
Sub ShowAddWord()
    Dim strSource As String
    strSource = misspellings(intCurrentError).Text
    frmAddWord.txtWord.value = strSource
    frmAddWord.txtAffixes.value = ""
    frmAddWord.Show (1)
End Sub
