VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InlineEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub ClearContextMenu()
    CommandBars("Text").Reset
    CommandBars("Table Text").Reset
End Sub

Sub CreateContextMenu(rng As Range, rngMisspelling As MisspellingRange, suggestions() As String)
    'On Error Resume Next
    Dim MenuButton As CommandBarButton
    Dim suggestionMenu As CommandBarPopup
    Dim suggestionMenuButton As CommandBarButton
    Dim ContextMenu As CommandBar
    Dim i As Long
    
    CustomizationContext = ActiveDocument
    
    If rng.Information(wdWithInTable) Then
        Set ContextMenu = CommandBars("Table Text")
    Else
        Set ContextMenu = CommandBars("Text")
    End If
    
    Set MenuButton = ContextMenu.Controls.Add(Type:=msoControlButton, Before:=1, Temporary:=True)
    
    With MenuButton
        .Caption = Localization.GetLocalizedString("btnIgnoreOnce", "Ignore Once")
        .FaceId = 9769
        .OnAction = "Spelling.IgnoreOnce"
    End With
    
    Set MenuButton = ContextMenu.Controls.Add(Type:=msoControlButton, Before:=2, Temporary:=True)
    
    With MenuButton
        .Caption = Localization.GetLocalizedString("btnIgnoreAll", "Ignore All")
        .FaceId = 9767
        .OnAction = "Spelling.IgnoreAll"
    End With
    
    Dim suggestionCount As Long
    suggestionCount = 0
    
    For i = LBound(suggestions) To UBound(suggestions)
        If i < 5 And Len(suggestions(i)) > 0 Then
            Set suggestionMenu = ContextMenu.Controls.Add(Type:=msoControlPopup, Before:=suggestionCount + 3, Temporary:=True)
            With suggestionMenu
                .Caption = suggestions(i)
            End With
            
            Set suggestionMenuButton = suggestionMenu.Controls.Add(Type:=msoControlButton, Before:=1, Temporary:=True)
            With suggestionMenuButton
                .Caption = Localization.GetLocalizedString("btnChange", "Change")
                .OnAction = "Spelling.RunFunction"
                .Parameter = "Spelling.Change|" & CStr(i)
            End With
            
            Set suggestionMenuButton = suggestionMenu.Controls.Add(Type:=msoControlButton, Before:=2, Temporary:=True)
            With suggestionMenuButton
                .Caption = Localization.GetLocalizedString("btnChangeAll", "Change All")
                .OnAction = "Spelling.RunFunction"
                .Parameter = "Spelling.ChangeAll|" & CStr(i)
            End With
            suggestionCount = suggestionCount + 1
        Else
            Exit For
        End If
    Next
    
    Set MenuButton = ContextMenu.Controls.Add(Type:=msoControlButton, Before:=suggestionCount + 3, Temporary:=True)
    
    With MenuButton
        .Caption = Localization.GetLocalizedString("btnClear", "Clear")
        .FaceId = 358
        .OnAction = "Spelling.ResetCheck"
    End With
    
    Set MenuButton = ContextMenu.Controls.Add(Type:=msoControlButton, Before:=suggestionCount + 4, Temporary:=True)
    
    With MenuButton
        .Caption = Localization.GetLocalizedString("btnAddToDictionary", "Add to Dictionary")
        .FaceId = 44
        .OnAction = "Spelling.ShowAddWord"
    End With
    
End Sub

Public Sub UpdateContextMenu(rng As Range, rngMisspelling As MisspellingRange, suggestions() As String)
    'On Error Resume Next
    'Exit Sub
    ClearContextMenu
    CreateContextMenu rng, rngMisspelling, suggestions
End Sub

