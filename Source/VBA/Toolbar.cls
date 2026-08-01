VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Toolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.Toolbar
'
' Description:
'
' Author:
'   Paul Bender <pbender@alumni.ucsd.edu>
'
' Copyright:
'   Copyright (c) 2000,2001,2002,2003 Paul Bender
'
'   All rights reserved.
'
'   Permission is hereby granted, free of charge, to any person obtaining a
'   copy of this software and associated documentation files (the
'   "Software"), to deal in the Software without restriction, including
'   without limitation the rights to use, copy, modify, merge, publish,
'   distribute, and/or sell copies of the Software, and to permit persons
'   to whom the Software is furnished to do so, provided that the above
'   copyright notice(s) and this permission notice appear in all copies of
'   the Software and that both the above copyright notice(s) and this
'   permission notice appear in supporting documentation.
'
'   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
'   OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
'   MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT
'   OF THIRD PARTY RIGHTS. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR
'   HOLDERS INCLUDED IN THIS NOTICE BE LIABLE FOR ANY CLAIM, OR ANY SPECIAL
'   INDIRECT OR CONSEQUENTIAL DAMAGES, OR ANY DAMAGES WHATSOEVER RESULTING
'   FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT,
'   NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION
'   WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
'
'   Except as contained in this notice, the name of a copyright holder
'   shall not be used in advertising or otherwise to promote the sale, use
'   or other dealings in this Software without prior written authorization
'   of the copyright holder.
'
' Change History:
'   2.01.0000:
'     (1) Add support for "Set Template" menu item in "Song Edit" menu.
'   2.00.0000:
'     (1) Reset change history.
'===============================================================================


'===============================================================================
' Options.
'===============================================================================
Option Explicit
Option Compare Text
Option Base 0


'===============================================================================
' Public Constants.
'===============================================================================


'===============================================================================
' Private Constants.
'===============================================================================


'===============================================================================
' Public Variables.
'===============================================================================
Private mblnInvariant           As Boolean
Private mudtApplication         As WorshipServiceAssistant.Application
Private mudtParent              As WorshipServiceAssistant.Application
Private mtlb                    As Office.CommandBar
Private mmnuMode                As Office.CommandBarPopup
Private mcmdModeOperator        As Office.CommandBarButton
Private mcmdModeEditor          As Office.CommandBarButton
Private mcmdModeDeveloper       As Office.CommandBarButton
Private mcmdNavigator           As Office.CommandBarButton
Private mmnuSongEdit            As Office.CommandBarPopup
Private mmnuSongEditSetCategory As Office.CommandBarPopup
Private mcmdSongEditSetTemplate As Office.CommandBarButton
Private mcmdSongEditSort        As Office.CommandBarButton
Private mcmdSongEditCreateIndex As Office.CommandBarButton
Private mcboCategory            As Office.CommandBarComboBox
Private mmnuDebug               As Office.CommandBarPopup
Private mmnuDebugSSWMonitor     As Office.CommandBarPopup
Private mmnuDebugSSWSize        As Office.CommandBarPopup
Private mmnuHelp                As Office.CommandBarPopup


'===============================================================================
' Private Variables.
'===============================================================================


'===============================================================================
' Public Properties and Methods.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Initialize _
( _
    ByRef udtApplication As WorshipServiceAssistant.Application, _
    ByRef udtParent As WorshipServiceAssistant.Application _
)
    mblnInvariant = False
    
    Set mudtApplication = udtApplication
    Set mudtParent = udtParent
    
    Toolbar_Remove
    Toolbar_Add
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Terminate _
( _
)
    Toolbar_Remove
    
    Set mudtParent = Nothing
    Set mudtApplication = Nothing
    
    mblnInvariant = False
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Invariant _
( _
) As Boolean
    Invariant = mblnInvariant
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Let Invariant _
( _
    ByRef blnInvariant As Boolean _
)
    If (Not blnInvariant = mblnInvariant) Then
        If (blnInvariant = True) Then
            mblnInvariant = blnInvariant
        End If
        If (blnInvariant = False) Then
            mblnInvariant = blnInvariant
        End If
    End If
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Doppelganger _
( _
) As Office.CommandBar
    Set Doppelganger = mtlb
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Application _
( _
) As WorshipServiceAssistant.Application
    Set Application = mudtApplication
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Parent _
( _
) As WorshipServiceAssistant.Application
    Set Parent = mudtParent
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Update _
( _
)
    Toolbar_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Control_ModeValue_Action _
( _
)
    Select Case mudtApplication.Doppelganger.CommandBars.ActionControl.Parameter
        Case "Operator"
            mudtApplication.Mode = WorshipServiceAssistant.wsaModeOperator
        Case "Editor"
            mudtApplication.Mode = WorshipServiceAssistant.wsaModeEditor
        Case "Developer"
            mudtApplication.Mode = WorshipServiceAssistant.wsaModeDeveloper
    End Select
    
    If (Not mudtApplication.Mode = WorshipServiceAssistant.wsaModeDeveloper) Then
        With mudtApplication.Doppelganger.CommandBars(WorshipServiceAssistant.wsaApplicationName).Controls("Debug")
            mudtApplication.SlideShowWindowMonitor = _
                .Controls("SSW Monitor").Controls("<Default>").Parameter
            mudtApplication.SlideShowWindowSize = _
                VBA.CLng(.Controls("SSW Size").Controls("<Default>").Parameter)
        End With
    End If
    
    Toolbar_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Control_Navigator_Action _
( _
)
    mudtApplication.ActivateNavigator
    
    Toolbar_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Control_SongEditSetCategoryValue_Action _
( _
)
    Dim udtSlideSelection As WorshipServiceAssistant.SlideRange
    
    Set udtSlideSelection = mudtApplication.ActiveSlideSelection
    If (Not udtSlideSelection Is Nothing) Then
        udtSlideSelection.Category = CommandBars.ActionControl.Parameter
    End If
    
    Toolbar_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Control_SongEditSetTemplate_Action _
( _
)
    Dim udtPresentation As WorshipServiceAssistant.Presentation
    
    Set udtPresentation = mudtApplication.ActivePresentation
    If (Not udtPresentation Is Nothing) Then
        With Application.Doppelganger.FileDialog(Type:=msoFileDialogOpen)
            .InitialFileName = CurDir()
            .InitialView = msoFileDialogViewList
            .AllowMultiSelect = False
            If (.Show = -1) Then
                With .SelectedItems
                    If (.Count = 1) Then
                        udtPresentation.Doppelganger.ApplyTemplate FileName:=.Item(1)
                    End If
                End With
            End If
        End With
    End If
    
    Toolbar_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Control_SongEditSort_Action _
( _
)
    Dim udtPresentation As WorshipServiceAssistant.Presentation
    
    Set udtPresentation = mudtApplication.ActivePresentation
    If (Not udtPresentation Is Nothing) Then
        udtPresentation.Slides.SortByTitle
    End If
    
    VBA.MsgBox _
        Buttons:= _
            VBA.vbInformation, _
        Title:= _
            WorshipServiceAssistant.wsaApplicationNamePretty, _
        Prompt:= _
            "Slide sort is complete."
            
    Toolbar_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Control_SongEditIndex_Action _
( _
)
    Dim udtPresentation As WorshipServiceAssistant.Presentation
    
    Set udtPresentation = mudtApplication.ActivePresentation
    If (Not udtPresentation Is Nothing) Then
        udtPresentation.CreateIndex
    End If
    
    VBA.MsgBox _
        Buttons:= _
            VBA.vbInformation, _
        Title:= _
            WorshipServiceAssistant.wsaApplicationNamePretty, _
        Prompt:= _
            "Slide index generation is complete."
            
    Toolbar_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Control_DebugSSWMonitorValue_Action _
( _
)
    mudtApplication.SlideShowWindowMonitor = VBA.CStr(CommandBars.ActionControl.Parameter)
    
    Toolbar_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Control_DebugSSWSizeValue_Action _
( _
)
    mudtApplication.SlideShowWindowSize = VBA.CStr(CommandBars.ActionControl.Parameter)
    
    Toolbar_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Control_HelpTopicApplication_Action _
( _
)
    mudtApplication.Help.ShowTopic WorshipServiceAssistant.wsaHelpTopicApplication
    
    Toolbar_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Control_HelpTopicHowTo_Action _
( _
)
    mudtApplication.Help.ShowTopic WorshipServiceAssistant.wsaHelpTopicHowTo
    
    Toolbar_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Control_HelpTopicToolbar_Action _
( _
)
    mudtApplication.Help.ShowTopic WorshipServiceAssistant.wsaHelpTopicToolbar
    
    Toolbar_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Control_HelpTopicKnownIssues_Action _
( _
)
    mudtApplication.Help.ShowTopic WorshipServiceAssistant.wsaHelpTopicKnownIssues
    
    Toolbar_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Control_HelpTopicCopyrightPermission_Action _
( _
)
    mudtApplication.Help.ShowTopic WorshipServiceAssistant.wsaHelpTopicCopyrightPermission
    
    Toolbar_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Control_HelpVisitHomepage_Action _
( _
)
    mudtApplication.Help.VisitHomepage
    
    Toolbar_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Control_HelpEmailAuthor_Action _
( _
)
    mudtApplication.Help.EmailAuthor
    
    Toolbar_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Control_HelpAbout_Action _
( _
)
    mudtApplication.Help.ShowAbout
    
    Toolbar_Update
End Sub


'===============================================================================
' Private Properties and Methods.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub Toolbar_Remove _
( _
)
    Dim tlb As Office.CommandBar
    
    ' Uninstall any already installed toolbar instances.
    For Each tlb In mudtApplication.Doppelganger.CommandBars
        If (tlb.BuiltIn = False) Then
            If (tlb.Name = WorshipServiceAssistant.wsaApplicationName) Then
                tlb.Delete
            End If
        End If
    Next
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub Toolbar_Add _
( _
)
    Dim tlb As Office.CommandBar
    Dim lngBarRowIndex As Long
    
    ' Find the end of command bar list.
    ' The toolbar will be placed at the end of the command bar list.
    For Each tlb In mudtApplication.Doppelganger.CommandBars
        If (tlb.RowIndex > lngBarRowIndex) Then
            lngBarRowIndex = tlb.RowIndex
        End If
    Next
    lngBarRowIndex = lngBarRowIndex + 1
    
    ' Install and configure the toolbar.
    With mudtApplication.Doppelganger.CommandBars.Add(Temporary:=True)
        .Name = WorshipServiceAssistant.wsaApplicationName
        .Position = msoBarTop
        .RowIndex = lngBarRowIndex
        .Visible = True
        .Protection = _
            Office.msoBarNoChangeDock + _
            Office.msoBarNoChangeVisible + _
            Office.msoBarNoCustomize + _
            Office.msoBarNoMove + _
            Office.msoBarNoResize + _
            Office.msoBarNoVerticalDock
    End With
    Set mtlb = mudtApplication.Doppelganger.CommandBars(WorshipServiceAssistant.wsaApplicationName)
        
    ' Install the toolbar items.
    Control_Mode_Add
    Control_Navigator_Add
    Control_SongEdit_Add
    Control_Category_Add
    Control_Debug_Add
    Control_Help_Add
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub Control_Mode_Add _
( _
)
    With mtlb.Controls.Add(Type:=Office.msoControlPopup, Temporary:=True)
        .Caption = "Mode"
        .TooltipText = "Select the Worship Service Assistant mode"
        .BeginGroup = True
        .Width = 72
        With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
            .Caption = "&Operator"
            .TooltipText = "Set the mode to Operator"
            .OnAction = "ToolbarAction.ModeButton"
            .Parameter = "Operator"
            .BeginGroup = False
        End With
        With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
            .Caption = "&Editor"
            .TooltipText = "Set the mode to Editor"
            .OnAction = "ToolbarAction.ModeButton"
            .Parameter = "Editor"
            .BeginGroup = False
        End With
        With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
            .Caption = "&Developer"
            .TooltipText = "Set the mode to Developer"
            .OnAction = "ToolbarAction.ModeButton"
            .Parameter = "Developer"
            .BeginGroup = False
        End With
    End With
    Set mmnuMode = mtlb.Controls("Mode")
    Set mcmdModeOperator = mmnuMode.Controls("Operator")
    Set mcmdModeEditor = mmnuMode.Controls("Editor")
    Set mcmdModeDeveloper = mmnuMode.Controls("Developer")
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub Control_Navigator_Add _
( _
)
    With mtlb.Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
        .Style = Office.msoButtonCaption
        .Caption = "Navigator"
        .TooltipText = "Launch the Navigator"
        .OnAction = "ToolbarAction.Navigator"
        .BeginGroup = True
        .Width = 72
    End With
    Set mcmdNavigator = mtlb.Controls("Navigator")
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub Control_SongEdit_Add _
( _
)
    Dim lng As Long
    
    With mtlb.Controls.Add(Type:=Office.msoControlPopup, Temporary:=True)
        .Caption = "Song Edit"
        .TooltipText = "View the 'Song Edit' menu items"
        .BeginGroup = True
        .Width = 72
        With .Controls.Add(Type:=Office.msoControlPopup, Temporary:=True)
            .Caption = "Set Category"
            .TooltipText = "Set the category of the selected slides"
            .BeginGroup = False
            For lng = 1 To mudtApplication.Options.CategoryList.Count Step 1
                With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
                    .Caption = mudtApplication.Options.CategoryList.Item(lng)
                    .OnAction = "ToolbarAction.SongEditSetCategoryButton"
                    .Parameter = mudtApplication.Options.CategoryList.Item(lng)
                    .BeginGroup = False
                End With
            Next
            If (.Controls.Count >= 2) Then
                .Controls(2).BeginGroup = True
            End If
        End With
        With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
            .Caption = "Set Template"
            .TooltipText = "Set Song File Template"
            .OnAction = "ToolbarAction.SongEditSetTemplate"
            .BeginGroup = False
        End With
        With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
            .Caption = "Sort"
            .TooltipText = "Sort slides aphabetically"
            .OnAction = "ToolbarAction.SongEditSort"
            .BeginGroup = False
        End With
        With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
            .Caption = "Create Index"
            .TooltipText = "Generate slide index"
            .OnAction = "ToolbarAction.SongEditIndex"
            .BeginGroup = False
        End With
    End With
    Set mmnuSongEdit = mtlb.Controls("Song Edit")
    Set mmnuSongEditSetCategory = mmnuSongEdit.Controls("Set Category")
    Set mcmdSongEditSetTemplate = mmnuSongEdit.Controls("Set Template")
    Set mcmdSongEditSort = mmnuSongEdit.Controls("Sort")
    Set mcmdSongEditCreateIndex = mmnuSongEdit.Controls("Create Index")
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub Control_Category_Add _
( _
)
    With mtlb.Controls.Add(Type:=Office.msoControlEdit, Temporary:=True)
        .Caption = "Category"
        .Text = ""
        .TooltipText = "Slide Category"
        .BeginGroup = True
        .Width = 72
    End With
    Set mcboCategory = mtlb.Controls("Category")
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub Control_Debug_Add _
( _
)
    Dim lng As Long
    
    With mtlb.Controls.Add(Type:=Office.msoControlPopup, Temporary:=True)
        .Caption = "Debug"
        .TooltipText = "View the 'Debug' menu items"
        .BeginGroup = True
        .Width = 72
        With .Controls.Add(Type:=Office.msoControlPopup, Temporary:=True)
            .Caption = "SSW Monitor"
            .TooltipText = "Set the slide show window monitor"
            .BeginGroup = False
            With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
                .Caption = "<Default>"
                .OnAction = "ToolbarAction.DebugSSWMonitorButton"
                .BeginGroup = False
                .Parameter = mudtApplication.Options.DisplayMonitorList.Item(mudtApplication.Options.DisplayMonitorList.Count)
            End With
            For lng = 1 To mudtApplication.Options.DisplayMonitorList.Count Step 1
                With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
                    .Caption = mudtApplication.Options.DisplayMonitorList.Item(lng)
                    .OnAction = "ToolbarAction.DebugSSWMonitorButton"
                    .BeginGroup = False
                    .Parameter = mudtApplication.Options.DisplayMonitorList.Item(lng)
                End With
            Next
            If (.Controls.Count >= 2) Then
                .Controls(2).BeginGroup = True
            End If
        End With
        With .Controls.Add(Type:=Office.msoControlPopup, Temporary:=True)
            .Caption = "SSW Size"
            .TooltipText = "Set the slide show window size"
            .BeginGroup = False
            With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
                .Caption = "<Default>"
                .OnAction = "ToolbarAction.DebugSSWSizeButton"
                .BeginGroup = False
                .State = Office.msoButtonUp
                .Parameter = VBA.CStr(1)
            End With
            lng = 1
            Do While (lng <= 8)
                With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
                    .Caption = "1/" & VBA.CStr(lng * lng)
                    .OnAction = "ToolbarAction.DebugSSWSizeButton"
                    .BeginGroup = False
                    .State = Office.msoButtonUp
                    .Parameter = VBA.CStr(lng)
                End With
                lng = lng * 2
            Loop
            If (.Controls.Count >= 2) Then
                .Controls(2).BeginGroup = True
            End If
        End With
    End With
    Set mmnuDebug = mtlb.Controls("Debug")
    Set mmnuDebugSSWMonitor = mmnuDebug.Controls("SSW Monitor")
    Set mmnuDebugSSWSize = mmnuDebug.Controls("SSW Size")
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub Control_Help_Add _
( _
)
    With mtlb.Controls.Add(Type:=Office.msoControlPopup, Temporary:=True)
        .Caption = "Help"
        .TooltipText = "View the 'Help' menu items"
        .BeginGroup = True
        .Width = 72
        Dim mnu As Control
        With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
            .Caption = "&Help"
            .TooltipText = "View the '" & WorshipServiceAssistant.wsaApplicationNamePretty & "' help"
            .OnAction = "ToolbarAction.HelpTopicApplication"
            .BeginGroup = False
        End With
        With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
            .Caption = "Help on How To ..."
            .TooltipText = "View the '" & WorshipServiceAssistant.wsaApplicationNamePretty & "' How To ... help"
            .OnAction = "ToolbarAction.HelpTopicHowTo"
            .BeginGroup = False
        End With
        With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
            .Caption = "Help on Toolbar"
            .TooltipText = "View the '" & WorshipServiceAssistant.wsaApplicationNamePretty & "' Toolbar help"
            .OnAction = "ToolbarAction.HelpTopicToolbar"
            .BeginGroup = False
        End With
        With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
            .Caption = "Help on Known Issues"
            .TooltipText = "View the '" & WorshipServiceAssistant.wsaApplicationNamePretty & "' known issues"
            .OnAction = "ToolbarAction.HelpTopicKnownIssues"
            .BeginGroup = False
        End With
        With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
            .Caption = "Help on Copyright and Permisison"
            .TooltipText = "View the '" & WorshipServiceAssistant.wsaApplicationNamePretty & "' copyright and permission notice"
            .OnAction = "ToolbarAction.HelpTopicCopyrightPermission"
            .BeginGroup = False
        End With
        With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
            .Caption = "Visit the Homepage"
            .TooltipText = "Visit the '" & WorshipServiceAssistant.wsaApplicationNamePretty & "' homepage"
            .OnAction = "ToolbarAction.HelpVisitHomepage"
            .BeginGroup = True
        End With
        With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
            .Caption = "Email the Author"
            .TooltipText = "Email the '" & WorshipServiceAssistant.wsaApplicationNamePretty & "' author"
            .OnAction = "ToolbarAction.HelpEmailAuthor"
            .BeginGroup = False
        End With
        With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
            .Caption = "&About"
            .TooltipText = "View the '" & WorshipServiceAssistant.wsaApplicationNamePretty & "' about box"
            .OnAction = "ToolbarAction.HelpAbout"
            .BeginGroup = True
        End With
    End With
    Set mmnuHelp = mtlb.Controls("Help")
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub Toolbar_Update _
( _
)
    With mtlb
        With mmnuMode
            .Visible = _
                True
            If (.Visible) Then Control_Mode_Update
        End With
        With mcmdNavigator
            .Visible = _
                (mudtApplication.Mode = WorshipServiceAssistant.wsaModeOperator) Or _
                (mudtApplication.Mode = WorshipServiceAssistant.wsaModeEditor) Or _
                (mudtApplication.Mode = WorshipServiceAssistant.wsaModeDeveloper)
            If (.Visible) Then Control_Navigator_Update
        End With
        With mmnuSongEdit
            .Visible = _
                (mudtApplication.Mode = WorshipServiceAssistant.wsaModeEditor) Or _
                (mudtApplication.Mode = WorshipServiceAssistant.wsaModeDeveloper)
            If (.Visible) Then Control_SongEdit_Update
        End With
        With mcboCategory
            .Visible = _
                (mudtApplication.Mode = WorshipServiceAssistant.wsaModeEditor) Or _
                (mudtApplication.Mode = WorshipServiceAssistant.wsaModeDeveloper)
            If (.Visible) Then Control_Category_Update
        End With
        With mmnuDebug
            .Visible = _
                (mudtApplication.Mode = WorshipServiceAssistant.wsaModeDeveloper)
            If (.Visible) Then Control_Debug_Update
        End With
        With mmnuHelp
            .Visible = _
                True
            If (.Visible) Then Control_Help_Update
        End With
    End With
End Sub

Private Sub Control_Mode_Update _
( _
)
    With mmnuMode
        With mcmdModeOperator
            If (mudtApplication.Mode = WorshipServiceAssistant.wsaModeOperator) Then
                .State = Office.msoButtonDown
            Else
                .State = Office.msoButtonUp
            End If
        End With
        With mcmdModeEditor
            If (mudtApplication.Mode = WorshipServiceAssistant.wsaModeEditor) Then
                .State = Office.msoButtonDown
            Else
                .State = Office.msoButtonUp
            End If
        End With
        With mcmdModeDeveloper
            If (mudtApplication.Mode = WorshipServiceAssistant.wsaModeDeveloper) Then
                .State = Office.msoButtonDown
            Else
                .State = Office.msoButtonUp
            End If
        End With
    End With
End Sub

Private Sub Control_Navigator_Update _
( _
)
End Sub

Private Sub Control_SongEdit_Update _
( _
)
    Dim udtActiveWindow As WorshipServiceAssistant.DocumentWindow
    Dim blnHasWindow As Boolean
    Dim blnHasSlide As Boolean
    Dim blnHasSlideSelection As Boolean

    blnHasWindow = False
    blnHasSlide = False
    blnHasSlideSelection = False
    
    Set udtActiveWindow = mudtApplication.ActiveWindow
    If (Not udtActiveWindow Is Nothing) Then
        blnHasWindow = True
        If (udtActiveWindow.Presentation.Slides.Count > 0) Then
            blnHasSlide = True
        End If
        If (udtActiveWindow.ActiveSlideSelection.Count > 0) Then
            blnHasSlideSelection = True
        End If
    End If

    With mmnuSongEdit
        .Enabled = blnHasWindow And blnHasSlide
        With mmnuSongEditSetCategory
            .Enabled = blnHasWindow And blnHasSlideSelection
        End With
        With mcmdSongEditSetTemplate
            .Enabled = blnHasWindow
        End With
        With mcmdSongEditSort
            .Enabled = blnHasWindow And blnHasSlide
        End With
        With mcmdSongEditCreateIndex
            .Enabled = blnHasWindow And blnHasSlide
        End With
    End With
End Sub

Private Sub Control_Category_Update _
( _
)
    With mcboCategory
        .Enabled = True
        .Text = mudtApplication.ActiveSlideSelection.Category
        .Enabled = False
    End With
End Sub

Private Sub Control_Debug_Update _
( _
)
    Dim cmd As Office.CommandBarButton
    
    With mmnuDebug
        With mmnuDebugSSWMonitor
            For Each cmd In .Controls
                If (cmd.Parameter = mudtApplication.SlideShowWindowMonitor) Then
                    cmd.State = Office.msoButtonDown
                Else
                    cmd.State = Office.msoButtonUp
                End If
            Next
        End With
        With mmnuDebugSSWSize
            For Each cmd In .Controls
                If (cmd.Parameter = VBA.CStr(mudtApplication.SlideShowWindowSize)) Then
                    cmd.State = Office.msoButtonDown
                Else
                    cmd.State = Office.msoButtonUp
                End If
            Next
        End With
    End With
End Sub

Private Sub Control_Help_Update _
( _
)
End Sub


'===============================================================================
' Private Event Handlers.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub Class_Initialize _
( _
)
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub Class_Terminate _
( _
)
End Sub

