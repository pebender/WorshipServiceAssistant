Attribute VB_Name = "modToolbar"
'===============================================================================
' Name:
'   WorshipServiceAssistant.modToolbar
'
' Description:
'   This Module creates the 'Worship Service Assistant' toolbar with the
'   following controls:
'     'Navigator':
'       Clicking this control causes the Navigator to be displayed.
'     'Song Edit':
'       Clicking this control causes the 'Song Edit' menu to be displayed.
'       The special menu has functions for categorizing, sorting and indexing
'       slides.
'     'Category':
'       This control displays the slide category.
'     'Debug':
'       Clicking this control cause the 'Debug' menu to be displayed.
'     'Help':
'       Clicking this control causes the 'Help' menu to be displayed.
'
' Author:
'   Paul Bender <pbender@alumni.ucsd.edu>
'
' Copyright:
'   Copyright (c) 2000,2001,2002 Paul Bender
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
'   1.04.0001:
'     (1) Changed the name of the "Debug" menu's menu "SSW Display" to
'         "SSW Monitor".
'   1.04.0000:
'     (1) Removed the checks from the pop-up menu actions, because deleting
'         the pop-up menu from inside the action causes PowerPoint 2002 to
'         crash.
'   1.03.0002:
'     (1) Made changes to the source code so that it follows Microsoft's
'         Visual Basic coding conventions.
'   1.02.0000:
'     (1) Fixed bug that caused the "Debug" "Display" "Default" menu item
'         to set to default the slide show's size rather than the slide show's
'         display.
'   1.01.0007:
'     (1) Added 'Help on Known Issues' to the 'Help' menu.
'   1.01.0001:
'     (1) Changed 'Help' menu to merge 'Help on Copyright' and 'Help on License'
'         menus into a single 'Help on Copyright and Permission' menu.
'   1.01.0000:
'     (1) Added support for banner display.
'     (2) Added "Children" and "Liturgy" categories.
'     (3) Removed "Load" and "Hide" controls.
'   1.00.0002:
'     (1) Added a check to make sure the menu was refreshed before processing
'         menu controls.
'   1.00.0001:
'     (1) Reworked "Song Edit" menu's "Set Category" menu
'         to make it more flexible.
'     (2) Reworked "Debug" menu's "SSW Display" and "SSW Size" menus
'         to make them more flexible.
'     (3) Changed the "Help" menu's "Visit the Homepage" and "Email the Author"
'         menu items to use the ShellExecute Win32 API function call.
'     (4) Changed "Help" menu help topics so that they use the HTML Help
'         control API rather than the PowerPoint help interface.
'   1.00.0000:
'     Initial revision.
'===============================================================================


'===============================================================================
' Options.
'===============================================================================
Option Private Module
Option Explicit
Option Compare Text
Option Base 0


'===============================================================================
' Public Constants.
'===============================================================================


'===============================================================================
' Public Variables.
'===============================================================================


'===============================================================================
' Private Constants.
'===============================================================================
Private Const MstrToolbarName As String = "Worship Service Assistant"


'===============================================================================
' Private Variables.
'===============================================================================

' These variables are direct pointers to the toolbar, the menus and the buttons.
Private mtlbWSA                     As Office.CommandBar
Private mcmdWSANavigator            As Office.CommandBarButton
Private mmnuWSASongEdit             As Office.CommandBarPopup
Private mmnuWSASongEditSetCategory  As Office.CommandBarPopup
Private mcmdWSASongEditSort         As Office.CommandBarButton
Private mcmdWSASongEditCreateIndex  As Office.CommandBarButton
Private mcboWSACategory             As Office.CommandBarComboBox
Private mmnuWSADebug                As Office.CommandBarPopup
Private mmnuWSADebugSSWMonitor      As Office.CommandBarPopup
Private mmnuWSADebugSSWSize         As Office.CommandBarPopup
Private mmnuWSAHelp                 As Office.CommandBarPopup
Private mcmdWSAHelpDebug            As Office.CommandBarButton


'===============================================================================
' Public Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
'   Unistalls the command bar.
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gUnload _
( _
)
    Dim tlbTemp As Office.CommandBar
    Dim ctrTemp As Office.CommandBarControl
    
    '
    ' Uninstall any "Worship Service Assistant" command bars.
    '
    For Each tlbTemp In Application.CommandBars
        If (tlbTemp.BuiltIn = False) Then
            If (VBA.Left(tlbTemp.Name, VBA.Len(MstrToolbarName)) = MstrToolbarName) Then
                tlbTemp.Delete
            End If
        End If
    Next
    For Each tlbTemp In Application.CommandBars
        If (tlbTemp.BuiltIn = False) Then
            For Each ctrTemp In tlbTemp.Controls
                If (ctrTemp.BuiltIn = False) Then
                    If (VBA.Left(ctrTemp.Tag, VBA.Len(MstrToolbarName)) = MstrToolbarName) Then
                        ctrTemp.Delete
                    End If
                End If
            Next
        End If
    Next
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Installs the command bar.
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gLoad _
( _
)
    Dim tlbTemp As Office.CommandBar
    Dim lngBarRowIndex As Long
    
    ' Uninstall any pre-existing "Worship Service Assistant" command bars.
    modToolbar.gUnload
    
    ' Find the end of command bar list.  The "Worship Service Assistant" will
    ' be placed at the end of the command bar list.
    For Each tlbTemp In Application.CommandBars
        If (tlbTemp.RowIndex > lngBarRowIndex) Then
            lngBarRowIndex = tlbTemp.RowIndex
        End If
    Next
    lngBarRowIndex = lngBarRowIndex + 1
    
    ' Install the 'Worship Service Assistant' command bar.
    Set mtlbWSA = Application.CommandBars.Add( _
        Temporary:=True)
    
    ' Configure the 'Worship Service Assistant' command bar.
    With mtlbWSA
        .Name = MstrToolbarName
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
        
    ' Install the 'Worship Service Assistant' command bar items.
    NavigatorControlAdd
    SongEditControlAdd
    CategoryControlAdd
    DebugControlAdd
    HelpControlAdd
    
    modToolbar.gRefresh
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gDisable _
( _
)
    Dim lngIndex As Long
    
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    mcboWSACategory.Enabled = True
    mcboWSACategory.Text = ""
    mcboWSACategory.Enabled = False
    
    ' Disable all controls.
    mcmdWSANavigator.Enabled = False
    mmnuWSASongEdit.Enabled = False
    mmnuWSASongEditSetCategory.Enabled = False
    mcmdWSASongEditSort.Enabled = True
    mcmdWSASongEditCreateIndex.Enabled = False
    mcboWSACategory.Enabled = False
    mmnuWSAHelp.Enabled = False
    mcmdWSAHelpDebug.Enabled = False
    
    mcmdWSAHelpDebug.State = Office.msoButtonUp
    
    If (mmnuWSADebug.Visible = True) Then
        mmnuWSADebug.Enabled = False
        mmnuWSADebugSSWMonitor.Enabled = False
        mmnuWSADebugSSWSize.Enabled = False
        For lngIndex = 1 To mmnuWSADebugSSWMonitor.Controls.Count Step 1
            mmnuWSADebugSSWMonitor.Controls(lngIndex).State = Office.msoButtonUp
        Next
        For lngIndex = 1 To mmnuWSADebugSSWSize.Controls.Count Step 1
            mmnuWSADebugSSWSize.Controls(lngIndex).State = Office.msoButtonUp
        Next
    End If
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gRefresh _
( _
)
    Dim dwDocumentWindow As PowerPoint.DocumentWindow
    
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    modToolbar.gDisable
    
    ' By default, enable all visible controls (except Category).
    mtlbWSA.Enabled = True
    mcmdWSANavigator.Enabled = True
    mmnuWSASongEdit.Enabled = True
    mmnuWSASongEditSetCategory.Enabled = True
    mcmdWSASongEditSort.Enabled = True
    mcmdWSASongEditCreateIndex.Enabled = True
    mcboWSACategory.Enabled = False
    If (mmnuWSADebug.Visible = True) Then
        mmnuWSADebug.Enabled = True
        mmnuWSADebugSSWMonitor.Enabled = True
        mmnuWSADebugSSWSize.Enabled = True
    End If
    mmnuWSAHelp.Enabled = True
    mcmdWSAHelpDebug.Enabled = True
    
    If (modActive.gblnActiveWindowExists = False) Then
        mmnuWSASongEdit.Enabled = False
        mmnuWSASongEditSetCategory.Enabled = False
        mcmdWSASongEditSort.Enabled = False
        mcmdWSASongEditCreateIndex.Enabled = False
        
        mcboWSACategory.Enabled = True
        mcboWSACategory.Text = ""
        mcboWSACategory.Enabled = False
    Else
        Set dwDocumentWindow = Application.ActiveWindow
        
        If (modActive.gblnActiveWindowSlideExists(dwDocumentWindow) = False) Then
            mmnuWSASongEdit.Enabled = False
            mmnuWSASongEditSetCategory.Enabled = False
            mcmdWSASongEditSort.Enabled = False
            mcmdWSASongEditCreateIndex.Enabled = False
        End If
        
        If (modActive.gblnActiveSlideExists(dwDocumentWindow) = False) Then
            mmnuWSASongEditSetCategory.Enabled = False
            mcmdWSASongEditSort.Enabled = False
        End If
        
        mcboWSACategory.Enabled = True
        mcboWSACategory.Text = CategoryGet(dwDocumentWindow)
        mcboWSACategory.Enabled = False
    End If
    
    If (mmnuWSADebug.Visible = True) Then
        mcmdWSAHelpDebug.State = Office.msoButtonDown
        
        mmnuWSADebug.Enabled = True
        mmnuWSADebugSSWMonitor.Enabled = True
        mmnuWSADebugSSWSize.Enabled = True
        
        mmnuWSADebugSSWMonitor.Controls(modProject.glngSlideShowWindowMonitor + 1).State = Office.msoButtonDown
        mmnuWSADebugSSWSize.Controls(modProject.glngSlideShowWindowSize + 1).State = Office.msoButtonDown
    End If
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Launchs the slide navigator. In order to exit, the slide navigator must be
'   unloaded. If the slide navigator is only hidden, then this routine will
'   automatically refresh it and re-show it.  This is a hack to work around
'   some focus problems resulting from activating slide shows and activating
'   new presentations.
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gOnActionNavigator _
( _
)
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    ' Update the menu bar.
    modToolbar.gRefresh
    
    ' Exit if the control is not enabled.
    If (CommandBars.ActionControl.Enabled = False) Then
        Exit Sub
    End If
    
    modNavigator.gRun
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gOnActionSongEdit _
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
Public Sub gOnActionSongEditSetCategory _
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
Public Sub gOnActionSongEditSetCategoryButton _
( _
)
    Dim dwDocumentWindow As PowerPoint.DocumentWindow
    
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    If (modActive.gblnActiveWindowExists = False) Then
        Exit Sub
    End If
    
    Set dwDocumentWindow = Application.ActiveWindow
    
    CategorySet dwDocumentWindow, CommandBars.ActionControl.Caption
    modToolbar.gRefresh
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gOnActionSongEditSort _
( _
)
    Dim dwDocumentWindow As PowerPoint.DocumentWindow
    
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    If (modActive.gblnActiveWindowExists = False) Then
        Exit Sub
    End If
    
    Set dwDocumentWindow = Application.ActiveWindow
    
    modSort.gRun dwDocumentWindow
    
    VBA.MsgBox _
        buttons:= _
            VBA.vbInformation, _
        Title:= _
            modProject.GstrNamePretty, _
        Prompt:= _
            "Slide sort is complete."
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gOnActionSongEditIndex _
( _
)
    Dim dwDocumentWindow As PowerPoint.DocumentWindow
    
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    If (modActive.gblnActiveWindowExists = False) Then
        Exit Sub
    End If
    
    Set dwDocumentWindow = Application.ActiveWindow
    
    modIndex.gRun dwDocumentWindow
    
    VBA.MsgBox _
        buttons:= _
            VBA.vbInformation, _
        Title:= _
            modProject.GstrNamePretty, _
        Prompt:= _
            "Slide index generation is complete."
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gOnActionDebug _
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
Public Sub gOnActionDebugSSWMonitor _
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
Public Sub gOnActionDebugSSWMonitorButton _
( _
)
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    modProject.glngSlideShowWindowMonitor = CommandBars.ActionControl.Index - 1
    modToolbar.gRefresh
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gOnActionDebugSSWSize _
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
Public Sub gOnActionDebugSSWSizeButton _
( _
)
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    modProject.glngSlideShowWindowSize = CommandBars.ActionControl.Index - 1
    modToolbar.gRefresh
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gOnActionHelp _
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
Public Sub gOnActionHelpHelp _
( _
)
    Dim strHelpFile As String
    
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    strHelpFile = modHelp.gstrFileNameGet(True)
    
    If (strHelpFile = "") Then
        Exit Sub
     End If
    
    Call modHelp.HtmlHelp( _
        0&, _
        strHelpFile, _
        modHelp.HH_DISPLAY_TOPIC, _
        modHelp.GstrIDH_TopicPath_WSA)
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gOnActionHelpHelpHowTo _
( _
)
    Dim strHelpFile As String
    
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    strHelpFile = modHelp.gstrFileNameGet(True)
    
    If (strHelpFile = "") Then
        Exit Sub
     End If
        
    Call modHelp.HtmlHelp( _
        0&, _
        strHelpFile, _
        modHelp.HH_DISPLAY_TOPIC, _
        modHelp.GstrIDH_TopicPath_WSAHowTo)
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gOnActionHelpHelpCommandBar _
( _
)
    Dim strHelpFile As String
    
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    strHelpFile = modHelp.gstrFileNameGet(True)
    
    If (strHelpFile = "") Then
        Exit Sub
     End If
    
    Call modHelp.HtmlHelp( _
        0&, _
        strHelpFile, _
        modHelp.HH_DISPLAY_TOPIC, _
        modHelp.GstrIDH_TopicPath_WSACommandBar)
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gOnActionHelpHelpKnownIssues _
( _
)
    Dim strHelpFile As String
    
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    strHelpFile = modHelp.gstrFileNameGet(True)
    
    If (strHelpFile = "") Then
        Exit Sub
     End If
    
    Call modHelp.HtmlHelp( _
        0&, _
        strHelpFile, _
        modHelp.HH_DISPLAY_TOPIC, _
        modHelp.GstrIDH_TopicPath_WSAKnownIssues)
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gOnActionHelpHelpCopyrightPermission _
( _
)
    Dim strHelpFile As String
    
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    strHelpFile = modHelp.gstrFileNameGet(True)
    
    If (strHelpFile = "") Then
        Exit Sub
     End If
    
    Call modHelp.HtmlHelp( _
        0&, _
        strHelpFile, _
        modHelp.HH_DISPLAY_TOPIC, _
        modHelp.GstrIDH_TopicPath_WSACopyrightPermission)
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gOnActionHelpVisitHomepage _
( _
)
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    modWin32Shell32.ShellExecute _
        0&, _
        VBA.vbNullString, _
        modProject.GstrHomepage, _
        VBA.vbNullString, _
        VBA.vbNullString, _
        modWin32Shell32.SW_SHOWNORMAL
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gOnActionHelpEmailAuthor _
( _
)
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    modWin32Shell32.ShellExecute _
        0&, _
        VBA.vbNullString, _
        modProject.GstrEmail & "?subject=" & modProject.GstrName & "%20" & modProject.GstrVersion & ":%20", _
        VBA.vbNullString, _
        VBA.vbNullString, _
        modWin32Shell32.SW_SHOWNORMAL
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gOnActionHelpDebug _
( _
)
    Dim intResponse As VBA.VbMsgBoxResult
    
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    If (mmnuWSADebug.Visible = False) Then
        intResponse = VBA.MsgBox( _
            buttons:= _
                VBA.vbYesNo + VBA.vbDefaultButton2 + VBA.vbExclamation, _
            Title:= _
                modProject.GstrNamePretty, _
            Prompt:= _
                "Enabling the 'Debug' menu can cause " & _
                "Worship Service Assistant to function incorrectly. " & _
                VBA.vbCrLf & _
                VBA.vbCrLf & _
                "Are you sure you want to enable the 'Debug' menu?")
        If (intResponse = VBA.vbNo) Then
            Exit Sub
        End If
    End If
    
    mmnuWSADebug.Visible = Not mmnuWSADebug.Visible
    
    modProject.glngSlideShowWindowMonitor = 0
    modProject.glngSlideShowWindowSize = 0
    modToolbar.gRefresh
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gOnActionHelpAbout _
( _
)
    Dim strPrompt As String
    
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    strPrompt = _
        modProject.GstrNamePretty & " (version " & modProject.GstrVersion & ")" & _
        VBA.vbCrLf & VBA.vbCrLf & _
        modProject.GstrCopyright & _
        VBA.vbCrLf & VBA.vbCrLf & _
        "Worship Service Assistant is a Microsoft PowerPoint add-in " & _
        "that is designed to make PowerPoint a more useful tool in a " & _
        "worship service. " & _
        VBA.vbCrLf & VBA.vbCrLf & _
        modProject.GstrNamePretty & " comes with ABSOLUTELY NO WARRANTY; " & _
        "for details view the 'Help on Copyright and Permission' in the 'Help' menu. " & _
        "This is free software, " & _
        "and you are welcome to redistribute it under certain conditions; " & _
        "for details view the 'Help on Copyright and Permission' in the 'Help' menu. "

    VBA.MsgBox _
        buttons:= _
            VBA.vbOKOnly + VBA.vbMsgBoxHelpButton, _
        Title:= _
            "About " & modProject.GstrNamePretty, _
        Prompt:= _
            strPrompt
End Sub


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
'   Installs the 'Navigator' control on the command bar specified by the
'   variable mtlbWSA.
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub NavigatorControlAdd _
( _
)
    ' Install the 'Navigator' control.
    Set mcmdWSANavigator = mtlbWSA.Controls.Add( _
        Type:=Office.msoControlButton, _
        Temporary:=True)
    
    ' Configure the 'Navigator' control.
    With mcmdWSANavigator
        .Style = Office.msoButtonCaption
        .Caption = "Navigator"
        .TooltipText = "Launch the Navigator"
        .OnAction = "modToolbar.gOnActionNavigator"
        .BeginGroup = True
        .Width = 72
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Installs the 'SongEdit' control on the command bar specified by the
'   variable mtlbWSA.
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub SongEditControlAdd _
( _
)
    Dim cmdMenuItem As Office.CommandBarButton
    Dim lngCategory As Long
    
    ' Install the 'SongEdit' control.
    Set mmnuWSASongEdit = mtlbWSA.Controls.Add( _
        Type:=Office.msoControlPopup, _
        Temporary:=True)
    
    ' Configure the 'SongEdit' control.
    With mmnuWSASongEdit
        .Caption = "Song Edit"
        .TooltipText = "View the 'Song Edit' menu items"
        .OnAction = "modToolbar.gOnActionSongEdit"
        .BeginGroup = True
        .Width = 72
        
        Set mmnuWSASongEditSetCategory = .Controls.Add( _
            Type:=Office.msoControlPopup, _
            Temporary:=True)
        mmnuWSASongEditSetCategory.Caption = "Set Category"
        mmnuWSASongEditSetCategory.TooltipText = "Set the category of the selected slides"
        mmnuWSASongEditSetCategory.BeginGroup = False
        
        Set mcmdWSASongEditSort = .Controls.Add( _
            Type:=Office.msoControlButton, _
            Temporary:=True)
        mcmdWSASongEditSort.Caption = "Sort"
        mcmdWSASongEditSort.TooltipText = "Sort slides aphabetically"
        mcmdWSASongEditSort.OnAction = "modToolbar.gOnActionSongEditSort"
        mcmdWSASongEditSort.BeginGroup = False
        
        Set mcmdWSASongEditCreateIndex = .Controls.Add( _
            Type:=Office.msoControlButton, _
            Temporary:=True)
        mcmdWSASongEditCreateIndex.Caption = "Create Index"
        mcmdWSASongEditCreateIndex.TooltipText = "Generate slide index"
        mcmdWSASongEditCreateIndex.OnAction = "modToolbar.gOnActionSongEditIndex"
        mcmdWSASongEditCreateIndex.BeginGroup = False
    End With
    
    ReDim modProject.gastrCategories(5)
    modProject.gastrCategories(0) = "Worship"
    modProject.gastrCategories(1) = "Choir"
    modProject.gastrCategories(2) = "Hymn"
    modProject.gastrCategories(3) = "Carol"
    modProject.gastrCategories(4) = "Children"
    modProject.gastrCategories(5) = "Liturgy"
    With mmnuWSASongEditSetCategory
        Set cmdMenuItem = .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
        cmdMenuItem.Caption = "<none>"
        cmdMenuItem.TooltipText = "Set slide category to '<none>'"
        cmdMenuItem.OnAction = "modToolbar.gOnActionSongEditSetCategoryButton"
        cmdMenuItem.BeginGroup = False
    
        For lngCategory = LBound(modProject.gastrCategories) To UBound(modProject.gastrCategories) Step 1
            Set cmdMenuItem = .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
            cmdMenuItem.Caption = modProject.gastrCategories(lngCategory)
            cmdMenuItem.OnAction = "modToolbar.gOnActionSongEditSetCategoryButton"
            cmdMenuItem.BeginGroup = False
        Next
        If (.Controls.Count >= 2) Then
            .Controls(2).BeginGroup = True
        End If
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Installs the 'Category' control on the command bar specified by the
'   variable mtlbWSA.
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub CategoryControlAdd _
( _
)

    ' Install the 'Category' control.
    Set mcboWSACategory = mtlbWSA.Controls.Add( _
        Type:=Office.msoControlEdit, _
        Temporary:=True)
    
    ' Configure the 'Category' control.
    With mcboWSACategory
        .Caption = "Category"
        .Text = ""
        .TooltipText = "Slide Category"
        .BeginGroup = True
        .Width = 72
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Installs the 'Debug' control on the command bar specified by the
'   variable mtlbWSA.
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub DebugControlAdd _
( _
)
    Dim lngMonitor As Long
    Dim lngSize As Long
    Dim cmdMenuItem As Office.CommandBarButton
    
    ' Install the 'Debug' control.
    Set mmnuWSADebug = mtlbWSA.Controls.Add( _
        Type:=Office.msoControlPopup, _
        Temporary:=True)
    
    ' Configure the 'Debug' control.
    With mmnuWSADebug
        .Caption = "Debug"
        .TooltipText = "View the 'Debug' menu items"
        .OnAction = "modToolbar.gOnActionDebug"
        .BeginGroup = True
        .Width = 72
        
        Set mmnuWSADebugSSWMonitor = .Controls.Add( _
            Type:=Office.msoControlPopup, _
            Temporary:=True)
        mmnuWSADebugSSWMonitor.Caption = "SSW Monitor"
        mmnuWSADebugSSWMonitor.TooltipText = "Set the slide show window monitor"
        mmnuWSADebugSSWMonitor.BeginGroup = False
        
        Set mmnuWSADebugSSWSize = .Controls.Add( _
            Type:=Office.msoControlPopup, _
            Temporary:=True)
        mmnuWSADebugSSWSize.Caption = "SSW Size"
        mmnuWSADebugSSWSize.TooltipText = "Set the slide show window size"
        mmnuWSADebugSSWSize.BeginGroup = False
        
    End With
    
    With mmnuWSADebugSSWMonitor
        Set cmdMenuItem = .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
        cmdMenuItem.Caption = "Default"
        cmdMenuItem.OnAction = "modToolbar.gOnActionDebugSSWMonitorButton"
        cmdMenuItem.BeginGroup = False
        cmdMenuItem.State = Office.msoButtonUp
        lngMonitor = 1
        While lngMonitor <= modWin32User32.GetSystemMetrics(modWin32User32.SM_CMONITORS)
            Set cmdMenuItem = .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
            cmdMenuItem.Caption = "Monitor " & lngMonitor
            cmdMenuItem.OnAction = "modToolbar.gOnActionDebugSSWMonitorButton"
            cmdMenuItem.BeginGroup = False
            cmdMenuItem.State = Office.msoButtonUp
            lngMonitor = lngMonitor + 1
        Wend
        If (.Controls.Count >= 2) Then
            .Controls(2).BeginGroup = True
        End If
    End With
    
    With mmnuWSADebugSSWSize
        Set cmdMenuItem = .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
        cmdMenuItem.Caption = "Default"
        cmdMenuItem.OnAction = "modToolbar.gOnActionDebugSSWSizeButton"
        cmdMenuItem.BeginGroup = False
        cmdMenuItem.State = Office.msoButtonUp
        lngSize = 1
        While lngSize <= 64
            Set cmdMenuItem = .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
            cmdMenuItem.Caption = "1/" & lngSize
            cmdMenuItem.OnAction = "modToolbar.gOnActionDebugSSWSizeButton"
            cmdMenuItem.BeginGroup = False
            cmdMenuItem.State = Office.msoButtonUp
            lngSize = lngSize * 2
        Wend
        If (.Controls.Count >= 2) Then
            .Controls(2).BeginGroup = True
        End If
    End With
    
    mmnuWSADebug.Visible = False
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Installs the 'Help' control on the command bar specified by the
'   variable mtlbWSA.
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function HelpControlAdd _
( _
) As Boolean
    Dim cmdMenuItem As Office.CommandBarButton
    
    ' Install the 'Help' control.
    Set mmnuWSAHelp = mtlbWSA.Controls.Add( _
        Type:=Office.msoControlPopup, _
        Temporary:=True)
    
    ' Configure the 'Help' control.
    With mmnuWSAHelp
        .Caption = "Help"
        .TooltipText = "View the 'Help' menu items"
        .OnAction = "modToolbar.gOnActionHelp"
        .BeginGroup = True
        .Width = 72
        
        Set cmdMenuItem = .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
        cmdMenuItem.Caption = "&Help"
        cmdMenuItem.TooltipText = "View the '" & modProject.GstrNamePretty & "' help"
        cmdMenuItem.OnAction = "modToolbar.gOnActionHelpHelp"
        cmdMenuItem.BeginGroup = False
        
        Set cmdMenuItem = .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
        cmdMenuItem.Caption = "Help on How To ..."
        cmdMenuItem.TooltipText = "View the '" & modProject.GstrNamePretty & "' How To ... help"
        cmdMenuItem.OnAction = "modToolbar.gOnActionHelpHelpHowTo"
        cmdMenuItem.BeginGroup = False
        
        Set cmdMenuItem = .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
        cmdMenuItem.Caption = "Help on Command Bar"
        cmdMenuItem.TooltipText = "View the '" & modProject.GstrNamePretty & "' Command Bar help"
        cmdMenuItem.OnAction = "modToolbar.gOnActionHelpHelpCommandBar"
        cmdMenuItem.BeginGroup = False
        
        Set cmdMenuItem = .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
        cmdMenuItem.Caption = "Help on Known Issues"
        cmdMenuItem.TooltipText = "View the '" & modProject.GstrNamePretty & "' known issues"
        cmdMenuItem.OnAction = "modToolbar.gOnActionHelpHelpKnownIssues"
        cmdMenuItem.BeginGroup = False
        
        Set cmdMenuItem = .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
        cmdMenuItem.Caption = "Help on Copyright and Permisison"
        cmdMenuItem.TooltipText = "View the '" & modProject.GstrNamePretty & "' copyright and permission notice"
        cmdMenuItem.OnAction = "modToolbar.gOnActionHelpHelpCopyrightPermission"
        cmdMenuItem.BeginGroup = False
        
        Set cmdMenuItem = .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
        cmdMenuItem.Caption = "Visit the Homepage"
        cmdMenuItem.TooltipText = "Visit the '" & modProject.GstrNamePretty & "' homepage"
        cmdMenuItem.OnAction = "modToolbar.gOnActionHelpVisitHomepage"
        cmdMenuItem.BeginGroup = True
               
        Set cmdMenuItem = .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
        cmdMenuItem.Caption = "Email the Author"
        cmdMenuItem.TooltipText = "Email the '" & modProject.GstrNamePretty & "' author"
        cmdMenuItem.OnAction = "modToolbar.gOnActionHelpEmailAuthor"
        cmdMenuItem.BeginGroup = False
               
        Set mcmdWSAHelpDebug = .Controls.Add( _
            Type:=Office.msoControlButton, _
            Temporary:=True)
        mcmdWSAHelpDebug.Caption = "Debug"
        mcmdWSAHelpDebug.TooltipText = "Show or hide the Debug menu"
        mcmdWSAHelpDebug.OnAction = "modToolbar.gOnActionHelpDebug"
        mcmdWSAHelpDebug.BeginGroup = True
        mcmdWSAHelpDebug.State = Office.msoButtonUp
        
        Set cmdMenuItem = .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
        cmdMenuItem.Caption = "&About"
        cmdMenuItem.TooltipText = "View the '" & modProject.GstrNamePretty & "' about box"
        cmdMenuItem.OnAction = "modToolbar.gOnActionHelpAbout"
        cmdMenuItem.BeginGroup = True
    End With
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub CategorySet _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow, _
    ByRef Category As String _
)
    Dim sldSlide As PowerPoint.Slide
    
    If ((modActive.gblnActiveSelectionExists(dwDocumentWindow) = False) And _
        (modActive.gblnActiveSlideExists(dwDocumentWindow) = True)) Then
        Set sldSlide = modActive.gsldActiveSlideGet(dwDocumentWindow)
        If (Category = "") Then
            sldSlide.Tags.Delete "Category"
        ElseIf (VBA.LCase(Category) = "<none>") Then
            sldSlide.Tags.Delete "Category"
        Else
            sldSlide.Tags.Add "Category", Category
        End If
    Else
        For Each sldSlide In dwDocumentWindow.Selection.SlideRange
            If (Category = "") Then
                sldSlide.Tags.Delete "Category"
            ElseIf (VBA.LCase(Category) = "<none>") Then
                sldSlide.Tags.Delete "Category"
            Else
                sldSlide.Tags.Add "Category", Category
            End If
        Next
    End If
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function CategoryGet _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
) As String
    Dim sldSlide As PowerPoint.Slide
    
    CategoryGet = ""
    
    If (modActive.gblnActiveWindowSlideExists(dwDocumentWindow) = False) Then
        CategoryGet = ""
    ElseIf (modActive.gblnActiveSelectionExists(dwDocumentWindow) = False) Then
        If (modActive.gblnActiveSlideExists(dwDocumentWindow) = True) Then
            CategoryGet = modActive.gsldActiveSlideGet(dwDocumentWindow).Tags("Category")
        Else
            CategoryGet = ""
        End If
    ElseIf (dwDocumentWindow.Selection.SlideRange.Count = 0) Then
        If (modActive.gblnActiveSlideExists(dwDocumentWindow) = True) Then
            CategoryGet = modActive.gsldActiveSlideGet(dwDocumentWindow).Tags("Category")
        Else
            CategoryGet = ""
        End If
    Else
        CategoryGet = dwDocumentWindow.Selection.SlideRange(1).Tags("Category")
        For Each sldSlide In dwDocumentWindow.Selection.SlideRange
            If (sldSlide.Tags("Category") <> CategoryGet) Then
                    CategoryGet = ""
                Exit For
            End If
        Next
    End If
End Function
