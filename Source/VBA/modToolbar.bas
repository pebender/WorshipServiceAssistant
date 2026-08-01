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
'   1.04.0003:
'     (1) Changed impacted gWSAHelp_OnAction* routines to use the
'         modHelp.gTopicShow, modHelp.gHomepageGoto and modHelp.gAuthorEmail
'         routines.
'     (2) Moved category routines to the modCategory module.
'     (3) Cleaned up the code.
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
' Private Constants.
'===============================================================================


'===============================================================================
' Public Variables.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================


'===============================================================================
' Public Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gUnload _
( _
)
    mWSA_Delete
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gLoad _
( _
)
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    mWSA_Add
    mWSA_Update
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
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    mWSA_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gWSANavigator_OnAction _
( _
)
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    modNavigator.gRun
    
    mWSA_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gWSASongEditSetCategoryButton_OnAction _
( _
)
    Dim dwCurrent As PowerPoint.DocumentWindow
    Dim sldCurrent As PowerPoint.SlideRange
    
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    Set dwCurrent = mValidWindow(mActiveWindow)
    Set sldCurrent = mSlideRange(dwCurrent)
    
    modCategory.CategorySet sldCurrent, CommandBars.ActionControl.Caption
    
    mWSA_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gWSASongEditSort_OnAction _
( _
)
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    modSort.gRun mValidWindow(mActiveWindow)
    
    VBA.MsgBox _
        buttons:= _
            VBA.vbInformation, _
        Title:= _
            modProject.GstrNamePretty, _
        Prompt:= _
            "Slide sort is complete."
            
    mWSA_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gWSASongEditIndex_OnAction _
( _
)
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    modIndex.gRun mValidWindow(mActiveWindow)
    
    VBA.MsgBox _
        buttons:= _
            VBA.vbInformation, _
        Title:= _
            modProject.GstrNamePretty, _
        Prompt:= _
            "Slide index generation is complete."
            
    mWSA_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gWSADebugSSWMonitorButton_OnAction _
( _
)
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    modProject.gstrSlideShowWindowMonitor = VBA.CStr(CommandBars.ActionControl.Parameter)
    
    mWSA_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gWSADebugSSWSizeButton_OnAction _
( _
)
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    modProject.glngSlideShowWindowSize = VBA.CStr(CommandBars.ActionControl.Parameter)
    
    mWSA_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gWSAHelpHelp_OnAction _
( _
)
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    modHelp.gTopicShow HelpTopic.WSA
    
    mWSA_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gWSAHelpHelpHowTo_OnAction _
( _
)
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    modHelp.gTopicShow HelpTopic.WSAHowTo
    
    mWSA_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gWSAHelpHelpCommandBar_OnAction _
( _
)
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    modHelp.gTopicShow HelpTopic.WSACommandBar
    
    mWSA_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gWSAHelpHelpKnownIssues_OnAction _
( _
)
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    modHelp.gTopicShow HelpTopic.WSAKnownIssues
    
    mWSA_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gWSAHelpHelpCopyrightPermission_OnAction _
( _
)
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    modHelp.gTopicShow HelpTopic.WSACopyrightPermission
    
    mWSA_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gWSAHelpVisitHomepage_OnAction _
( _
)
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    modHelp.gHomepageVisit
    
    mWSA_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gWSAHelpEmailAuthor_OnAction _
( _
)
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    modHelp.gAuthorEmail
    
    mWSA_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gWSAHelpDebug_OnAction _
( _
)
    Dim intResponse As VBA.VbMsgBoxResult
    
    ' Load the project if the project is not loaded
    If (modProject.gblnLoaded = False) Then
        modProject.gLoad
        Exit Sub
    End If
    
    If (modProject.gblnDebugEnabled = False) Then
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
    
    modProject.gblnDebugEnabled = Not modProject.gblnDebugEnabled
    With Application.CommandBars(modProject.GstrName).Controls("Debug")
        modProject.gstrSlideShowWindowMonitor = _
            .Controls("SSW Monitor").Controls("<Default>").Parameter
        modProject.glngSlideShowWindowSize = _
            VBA.CLng(.Controls("SSW Size").Controls("<Default>").Parameter)
    End With
    
    mWSA_Update
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gWSAHelpAbout_OnAction _
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
    
    mWSA_Update
End Sub


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub mWSA_Delete _
( _
)
    Dim tlb As Office.CommandBar
    
    ' Uninstall any "Worship Service Assistant" command bars.
    For Each tlb In Application.CommandBars
        If (tlb.BuiltIn = False) Then
            If (tlb.Name = modProject.GstrName) Then
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
Private Sub mWSA_Add _
( _
)
    Dim tlb As Office.CommandBar
    Dim lngBarRowIndex As Long
    
    ' Delete any pre-existing "Worship Service Assistant" command bars.
    mWSA_Delete
    
    ' Find the end of command bar list.  The "Worship Service Assistant" will
    ' be placed at the end of the command bar list.
    For Each tlb In Application.CommandBars
        If (tlb.RowIndex > lngBarRowIndex) Then
            lngBarRowIndex = tlb.RowIndex
        End If
    Next
    lngBarRowIndex = lngBarRowIndex + 1
    
    ' Install and configure the 'Worship Service Assistant' command bar.
    With Application.CommandBars.Add(Temporary:=True)
        .Name = modProject.GstrName
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
    mWSANavigator_Add
    mWSASongEdit_Add
    mWSACategory_Add
    mWSADebug_Add
    mWSAHelp_Add
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub mWSANavigator_Add _
( _
)
    With Application.CommandBars(modProject.GstrName)
        With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
            .Style = Office.msoButtonCaption
            .Caption = "Navigator"
            .TooltipText = "Launch the Navigator"
            .OnAction = "modToolbar.gWSANavigator_OnAction"
            .BeginGroup = True
            .Width = 72
        End With
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub mWSASongEdit_Add _
( _
)
    Dim lng As Long
    
    With Application.CommandBars(modProject.GstrName)
        With .Controls.Add(Type:=Office.msoControlPopup, Temporary:=True)
            .Caption = "Song Edit"
            .TooltipText = "View the 'Song Edit' menu items"
            .BeginGroup = True
            .Width = 72
            With .Controls.Add(Type:=Office.msoControlPopup, Temporary:=True)
                .Caption = "Set Category"
                .TooltipText = "Set the category of the selected slides"
                .BeginGroup = False
                For lng = 1 To modCategory.Count Step 1
                    With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
                        .Caption = modCategory.Item(lng)
                        .OnAction = "modToolbar.gWSASongEditSetCategoryButton_OnAction"
                        .BeginGroup = False
                    End With
                Next
                If (.Controls.Count >= 2) Then
                    .Controls(2).BeginGroup = True
                End If
            End With
            With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
                .Caption = "Sort"
                .TooltipText = "Sort slides aphabetically"
                .OnAction = "modToolbar.gWSASongEditSort_OnAction"
                .BeginGroup = False
            End With
            With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
                .Caption = "Create Index"
                .TooltipText = "Generate slide index"
                .OnAction = "modToolbar.gWSASongEditIndex_OnAction"
                .BeginGroup = False
            End With
        End With
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub mWSACategory_Add _
( _
)
    With Application.CommandBars(modProject.GstrName)
        With .Controls.Add(Type:=Office.msoControlEdit, Temporary:=True)
            .Caption = "Category"
            .Text = ""
            .TooltipText = "Slide Category"
            .BeginGroup = True
            .Width = 72
        End With
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub mWSADebug_Add _
( _
)
    Dim lng As Long
    Dim udtOptions As New WorshipServiceAssistant.WSAApplicationOptions
    Dim astrMonitorList() As String
    
    ' Get monitor names
    astrMonitorList = udtOptions.DisplayMonitorList
    
    With Application.CommandBars(modProject.GstrName)
        With .Controls.Add(Type:=Office.msoControlPopup, Temporary:=True)
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
                    .OnAction = "modToolbar.gWSADebugSSWMonitorButton_OnAction"
                    .BeginGroup = False
                    .Parameter = VBA.CStr(astrMonitorList(UBound(astrMonitorList)))
                End With
                For lng = LBound(astrMonitorList) To UBound(astrMonitorList) Step 1
                    With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
                        .Caption = astrMonitorList(lng)
                        .OnAction = "modToolbar.gWSADebugSSWMonitorButton_OnAction"
                        .BeginGroup = False
                        .Parameter = astrMonitorList(lng)
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
                    .OnAction = "modToolbar.gWSADebugSSWSizeButton_OnAction"
                    .BeginGroup = False
                    .State = Office.msoButtonUp
                    .Parameter = VBA.CStr(1)
                End With
                lng = 1
                While lng <= 8
                    With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
                        .Caption = "1/" & VBA.CStr(lng * lng)
                        .OnAction = "modToolbar.gWSADebugSSWSizeButton_OnAction"
                        .BeginGroup = False
                        .State = Office.msoButtonUp
                        .Parameter = VBA.CStr(lng)
                    End With
                    lng = lng * 2
                Wend
                If (.Controls.Count >= 2) Then
                    .Controls(2).BeginGroup = True
                End If
            End With
        End With
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub mWSAHelp_Add _
( _
)
    With Application.CommandBars(modProject.GstrName)
        With .Controls.Add(Type:=Office.msoControlPopup, Temporary:=True)
            .Caption = "Help"
            .TooltipText = "View the 'Help' menu items"
            .BeginGroup = True
            .Width = 72
            Dim mnu As Control
            With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
                .Caption = "&Help"
                .TooltipText = "View the '" & modProject.GstrNamePretty & "' help"
                .OnAction = "modToolbar.gWSAHelpHelp_OnAction"
                .BeginGroup = False
            End With
            With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
                .Caption = "Help on How To ..."
                .TooltipText = "View the '" & modProject.GstrNamePretty & "' How To ... help"
                .OnAction = "modToolbar.gWSAHelpHelpHowTo_OnAction"
                .BeginGroup = False
            End With
            With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
                .Caption = "Help on Command Bar"
                .TooltipText = "View the '" & modProject.GstrNamePretty & "' Command Bar help"
                .OnAction = "modToolbar.gWSAHelpHelpCommandBar_OnAction"
                .BeginGroup = False
            End With
            With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
                .Caption = "Help on Known Issues"
                .TooltipText = "View the '" & modProject.GstrNamePretty & "' known issues"
                .OnAction = "modToolbar.gWSAHelpHelpKnownIssues_OnAction"
                .BeginGroup = False
            End With
            With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
                .Caption = "Help on Copyright and Permisison"
                .TooltipText = "View the '" & modProject.GstrNamePretty & "' copyright and permission notice"
                .OnAction = "modToolbar.gWSAHelpHelpCopyrightPermission_OnAction"
                .BeginGroup = False
            End With
            With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
                .Caption = "Visit the Homepage"
                .TooltipText = "Visit the '" & modProject.GstrNamePretty & "' homepage"
                .OnAction = "modToolbar.gWSAHelpVisitHomepage_OnAction"
                .BeginGroup = True
            End With
            With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
                .Caption = "Email the Author"
                .TooltipText = "Email the '" & modProject.GstrNamePretty & "' author"
                .OnAction = "modToolbar.gWSAHelpEmailAuthor_OnAction"
                .BeginGroup = False
            End With
            With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
                .Caption = "Debug"
                .TooltipText = "Show or hide the Debug menu"
                .OnAction = "modToolbar.gWSAHelpDebug_OnAction"
                .BeginGroup = True
            End With
            With .Controls.Add(Type:=Office.msoControlButton, Temporary:=True)
                .Caption = "&About"
                .TooltipText = "View the '" & modProject.GstrNamePretty & "' about box"
                .OnAction = "modToolbar.gWSAHelpAbout_OnAction"
                .BeginGroup = True
            End With
        End With
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub mWSA_Update _
( _
)
    Dim dwCurrent As PowerPoint.DocumentWindow
    Dim sldCurrent As PowerPoint.SlideRange
    Dim cmd As Office.CommandBarButton
    
    Set dwCurrent = mValidWindow(mActiveWindow)
    Set sldCurrent = mSlideRange(dwCurrent)
    
    With Application.CommandBars(modProject.GstrName)
        With .Controls("Song Edit")
            .Enabled = True
            If ((dwCurrent Is Nothing) = True) Then
                .Enabled = False
            Else
                If (dwCurrent.Presentation.Slides.Count = 0) Then
                    .Enabled = False
                End If
            End If
            With .Controls("Set Category")
                .Enabled = True
                If ((dwCurrent Is Nothing) = True) Then
                    .Enabled = False
                Else
                    If (dwCurrent.Presentation.Slides.Count = 0) Then
                        .Enabled = False
                    End If
                    If ((sldCurrent Is Nothing) = True) Then
                        .Enabled = False
                    End If
                End If
            End With
            With .Controls("Sort")
                .Enabled = True
                If ((dwCurrent Is Nothing) = True) Then
                    .Enabled = False
                Else
                    If (dwCurrent.Presentation.Slides.Count = 0) Then
                        .Enabled = False
                    End If
                End If
            End With
            With .Controls("Create Index")
                .Enabled = True
                If ((dwCurrent Is Nothing) = True) Then
                    .Enabled = False
                Else
                    If (dwCurrent.Presentation.Slides.Count = 0) Then
                        .Enabled = False
                    End If
                End If
            End With
        End With
        With .Controls("Category")
            .Enabled = True
            .Text = modCategory.CategoryGet(sldCurrent)
            .Enabled = False
        End With
        With .Controls("Debug")
            If (modProject.gblnDebugEnabled = True) Then
                .Visible = True
                .Enabled = True
            Else
                .Visible = False
                .Enabled = False
            End If
            With .Controls("SSW Monitor")
                For Each cmd In .Controls
                    If (cmd.Parameter = VBA.CStr(modProject.gstrSlideShowWindowMonitor)) Then
                        cmd.State = Office.msoButtonDown
                    Else
                        cmd.State = Office.msoButtonUp
                    End If
                Next
            End With
            With .Controls("SSW Size")
                For Each cmd In .Controls
                    If (cmd.Parameter = VBA.CStr(modProject.glngSlideShowWindowSize)) Then
                        cmd.State = Office.msoButtonDown
                    Else
                        cmd.State = Office.msoButtonUp
                    End If
                Next
            End With
        End With
        With .Controls("Help")
            With .Controls("Debug")
                If (modProject.gblnDebugEnabled = True) Then
                    .State = Office.msoButtonDown
                End If
            End With
        End With
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Function mActiveWindow _
( _
) As PowerPoint.DocumentWindow
    On Error GoTo mActiveWindow_Exit
    
    Set mActiveWindow = Nothing
    
    If ((PowerPoint.Application.ActiveWindow Is Nothing) = False) Then
        Set mActiveWindow = PowerPoint.Application.ActiveWindow
    End If
    
mActiveWindow_Exit:
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mValidWindow _
( _
    ByVal dwCurrent As PowerPoint.DocumentWindow _
) As PowerPoint.DocumentWindow
    Dim ppActiveWindowViewType As PowerPoint.PpViewType
    
    Set mValidWindow = Nothing
    
    If ((dwCurrent Is Nothing) = False) Then
        With dwCurrent
            ppActiveWindowViewType = .ViewType
            If (.ViewType = PowerPoint.ppViewNormal) Then
                Select Case .Panes(2).ViewType
                    Case PowerPoint.ppViewSlide
                        ppActiveWindowViewType = PowerPoint.ppViewNormal
                    Case PowerPoint.ppViewSlideMaster
                        ppActiveWindowViewType = PowerPoint.ppViewSlideMaster
                End Select
            End If
        End With
    
        If ((ppActiveWindowViewType = PowerPoint.ppViewSlide) Or _
            (ppActiveWindowViewType = PowerPoint.ppViewNormal) Or _
            (ppActiveWindowViewType = PowerPoint.ppViewSlideSorter)) Then
            Set mValidWindow = dwCurrent
        End If
    End If
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mSlideRange _
( _
    ByRef dwCurrent As PowerPoint.DocumentWindow _
) As PowerPoint.SlideRange
    On Error GoTo mSlideRange_Exit
    
    Set mSlideRange = Nothing
    If ((mValidWindow(dwCurrent) Is Nothing) = False) Then
        If (dwCurrent.Selection.Type <> PowerPoint.ppSelectionNone) Then
            Set mSlideRange = dwCurrent.Selection.SlideRange
        ElseIf (VBA.IsNull(dwCurrent.View.Slide) = False) Then
            Set mSlideRange = dwCurrent.Presentation.Slides.Range(dwCurrent.View.Slide.SlideIndex)
        End If
    End If

mSlideRange_Exit:
End Function
