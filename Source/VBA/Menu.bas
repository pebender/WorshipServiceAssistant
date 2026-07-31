Attribute VB_Name = "Menu"
'===============================================================================
' Name:
'   WorshipServiceAssistant.Menu
'
' Description:
'   This Module creates the 'Worship Service Assistant' command bar (menu) with
'   the following controls:
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
'   Copyright (c) 2000,2001 Paul Bender
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
'   1.01.0000:
'     (1) Changed Help menu to merge 'Help on Copyright' and 'Help on License'
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
Private Const MenuName As String = "Worship Service Assistant"


'===============================================================================
' Private Variables.
'===============================================================================

'-------------------------------------------------------------------------------
' These variables are direct pointers to the menu and the menu controls.
'-------------------------------------------------------------------------------
Private MenuCommandBar              As Office.CommandBar
Private NavigatorControl            As Office.CommandBarButton
Private SongEditControl             As Office.CommandBarPopup
Private SongEditSetCategoryControl  As Office.CommandBarPopup
Private SongEditSortControl         As Office.CommandBarButton
Private SongEditCreateIndexControl  As Office.CommandBarButton
Private CategoryControl             As Office.CommandBarComboBox
Private DebugControl                As Office.CommandBarPopup
Private DebugSSWDisplayControl      As Office.CommandBarPopup
Private DebugSSWSizeControl         As Office.CommandBarPopup
Private HelpControl                 As Office.CommandBarPopup
Private HelpDebugControl            As Office.CommandBarButton


'===============================================================================
' Public Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'   Uninstalls the command bar.
'-------------------------------------------------------------------------------
Public Sub Menu_Unload()
    Dim Bar As Office.CommandBar
    Dim Control As Office.CommandBarControl
    
    '
    ' Uninstall any "Worship Service Assistant" command bars.
    '
    For Each Bar In Application.CommandBars
        If (Bar.BuiltIn = False) Then
            If (Left(Bar.Name, Len(MenuName)) = MenuName) Then
                Bar.Delete
            End If
        End If
    Next
    For Each Bar In Application.CommandBars
        If (Bar.BuiltIn = False) Then
            For Each Control In Bar.controls
                If (Control.BuiltIn = False) Then
                    If (Left(Control.Tag, Len(MenuName)) = MenuName) Then
                        Control.Delete
                    End If
                End If
            Next
        End If
    Next
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Installs the command bar.
'-------------------------------------------------------------------------------
Public Sub Menu_Load()
    Dim Bar As Office.CommandBar
    Dim BarRowIndex As Long
    
    '
    ' Uninstall any pre-existing "Worship Service Assistant" command bars.
    '
    Menu_Unload
    
    '
    ' Find the end of command bar list.  The "Worship Service Assistant" will
    ' be placed at the end of the command bar list.
    '
    For Each Bar In Application.CommandBars
        If (Bar.RowIndex > BarRowIndex) Then
            BarRowIndex = Bar.RowIndex
        End If
    Next
    BarRowIndex = BarRowIndex + 1
    
    '
    ' Install the 'Worship Service Assistant' command bar.
    '
    Set MenuCommandBar = Application.CommandBars.Add( _
        Temporary:=True)
    
    '
    ' Configure the 'Worship Service Assistant' command bar.
    '
    With MenuCommandBar
        .Name = MenuName
        .Position = msoBarTop
        .RowIndex = BarRowIndex
        .Visible = True
        .Protection = _
            msoBarNoChangeDock + _
            msoBarNoChangeVisible + _
            msoBarNoCustomize + _
            msoBarNoMove + _
            msoBarNoResize + _
            msoBarNoVerticalDock
    End With
        
    '
    ' Install the 'Worship Service Assistant' command bar items.
    '
    AddNavigatorControl
    AddSongEditControl
    AddCategoryControl
    AddDebugControl
    AddHelpControl
    
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_Disable()
    Dim Index As Long
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    
    CategoryControl.Enabled = True
    CategoryControl.Text = ""
    CategoryControl.Enabled = False
    '
    ' Disable all controls.
    '
    NavigatorControl.Enabled = False
    SongEditControl.Enabled = False
    SongEditSetCategoryControl.Enabled = False
    SongEditSortControl.Enabled = True
    SongEditCreateIndexControl.Enabled = False
    CategoryControl.Enabled = False
    HelpControl.Enabled = False
    HelpDebugControl.Enabled = False
    
    HelpDebugControl.State = msoButtonUp
    
    If (DebugControl.Visible = True) Then
        DebugControl.Enabled = False
        DebugSSWDisplayControl.Enabled = False
        DebugSSWSizeControl.Enabled = False
        For Index = 1 To DebugSSWDisplayControl.controls.Count Step 1
            DebugSSWDisplayControl.controls(Index).State = msoButtonUp
        Next
        For Index = 1 To DebugSSWSizeControl.controls.Count Step 1
            DebugSSWSizeControl.controls(Index).State = msoButtonUp
        Next
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_Refresh()
    Dim W As PowerPoint.DocumentWindow
    
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    Menu_Disable
    
    '
    ' By default, enable all visible controls (except Category).
    '
    MenuCommandBar.Enabled = True
    NavigatorControl.Enabled = True
    SongEditControl.Enabled = True
    SongEditSetCategoryControl.Enabled = True
    SongEditSortControl.Enabled = True
    SongEditCreateIndexControl.Enabled = True
    CategoryControl.Enabled = False
    If (DebugControl.Visible = True) Then
        DebugControl.Enabled = True
        DebugSSWDisplayControl.Enabled = True
        DebugSSWSizeControl.Enabled = True
    End If
    HelpControl.Enabled = True
    HelpDebugControl.Enabled = True
    
    If (Presentation.Exists = False) Then
        NavigatorControl.Enabled = False
    End If
    
    If (ActiveWindowExists = False) Then
        SongEditControl.Enabled = False
        SongEditSetCategoryControl.Enabled = False
        SongEditSortControl.Enabled = False
        SongEditCreateIndexControl.Enabled = False
        
        CategoryControl.Enabled = True
        CategoryControl.Text = ""
        CategoryControl.Enabled = False
    Else
        Set W = Application.ActiveWindow
        
        If (ActiveWindowSlideExists(W) = False) Then
            SongEditControl.Enabled = False
            SongEditSetCategoryControl.Enabled = False
            SongEditSortControl.Enabled = False
            SongEditCreateIndexControl.Enabled = False
        End If
        
        If (ActiveSlideExists(W) = False) Then
            SongEditSetCategoryControl.Enabled = False
            SongEditSortControl.Enabled = False
        End If
        
        CategoryControl.Enabled = True
        CategoryControl.Text = GetCategory(W)
        CategoryControl.Enabled = False
    End If
    
    If (DebugControl.Visible = True) Then
        HelpDebugControl.State = msoButtonDown
        
        DebugControl.Enabled = True
        DebugSSWDisplayControl.Enabled = True
        DebugSSWSizeControl.Enabled = True
        
        DebugSSWDisplayControl.controls(SlideShow_GetWindowDisplay + 1).State = msoButtonDown
        DebugSSWSizeControl.controls(SlideShow_GetWindowSize + 1).State = msoButtonDown
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Launch the slide navigator.  In order to exit, the slide navigator must be
'   unloaded.  If the slide navigator is only hidden, then this routine will
'   automatically refresh it and re-show it.  This is a hack to work around
'   some focus problems resulting from activating slide shows and activating
'   new presentations.
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionNavigator()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    '
    ' Update the menu bar.
    '
    Menu_Refresh
    '
    ' Exit if the control is not enabled.
    '
    If (CommandBars.ActionControl.Enabled = False) Then
        Exit Sub
    End If
    
    Navigator_Run
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionSongEdit()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    '
    ' Update the menu bar.
    '
    Menu_Refresh
    '
    ' Exit if the control is not enabled.
    '
    If (CommandBars.ActionControl.Enabled = False) Then
        Exit Sub
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionSongEditSetCategory()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionSongEditSetCategoryButton()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    If (ActiveWindowExists = False) Then
        Exit Sub
    End If
    
    Dim W As PowerPoint.DocumentWindow
    Set W = Application.ActiveWindow
    
    SetCategory W, CommandBars.ActionControl.Caption
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionSongEditSort()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    If (ActiveWindowExists = False) Then
        Exit Sub
    End If
    
    Dim W As PowerPoint.DocumentWindow
    Set W = Application.ActiveWindow
    
    Sort_Run W
    
    MsgBox _
        buttons:= _
            vbInformation, _
        Title:= _
            ProjectNamePretty, _
        Prompt:= _
            "Slide sort is complete."
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionSongEditIndex()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    If (ActiveWindowExists = False) Then
        Exit Sub
    End If
    
    Dim W As PowerPoint.DocumentWindow
    Set W = Application.ActiveWindow
    
    Index_Run W
    
    MsgBox _
        buttons:= _
            vbInformation, _
        Title:= _
            ProjectNamePretty, _
        Prompt:= _
            "Slide index generation is complete."
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionDebug()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    '
    ' Update the menu bar.
    '
    Menu_Refresh
    '
    ' Exit if the control is not enabled.
    '
    If (CommandBars.ActionControl.Enabled = False) Then
        Exit Sub
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionDebugSSWDisplay()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionDebugSSWDisplayButton()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    SlideShow_SetWindowDisplay CommandBars.ActionControl.Index - 1
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionDebugSSWSize()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionDebugSSWSizeButton()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    SlideShow_SetWindowSize CommandBars.ActionControl.Index - 1
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionHelp()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    '
    ' Update the menu bar.
    '
    Menu_Refresh
    '
    ' Exit if the control is not enabled.
    '
    If (CommandBars.ActionControl.Enabled = False) Then
        Exit Sub
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionHelpHelp()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    Dim HelpFile As String
    
    HelpFile = Help_GetHelpFileName(True)
    
    If (HelpFile = "") Then
        Exit Sub
     End If
    
    Call Help.HtmlHelp( _
        0&, _
        HelpFile, _
        Help.HH_DISPLAY_TOPIC, _
        Help.IDH_TopicPath_WSA)
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionHelpHelpHowTo()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    Dim HelpFile As String
    
    HelpFile = Help_GetHelpFileName(True)
    
    If (HelpFile = "") Then
        Exit Sub
     End If
        
    Call Help.HtmlHelp( _
        0&, _
        HelpFile, _
        Help.HH_DISPLAY_TOPIC, _
        Help.IDH_TopicPath_WSAHowTo)
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionHelpHelpCommandBar()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    Dim HelpFile As String
    
    HelpFile = Help_GetHelpFileName(True)
    
    If (HelpFile = "") Then
        Exit Sub
     End If
    
    Call Help.HtmlHelp( _
        0&, _
        HelpFile, _
        Help.HH_DISPLAY_TOPIC, _
        Help.IDH_TopicPath_WSACommandBar)
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionHelpHelpCopyrightPermission()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    Dim HelpFile As String
    
    HelpFile = Help_GetHelpFileName(True)
    
    If (HelpFile = "") Then
        Exit Sub
     End If
    
    Call Help.HtmlHelp( _
        0&, _
        HelpFile, _
        Help.HH_DISPLAY_TOPIC, _
        Help.IDH_TopicPath_WSACopyrightPermission)
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionHelpVisitHomepage()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    Win32_Shell32.ShellExecute _
        0&, _
        vbNullString, _
        ProjectHomepage, _
        vbNullString, _
        vbNullString, _
        Win32_Shell32.SW_SHOWNORMAL
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionHelpEmailAuthor()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    Win32_Shell32.ShellExecute _
        0&, _
        vbNullString, _
        ProjectEmail & "?subject=" & ProjectName & "%20" & ProjectVersion & ":%20", _
        vbNullString, _
        vbNullString, _
        Win32_Shell32.SW_SHOWNORMAL
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionHelpDebug()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    Dim Response As Long
    
    If (DebugControl.Visible = False) Then
        Response = MsgBox( _
            buttons:= _
                vbYesNo + vbDefaultButton2 + vbExclamation, _
            Title:= _
                ProjectNamePretty, _
            Prompt:= _
                "Enabling the 'Debug' menu can cause " & _
                "Worship Service Assistant to function incorrectly. " & _
                Chr(13) & Chr(10) & _
                Chr(13) & Chr(10) & _
                "Are you sure you want to enable the 'Debug' menu?")
        If (Response = vbNo) Then
            Exit Sub
        End If
    End If
    
    DebugControl.Visible = Not DebugControl.Visible
    SlideShow_Initialize
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionHelpAbout()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    Dim Prompt As String
    
    Prompt = _
        ProjectNamePretty & " (version " & ProjectVersion & ")" & _
        Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
        ProjectCopyright & _
        Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
        "Worship Service Assistant is a Microsoft PowerPoint add-in " & _
        "that is designed to make PowerPoint a more useful tool in a " & _
        "worship service. " & _
        Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
        ProjectNamePretty & " comes with ABSOLUTELY NO WARRANTY; " & _
        "for details view the 'Help on Copyright and Permission' in the 'Help' menu. " & _
        "This is free software, " & _
        "and you are welcome to redistribute it under certain conditions; " & _
        "for details view the 'Help on Copyright and Permission' in the 'Help' menu. "

    MsgBox _
        buttons:= _
            vbOKOnly + vbMsgBoxHelpButton, _
        Title:= _
            "About " & ProjectNamePretty, _
        Prompt:= _
            Prompt
End Sub


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'   Installs the 'Navigator' control on the command bar specified by the
'   variable MenuCommandBar.
'-------------------------------------------------------------------------------
Private Sub AddNavigatorControl()
    '
    ' Install the 'Navigator' control.
    '
    Set NavigatorControl = MenuCommandBar.controls.Add( _
        Type:=msoControlButton, _
        Temporary:=True)
    
    '
    ' Configure the 'Navigator' control.
    '
    With NavigatorControl
        .Style = msoButtonCaption
        .Caption = "Navigator"
        .TooltipText = "Launch the Navigator"
        .OnAction = "Menu_OnActionNavigator"
        .BeginGroup = True
        .Width = 72
    End With
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Installs the 'SongEdit' control on the command bar specified by the
'   variable MenuCommandBar.
'-------------------------------------------------------------------------------
Private Sub AddSongEditControl()
    Dim Button As Office.CommandBarButton
    Dim Category As Integer
    
    '
    ' Install the 'SongEdit' control.
    '
    Set SongEditControl = MenuCommandBar.controls.Add( _
        Type:=msoControlPopup, _
        Temporary:=True)
    
    '
    ' Configure the 'SongEdit' control.
    '
    With SongEditControl
        .Caption = "Song Edit"
        .TooltipText = "View the 'Song Edit' menu items"
        .OnAction = "Menu_OnActionSongEdit"
        .BeginGroup = True
        .Width = 72
        
        Set SongEditSetCategoryControl = .controls.Add( _
            Type:=msoControlPopup, _
            Temporary:=True)
        SongEditSetCategoryControl.Caption = "Set Category"
        SongEditSetCategoryControl.TooltipText = "Set the category of the selected slides"
        SongEditSetCategoryControl.BeginGroup = False
        
        Set SongEditSortControl = .controls.Add( _
            Type:=msoControlButton, _
            Temporary:=True)
        SongEditSortControl.Caption = "Sort"
        SongEditSortControl.TooltipText = "Sort slides aphabetically"
        SongEditSortControl.OnAction = "Menu_OnActionSongEditSort"
        SongEditSortControl.BeginGroup = False
        
        Set SongEditCreateIndexControl = .controls.Add( _
            Type:=msoControlButton, _
            Temporary:=True)
        SongEditCreateIndexControl.Caption = "Create Index"
        SongEditCreateIndexControl.TooltipText = "Generate slide index"
        SongEditCreateIndexControl.OnAction = "Menu_OnActionSongEditIndex"
        SongEditCreateIndexControl.BeginGroup = False
    End With
    
    ReDim Project_Categories(5)
    Project_Categories(0) = "Worship"
    Project_Categories(1) = "Choir"
    Project_Categories(2) = "Hymn"
    Project_Categories(3) = "Carol"
    Project_Categories(4) = "Children"
    Project_Categories(5) = "Liturgy"
    With SongEditSetCategoryControl
        Set Button = .controls.Add(Type:=msoControlButton, Temporary:=True)
        Button.Caption = "<none>"
        Button.TooltipText = "Set slide category to '<none>'"
        Button.OnAction = "Menu_OnActionSongEditSetCategoryButton"
        Button.BeginGroup = False
    
        For Category = LBound(Project_Categories) To UBound(Project_Categories) Step 1
            Set Button = .controls.Add(Type:=msoControlButton, Temporary:=True)
            Button.Caption = Project_Categories(Category)
            Button.OnAction = "Menu_OnActionSongEditSetCategoryButton"
            Button.BeginGroup = False
        Next
        If (.controls.Count >= 2) Then
            .controls(2).BeginGroup = True
        End If
    End With
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Installs the 'Category' control on the command bar specified by the
'   variable MenuCommandBar.
'-------------------------------------------------------------------------------
Private Sub AddCategoryControl()

    '
    ' Install the 'Category' control.
    '
    Set CategoryControl = MenuCommandBar.controls.Add( _
        Type:=msoControlEdit, _
        Temporary:=True)
    
    '
    ' Configure the 'Category' control.
    '
    With CategoryControl
        .Caption = "Category"
        .Text = ""
        .TooltipText = "Slide Category"
        .BeginGroup = True
        .Width = 72
    End With
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Installs the 'Debug' control on the command bar specified by the
'   variable MenuCommandBar.
'-------------------------------------------------------------------------------
Private Sub AddDebugControl()
    Dim Display As Long
    Dim Size As Long
    Dim Button As Office.CommandBarButton
    
    '
    ' Install the 'Debug' control.
    '
    Set DebugControl = MenuCommandBar.controls.Add( _
        Type:=msoControlPopup, _
        Temporary:=True)
    
    '
    ' Configure the 'Debug' control.
    '
    With DebugControl
        .Caption = "Debug"
        .TooltipText = "View the 'Debug' menu items"
        .OnAction = "Menu_OnActionDebug"
        .BeginGroup = True
        .Width = 72
        
        Set DebugSSWDisplayControl = .controls.Add( _
            Type:=msoControlPopup, _
            Temporary:=True)
        DebugSSWDisplayControl.Caption = "SSW Display"
        DebugSSWDisplayControl.TooltipText = "Set the slide show window monitor"
        DebugSSWDisplayControl.BeginGroup = False
        
        Set DebugSSWSizeControl = .controls.Add( _
            Type:=msoControlPopup, _
            Temporary:=True)
        DebugSSWSizeControl.Caption = "SSW Size"
        DebugSSWSizeControl.TooltipText = "Set the slide show window size"
        DebugSSWSizeControl.BeginGroup = False
        
    End With
    
    With DebugSSWDisplayControl
        Set Button = .controls.Add(Type:=msoControlButton, Temporary:=True)
        Button.Caption = "Default"
        Button.OnAction = "Menu_OnActionDebugSSWSizeButton"
        Button.BeginGroup = False
        Button.State = msoButtonUp
        Display = 1
        While Display <= Win32_User32.GetSystemMetrics(Win32_User32.SM_CMONITORS)
            Set Button = .controls.Add(Type:=msoControlButton, Temporary:=True)
            Button.Caption = "Monitor " & Display
            Button.OnAction = "Menu_OnActionDebugSSWDisplayButton"
            Button.BeginGroup = False
            Button.State = msoButtonUp
            Display = Display + 1
        Wend
        If (.controls.Count >= 2) Then
            .controls(2).BeginGroup = True
        End If
    End With
    
    With DebugSSWSizeControl
        Set Button = .controls.Add(Type:=msoControlButton, Temporary:=True)
        Button.Caption = "Default"
        Button.OnAction = "Menu_OnActionDebugSSWSizeButton"
        Button.BeginGroup = False
        Button.State = msoButtonUp
        Size = 1
        While Size <= 64
            Set Button = .controls.Add(Type:=msoControlButton, Temporary:=True)
            Button.Caption = "1/" & Size
            Button.OnAction = "Menu_OnActionDebugSSWSizeButton"
            Button.BeginGroup = False
            Button.State = msoButtonUp
            Size = Size * 2
        Wend
        If (.controls.Count >= 2) Then
            .controls(2).BeginGroup = True
        End If
    End With
    
    DebugControl.Visible = False
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Installs the 'Help' control on the command bar specified by the
'   variable MenuCommandBar.
'-------------------------------------------------------------------------------
Private Function AddHelpControl() As Boolean
    Dim MenuItem As Office.CommandBarButton
    
    '
    ' Install the 'Help' control.
    '
    Set HelpControl = MenuCommandBar.controls.Add( _
        Type:=msoControlPopup, _
        Temporary:=True)
    
    '
    ' Configure the 'Help' control.
    '
    With HelpControl
        .Caption = "Help"
        .TooltipText = "View the 'Help' menu items"
        .OnAction = "Menu_OnActionHelp"
        .BeginGroup = True
        .Width = 72
        
        Set MenuItem = .controls.Add(Type:=msoControlButton, Temporary:=True)
        MenuItem.Caption = "&Help"
        MenuItem.TooltipText = "View the '" & ProjectNamePretty & "' help"
        MenuItem.OnAction = "Menu_OnActionHelpHelp"
        MenuItem.BeginGroup = False
        
        Set MenuItem = .controls.Add(Type:=msoControlButton, Temporary:=True)
        MenuItem.Caption = "Help on How To ..."
        MenuItem.TooltipText = "View the '" & ProjectNamePretty & "' How To ... help"
        MenuItem.OnAction = "Menu_OnActionHelpHelpHowTo"
        MenuItem.BeginGroup = False
        
        Set MenuItem = .controls.Add(Type:=msoControlButton, Temporary:=True)
        MenuItem.Caption = "Help on Command Bar"
        MenuItem.TooltipText = "View the '" & ProjectNamePretty & "' Command Bar help"
        MenuItem.OnAction = "Menu_OnActionHelpHelpCommandBar"
        MenuItem.BeginGroup = False
        
        Set MenuItem = .controls.Add(Type:=msoControlButton, Temporary:=True)
        MenuItem.Caption = "Help on Copyright and Permisison"
        MenuItem.TooltipText = "View the '" & ProjectNamePretty & "' copyright and permission notice"
        MenuItem.OnAction = "Menu_OnActionHelpHelpCopyrightPermission"
        MenuItem.BeginGroup = False
        
        Set MenuItem = .controls.Add(Type:=msoControlButton, Temporary:=True)
        MenuItem.Caption = "Visit the Homepage"
        MenuItem.TooltipText = "Visit the '" & ProjectNamePretty & "' homepage"
        MenuItem.OnAction = "Menu_OnActionHelpVisitHomepage"
        MenuItem.BeginGroup = True
               
        Set MenuItem = .controls.Add(Type:=msoControlButton, Temporary:=True)
        MenuItem.Caption = "Email the Author"
        MenuItem.TooltipText = "Email the '" & ProjectNamePretty & "' author"
        MenuItem.OnAction = "Menu_OnActionHelpEmailAuthor"
        MenuItem.BeginGroup = False
               
        Set HelpDebugControl = .controls.Add( _
            Type:=msoControlButton, _
            Temporary:=True)
        HelpDebugControl.Caption = "Debug"
        HelpDebugControl.TooltipText = "Show or hide the Debug menu"
        HelpDebugControl.OnAction = "Menu_OnActionHelpDebug"
        HelpDebugControl.BeginGroup = True
        HelpDebugControl.State = msoButtonUp
        
        Set MenuItem = .controls.Add(Type:=msoControlButton, Temporary:=True)
        MenuItem.Caption = "&About"
        MenuItem.TooltipText = "View the '" & ProjectNamePretty & "' about box"
        MenuItem.OnAction = "Menu_OnActionHelpAbout"
        MenuItem.BeginGroup = True
    End With
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SetCategory(ByVal W As PowerPoint.DocumentWindow, ByVal Category As String)
    Dim S As PowerPoint.Slide
    
    If ((ActiveSelectionExists(W) = False) And _
        (ActiveSlideExists(W) = True)) Then
        Set S = ActiveSlide(W)
        If (Category = "") Then
            S.Tags.Delete "Category"
        ElseIf (LCase(Category) = "<none>") Then
            S.Tags.Delete "Category"
        Else
            S.Tags.Add "Category", Category
        End If
    Else
        For Each S In W.Selection.SlideRange
            If (Category = "") Then
                S.Tags.Delete "Category"
            ElseIf (LCase(Category) = "<none>") Then
                S.Tags.Delete "Category"
            Else
                S.Tags.Add "Category", Category
            End If
        Next
    End If
End Sub

Private Function GetCategory(ByVal W As PowerPoint.DocumentWindow) As String
    Dim Slide As PowerPoint.Slide
    GetCategory = ""
    
    If (ActiveWindowSlideExists(W) = False) Then
        GetCategory = ""
    ElseIf (ActiveSelectionExists(W) = False) Then
        If (ActiveSlideExists(W) = True) Then
            GetCategory = ActiveSlide(W).Tags("Category")
        Else
            GetCategory = ""
        End If
    ElseIf (W.Selection.SlideRange.Count = 0) Then
        If (ActiveSlideExists(W) = True) Then
            GetCategory = ActiveSlide(W).Tags("Category")
        Else
            GetCategory = ""
        End If
    Else
        GetCategory = W.Selection.SlideRange(1).Tags("Category")
        For Each Slide In W.Selection.SlideRange
            If (Slide.Tags("Category") <> GetCategory) Then
                    GetCategory = ""
                Exit For
            End If
        Next
    End If
End Function
