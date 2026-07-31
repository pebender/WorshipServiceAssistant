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
'     'Load':
'       Clicking this control causes the slide selected in the normal view
'       (on the primary monitor) to be loaded into the slide view (on the
'       secondary monitor).  If there is no slide show associated with the
'       presentation, then a slide show is started on the secondary monitor.
'     'Hide':
'       Clicking this control causes the slide show to toggle between
'       hidden and shown.
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
'   Paul Bender <pebender@san.rr.com>
'
' Copyright:
'   Copyright (C) 2000, 2001 Paul Bender
'
'   This program is free software; you can redistribute it and/or
'   modify it under the terms of the GNU General Public License
'   as published by the Free Software Foundation; version 2 of the License.
'
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with this program; if not, write to the Free Software
'   Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'
' Change History:
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

'-------------------------------------------------------------------------------
' These constants are the names used to identify the menus and controls.
'-------------------------------------------------------------------------------
Private Const MenuName As String _
    = "Worship Service Assistant"
Private Const NavigatorControlName As String _
    = MenuName & " - Navigator"
Private Const LoadControlName As String _
    = MenuName & " - Load"
Private Const HideControlName As String _
    = MenuName & " - Hide"
Private Const SongEditControlName As String _
    = MenuName & " - Song Edit"
Private Const SongEditSetCategoryControlName As String _
    = SongEditControlName & " - Set Category"
Private Const SongEditSortControlName As String _
    = SongEditControlName & " - Sort"
Private Const SongEditCreateIndexControlName As String _
    = SongEditControlName & " - Create Index"
Private Const CategoryControlName As String _
    = MenuName & " - Category"
Private Const DebugControlName As String _
    = MenuName & " - Debug"
Private Const DebugSSWDisplayControlName As String _
    = DebugControlName & " - SSW Display"
Private Const DebugSSWSizeControlName As String _
    = DebugControlName & " - SSW Size"
Private Const HelpControlName As String _
    = MenuName & " - Help"
Private Const HelpDebugControlName As String _
    = HelpControlName & " - Debug"

'===============================================================================
' Private Variables.
'===============================================================================

'-------------------------------------------------------------------------------
' These variables are direct pointers to the menu and the menu controls.
'-------------------------------------------------------------------------------
Private MenuCommandBar              As CommandBar
Private NavigatorControl            As CommandBarButton
Private LoadControl                 As CommandBarButton
Private HideControl                 As CommandBarButton
Private SongEditControl             As CommandBarPopup
Private SongEditSetCategoryControl  As CommandBarPopup
Private SongEditSortControl         As CommandBarButton
Private SongEditCreateIndexControl  As CommandBarButton
Private CategoryControl             As CommandBarComboBox
Private DebugControl                As CommandBarPopup
Private DebugSSWDisplayControl      As CommandBarPopup
Private DebugSSWDisplayLastControl  As CommandBarButton
Private DebugSSWDisplayFirstControl As CommandBarButton
Private DebugSSWDisplay01Control    As CommandBarButton
Private DebugSSWDisplay02Control    As CommandBarButton
Private DebugSSWDisplay03Control    As CommandBarButton
Private DebugSSWDisplay04Control    As CommandBarButton
Private DebugSSWSizeControl         As CommandBarPopup
Private DebugSSWSizeFullControl     As CommandBarButton
Private DebugSSWSizeHalfControl     As CommandBarButton
Private DebugSSWSizeQuarterControl  As CommandBarButton
Private HelpControl                 As CommandBarPopup
Private HelpDebugControl            As CommandBarButton


'===============================================================================
' Public Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'   Uninstalls the command bar.
'-------------------------------------------------------------------------------
Public Sub Menu_Unload()
    Dim Bar As CommandBar
    Dim Control As CommandBarControl
    
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
    Dim Bar As CommandBar
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
    AddLoadControl
    AddHideControl
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
    LoadControl.Enabled = False
    HideControl.Enabled = False
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
        
        DebugSSWDisplayLastControl.State = msoButtonUp
        DebugSSWDisplayFirstControl.State = msoButtonUp
        DebugSSWDisplay01Control.State = msoButtonUp
        DebugSSWDisplay02Control.State = msoButtonUp
        DebugSSWDisplay03Control.State = msoButtonUp
        DebugSSWDisplay04Control.State = msoButtonUp
        DebugSSWSizeFullControl.State = msoButtonUp
        DebugSSWSizeHalfControl.State = msoButtonUp
        DebugSSWSizeQuarterControl.State = msoButtonUp
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_Refresh()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    Dim Slide As Slide
    
    Menu_Disable
    
    MenuCommandBar.Enabled = True
    If (DebugControl.Visible = True) Then
        DebugControl.Enabled = True
        DebugSSWDisplayControl.Enabled = True
        DebugSSWSizeControl.Enabled = True
    End If
    HelpControl.Enabled = True
    
    If (ActiveWindowExists = False) Then
        Exit Sub
    End If
    
    Dim W As DocumentWindow
    Set W = Application.ActiveWindow
    
    '
    ' Set the category value based on the selected slides.
    '
    CategoryControl.Enabled = True
    If (ActiveWindowSlideExists(W) = False) Then
        CategoryControl.Text = ""
    ElseIf (ActiveSelectionExists(W) = False) Then
        If (ActiveSlideExists(W) = True) Then
            CategoryControl.Text = ActiveSlide(W).Tags("Category")
        Else
            CategoryControl.Text = ""
        End If
    ElseIf (W.Selection.SlideRange.Count = 0) Then
        If (ActiveSlideExists(W) = True) Then
            CategoryControl.Text = ActiveSlide(W).Tags("Category")
        Else
            CategoryControl.Text = ""
        End If
    Else
        CategoryControl.Text = W.Selection.SlideRange(1).Tags("Category")
        For Each Slide In W.Selection.SlideRange
            If (Slide.Tags("Category") <> CategoryControl.Text) Then
                    CategoryControl.Text = ""
                Exit For
            End If
        Next
    End If
    CategoryControl.Enabled = False
    
    '
    ' By default, enable all controls (except Category).
    '
    MenuCommandBar.Enabled = True
    NavigatorControl.Enabled = True
    LoadControl.Enabled = True
    HideControl.Enabled = True
    SongEditControl.Enabled = True
    SongEditSetCategoryControl.Enabled = True
    SongEditSortControl.Enabled = True
    SongEditCreateIndexControl.Enabled = True
    HelpControl.Enabled = True
    HelpDebugControl.Enabled = True
    
    If (ActiveWindowSlideExists(W) = False) Then
        NavigatorControl.Enabled = False
        LoadControl.Enabled = False
        HideControl.Enabled = False
        SongEditControl.Enabled = False
        SongEditSetCategoryControl.Enabled = False
        SongEditSortControl.Enabled = False
        SongEditCreateIndexControl.Enabled = False
    End If
    If (ActiveSlideExists(W) = False) Then
        LoadControl.Enabled = False
        HideControl.Enabled = False
        SongEditSetCategoryControl.Enabled = False
        SongEditSortControl.Enabled = False
    End If
    
    If (DebugControl.Visible = True) Then
        HelpDebugControl.State = msoButtonDown
        
        DebugControl.Enabled = True
        DebugSSWDisplayControl.Enabled = True
        DebugSSWSizeControl.Enabled = True
        
        Select Case SlideShow_GetWindowDisplay
            Case SlideShow_WindowDisplay.Last:
                DebugSSWDisplayLastControl.State = msoButtonDown
            Case SlideShow_WindowDisplay.First:
                DebugSSWDisplayFirstControl.State = msoButtonDown
            Case SlideShow_WindowDisplay.Display01:
                DebugSSWDisplay01Control.State = msoButtonDown
            Case SlideShow_WindowDisplay.Display02:
                DebugSSWDisplay02Control.State = msoButtonDown
            Case SlideShow_WindowDisplay.Display03:
                DebugSSWDisplay03Control.State = msoButtonDown
            Case SlideShow_WindowDisplay.Display04:
                DebugSSWDisplay04Control.State = msoButtonDown
        End Select
        Select Case SlideShow_GetWindowSize
            Case SlideShow_WindowSize.Full:
                DebugSSWSizeFullControl.State = msoButtonDown
            Case SlideShow_WindowSize.Half:
                DebugSSWSizeHalfControl.State = msoButtonDown
            Case SlideShow_WindowSize.Quarter:
                DebugSSWSizeQuarterControl.State = msoButtonDown
        End Select
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
    
    Navigator_Run
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionLoad()
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
    
    Dim W As DocumentWindow
    Set W = Application.ActiveWindow
    
    If (ActiveSlideShowExists(W) = False) Then
        SlideShow_End
        SlideShow_Setup W.Presentation
        SlideShow_Begin W
    End If
    
    SlideShow_Load W
    W.Activate
    W.Presentation.SlideShowWindow.Activate
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionHide()
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
    
    Dim W As DocumentWindow
    Set W = Application.ActiveWindow
    
    If (ActiveSlideShowExists(W) = False) Then
        Exit Sub
    End If
    
    SlideShow_Hide W
    W.Activate
    W.Presentation.SlideShowWindow.Activate
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
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionSongEditSetCategory01()
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
    
    Dim W As DocumentWindow
    Set W = Application.ActiveWindow
    
    SetCategory W, SongEditSetCategoryControl.controls(1).Caption
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionSongEditSetCategory02()
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
    
    Dim W As DocumentWindow
    Set W = Application.ActiveWindow
    
    SetCategory W, SongEditSetCategoryControl.controls(2).Caption
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionSongEditSetCategory03()
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
    
    Dim W As DocumentWindow
    Set W = Application.ActiveWindow
    
    SetCategory W, SongEditSetCategoryControl.controls(3).Caption
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionSongEditSetCategory04()
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
    
    Dim W As DocumentWindow
    Set W = Application.ActiveWindow
    
    SetCategory W, SongEditSetCategoryControl.controls(4).Caption
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionSongEditSetCategory05()
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
    
    Dim W As DocumentWindow
    Set W = Application.ActiveWindow
    
    SetCategory W, SongEditSetCategoryControl.controls(5).Caption
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
    
    Dim W As DocumentWindow
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
    
    Dim W As DocumentWindow
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
Public Sub Menu_OnActionDebugSSWDisplayLast()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    SlideShow_SetWindowDisplay SlideShow_WindowDisplay.Last
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionDebugSSWDisplayFirst()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    SlideShow_SetWindowDisplay SlideShow_WindowDisplay.First
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionDebugSSWDisplay01()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    SlideShow_SetWindowDisplay SlideShow_WindowDisplay.Display01
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionDebugSSWDisplay02()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    SlideShow_SetWindowDisplay SlideShow_WindowDisplay.Display02
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionDebugSSWDisplay03()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    SlideShow_SetWindowDisplay SlideShow_WindowDisplay.Display03
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionDebugSSWDisplay04()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    SlideShow_SetWindowDisplay SlideShow_WindowDisplay.Display04
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
Public Sub Menu_OnActionDebugSSWSizeFull()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    SlideShow_SetWindowSize SlideShow_WindowSize.Full
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionDebugSSWSizeHalf()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    SlideShow_SetWindowSize SlideShow_WindowSize.Half
    Menu_Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionDebugSSWSizeQuarter()
    '
    ' Load the project if the project is not loaded
    '
    If (Project_Loaded = False) Then
        Project_Load
        Exit Sub
    End If
    
    SlideShow_SetWindowSize SlideShow_WindowSize.Quarter
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
    
    Application.Help HelpFile, IDH_Topic_WSA
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
        
    Application.Help HelpFile, IDH_Topic_WSAHowTO
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
    
    Application.Help HelpFile, IDH_Topic_WSACommandBar
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionHelpHelpCopyright()
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
    
    Application.Help HelpFile, IDH_Topic_WSACopyright
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Menu_OnActionHelpHelpLicense()
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
    
    Application.Help HelpFile, IDH_Topic_WSALicense
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
    
    Dim IE As Object
    
    Set IE = CreateObject("InternetExplorer.application")
    IE.Visible = True
    IE.Navigate ProjectHomepage
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
    
    Dim IE As Object
    
    Set IE = CreateObject("InternetExplorer.application")
    IE.Navigate ProjectEmail
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
        "for details view the 'Help on Copyright' in the 'Help' menu. " & _
        "This is free software, " & _
        "and you are welcome to redistribute it under certain conditions; " & _
        "for details view the 'Help on License' in the 'Help' menu. "

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
        .Tag = NavigatorControlName
        .TooltipText = "Launch the Navigator"
        .OnAction = "Menu_OnActionNavigator"
        .BeginGroup = True
        .Width = 72
    End With
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Installs the 'Load' control on the command bar specified by the
'   variable MenuCommandBar.
'-------------------------------------------------------------------------------
Private Sub AddLoadControl()
    '
    ' Install the 'Load' control.
    '
    Set LoadControl = MenuCommandBar.controls.Add( _
        Type:=msoControlButton, _
        Temporary:=True)
    
    '
    ' Configure the 'Load' control.
    '
    With LoadControl
        .Style = msoButtonCaption
        .Caption = "Load"
        .Tag = LoadControlName
        .TooltipText = "Load the selected slide into the slide show"
        .OnAction = "Menu_OnActionLoad"
        .BeginGroup = True
        .Width = 72
    End With
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Installs the 'Hide' control on the command bar specified by the
'   variable MenuCommandBar.
'-------------------------------------------------------------------------------
Private Sub AddHideControl()
    '
    ' Install the 'Hide' control.
    '
    Set HideControl = MenuCommandBar.controls.Add( _
        Type:=msoControlButton, _
        Temporary:=True)
    
    '
    ' Configure the 'Hide' control.
    '
    With HideControl
        .Style = msoButtonCaption
        .Caption = "Hide"
        .Tag = HideControlName
        .TooltipText = "Toggle the slide show between hidden and shown"
        .OnAction = "Menu_OnActionHide"
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
    Dim MenuItem As CommandBarButton
    Dim Category As String
    
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
        .Tag = SongEditControlName
        .TooltipText = "View the 'Song Edit' menu items"
        .OnAction = "Menu_OnActionSongEdit"
        .BeginGroup = True
        .Width = 72
        
        Set SongEditSetCategoryControl = .controls.Add( _
            Type:=msoControlPopup, _
            Temporary:=True)
        SongEditSetCategoryControl.Caption = "Set Category"
        SongEditSetCategoryControl.Tag = SongEditSetCategoryControlName
        SongEditSetCategoryControl.TooltipText = "Set the category of the selected slides"
        SongEditSetCategoryControl.BeginGroup = False
        
        Set SongEditSortControl = .controls.Add( _
            Type:=msoControlButton, _
            Temporary:=True)
        SongEditSortControl.Caption = "Sort"
        SongEditSortControl.Tag = SongEditSortControlName
        SongEditSortControl.TooltipText = "Sort slides aphabetically"
        SongEditSortControl.OnAction = "Menu_OnActionSongEditSort"
        SongEditSortControl.BeginGroup = False
        
        Set SongEditCreateIndexControl = .controls.Add( _
            Type:=msoControlButton, _
            Temporary:=True)
        SongEditCreateIndexControl.Caption = "Create Index"
        SongEditCreateIndexControl.Tag = SongEditCreateIndexControlName
        SongEditCreateIndexControl.TooltipText = "Generate slide index"
        SongEditCreateIndexControl.OnAction = "Menu_OnActionSongEditIndex"
        SongEditCreateIndexControl.BeginGroup = False
    End With
    
    ReDim Project_Categories(3)
    Project_Categories(0) = "Worship"
    Project_Categories(1) = "Choir"
    Project_Categories(2) = "Hymn"
    Project_Categories(3) = "Carol"
    With SongEditSetCategoryControl
        Set MenuItem = .controls.Add(Type:=msoControlButton, Temporary:=True)
        MenuItem.Caption = "<none>"
        MenuItem.TooltipText = "Set slide category to '<none>'"
        MenuItem.OnAction = "Menu_OnActionSongEditSetCategory01"
        MenuItem.BeginGroup = False
    
        Set MenuItem = .controls.Add(Type:=msoControlButton, Temporary:=True)
        MenuItem.Caption = Project_Categories(0)
        MenuItem.TooltipText = "Set slide category to '" & Project_Categories(0) & "'"
        MenuItem.OnAction = "Menu_OnActionSongEditSetCategory02"
        MenuItem.BeginGroup = True
        
        Set MenuItem = .controls.Add(Type:=msoControlButton, Temporary:=True)
        MenuItem.Caption = Project_Categories(1)
        MenuItem.TooltipText = "Set slide category to '" & Project_Categories(1) & "'"
        MenuItem.OnAction = "Menu_OnActionSongEditSetCategory03"
        MenuItem.BeginGroup = False
        
        Set MenuItem = .controls.Add(Type:=msoControlButton, Temporary:=True)
        MenuItem.Caption = Project_Categories(2)
        MenuItem.TooltipText = "Set slide category to '" & Project_Categories(2) & "'"
        MenuItem.OnAction = "Menu_OnActionSongEditSetCategory04"
        MenuItem.BeginGroup = False
        
        Set MenuItem = .controls.Add(Type:=msoControlButton, Temporary:=True)
        MenuItem.Caption = Project_Categories(3)
        MenuItem.TooltipText = "Set slide category to '" & Project_Categories(3) & "'"
        MenuItem.OnAction = "Menu_OnActionSongEditSetCategory05"
        MenuItem.BeginGroup = False
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
        .Tag = CategoryControlName
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
        .Tag = DebugControlName
        .TooltipText = "View the 'Debug' menu items"
        .OnAction = "Menu_OnActionDebug"
        .BeginGroup = True
        .Width = 72
        
        Set DebugSSWDisplayControl = .controls.Add( _
            Type:=msoControlPopup, _
            Temporary:=True)
        DebugSSWDisplayControl.Caption = "SSW Display"
        DebugSSWDisplayControl.Tag = DebugSSWDisplayControlName
        DebugSSWDisplayControl.TooltipText = "Set the slide show window display"
        DebugSSWDisplayControl.BeginGroup = False
        
        Set DebugSSWSizeControl = .controls.Add( _
            Type:=msoControlPopup, _
            Temporary:=True)
        DebugSSWSizeControl.Caption = "SSW Size"
        DebugSSWSizeControl.Tag = DebugSSWSizeControlName
        DebugSSWSizeControl.TooltipText = "Set the slide show window size"
        DebugSSWSizeControl.BeginGroup = False
    End With
    
    With DebugSSWDisplayControl
        Set DebugSSWDisplayLastControl = .controls.Add(Type:=msoControlButton, Temporary:=True)
        DebugSSWDisplayLastControl.Caption = "Last"
        DebugSSWDisplayLastControl.TooltipText = "Display the slide show window on the last display"
        DebugSSWDisplayLastControl.OnAction = "Menu_OnActionDebugSSWDisplayLast"
        DebugSSWDisplayLastControl.BeginGroup = False
        DebugSSWDisplayLastControl.State = msoButtonDown
        
        Set DebugSSWDisplayFirstControl = .controls.Add(Type:=msoControlButton, Temporary:=True)
        DebugSSWDisplayFirstControl.Caption = "First"
        DebugSSWDisplayFirstControl.TooltipText = "Use the first display for the slide show window"
        DebugSSWDisplayFirstControl.OnAction = "Menu_OnActionDebugSSWDisplayFirst"
        DebugSSWDisplayFirstControl.BeginGroup = False
        DebugSSWDisplayFirstControl.State = msoButtonUp
        
        Set DebugSSWDisplay01Control = .controls.Add(Type:=msoControlButton, Temporary:=True)
        DebugSSWDisplay01Control.Caption = "1"
        DebugSSWDisplay01Control.TooltipText = "Use the display 1 for the slide show window"
        DebugSSWDisplay01Control.OnAction = "Menu_OnActionDebugSSWDisplay01"
        DebugSSWDisplay01Control.BeginGroup = True
        DebugSSWDisplay01Control.State = msoButtonUp
        
        Set DebugSSWDisplay02Control = .controls.Add(Type:=msoControlButton, Temporary:=True)
        DebugSSWDisplay02Control.Caption = "2"
        DebugSSWDisplay02Control.TooltipText = "Use the display 2 for the slide show window"
        DebugSSWDisplay02Control.OnAction = "Menu_OnActionDebugSSWDisplay02"
        DebugSSWDisplay02Control.BeginGroup = False
        DebugSSWDisplay02Control.State = msoButtonUp
        
        Set DebugSSWDisplay03Control = .controls.Add(Type:=msoControlButton, Temporary:=True)
        DebugSSWDisplay03Control.Caption = "3"
        DebugSSWDisplay03Control.TooltipText = "Use the display 3 for the slide show window"
        DebugSSWDisplay03Control.OnAction = "Menu_OnActionDebugSSWDisplay03"
        DebugSSWDisplay03Control.BeginGroup = False
        DebugSSWDisplay03Control.State = msoButtonUp
        
        Set DebugSSWDisplay04Control = .controls.Add(Type:=msoControlButton, Temporary:=True)
        DebugSSWDisplay04Control.Caption = "4"
        DebugSSWDisplay04Control.TooltipText = "Use the display 4 for the slide show window"
        DebugSSWDisplay04Control.OnAction = "Menu_OnActionDebugSSWDisplay04"
        DebugSSWDisplay04Control.BeginGroup = False
        DebugSSWDisplay04Control.State = msoButtonUp
    End With
    
    With DebugSSWSizeControl
        Set DebugSSWSizeFullControl = .controls.Add(Type:=msoControlButton, Temporary:=True)
        DebugSSWSizeFullControl.Caption = "Full"
        DebugSSWSizeFullControl.TooltipText = "Use the last display for the slide show window"
        DebugSSWSizeFullControl.OnAction = "Menu_OnActionDebugSSWSizeFull"
        DebugSSWSizeFullControl.BeginGroup = False
        DebugSSWSizeFullControl.State = msoButtonDown
    
        Set DebugSSWSizeHalfControl = .controls.Add(Type:=msoControlButton, Temporary:=True)
        DebugSSWSizeHalfControl.Caption = "Half"
        DebugSSWSizeHalfControl.TooltipText = "Use a half size slide show window"
        DebugSSWSizeHalfControl.OnAction = "Menu_OnActionDebugSSWSizeHalf"
        DebugSSWSizeHalfControl.BeginGroup = False
        DebugSSWSizeHalfControl.State = msoButtonUp
    
        Set DebugSSWSizeQuarterControl = .controls.Add(Type:=msoControlButton, Temporary:=True)
        DebugSSWSizeQuarterControl.Caption = "Quarter"
        DebugSSWSizeQuarterControl.TooltipText = "Use a quarter size slide show window"
        DebugSSWSizeQuarterControl.OnAction = "Menu_OnActionDebugSSWSizeQuarter"
        DebugSSWSizeQuarterControl.BeginGroup = False
        DebugSSWSizeQuarterControl.State = msoButtonUp
    End With
    
    DebugControl.Visible = False
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Installs the 'Help' control on the command bar specified by the
'   variable MenuCommandBar.
'-------------------------------------------------------------------------------
Private Function AddHelpControl() As Boolean
    Dim MenuItem As CommandBarButton
    
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
        .Tag = HelpControlName
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
        MenuItem.Caption = "Help on Copyright"
        MenuItem.TooltipText = "View the '" & ProjectNamePretty & "' copyright"
        MenuItem.OnAction = "Menu_OnActionHelpHelpCopyright"
        MenuItem.BeginGroup = False
        
        Set MenuItem = .controls.Add(Type:=msoControlButton, Temporary:=True)
        MenuItem.Caption = "Help on License"
        MenuItem.TooltipText = "View the '" & ProjectNamePretty & "' license"
        MenuItem.OnAction = "Menu_OnActionHelpHelpLicense"
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
        HelpDebugControl.Tag = HelpDebugControlName
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
Private Sub SetCategory(ByVal W As DocumentWindow, ByVal Category As String)
    Dim S As Slide
    
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


