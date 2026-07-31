VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NavigatorForm 
   Caption         =   "Navigator"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "NavigatorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NavigatorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.NavigatorForm
'
' Description:
'   The form implements the Navigator.
'
'   The form contains an empty frame (FrameEmpty) to which the focus is locked.
'   The empty frame is not visible to the user because its height and width
'   are both 0.  The form handles all if the keyboard input processing using
'   its KeyPress and KeyDown event handlers.  The KeyPress event handler
'   processes the alphanumeric input for the slide filters.  The KeyDown event
'   handler processes the shortcut inputs.
'
'   Since the form is modal, the user cannot change any of the presentations
'   while the form is running.  As a result, the form can perform all of the
'   configurations of the presentations when it is initialized.  This makes
'   the lauching of the form take more time.  However, after that, the form
'   responds much more quickly.
'
'   The presentation intialization includes: reconfiguring the presentation
'   view, setting up the slide show view and cleaning up the slide titles.
'   Information associated with a presentation or a slide is stored in the
'   Tags property of the presentation or slide.
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
'   1.00.0001:
'     (1) Changed CleanEverything routine to remove quotes and match
'         commonly interchanged words.
'     (2) Changed SlideShowPrev and SlideShowNext to SlideShowPrevEffect and
'         SlideShowNextEffect in order to make them more consistant with the
'         button names.
'     (3) Fixed a problem where the part of the presentation window would
'         appear on the the slide show display when under certain conditions
'         while the Navigator is launching.
'   1.00.0000:
'     Initial revision.
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
' Public Variables.
'===============================================================================


'===============================================================================
' Private Constants.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================
'
' Indicates whether or not the Navigator form's focus is locked to the empty
' frame.  Normally the focus is locked to the empty frame.  However, sometimes
' it is necessary to change the focus in order to force the Navigator form
' to be the active window.  During this time, the focus must be unlocked.
'
Private NavigatorFormLocked As Boolean

'===============================================================================
' Public Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Refresh()
    If (NavigatorValid = True) Then
        Dim W As DocumentWindow
        Set W = Application.ActiveWindow
        UpdatePresentationName W
        UpdateSlideList W
        UpdateEnabledControls W
    End If
End Sub


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub UpdateEnabledControls(ByVal W As DocumentWindow)
    '
    ' Enable everything by default
    '
    NavigatorForm.ControlSlideShowLoad.Enabled = True
    NavigatorForm.ControlSlideShowHide.Enabled = True
    NavigatorForm.ControlSlideShowHide.Caption = "Hide"
    NavigatorForm.ControlSlideShowRun.Enabled = True
    NavigatorForm.ControlSlideShowPause.Enabled = True
    NavigatorForm.ControlSlideShowPrevEffect.Enabled = True
    NavigatorForm.ControlSlideShowNextEffect.Enabled = True
    NavigatorForm.ControlGeneralExit.Enabled = True
    NavigatorForm.ControlGeneralHelp.Enabled = True
    NavigatorForm.ControlPresentationSelectionPrev.Enabled = True
    NavigatorForm.ControlPresentationSelectionNext.Enabled = True
    NavigatorForm.ControlSlideSelectionClear.Enabled = True
    NavigatorForm.ControlSlideSelectionList.Enabled = True
    
    '
    ' Since there are no slides,
    ' disable slide controls.
    '
    If (ActiveWindowSlideExists(W) = False) Then
        NavigatorForm.ControlSlideShowLoad.Enabled = False
        NavigatorForm.ControlSlideShowHide.Enabled = False
        NavigatorForm.ControlSlideShowRun.Enabled = False
        NavigatorForm.ControlSlideShowPause.Enabled = False
        NavigatorForm.ControlSlideShowPrevEffect.Enabled = False
        NavigatorForm.ControlSlideShowNextEffect.Enabled = False
    End If
    '
    ' Since there are no slides in the slide list,
    ' disable slide show controls.
    '
    If (NavigatorForm.ControlSlideSelectionList.ListCount = 0) Then
        NavigatorForm.ControlSlideShowLoad.Enabled = False
        NavigatorForm.ControlSlideShowHide.Enabled = False
        NavigatorForm.ControlSlideShowRun.Enabled = False
        NavigatorForm.ControlSlideShowPause.Enabled = False
        NavigatorForm.ControlSlideShowPrevEffect.Enabled = False
        NavigatorForm.ControlSlideShowNextEffect.Enabled = False
    End If
    '
    ' Since there are no slides in the slide list,
    ' disable slide load control.
    '
    If (ActiveSlideShowExists(W) = False) Then
        NavigatorForm.ControlSlideShowHide.Enabled = False
        NavigatorForm.ControlSlideShowRun.Enabled = False
        NavigatorForm.ControlSlideShowPause.Enabled = False
        NavigatorForm.ControlSlideShowPrevEffect.Enabled = False
        NavigatorForm.ControlSlideShowNextEffect.Enabled = False
    End If
    '
    ' Since the slide show is hidden,
    ' relable the hide control as Show.
    '
    If (ActiveSlideShowExists(W) = True) Then
        If (W.Presentation.SlideShowWindow.View.State = ppSlideShowBlackScreen) Then
            NavigatorForm.ControlSlideShowHide.Caption = "Show"
        End If
    End If
    '
    ' Since the slide show is running,
    ' disable the run control.
    '
    If (ActiveSlideShowExists(W) = True) Then
        If (W.Presentation.SlideShowWindow.View.State = ppSlideShowRunning) Then
            NavigatorForm.ControlSlideShowRun.Enabled = False
        End If
    End If
    '
    ' Since the slide show is paused,
    ' disable the pause control.
    '
    If (ActiveSlideShowExists(W) = True) Then
        If (W.Presentation.SlideShowWindow.View.State = ppSlideShowPaused) Then
            NavigatorForm.ControlSlideShowPause.Enabled = False
        End If
    End If
    '
    ' Since the slide filter is clear,
    ' disable the slide filter clear control.
    '
    If ((NavigatorForm.ControlSlideSelectionNumber.Caption = "") And _
        (NavigatorForm.ControlSlideSelectionTitle.Caption = "")) Then
        NavigatorForm.ControlSlideSelectionClear.Enabled = False
    End If
    
    '
    ' Set focus and try to make sure that the Navigator is the
    ' active window. Changing the focus seems to accomplish it.
    '
    NavigatorFormLocked = False
    NavigatorForm.FrameGeneral.SetFocus
    NavigatorForm.FrameEmpty.SetFocus
    NavigatorFormLocked = True
    
    '
    ' Set pointer.
    '
    NavigatorForm.MousePointer = fmMousePointerArrow
    NavigatorForm.Repaint
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SlideShowLoad(ByVal W As DocumentWindow)
    If ((NavigatorForm.ControlSlideSelectionNumber.Caption <> "") Or _
        (NavigatorForm.ControlSlideSelectionTitle.Caption <> "")) Then
        NavigatorForm.ControlSlideSelectionNumber.Caption = ""
        NavigatorForm.ControlSlideSelectionTitle.Caption = ""
        UpdateSlideList W
    End If
    
    If (ActiveSlideShowExists(W) = False) Then
        NavigatorForm.Hide
        SlideShow_End
        SlideShow_Begin W
    End If
    
    SlideShow_Load W
        
    W.Activate
    
    UpdateEnabledControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SlideShowHide(ByVal W As DocumentWindow)
    If (ActiveSlideShowExists(W) = False) Then
        Exit Sub
    End If
    
    SlideShow_Hide W
    
    UpdateEnabledControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SlideShowRun(ByVal W As DocumentWindow)
    If (ActiveSlideShowExists(W) = False) Then
        Exit Sub
    End If
    
    SlideShow_Run W
    
    UpdateEnabledControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SlideShowPause(ByVal W As DocumentWindow)
    If (ActiveSlideShowExists(W) = False) Then
        Exit Sub
    End If
    
    SlideShow_Pause W
    
    UpdateEnabledControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SlideShowPrevEffect(ByVal W As DocumentWindow)
    If (ActiveSlideShowExists(W) = False) Then
        Exit Sub
    End If
    
    SlideShow_Prev W
    
    ControlSlideSelectionClear_Click
    
    If (ActiveSlideExists(W) = True) Then
        NavigatorForm.ControlSlideSelectionList.ListIndex = ActiveSlide(W).SlideIndex - 1
    End If
    
    UpdateEnabledControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SlideShowNextEffect(ByVal W As DocumentWindow)
    If (ActiveSlideShowExists(W) = False) Then
        Exit Sub
    End If
    
    SlideShow_Next W
    
    ControlSlideSelectionClear_Click
    
    If (ActiveSlideExists(W) = True) Then
        NavigatorForm.ControlSlideSelectionList.ListIndex = ActiveSlide(W).SlideIndex - 1
    End If
    
    UpdateEnabledControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PresentationSelectionPrev(ByVal W As DocumentWindow)
    If (Application.Presentations.Count <= 1) Then
        Exit Sub
    End If
    
    Dim I As Long
    Dim J As Long
    
    For I = Application.Presentations.Count To 1 Step -1
        If (Application.Presentations(I) Is W.Presentation) Then
            Exit For
        End If
    Next
    
    J = I
    Do
        J = J - 1
        If (J < 1) Then
            J = Application.Presentations.Count
        End If
    Loop While ((J <> I) And (ActiveWindowSlideExists(Application.Presentations(J).Windows(1)) = False))
    
    If (J <> I) Then
        Application.Presentations(J).Windows(1).Activate
        Set W = Application.ActiveWindow
        UpdatePresentationName W
        UpdateSlideList W
        UpdateEnabledControls W
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PresentationSelectionNext(ByVal W As DocumentWindow)
    If (Application.Presentations.Count <= 1) Then
        Exit Sub
    End If
    
    Dim I As Long
    Dim J As Long
    
    For I = 1 To Application.Presentations.Count Step 1
        If (Application.Presentations(I) Is W.Presentation) Then
            Exit For
        End If
    Next
    
    J = I
    Do
        J = J + 1
        If (J > Application.Presentations.Count) Then
            J = 1
        End If
    Loop While ((J <> I) And (ActiveWindowSlideExists(Application.Presentations(J).Windows(1)) = False))
    
    If (J <> I) Then
        Application.Presentations(J).Windows(1).Activate
        Set W = Application.ActiveWindow
        UpdatePresentationName W
        UpdateSlideList W
        UpdateEnabledControls W
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SlideSelectionClear(ByVal W As DocumentWindow)
    If ((NavigatorForm.ControlSlideSelectionNumber.Caption <> "") Or _
        (NavigatorForm.ControlSlideSelectionTitle.Caption <> "")) Then
        NavigatorForm.ControlSlideSelectionNumber.Caption = ""
        NavigatorForm.ControlSlideSelectionTitle.Caption = ""
        UpdateSlideList W
    End If
    UpdateEnabledControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SlideSelectionUpdate(ByVal W As DocumentWindow)
    Dim sIndex As Long
    '
    ' Set slide selected in the presentation to match the
    ' slide selected in the slide list control
    '
    If (NavigatorForm.ControlSlideSelectionList.ListIndex >= 0) Then
        sIndex = NavigatorForm.ControlSlideSelectionList.Text
        W.View.Slide = W.Presentation.slides(sIndex)
    End If
    UpdateEnabledControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SlideSelectionPrev(ByVal W As DocumentWindow)
    With NavigatorForm.ControlSlideSelectionList
        If (.ListCount <= 1) Then
            Exit Sub
        End If
        If (.ListIndex > 0) Then
            .ListIndex = .ListIndex - 1
        Else
            .ListIndex = .ListCount - 1
        End If
    End With
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SlideSelectionNext(ByVal W As DocumentWindow)
    With NavigatorForm.ControlSlideSelectionList
        If (.ListCount <= 1) Then
            Exit Sub
        End If
        If (.ListIndex < .ListCount - 1) Then
            .ListIndex = .ListIndex + 1
        Else
            .ListIndex = 0
        End If
    End With
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Function NavigatorValid() As Boolean
    NavigatorValid = False
    '
    ' Automatically hide the NavigatorForm form if there is no active window.
    '
    If (ActiveWindowExists = False) Then
        NavigatorForm.Hide
        Exit Function
    End If
    NavigatorValid = True
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub UpdateSlideTitles(ByVal P As Presentation)
    Dim S As Slide
    Dim Title As String
    Dim Saved As Boolean
    
    Saved = P.Saved
    For Each S In P.slides
        If (S.Shapes.HasTitle = msoTrue) Then
            Title = S.Shapes.Title.TextFrame.TextRange.Text
        Else
            Title = ""
        End If
        S.Tags.Add "WorshipServiceAssistant_TitleDisplay", CleanWhiteSpace(Title)
        S.Tags.Add "WorshipServiceAssistant_TitleMatch", CleanEverything(Title)
    Next
    P.Saved = Saved
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Function CleanWhiteSpace(Title As String) As String
    Dim T As String
    
    T = Title
    
    '
    ' Fix white space.
    '
    T = Replace(T, Chr(9), " ")        ' replace tab with space
    T = Replace(T, Chr(11), " ")       ' replace line feed with space
    T = Replace(T, Chr(13), " ")       ' replace return with space
    '
    ' Fix quotes.
    '
    T = Replace(T, "`", Chr(39))       ' replace back single quote with single quote
    T = Replace(T, Chr(145), Chr(39))  ' replace open single quote with single quote
    T = Replace(T, Chr(146), Chr(39))  ' replace close single quote with single quote
    T = Replace(T, Chr(147), Chr(34))  ' replace open double quote with double quote
    T = Replace(T, Chr(148), Chr(34))  ' replace close double quote with double quote

    While (InStr(T, "  ") > 0)
        T = Replace(T, "  ", " ")
    Wend
    
    CleanWhiteSpace = T
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Function CleanEverything(Title As String) As String
    Dim T As String
    
    T = LCase(Title)
    
    '
    ' Fix white space.
    '
    T = Replace(T, Chr(9), " ")        ' replace tab with space
    T = Replace(T, Chr(11), " ")       ' replace line feed with space
    T = Replace(T, Chr(13), " ")       ' replace return with space
    '
    ' Fix quotes.
    '
    T = Replace(T, "`", Chr(39))       ' replace back single quote with single quote
    T = Replace(T, Chr(145), Chr(39))  ' replace open single quote with single quote
    T = Replace(T, Chr(146), Chr(39))  ' replace close single quote with single quote
    T = Replace(T, Chr(147), Chr(34))  ' replace open double quote with double quote
    T = Replace(T, Chr(148), Chr(34))  ' replace close double quote with double quote
    
    T = Replace(T, "'s", "s")          ' ignore ' in 's
    T = Replace(T, "'t", "t")          ' ignore ' in 't
    T = Replace(T, "'ve", "ve")        ' ignore ' in 've
    
    '
    ' Remove basic punctuation.
    '
    T = Replace(T, "-", " ")
    T = Replace(T, ",", " ")
    T = Replace(T, ";", " ")
    T = Replace(T, ":", " ")
    T = Replace(T, ".", " ")
    T = Replace(T, "!", " ")
    T = Replace(T, "?", " ")
    T = Replace(T, "/", " ")
    T = Replace(T, Chr(39), " ")
    T = Replace(T, Chr(34), " ")
    
    While (InStr(T, "  ") > 0)
        T = Replace(T, "  ", " ")
    Wend
    
    '
    ' Replace commonly interchanged words.
    '
    T = " " & T & " "
    T = Replace(T, " oh ", " o ")
    T = Replace(T, " alleluia ", " hallelujah ")
    T = Replace(T, " allelujah ", " hallelujah ")
    T = Replace(T, " emanuel ", " immanuel ")
    T = Replace(T, " emmanuel ", " immanuel ")
    T = Replace(T, " imanuel ", " immanuel ")
    If (Len(T) >= 2) Then
        T = Left(T, Len(T) - 1)
        T = Right(T, Len(T) - 1)
    End If
    
    CleanEverything = T
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub UpdateApplicationView(ByVal Save As Boolean)
    Static WindowWindowStateMax As Boolean
    Static WindowState As PpWindowState
    Static Top As Long
    Static Left As Long
    Static Height As Long
    Static Width As Long
    Dim W As DocumentWindow
    
    '
    ' Save current view.
    '
    If (Save = True) Then
        WindowWindowStateMax = False
        For Each W In Application.Windows
            If (W.WindowState = ppWindowMaximized) Then
                WindowWindowStateMax = True
            End If
        Next
        WindowState = Application.WindowState
        Application.WindowState = ppWindowNormal
        Top = Application.Top
        Left = Application.Left
        Height = Application.Height
        Width = Application.Width
        Application.WindowState = WindowState
    '
    ' Restore old view.
    '
    Else
        If (WindowWindowStateMax = True) Then
            Application.Windows(1).WindowState = ppWindowMaximized
        End If
        Application.WindowState = ppWindowNormal
        Application.Height = 0
        Application.Width = 0
        Application.Top = Top
        Application.Left = Left
        Application.Height = Height
        Application.Width = Width
        Application.WindowState = WindowState
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SetApplicationView()
    Dim Top As Long
    Dim Left As Long
    Dim Height As Long
    Dim Width As Long
    
    Application.WindowState = ppWindowMaximized
    Top = Application.Top + 3
    Left = Application.Left + 3
    Height = Application.Height - 6
    Width = Application.Width - 6
    Application.WindowState = ppWindowNormal
    Application.Height = 0
    Application.Width = 0
    Application.Top = Top
    Application.Left = Left + NavigatorForm.Width
    Application.Height = Height
    Application.Width = Width - NavigatorForm.Width
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub LoadPresentationView(ByVal P As Presentation)
    Dim WindowState As PpWindowState
    Dim ViewType As PpViewType
    Dim SplitHorizontal As Long
    Dim SplitVertical As Long
    Dim W As DocumentWindow
    Dim Saved As MsoTriState
    
    Set W = P.Windows(1)
    
    '
    ' Save current view.
    '
    WindowState = W.WindowState
    ViewType = W.ViewType
    If (ViewType = ppViewNormal) Then
        SplitHorizontal = W.SplitHorizontal
        SplitVertical = W.SplitVertical
    End If
    
    Saved = P.Saved
    P.Tags.Add "WorshipServiceAssistant_Window_WindowState", WindowState
    P.Tags.Add "WorshipServiceAssistant_Window_ViewType", ViewType
    P.Tags.Add "WorshipServiceAssistant_Window_SplitHorizontal", SplitHorizontal
    P.Tags.Add "WorshipServiceAssistant_Window_SplitVertical", SplitVertical
    P.Saved = Saved
    
    '
    ' Set view.
    '
    W.WindowState = ppWindowMaximized
    W.ViewType = ppViewNormal
    W.SplitHorizontal = 0
    W.SplitVertical = 100
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub UnloadPresentationView(ByVal P As Presentation)
    Dim WindowState As PpWindowState
    Dim ViewType As PpViewType
    Dim SplitHorizontal As Long
    Dim SplitVertical As Long
    Dim W As DocumentWindow
    
    Set W = P.Windows(1)
    
    '
    ' Unload current view.
    '
    WindowState = P.Tags("WorshipServiceAssistant_Window_WindowState")
    ViewType = P.Tags("WorshipServiceAssistant_Window_ViewType")
    SplitHorizontal = P.Tags("WorshipServiceAssistant_Window_SplitHorizontal")
    SplitVertical = P.Tags("WorshipServiceAssistant_Window_SplitVertical")
    
    W.WindowState = WindowState
    W.ViewType = ViewType
    If (W.ViewType = ppViewNormal) Then
        W.SplitHorizontal = SplitHorizontal
        W.SplitVertical = SplitVertical
    End If
End Sub


'===============================================================================
' Private Event Handlers.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    If (ActiveWindowExists = False) Then
        Unload NavigatorForm
        Exit Sub
    End If
    
    NavigatorFormLocked = True
    
    Dim W As DocumentWindow
    Set W = Application.ActiveWindow
    
    Dim P As Presentation
    
    For Each P In Application.Presentations
        SlideShow_Setup P
    Next
    W.Activate
    For Each P In Application.Presentations
        UpdateSlideTitles P
    Next
    W.Activate
    For Each P In Application.Presentations
        LoadPresentationView P
    Next
    W.Activate
    
    UpdateApplicationView True
    
    SetApplicationView
    
    NavigatorForm.StartUpPosition = 0
    NavigatorForm.Left = Application.Left - NavigatorForm.Width
    NavigatorForm.Top = Application.Top
    NavigatorForm.Height = Application.Height
    NavigatorForm.FrameSlideSelection.Height = _
        NavigatorForm.Height - _
        NavigatorForm.FrameSlideSelection.Top - 18
    NavigatorForm.ControlSlideSelectionList.Height = _
        NavigatorForm.FrameSlideSelection.Height - _
        NavigatorForm.ControlSlideSelectionList.Top - 12
    NavigatorForm.ControlSlideSelectionNumber.Caption = ""
    NavigatorForm.ControlSlideSelectionTitle.Caption = ""
    
    NavigatorForm.Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim W As DocumentWindow
    Dim P As Presentation
    Dim Response As Long
    
    Response = MsgBox( _
        buttons:= _
            vbYesNo + vbDefaultButton2 + vbExclamation, _
        Title:= _
            ProjectNamePretty, _
        Prompt:= _
            "Are you sure you want to exit the Navigator?")

    If (Response = vbYes) Then
        Cancel = 0
        
        NavigatorForm.Hide
        
        Set W = Application.ActiveWindow
        For Each P In Application.Presentations
            UnloadPresentationView P
        Next
        W.Activate
        
        UpdateApplicationView False
    Else
        Cancel = 1
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub UserForm_Activate()
    NavigatorFormLocked = True
    NavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub ControlSlideShowLoad_Click()
    If (NavigatorValid = False) Then
        Exit Sub
    End If
    SlideShowLoad Application.ActiveWindow
    NavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub ControlSlideShowHide_Click()
    If (NavigatorValid = False) Then
        Exit Sub
    End If
    SlideShowHide Application.ActiveWindow
    NavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub ControlSlideShowRun_Click()
    If (NavigatorValid = False) Then
        Exit Sub
    End If
    SlideShowRun Application.ActiveWindow
    NavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub ControlSlideShowPause_Click()
    If (NavigatorValid = False) Then
        Exit Sub
    End If
    SlideShowPause Application.ActiveWindow
    NavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub ControlSlideShowPrevEffect_Click()
    If (NavigatorValid = False) Then
        Exit Sub
    End If
    SlideShowPrevEffect Application.ActiveWindow
    NavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub ControlSlideShowNextEffect_Click()
    If (NavigatorValid = False) Then
        Exit Sub
    End If
    SlideShowNextEffect Application.ActiveWindow
    NavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub ControlGeneralExit_Click()
    Unload NavigatorForm
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub ControlGeneralHelp_Click()
    Dim HelpFile As String
    
    HelpFile = Help_GetHelpFileName(True)
    
    If (HelpFile = "") Then
        Exit Sub
     End If
    
    Application.Help HelpFile, IDH_Topic_WSACommandBarNavigator
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub ControlPresentationSelectionPrev_Click()
    If (NavigatorValid = False) Then
        Exit Sub
    End If
    PresentationSelectionPrev Application.ActiveWindow
    NavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub ControlPresentationSelectionNext_Click()
    If (NavigatorValid = False) Then
        Exit Sub
    End If
    PresentationSelectionNext Application.ActiveWindow
    NavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub ControlSlideSelectionClear_Click()
    If (NavigatorValid = False) Then
        Exit Sub
    End If
    SlideSelectionClear Application.ActiveWindow
    NavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub ControlSlideSelectionList_Click()
    If (NavigatorValid = False) Then
        Exit Sub
    End If
    SlideSelectionUpdate Application.ActiveWindow
    NavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'   This event handler forces the 'Empty' frame to hold the focus.
'-------------------------------------------------------------------------------
Private Sub FrameEmpty_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = NavigatorFormLocked
    If (Cancel = True) Then
        NavigatorValid
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub FrameEmpty_KeyPress(ByVal KeyASCII As MSForms.ReturnInteger)
    If (NavigatorValid = False) Then
        Exit Sub
    End If
    
    Dim W As DocumentWindow
    Set W = Application.ActiveWindow
    
    Dim FilterText As String
    
    If (NavigatorForm.ControlSlideSelectionNumber <> "") Then
        FilterText = NavigatorForm.ControlSlideSelectionNumber.Caption
    Else
        FilterText = NavigatorForm.ControlSlideSelectionTitle.Caption
    End If
    
    Select Case KeyASCII
        Case 8:                   ' <backspace>
            If (Len(FilterText) > 0) Then
                FilterText = Left(FilterText, Len(FilterText) - 1)
            End If
        Case 32 To 127:           ' <space> or printable character
            FilterText = FilterText & Chr(KeyASCII)
        Case Else:
    End Select
            
    If (IsNumeric(FilterText) = True) Then
        NavigatorForm.ControlSlideSelectionNumber.Caption = FilterText
        NavigatorForm.ControlSlideSelectionTitle.Caption = ""
    Else
        NavigatorForm.ControlSlideSelectionNumber.Caption = ""
        NavigatorForm.ControlSlideSelectionTitle.Caption = FilterText
    End If
    
    UpdateSlideList W
    UpdateEnabledControls W
    NavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub FrameEmpty_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal KeyModifier As Integer)
    If (NavigatorValid = False) Then
        Exit Sub
    End If
    
    Dim W As DocumentWindow
    Set W = Application.ActiveWindow
    
    Dim SHIFT As Boolean
    Dim Control As Boolean
    Dim ALTERNATE As Boolean
    
    SHIFT = KeyModifier And 1
    Control = KeyModifier And 2
    ALTERNATE = KeyModifier And 4
    
    If ((SHIFT = False) And (Control = False) And (ALTERNATE = False)) Then
        Select Case KeyCode
            Case 13:                    ' RETURN
                SlideShowLoad W
            Case 27:                    ' ESCAPE
                Unload NavigatorForm
            Case 112:                   ' F1
                Menu_OnActionHelpHelp
            Case 46:                    ' DELETE
                SlideSelectionClear W
            Case 37:                    ' LEFT_ARROW
                PresentationSelectionPrev W
            Case 39:                    ' RIGHT_ARROW
                PresentationSelectionNext W
            Case 38:                    ' UP_ARROW
                SlideSelectionPrev W
            Case 40:                    ' DOWN_ARROW
                SlideSelectionNext W
        End Select
    ElseIf ((SHIFT = False) And (Control = True) And (ALTERNATE = False)) Then
        Select Case KeyCode
            Case 72, 83:                ' "H", "S"
                SlideShowHide W
            Case 82:                    ' "R"
                SlideShowRun W
            Case 80:                    ' "P"
                SlideShowPause W
            Case 38:                    ' UP_ARROW
                SlideShowPrevEffect W
            Case 40:                    ' DOWN_ARROW
                SlideShowNextEffect W
        End Select
    End If

    NavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub UpdatePresentationName(ByVal W As DocumentWindow)
    Dim Name As String
    If (ActiveWindowExists = False) Then
        Name = ""
    Else
        Name = W.Presentation.Name
    End If
    If (Len(Name) >= 4) Then
        If (LCase(Right(Name, 4)) = ".ppt") Then
            Name = Left(Name, Len(Name) - 4)
        End If
    End If
    NavigatorForm.ControlPresentationSelectionName.Caption = Name
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub UpdateSlideList(ByVal W As DocumentWindow)
    Dim P As Presentation
    Dim FilterText As String
    Dim FilterTextLen As Long
    Dim SlideIndex As Long
    Dim SlideTitle As String
    Dim SelectedSlideIndex As Long
    Dim ListIndex As Long
    Dim ListCount As Long
    
    Set P = W.Presentation
    
    NavigatorForm.ControlSlideSelectionList.Clear
    
    If (ActiveWindowSlideExists(W) = False) Then
        Exit Sub
    End If
    
    If (ActiveSlideExists(W) = False) Then
        W.View.Slide = P.slides(1)
    End If
    SelectedSlideIndex = ActiveSlide(W).SlideIndex
    If (NavigatorForm.ControlSlideSelectionNumber <> "") Then
        FilterText = NavigatorForm.ControlSlideSelectionNumber.Caption
    Else
        FilterText = NavigatorForm.ControlSlideSelectionTitle.Caption
    End If
    FilterText = CleanEverything(FilterText)
    FilterTextLen = Len(FilterText)
    ListIndex = -1
    '
    ' Since there is no filter text, all slides are in the list.
    ' As a result, the comparisons can be bypassed.
    '
    If (FilterTextLen = 0) Then
        ListCount = P.slides.Count
        ListIndex = SelectedSlideIndex - 1
        If (ListCount > 0) Then
            ReDim List(ListCount - 1, 1) As String
            For SlideIndex = P.slides.Count To 1 Step -1
                ListCount = ListCount - 1
                List(ListCount, 0) = SlideIndex
                List(ListCount, 1) = P.slides(SlideIndex).Tags("WorshipServiceAssistant_TitleDisplay")
            Next
            NavigatorForm.ControlSlideSelectionList.List() = List
        End If
    '
    ' Since the filter text is numeric, assume that it is a slide number.
    ' As a result, the looping and comparisons can be bypassed.
    '
    ElseIf (IsNumeric(FilterText)) Then
        If ((FilterText > 0) And (FilterText <= P.slides.Count)) Then
            ListCount = 1
        Else
            ListCount = 0
        End If
        ListIndex = 0
        If (ListCount > 0) Then
            ReDim List(0, 1) As String
            SlideIndex = FilterText
            ListCount = ListCount - 1
            List(ListCount, 0) = SlideIndex
            List(ListCount, 1) = P.slides(SlideIndex).Tags("WorshipServiceAssistant_TitleDisplay")
            NavigatorForm.ControlSlideSelectionList.List() = List
        End If
    '
    ' Since the filter text is non-zero and non-numeric, assume that it is a
    ' slide title.  As a result, all the looping an comparisons must be done.
    '
    Else
        '
        ' Adding new elements to a ListBox list takes time.  As a result, it is
        ' faster determine the number of elements, build an array with the elements,
        ' as assign the array to the ListBox list.
        '
        ListCount = 0
        For SlideIndex = 1 To P.slides.Count Step 1
            If (Left(P.slides(SlideIndex).Tags("WorshipServiceAssistant_TitleMatch"), FilterTextLen) = FilterText) Then
                ListCount = ListCount + 1
            End If
        Next
        If (ListCount > 0) Then
            ReDim List(ListCount - 1, 1) As String
            For SlideIndex = P.slides.Count To 1 Step -1
                If (Left(P.slides(SlideIndex).Tags("WorshipServiceAssistant_TitleMatch"), FilterTextLen) = FilterText) Then
                    ListCount = ListCount - 1
                    List(ListCount, 0) = SlideIndex
                    List(ListCount, 1) = P.slides(SlideIndex).Tags("WorshipServiceAssistant_TitleDisplay")
                    If (SlideIndex = SelectedSlideIndex) Then
                        ListIndex = ListCount
                    End If
                End If
            Next
            NavigatorForm.ControlSlideSelectionList.List() = List
        End If
    End If
    
    If (NavigatorForm.ControlSlideSelectionList.ListCount > 0) Then
        If (ListIndex >= 0) Then
            NavigatorForm.ControlSlideSelectionList.ListIndex = ListIndex
        Else
            NavigatorForm.ControlSlideSelectionList.ListIndex = 0
        End If
        NavigatorForm.ControlSlideSelectionList.TopIndex = _
            NavigatorForm.ControlSlideSelectionList.ListIndex
        W.View.Slide = P.slides(Val(NavigatorForm.ControlSlideSelectionList.Text))
    Else
        W.View.Slide = P.slides(1)
    End If
End Sub
