VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NavigatorForm 
   Caption         =   "Navigator"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
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
'     (1) Added controls for displaying a text banner.
'     (2) Changed so that the title will match the title displayed in the
'         list box.
'   1.00.0002:
'     (1) Fixed a bug that would allow the UserForm_QueryClose processing
'         even though the UserForm_Initialize processing failed.
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
' Indicates that the Navigator Form has loaded successfully.
'
Private NavigatorFormLoaded As Boolean

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
    If (NavigatorValid = False) Then
        Exit Sub
    End If
    Dim W As PowerPoint.DocumentWindow
    Set W = Application.ActiveWindow
    UpdatePresentationName W
    UpdateSlideList W
    UpdateControls W
End Sub


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub UpdateControls(ByVal W As DocumentWindow)
    '
    ' Mave general frame visible and enabled by default.
    '
    NavigatorForm.FrameGeneral.Visible = True
    NavigatorForm.ControlGeneralNextPage.Visible = True
    NavigatorForm.ControlGeneralHelp.Visible = True
    NavigatorForm.ControlGeneralExit.Visible = True
    NavigatorForm.Pages.Visible = True
    NavigatorForm.Pages("PagePresentation").Visible = True
    NavigatorForm.Pages("PageBanner").Visible = True
    NavigatorForm.FrameGeneral.Enabled = True
    NavigatorForm.ControlGeneralNextPage.Enabled = True
    NavigatorForm.ControlGeneralHelp.Enabled = True
    NavigatorForm.ControlGeneralExit.Enabled = True
    NavigatorForm.Pages.Enabled = True
    NavigatorForm.Pages("PagePresentation").Enabled = True
    NavigatorForm.Pages("PageBanner").Enabled = True
    '
    ' Disable control associated with active control.
    '
    If (NavigatorForm.Pages(NavigatorForm.Pages.Value).Name = "PagePresentation") Then
        UpdatePresentationControls W
        NavigatorForm.Caption = "Navigator - Presentation"
    ElseIf (NavigatorForm.Pages(NavigatorForm.Pages.Value).Name = "PageBanner") Then
        UpdateBannerControls W
        NavigatorForm.Caption = "Navigator - Banner"
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

Private Sub UpdatePresentationControls(ByVal W As PowerPoint.DocumentWindow)
    '
    ' Make everything visible by default.
    '
    NavigatorForm.FrameSlideShow.Visible = True
    NavigatorForm.ControlSlideShowLoad.Visible = True
    NavigatorForm.ControlSlideShowHide.Visible = True
    NavigatorForm.ControlSlideShowRun.Visible = True
    NavigatorForm.ControlSlideShowPause.Visible = True
    NavigatorForm.ControlSlideShowPrevEffect.Visible = True
    NavigatorForm.ControlSlideShowNextEffect.Visible = True
    NavigatorForm.FramePresentationSelection.Visible = True
    NavigatorForm.ControlPresentationSelectionName.Visible = True
    NavigatorForm.ControlPresentationSelectionPrev.Visible = True
    NavigatorForm.ControlPresentationSelectionNext.Visible = True
    NavigatorForm.FrameSlideSelection.Visible = True
    NavigatorForm.ControlSlideSelectionNumber.Visible = True
    NavigatorForm.ControlSlideSelectionTitle.Visible = True
    NavigatorForm.ControlSlideSelectionClear.Visible = True
    NavigatorForm.ControlSlideSelectionList.Visible = True
    
    '
    ' Make everything enabled by default.
    '
    NavigatorForm.FrameSlideShow.Enabled = True
    NavigatorForm.ControlSlideShowLoad.Enabled = True
    NavigatorForm.ControlSlideShowHide.Enabled = True
    NavigatorForm.ControlSlideShowRun.Enabled = True
    NavigatorForm.ControlSlideShowPause.Enabled = True
    NavigatorForm.ControlSlideShowPrevEffect.Enabled = True
    NavigatorForm.ControlSlideShowNextEffect.Enabled = True
    NavigatorForm.FramePresentationSelection.Enabled = True
    NavigatorForm.ControlPresentationSelectionName.Enabled = True
    NavigatorForm.ControlPresentationSelectionPrev.Enabled = True
    NavigatorForm.ControlPresentationSelectionNext.Enabled = True
    NavigatorForm.FrameSlideSelection.Enabled = True
    NavigatorForm.ControlSlideSelectionNumber.Enabled = True
    NavigatorForm.ControlSlideSelectionTitle.Enabled = True
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
    If (ActiveSlideShowExists(W.Presentation) = False) Then
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
    If (ActiveSlideShowExists(W.Presentation) = True) Then
        If (W.Presentation.SlideShowWindow.View.State = ppSlideShowBlackScreen) Then
            NavigatorForm.ControlSlideShowHide.Caption = "Show"
        End If
    End If
    '
    ' Since the slide show is running,
    ' disable the run control.
    '
    If (ActiveSlideShowExists(W.Presentation) = True) Then
        If (W.Presentation.SlideShowWindow.View.State = ppSlideShowRunning) Then
            NavigatorForm.ControlSlideShowRun.Enabled = False
        End If
    End If
    '
    ' Since the slide show is paused,
    ' disable the pause control.
    '
    If (ActiveSlideShowExists(W.Presentation) = True) Then
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
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub UpdateBannerControls(ByVal W As PowerPoint.DocumentWindow)
    '
    ' Make everything visible by default.
    '
    NavigatorForm.FrameBannerConfiguration.Visible = True
    NavigatorForm.ControlBannerConfigurationBannerDisable.Visible = True
    NavigatorForm.FrameBannerShow.Visible = True
    NavigatorForm.ControlBannerShowLoad.Visible = True
    NavigatorForm.ControlBannerShowHide.Visible = True
    NavigatorForm.FrameBannerSelection.Visible = True
    NavigatorForm.ControlBannerSelectionText.Visible = True
    NavigatorForm.ControlBannerSelectionClear.Visible = True
    
    '
    ' Make everything enabled by default.
    '
    NavigatorForm.FrameBannerConfiguration.Enabled = True
    NavigatorForm.ControlBannerConfigurationBannerDisable.Enabled = True
    NavigatorForm.FrameBannerShow.Enabled = True
    NavigatorForm.ControlBannerShowLoad.Enabled = True
    NavigatorForm.ControlBannerShowHide.Enabled = True
    NavigatorForm.FrameBannerSelection.Enabled = True
    NavigatorForm.ControlBannerSelectionText.Enabled = True
    NavigatorForm.ControlBannerSelectionClear.Enabled = True
    
    '
    ' Set Disable/Enable button caption.
    '
    If (Banner.Enabled = True) Then
        NavigatorForm.ControlBannerConfigurationBannerDisable.Caption = "Banner Disable"
    Else
        NavigatorForm.ControlBannerConfigurationBannerDisable.Caption = "Banner Enable"
    End If
    
    '
    ' Set Hide/Show button caption.
    '
    If (Banner.Visible = True) Then
        NavigatorForm.ControlBannerShowHide.Caption = "Hide"
    Else
        NavigatorForm.ControlBannerShowHide.Caption = "Show"
    End If
    
    If (Banner.Enabled = False) Then
        NavigatorForm.ControlBannerShowLoad.Enabled = False
        NavigatorForm.ControlBannerShowHide.Enabled = False
        NavigatorForm.FrameBannerShow.Enabled = False
        NavigatorForm.ControlBannerSelectionText = ""
        NavigatorForm.ControlBannerSelectionText.Enabled = False
        NavigatorForm.ControlBannerSelectionClear.Enabled = False
        NavigatorForm.FrameBannerSelection.Enabled = False
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub GeneralNextPage(ByVal W As PowerPoint.DocumentWindow)
    Dim Page As Integer
    Page = NavigatorForm.Pages.Value
    Page = Page + 1
    If (Page >= NavigatorForm.Pages.Count) Then
        Page = 0
    End If
    NavigatorFormLocked = False
    NavigatorForm.Pages.Value = Page
    NavigatorFormLocked = True
    
    UpdateControls W
    UpdateFormSize
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SlideShowLoad(ByVal W As PowerPoint.DocumentWindow)
    If ((NavigatorForm.ControlSlideSelectionNumber.Caption <> "") Or _
        (NavigatorForm.ControlSlideSelectionTitle.Caption <> "")) Then
        NavigatorForm.ControlSlideSelectionNumber.Caption = ""
        NavigatorForm.ControlSlideSelectionTitle.Caption = ""
        UpdateSlideList W
    End If
    
    If (ActiveSlideShowExists(W.Presentation) = False) Then
        NavigatorForm.Hide
        SlideShow_End
        SlideShow_Begin W
        Banner.Apply W.Presentation.SlideShowWindow
    Else
        NavigatorForm.Hide
        W.Presentation.SlideShowWindow.Activate
    End If
    
    SlideShow_Load W
        
    W.Activate
    
    UpdateControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SlideShowHide(ByVal W As PowerPoint.DocumentWindow)
    If (ActiveSlideShowExists(W.Presentation) = False) Then
        Exit Sub
    End If
    
    SlideShow_Hide W
    
    UpdateControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SlideShowRun(ByVal W As PowerPoint.DocumentWindow)
    If (ActiveSlideShowExists(W.Presentation) = False) Then
        Exit Sub
    End If
    
    SlideShow_Run W
    
    UpdateControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SlideShowPause(ByVal W As PowerPoint.DocumentWindow)
    If (ActiveSlideShowExists(W.Presentation) = False) Then
        Exit Sub
    End If
    
    SlideShow_Pause W
    
    UpdateControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SlideShowPrevEffect(ByVal W As PowerPoint.DocumentWindow)
    If (ActiveSlideShowExists(W.Presentation) = False) Then
        Exit Sub
    End If
    
    SlideShow_Prev W
    
    ControlSlideSelectionClear_Click
    
    If (ActiveSlideExists(W) = True) Then
        NavigatorForm.ControlSlideSelectionList.ListIndex = ActiveSlide(W).SlideIndex - 1
    End If
    
    UpdateControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SlideShowNextEffect(ByVal W As PowerPoint.DocumentWindow)
    If (ActiveSlideShowExists(W.Presentation) = False) Then
        Exit Sub
    End If
    
    SlideShow_Next W
    
    ControlSlideSelectionClear_Click
    
    If (ActiveSlideExists(W) = True) Then
        NavigatorForm.ControlSlideSelectionList.ListIndex = ActiveSlide(W).SlideIndex - 1
    End If
    
    UpdateControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PresentationSelectionPrev(ByVal W As PowerPoint.DocumentWindow)
    If (Application.Presentations.Count <= 1) Then
        Exit Sub
    End If
    
    Dim i As Long
    Dim J As Long
    
    For i = Application.Presentations.Count To 1 Step -1
        If (Application.Presentations(i) Is W.Presentation) Then
            Exit For
        End If
    Next
    
    J = i
    Do
        J = J - 1
        If (J < 1) Then
            J = Application.Presentations.Count
        End If
    Loop While ((J <> i) And (SlideShow_IsSlideShow(Application.Presentations(J)) = False))
    
    If (J <> i) Then
        Application.Presentations(J).Windows(1).Activate
        Set W = Application.ActiveWindow
        UpdatePresentationName W
        UpdateSlideList W
        UpdateControls W
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub PresentationSelectionNext(ByVal W As PowerPoint.DocumentWindow)
    If (Application.Presentations.Count <= 1) Then
        Exit Sub
    End If
    
    Dim i As Long
    Dim J As Long
    
    For i = 1 To Application.Presentations.Count Step 1
        If (Application.Presentations(i) Is W.Presentation) Then
            Exit For
        End If
    Next
    
    J = i
    Do
        J = J + 1
        If (J > Application.Presentations.Count) Then
            J = 1
        End If
    Loop While ((J <> i) And (SlideShow_IsSlideShow(Application.Presentations(J)) = False))
    
    If (J <> i) Then
        Application.Presentations(J).Windows(1).Activate
        Set W = Application.ActiveWindow
        UpdatePresentationName W
        UpdateSlideList W
        UpdateControls W
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SlideSelectionClear(ByVal W As PowerPoint.DocumentWindow)
    If ((NavigatorForm.ControlSlideSelectionNumber.Caption <> "") Or _
        (NavigatorForm.ControlSlideSelectionTitle.Caption <> "")) Then
        NavigatorForm.ControlSlideSelectionNumber.Caption = ""
        NavigatorForm.ControlSlideSelectionTitle.Caption = ""
        UpdateSlideList W
    End If
    UpdateControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SlideSelectionUpdate(ByVal W As PowerPoint.DocumentWindow)
    Dim sIndex As Long
    '
    ' Set slide selected in the presentation to match the
    ' slide selected in the slide list control
    '
    With NavigatorForm.ControlSlideSelectionList
        If (.ListIndex >= 0) Then
            sIndex = .List(.ListIndex, 0)
            W.View.Slide = W.Presentation.slides(sIndex)
        End If
    End With
    UpdateControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SlideSelectionPrev(ByVal W As PowerPoint.DocumentWindow)
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
Private Sub SlideSelectionNext(ByVal W As PowerPoint.DocumentWindow)
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
Private Sub BannerConfigurationBannerDisable(ByVal W As PowerPoint.DocumentWindow)
    Banner.Enabled = Not Banner.Enabled
    UpdateControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub BannerShowLoad(ByVal W As PowerPoint.DocumentWindow)
    If (Banner.Enabled = False) Then
        Exit Sub
    End If
    Banner.Load NavigatorForm.ControlBannerSelectionText.Caption
    BannerSelectionClear W
    UpdateControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub BannerShowHide(ByVal W As PowerPoint.DocumentWindow)
    If (Banner.Enabled = False) Then
        Exit Sub
    End If
    Banner.Visible = Not Banner.Visible
    UpdateControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub BannerSelectionClear(ByVal W As PowerPoint.DocumentWindow)
    If (Banner.Enabled = False) Then
        Exit Sub
    End If
    NavigatorForm.ControlBannerSelectionText.Caption = ""
    UpdateControls W
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Function NavigatorValid() As Boolean
    NavigatorValid = False
    If (ActiveWindowExists = False) Then
        NavigatorForm.Hide
        Exit Function
    End If
    NavigatorValid = True
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub UpdateSlideTitles(ByVal P As PowerPoint.Presentation)
    Dim S As PowerPoint.Slide
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
    Static WindowState As PowerPoint.PpWindowState
    Static Top As Long
    Static Left As Long
    Static Height As Long
    Static Width As Long
    Dim W As PowerPoint.DocumentWindow
    
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
Private Sub LoadPresentationView(ByVal P As PowerPoint.Presentation)
    Dim WindowState As PowerPoint.PpWindowState
    Dim ViewType As PowerPoint.PpViewType
    Dim SplitHorizontal As Long
    Dim SplitVertical As Long
    Dim W As PowerPoint.DocumentWindow
    Dim Saved As Office.MsoTriState
    
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
Private Sub UnloadPresentationView(ByVal P As PowerPoint.Presentation)
    Dim WindowState As PowerPoint.PpWindowState
    Dim ViewType As PowerPoint.PpViewType
    Dim SplitHorizontal As Long
    Dim SplitVertical As Long
    Dim W As PowerPoint.DocumentWindow
    
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
    Dim P As PowerPoint.Presentation
    Dim W As PowerPoint.DocumentWindow
    
    NavigatorFormLoaded = False
    
    If (Presentation.Exists = False) Then
        Unload NavigatorForm
        Exit Sub
    End If
    
    NavigatorFormLoaded = True
    
    NavigatorFormLocked = True
    
    If (Presentation.IsPresentation(Application.ActiveWindow.Presentation) = False) Then
        If (Presentation.Exists = True) Then
            For Each P In Application.Presentations
                If (Presentation.IsPresentation(P) = True) Then
                    If (P.Windows.Count > 0) Then
                        P.Windows(1).Activate
                    End If
                End If
            Next
        End If
    End If
    
    Set W = Application.ActiveWindow
    
    Banner.Create
    
    For Each P In Application.Presentations
        If (SlideShow_IsSlideShow(P) = True) Then
            SlideShow_Setup P
        End If
    Next
    W.Activate
    For Each P In Application.Presentations
        If (SlideShow_IsSlideShow(P) = True) Then
            UpdateSlideTitles P
        End If
    Next
    W.Activate
    For Each P In Application.Presentations
        If (SlideShow_IsSlideShow(P) = True) Then
            LoadPresentationView P
        End If
    Next
    W.Activate
    
    UpdateApplicationView True
    
    SetApplicationView
    
    NavigatorForm.Pages.Value = NavigatorForm.Pages("PagePresentation").Index
    
    NavigatorForm.ControlSlideSelectionNumber.Caption = ""
    NavigatorForm.ControlSlideSelectionTitle.Caption = ""
    NavigatorForm.ControlBannerSelectionText.Caption = ""
    
    NavigatorForm.StartUpPosition = 0
    NavigatorForm.Left = Application.Left - NavigatorForm.Width
    NavigatorForm.Top = Application.Top
    
    UpdateFormSize
    
    NavigatorForm.Refresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim W As PowerPoint.DocumentWindow
    Dim P As PowerPoint.Presentation
    Dim Response As Long
    
    '
    ' Exit without unloading form, because the form was never loaded.
    '
    If (NavigatorFormLoaded = False) Then
        Exit Sub
    End If
    
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
            If (SlideShow_IsSlideShow(P) = True) Then
                UnloadPresentationView P
            End If
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
        Call Help.HtmlHelp( _
        0&, _
        HelpFile, _
        Help.HH_DISPLAY_TOPIC, _
        Help.IDH_TopicPath_WSACommandBarNavigator)
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub ControlGeneralNextPage_Click()
    GeneralNextPage Application.ActiveWindow
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
'-------------------------------------------------------------------------------
Private Sub ControlBannerConfigurationBannerDisable_Click()
    If (NavigatorValid = False) Then
        Exit Sub
    End If
    BannerConfigurationBannerDisable Application.ActiveWindow
    NavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub ControlBannerShowLoad_Click()
    If (NavigatorValid = False) Then
        Exit Sub
    End If
    BannerShowLoad Application.ActiveWindow
    NavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub ControlBannerShowHide_Click()
    If (NavigatorValid = False) Then
        Exit Sub
    End If
    BannerShowHide Application.ActiveWindow
    NavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub ControlBannerSelectionClear_Click()
    If (NavigatorValid = False) Then
        Exit Sub
    End If
    BannerSelectionClear Application.ActiveWindow
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
    
    Dim W As PowerPoint.DocumentWindow
    Set W = Application.ActiveWindow
        
    If (NavigatorForm.Pages(NavigatorForm.Pages.Value).Name = "PagePresentation") Then
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
    ElseIf (NavigatorForm.Pages(NavigatorForm.Pages.Value).Name = "PageBanner") Then
        Dim BannerText As String
        
        If (Banner.Enabled = True) Then
            BannerText = NavigatorForm.ControlBannerSelectionText.Caption
            Select Case KeyASCII
                Case 8:                   ' <backspace>
                    If (Len(BannerText) > 0) Then
                        BannerText = Left(BannerText, Len(BannerText) - 1)
                    End If
                Case 32 To 127:           ' <space> or printable character
                    BannerText = BannerText & Chr(KeyASCII)
                Case Else:
            End Select
            NavigatorForm.ControlBannerSelectionText.Caption = BannerText
        End If
    End If
    
    UpdateControls W
    NavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub FrameEmpty_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal KeyModifier As Integer)
    If (NavigatorValid = False) Then
        Exit Sub
    End If
    
    Dim W As PowerPoint.DocumentWindow
    Set W = Application.ActiveWindow
    
    Dim SHIFT As Boolean
    Dim CONTROL As Boolean
    Dim ALTERNATE As Boolean
    
    SHIFT = KeyModifier And 1
    CONTROL = KeyModifier And 2
    ALTERNATE = KeyModifier And 4
    
    If (NavigatorForm.Pages(NavigatorForm.Pages.Value).Name = "PagePresentation") Then
        If ((SHIFT = False) And (CONTROL = False) And (ALTERNATE = False)) Then
            Select Case KeyCode
                Case 13:                    ' RETURN
                    SlideShowLoad W
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
        ElseIf ((SHIFT = False) And (CONTROL = True) And (ALTERNATE = False)) Then
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
    ElseIf (NavigatorForm.Pages(NavigatorForm.Pages.Value).Name = "PageBanner") Then
        If ((SHIFT = False) And (CONTROL = False) And (ALTERNATE = False)) Then
            Select Case KeyCode
                Case 13:                    ' RETURN
                    BannerShowLoad W
                Case 46:                    ' DELETE
                    BannerSelectionClear W
            End Select
        ElseIf ((SHIFT = False) And (CONTROL = True) And (ALTERNATE = False)) Then
            Select Case KeyCode
                Case 72, 83:                ' "H", "S"
                    BannerShowHide W
                Case 68, 69:                ' "D", "E"
                    BannerConfigurationBannerDisable W
            End Select
        End If
    End If
    If ((SHIFT = False) And (CONTROL = False) And (ALTERNATE = False)) Then
        Select Case KeyCode
            Case 9:                     ' TAB
                GeneralNextPage W
            Case 27:                    ' ESCAPE
                Unload NavigatorForm
            Case 112:                   ' F1
                ControlGeneralHelp_Click
        End Select
    End If
    NavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub UpdateFormSize()
    Const FormOverhead As Integer = 22
    Const PageOverhead As Integer = 0
    Const FrameOverhead As Integer = 14
    If (NavigatorForm.Pages(NavigatorForm.Pages.Value).Name = "PagePresentation") Then
        NavigatorForm.Height = Application.Height
        NavigatorForm.Pages.Height = _
            NavigatorForm.Height - _
            NavigatorForm.Pages.Top - _
            FormOverhead
        NavigatorForm.FrameSlideSelection.Height = _
            NavigatorForm.Pages.Height - _
            PageOverhead - _
            NavigatorForm.FrameSlideSelection.Top
        NavigatorForm.ControlSlideSelectionList.Height = _
            NavigatorForm.FrameSlideSelection.Height - _
            FrameOverhead - _
            NavigatorForm.ControlSlideSelectionList.Top
    ElseIf (NavigatorForm.Pages(NavigatorForm.Pages.Value).Name = "PageBanner") Then
        NavigatorForm.Height = _
            FormOverhead + _
            PageOverhead + _
            NavigatorForm.Pages.Top + _
            NavigatorForm.FrameBannerSelection.Top + _
            NavigatorForm.FrameBannerSelection.Height
        NavigatorForm.Pages.Height = _
            NavigatorForm.Height - _
            FormOverhead
    Else
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub UpdatePresentationName(ByVal W As PowerPoint.DocumentWindow)
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
Private Sub UpdateSlideList(ByVal W As PowerPoint.DocumentWindow)
    Dim P As PowerPoint.Presentation
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
            If ((Left(P.slides(SlideIndex).Tags("WorshipServiceAssistant_TitleMatch"), FilterTextLen) = FilterText) Or _
                (Left(P.slides(SlideIndex).Tags("WorshipServiceAssistant_TitleDisplay"), FilterTextLen) = FilterText)) Then
                ListCount = ListCount + 1
            End If
        Next
        If (ListCount > 0) Then
            ReDim List(ListCount - 1, 1) As String
            For SlideIndex = P.slides.Count To 1 Step -1
                If ((Left(P.slides(SlideIndex).Tags("WorshipServiceAssistant_TitleMatch"), FilterTextLen) = FilterText) Or _
                    (Left(P.slides(SlideIndex).Tags("WorshipServiceAssistant_TitleDisplay"), FilterTextLen) = FilterText)) Then
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
    
    With NavigatorForm.ControlSlideSelectionList
        If (.ListCount > 0) Then
            If (ListIndex >= 0) Then
                .ListIndex = ListIndex
            Else
                .ListIndex = 0
            End If
            .TopIndex = .ListIndex
            W.View.Slide = P.slides(Val(.List(.ListIndex, 0)))
        Else
            W.View.Slide = P.slides(1)
        End If
    End With
End Sub
