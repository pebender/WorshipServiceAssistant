VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNavigator 
   Caption         =   "Navigator"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   OleObjectBlob   =   "frmNavigator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.frmNavigator
'
' Description:
'   The form implements the Navigator.
'
'   The form contains an empty frame (fraEmpty) to which the focus is locked.
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
'   1.03.0002:
'     (1) Changed the source code so that it follows Microsoft's
'         Visual Basic coding conventions.
'     (2) Fixed a bug that the Slide Selection list to display the
'         incorrect slide titles.
'     (3) Changed the style of the frame borders.
'     (4) Made the Slide Selection frame border color dependent on
'         the Slide Selection mode.
'   1.03.0001:
'     (1) Fixed a bug where the slide body match was done even though the
'         the results were not be used.
'     (2) Modified LoadPresentationView to work around the PowerPoint 2002
'         VBA API bug described in Microsoft knowledge base article Q285436.
'     (2) Changed LoadPresentationView to set the view to ppViewSlide rather
'         than ppViewNormal, eliminating the need to save and restore the
'         window's SplitHorizontal and SplitVertical properties.
'   1.03.0000:
'     (1) Added support for matching the filter to the slide body
'         in the  Navigator Slide Selection frame.
'     (2) Added the "Mode" button and "Mode Name" display
'         to the Navigatory Slide Selection frame in order to allow the
'         operator to select the filter mode.
'     (3) Renamed Navigator Slide Selection Title to Navigator Slide
'         Selection Text.
'   1.02.0000:
'     (1) Added code that configures PowerPoint options to the most
'         appropiate values for running the Navigator.
'   1.01.0006:
'     (1) Worked around a PowerPoint 2002 bug that would cause
'         PowerPoint to prompt the user to save the banner presentation
'         even though it was already marked as saved.
'   1.01.0005:
'     (1) Changed UserForm_Initialize so that it hide all visible
'         floating and pop-up command bars, because floating and pop-up
'         command bars can get in the way of the Navigator form.
'     (2) Changed code so that default button would be highlighted in
'         both the Presentation page and the Banner page.
'     (3) Changed code so that 'm and m will match when searching
'         slide titles.
'   1.01.0004:
'     (1) Fixed Slide Show Show/Hide button that caused the button not
'         to update.
'     (2) Fixed banner checking bug in mSlideShowLoad.
'   1.01.0003:
'     (1) Fixed errors in control tip text.
'     (2) Fixed errors in form re-sizing in UpdateFormSize.
'     (3) Changed LoadPresentationView and UnloadPresentationView
'         so that the slide miniture window will not be displayed and
'         so that the slides will zoom to fit.
'     (4) Changed form so that it does not resize based on page.
'     (5) Changed mSlideShowLoad so that transitions between presentations
'         would be more seamless.
'   1.01.0002:
'     (1) Added banner color support.
'     (2) Replaced NavigatorForm with Me.
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
Private mblnNavigatorFormLoaded As Boolean

'
' Indicates whether or not the Navigator form's focus is locked to the empty
' frame.  Normally the focus is locked to the empty frame.  However, sometimes
' it is necessary to change the focus in order to force the Navigator form
' to be the active window.  During this time, the focus must be unlocked.
'
Private mblnNavigatorFormLocked As Boolean

Private msswActive As PowerPoint.SlideShowWindow

'===============================================================================
' Public Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub gRefresh _
( _
)
    Dim dwActive As PowerPoint.DocumentWindow
    
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If
    
    Set dwActive = Application.ActiveWindow
    mPresentationNameUpdate dwActive
    mSlideListUpdate dwActive
    mControlsUpdate dwActive
End Sub


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mControlsUpdate _
( _
    ByRef dwDocumentWindow As DocumentWindow _
)
    '
    ' Mave general frame visible and enabled by default.
    '
    Me.fraGeneral.Visible = True
    Me.cmdGeneralPageNext.Visible = True
    Me.cmdGeneralHelp.Visible = True
    Me.cmdGeneralExit.Visible = True
    Me.tabPages.Visible = True
    Me.tabPages("tabPresentation").Visible = True
    Me.tabPages("tabBanner").Visible = True
    Me.fraGeneral.Enabled = True
    Me.cmdGeneralPageNext.Enabled = True
    Me.cmdGeneralHelp.Enabled = True
    Me.cmdGeneralExit.Enabled = True
    Me.tabPages.Enabled = True
    Me.tabPages("tabPresentation").Enabled = True
    Me.tabPages("tabBanner").Enabled = True
    '
    ' Disable control associated with active control.
    '
    If (Me.tabPages(Me.tabPages.Value).Name = "tabPresentation") Then
        mPresentationControlsUpdate dwDocumentWindow
        Me.Caption = "Navigator - Presentation"
    ElseIf (Me.tabPages(Me.tabPages.Value).Name = "tabBanner") Then
        mBannerControlsUpdate dwDocumentWindow
        Me.Caption = "Navigator - Banner"
    End If
    
    '
    ' Set focus and try to make sure that the Navigator is the
    ' active window. Changing the focus seems to accomplish it.
    '
    mblnNavigatorFormLocked = False
    Me.fraGeneral.SetFocus
    Me.fraEmpty.SetFocus
    mblnNavigatorFormLocked = True
    
    '
    ' Set pointer.
    '
    Me.MousePointer = fmMousePointerArrow
    Me.Repaint
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mPresentationControlsUpdate _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    '
    ' Make everything visible by default.
    '
    Me.fraSlideShow.Visible = True
    Me.cmdSlideShowLoad.Visible = True
    Me.cmdSlideShowHide.Visible = True
    Me.cmdSlideShowRun.Visible = True
    Me.cmdSlideShowPause.Visible = True
    Me.cmdSlideShowEffectPrev.Visible = True
    Me.cmdSlideShowEffectNext.Visible = True
    Me.fraPresentationSelection.Visible = True
    Me.lblPresentationSelectionName.Visible = True
    Me.cmdPresentationSelectionPrev.Visible = True
    Me.cmdPresentationSelectionNext.Visible = True
    Me.fraSlideSelection.Visible = True
    Me.cmdSlideSelectionMode.Visible = True
    Me.lblSlideSelectionMode.Visible = True
    Me.lblSlideSelectionNumber.Visible = True
    Me.lblSlideSelectionText.Visible = True
    Me.cmdSlideSelectionClear.Visible = True
    Me.lstSlideSelectionList.Visible = True
    
    '
    ' Make everything enabled by default.
    '
    Me.fraSlideShow.Enabled = True
    Me.cmdSlideShowLoad.Enabled = True
    Me.cmdSlideShowHide.Enabled = True
    Me.cmdSlideShowRun.Enabled = True
    Me.cmdSlideShowPause.Enabled = True
    Me.cmdSlideShowEffectPrev.Enabled = True
    Me.cmdSlideShowEffectNext.Enabled = True
    Me.fraPresentationSelection.Enabled = True
    Me.lblPresentationSelectionName.Enabled = True
    Me.cmdPresentationSelectionPrev.Enabled = True
    Me.cmdPresentationSelectionNext.Enabled = True
    Me.fraSlideSelection.Enabled = True
    Me.cmdSlideSelectionMode.Enabled = True
    Me.lblSlideSelectionMode.Enabled = True
    Me.lblSlideSelectionNumber.Enabled = True
    Me.lblSlideSelectionText.Enabled = True
    Me.cmdSlideSelectionClear.Enabled = True
    Me.lstSlideSelectionList.Enabled = True
    
    '
    ' Set default button.
    '
    Me.cmdSlideShowLoad.Default = True
    
    '
    ' Set default button captions.
    '
    Me.cmdSlideShowHide.Caption = "Hide"
    
    '
    ' Since there are no slides,
    ' disable slide controls.
    '
    If (modActive.gblnActiveWindowSlideExists(dwDocumentWindow) = False) Then
        Me.cmdSlideShowLoad.Enabled = False
        Me.cmdSlideShowHide.Enabled = False
        Me.cmdSlideShowRun.Enabled = False
        Me.cmdSlideShowPause.Enabled = False
        Me.cmdSlideShowEffectPrev.Enabled = False
        Me.cmdSlideShowEffectNext.Enabled = False
    End If
    '
    ' Since there are no slides in the slide list,
    ' disable slide show controls.
    '
    If (Me.lstSlideSelectionList.ListCount = 0) Then
        Me.cmdSlideShowLoad.Enabled = False
        Me.cmdSlideShowHide.Enabled = False
        Me.cmdSlideShowRun.Enabled = False
        Me.cmdSlideShowPause.Enabled = False
        Me.cmdSlideShowEffectPrev.Enabled = False
        Me.cmdSlideShowEffectNext.Enabled = False
    End If
    '
    ' Since there are no slides in the slide list,
    ' disable slide load control.
    '
    If (modActive.gblnActiveSlideShowExists(dwDocumentWindow.Presentation) = False) Then
        Me.cmdSlideShowHide.Enabled = False
        Me.cmdSlideShowRun.Enabled = False
        Me.cmdSlideShowPause.Enabled = False
        Me.cmdSlideShowEffectPrev.Enabled = False
        Me.cmdSlideShowEffectNext.Enabled = False
    End If
    '
    ' Since the slide show is hidden,
    ' relable the hide control as Show.
    '
    If (modActive.gblnActiveSlideShowExists(dwDocumentWindow.Presentation) = True) Then
        If (dwDocumentWindow.Presentation.SlideShowWindow.View.State = PowerPoint.ppSlideShowBlackScreen) Then
            Me.cmdSlideShowHide.Caption = "Show"
        End If
    End If
    '
    ' Since the slide show is running,
    ' disable the run control.
    '
    If (modActive.gblnActiveSlideShowExists(dwDocumentWindow.Presentation) = True) Then
        If (dwDocumentWindow.Presentation.SlideShowWindow.View.State = PowerPoint.ppSlideShowRunning) Then
            Me.cmdSlideShowRun.Enabled = False
        End If
    End If
    '
    ' Since the slide show is paused,
    ' disable the pause control.
    '
    If (modActive.gblnActiveSlideShowExists(dwDocumentWindow.Presentation) = True) Then
        If (dwDocumentWindow.Presentation.SlideShowWindow.View.State = PowerPoint.ppSlideShowPaused) Then
            Me.cmdSlideShowPause.Enabled = False
        End If
    End If
    '
    ' Update Slide Selection Mode dependent controls.
    '
    Select Case Me.cmdSlideSelectionMode.Tag
        Case "0"
            Me.lblSlideSelectionMode.Caption = "Filter matching Number and Title"
            Me.fraSlideSelection.BorderColor = VBA.vbInactiveBorder
        Case "1"
            Me.lblSlideSelectionMode.Caption = "Filter matching Number and Title and Body"
            Me.fraSlideSelection.BorderColor = VBA.vbRed
        Case Else
            Me.lblSlideSelectionMode.Caption = ""
            Me.fraSlideSelection.BorderColor = VBA.vbInactiveBorder
    End Select
    '
    ' Since the slide filter is clear,
    ' disable the slide filter clear control.
    '
    If ((Me.lblSlideSelectionNumber.Caption = "") And _
        (Me.lblSlideSelectionText.Caption = "")) Then
        Me.cmdSlideSelectionClear.Enabled = False
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mBannerControlsUpdate _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    '
    ' Make everything visible by default.
    '
    Me.fraBannerConfiguration.Visible = True
    Me.cmdBannerConfigurationBannerDisable.Visible = True
    Me.fraBannerShow.Visible = True
    Me.cmdBannerShowLoad.Visible = True
    Me.cmdBannerShowHide.Visible = True
    Me.fraBannerColor.Visible = True
    Me.lblBannerColorName.Visible = True
    Me.cmdBannerColorPrev.Visible = True
    Me.cmdBannerColorNext.Visible = True
    Me.fraBannerSelection.Visible = True
    Me.lblBannerSelectionText.Visible = True
    Me.cmdBannerSelectionClear.Visible = True
    
    '
    ' Make everything enabled by default.
    '
    Me.fraBannerConfiguration.Enabled = True
    Me.cmdBannerConfigurationBannerDisable.Enabled = True
    Me.fraBannerShow.Enabled = True
    Me.cmdBannerShowLoad.Enabled = True
    Me.cmdBannerShowHide.Enabled = True
    Me.fraBannerColor.Enabled = True
    Me.lblBannerColorName.Enabled = True
    Me.cmdBannerColorPrev.Enabled = True
    Me.cmdBannerColorNext.Enabled = True
    Me.fraBannerSelection.Enabled = True
    Me.lblBannerSelectionText.Enabled = True
    Me.cmdBannerSelectionClear.Enabled = True
    
    '
    ' Set default button.
    '
    Me.cmdBannerShowLoad.Default = True
    
    '
    ' Set button captions.
    '
    If (modBanner.gblnEnabled = True) Then
        Me.cmdBannerConfigurationBannerDisable.Caption = "Banner Disable"
    Else
        Me.cmdBannerConfigurationBannerDisable.Caption = "Banner Enable"
    End If
    
    '
    ' Set Hide/Show button caption.
    '
    If (modBanner.gblnVisible = True) Then
        Me.cmdBannerShowHide.Caption = "Hide"
    Else
        Me.cmdBannerShowHide.Caption = "Show"
    End If
    
    '
    ' Set banner color name colors.
    '
    If (modBanner.gblnEnabled = True) Then
        With Me.lblBannerColorName
            .BackColor = VBA.RGB(0, 0, 0)
            .ForeColor = VBA.RGB(.List(.ListIndex, 1), .List(.ListIndex, 2), .List(.ListIndex, 3))
        End With
    Else
        With Me.lblBannerColorName
            .BackColor = VBA.RGB(0, 0, 0)
            .ForeColor = VBA.RGB(0, 0, 0)
        End With
    End If
    
    If (modBanner.gblnEnabled = False) Then
        Me.cmdBannerShowLoad.Enabled = False
        Me.cmdBannerShowHide.Enabled = False
        Me.fraBannerShow.Enabled = False
        Me.lblBannerColorName.Enabled = False
        Me.cmdBannerColorPrev.Enabled = False
        Me.cmdBannerColorNext.Enabled = False
        Me.lblBannerSelectionText = ""
        Me.lblBannerSelectionText.Enabled = False
        Me.cmdBannerSelectionClear.Enabled = False
        Me.fraBannerSelection.Enabled = False
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mGeneralPageNext _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    Dim lngPage As Long
    
    lngPage = Me.tabPages.Value
    lngPage = lngPage + 1
    If (lngPage >= Me.tabPages.Count) Then
        lngPage = 0
    End If
    mblnNavigatorFormLocked = False
    Me.tabPages.Value = lngPage
    mblnNavigatorFormLocked = True
    
    mControlsUpdate dwDocumentWindow
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mSlideShowLoad _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    Dim sswSlideShowWindow As PowerPoint.SlideShowWindow
    
    If ((Me.lblSlideSelectionNumber.Caption <> "") Or _
        (Me.lblSlideSelectionText.Caption <> "")) Then
        Me.lblSlideSelectionNumber.Caption = ""
        Me.lblSlideSelectionText.Caption = ""
        mSlideListUpdate dwDocumentWindow
    End If
    
    Me.Hide
    
    '
    ' Create slide show if one does not exist.
    '
    If (modActive.gblnActiveSlideShowExists(dwDocumentWindow.Presentation) = False) Then
        modSlideShow.gBegin dwDocumentWindow
        modBanner.gApply dwDocumentWindow.Presentation.SlideShowWindow
    End If
    
    modSlideShow.gLoad dwDocumentWindow
    
    dwDocumentWindow.Presentation.SlideShowWindow.Activate
        
' Exit Sub
    
    If (Not (dwDocumentWindow.Presentation.SlideShowWindow Is msswActive)) Then
        Set msswActive = dwDocumentWindow.Presentation.SlideShowWindow
        '
        ' Black other slide shows.
        ' This will stop the slide shows from running.
        '
        For Each sswSlideShowWindow In Application.SlideShowWindows
            If ((Not (sswSlideShowWindow Is dwDocumentWindow.Presentation.SlideShowWindow)) And _
                (Not modBanner.gblnIsBanner(sswSlideShowWindow.Presentation))) Then
                sswSlideShowWindow.View.State = PowerPoint.ppSlideShowBlackScreen
            End If
        Next
    End If
    
    dwDocumentWindow.Activate
    
    mControlsUpdate dwDocumentWindow
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mSlideShowHide _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    If (modActive.gblnActiveSlideShowExists(dwDocumentWindow.Presentation) = False) Then
        Exit Sub
    End If
    
    modSlideShow.gHide dwDocumentWindow
    
    mControlsUpdate dwDocumentWindow
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mSlideShowRun _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    If (modActive.gblnActiveSlideShowExists(dwDocumentWindow.Presentation) = False) Then
        Exit Sub
    End If
    
    modSlideShow.gRun dwDocumentWindow
    
    mControlsUpdate dwDocumentWindow
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mSlideShowPause _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    If (modActive.gblnActiveSlideShowExists(dwDocumentWindow.Presentation) = False) Then
        Exit Sub
    End If
    
    modSlideShow.gPause dwDocumentWindow
    
    mControlsUpdate dwDocumentWindow
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mSlideShowEffectPrev _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    If (modActive.gblnActiveSlideShowExists(dwDocumentWindow.Presentation) = False) Then
        Exit Sub
    End If
    
    modSlideShow.gEffectPrev dwDocumentWindow
    
    cmdSlideSelectionClear_Click
    
    If (modActive.gblnActiveSlideExists(dwDocumentWindow) = True) Then
        Me.lstSlideSelectionList.ListIndex = modActive.gppActiveSlideGet(dwDocumentWindow).SlideIndex - 1
    End If
    
    mControlsUpdate dwDocumentWindow
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mSlideShowEffectNext _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    If (modActive.gblnActiveSlideShowExists(dwDocumentWindow.Presentation) = False) Then
        Exit Sub
    End If
        
    modSlideShow.gEffectNext dwDocumentWindow
    
    cmdSlideSelectionClear_Click
    
    If (modActive.gblnActiveSlideExists(dwDocumentWindow) = True) Then
        Me.lstSlideSelectionList.ListIndex = modActive.gppActiveSlideGet(dwDocumentWindow).SlideIndex - 1
    End If
    
    mControlsUpdate dwDocumentWindow
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mPresentationSelectionPrev _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    Dim lngIndexCurr As Long
    Dim lngIndexPrev As Long
    
    If (Application.Presentations.Count <= 1) Then
        Exit Sub
    End If
    
    For lngIndexCurr = Application.Presentations.Count To 1 Step -1
        If (Application.Presentations(lngIndexCurr) Is dwDocumentWindow.Presentation) Then
            Exit For
        End If
    Next
    
    lngIndexPrev = lngIndexCurr
    Do
        lngIndexPrev = lngIndexPrev - 1
        If (lngIndexPrev < 1) Then
            lngIndexPrev = Application.Presentations.Count
        End If
    Loop While ((lngIndexPrev <> lngIndexCurr) And (modSlideShow.gblnIsSlideShow(Application.Presentations(lngIndexPrev)) = False))
    
    If (lngIndexPrev <> lngIndexCurr) Then
        Application.Presentations(lngIndexPrev).Windows(1).Activate
        Set dwDocumentWindow = Application.ActiveWindow
        mPresentationNameUpdate dwDocumentWindow
        mSlideListUpdate dwDocumentWindow
        mControlsUpdate dwDocumentWindow
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mPresentationSelectionNext _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    Dim lngIndexCurr As Long
    Dim lngIndexNext As Long
    
    If (Application.Presentations.Count <= 1) Then
        Exit Sub
    End If
    
    For lngIndexCurr = 1 To Application.Presentations.Count Step 1
        If (Application.Presentations(lngIndexCurr) Is dwDocumentWindow.Presentation) Then
            Exit For
        End If
    Next
    
    lngIndexNext = lngIndexCurr
    Do
        lngIndexNext = lngIndexNext + 1
        If (lngIndexNext > Application.Presentations.Count) Then
            lngIndexNext = 1
        End If
    Loop While ((lngIndexNext <> lngIndexCurr) And (modSlideShow.gblnIsSlideShow(Application.Presentations(lngIndexNext)) = False))
    
    If (lngIndexNext <> lngIndexCurr) Then
        Application.Presentations(lngIndexNext).Windows(1).Activate
        Set dwDocumentWindow = Application.ActiveWindow
        mPresentationNameUpdate dwDocumentWindow
        mSlideListUpdate dwDocumentWindow
        mControlsUpdate dwDocumentWindow
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mSlideSelectionMode _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    Select Case Me.cmdSlideSelectionMode.Tag
        Case "0"                        ' Title
            Me.cmdSlideSelectionMode.Tag = "1"
        Case "1"                        ' Title and Body
            Me.cmdSlideSelectionMode.Tag = "0"
        Case Else
            Me.cmdSlideSelectionMode.Tag = "0"
    End Select
    
    mSlideListUpdate dwDocumentWindow
    mControlsUpdate dwDocumentWindow
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mSlideSelectionClear _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    If ((Me.lblSlideSelectionNumber.Caption <> "") Or _
        (Me.lblSlideSelectionText.Caption <> "")) Then
        Me.lblSlideSelectionNumber.Caption = ""
        Me.lblSlideSelectionText.Caption = ""
        mSlideListUpdate dwDocumentWindow
    End If
    mControlsUpdate dwDocumentWindow
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mSlideSelectionUpdate _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    Dim lngIndex As Long
    
    '
    ' Set slide selected in the presentation to match the
    ' slide selected in the slide list control
    '
    With Me.lstSlideSelectionList
        If (.ListIndex >= 0) Then
            lngIndex = .List(.ListIndex, 0)
            dwDocumentWindow.View.Slide = dwDocumentWindow.Presentation.Slides(lngIndex)
        End If
    End With
    mControlsUpdate dwDocumentWindow
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mSlideSelectionPrev _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    With Me.lstSlideSelectionList
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
Private Sub mSlideSelectionNext _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    With Me.lstSlideSelectionList
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
Private Sub mBannerConfigurationBannerDisable _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    modBanner.gblnEnabled = Not modBanner.gblnEnabled
    mControlsUpdate dwDocumentWindow
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mBannerShowLoad _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    Dim intRed As Integer
    Dim intGreen As Integer
    Dim intBlue As Integer
    
    If (modBanner.gblnEnabled = False) Then
        Exit Sub
    End If
    
    With Me.lblBannerColorName
        If (.ListIndex >= 0) Then
            intRed = .List(.ListIndex, 1)
            intGreen = .List(.ListIndex, 2)
            intBlue = .List(.ListIndex, 3)
        Else
            intRed = 0
            intGreen = 0
            intBlue = 0
        End If
    End With
    modBanner.gLoad Me.lblBannerSelectionText.Caption, intRed, intGreen, intBlue
    mBannerSelectionClear dwDocumentWindow
    mControlsUpdate dwDocumentWindow
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mBannerShowHide _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    If (modBanner.gblnEnabled = False) Then
        Exit Sub
    End If
    modBanner.gblnVisible = Not modBanner.gblnVisible
    mControlsUpdate dwDocumentWindow
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mBannerColorPrev _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    With Me.lblBannerColorName
        If (.ListCount <= 1) Then
            Exit Sub
        End If
        If (.ListIndex > 0) Then
            .ListIndex = .ListIndex - 1
        Else
            .ListIndex = .ListCount - 1
        End If
    End With
    mControlsUpdate dwDocumentWindow
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mBannerColorNext _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    With Me.lblBannerColorName
        If (.ListCount <= 1) Then
            Exit Sub
        End If
        If (.ListIndex < .ListCount - 1) Then
            .ListIndex = .ListIndex + 1
        Else
            .ListIndex = 0
        End If
    End With
    mControlsUpdate dwDocumentWindow
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mBannerSelectionClear _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    If (modBanner.gblnEnabled = False) Then
        Exit Sub
    End If
    Me.lblBannerSelectionText.Caption = ""
    mControlsUpdate dwDocumentWindow
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Function mblnNavigatorValid _
( _
) As Boolean
    mblnNavigatorValid = False
    If (modActive.gblnActiveWindowExists = False) Then
        Me.Hide
        Exit Function
    End If
    mblnNavigatorValid = True
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mSlideTitlesUpdate _
( _
    ByRef prePresentation As PowerPoint.Presentation _
)
    Dim sldSlide As PowerPoint.Slide
    Dim strTitle As String
    Dim blnSaved As Boolean
    
    blnSaved = prePresentation.Saved
    For Each sldSlide In prePresentation.Slides
        If (sldSlide.Shapes.HasTitle = Office.msoTrue) Then
            strTitle = sldSlide.Shapes.Title.TextFrame.TextRange.Text
        Else
            strTitle = ""
        End If
        sldSlide.Tags.Add "WorshipServiceAssistant_TitleDisplay", mStringWhiteSpaceClean(strTitle)
        sldSlide.Tags.Add "WorshipServiceAssistant_TitleMatch", mStringEverythingClean(strTitle)
    Next
    prePresentation.Saved = blnSaved
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Function mStringWhiteSpaceClean _
( _
    ByRef strDirty As String _
) As String
    Dim strClean As String
    
    strClean = strDirty
    
    '
    ' Fix white space.
    '
    strClean = VBA.Replace(strClean, VBA.Chr(9), " ")        ' replace tab with space
    strClean = VBA.Replace(strClean, VBA.Chr(11), " ")       ' replace line feed with space
    strClean = VBA.Replace(strClean, VBA.Chr(13), " ")       ' replace return with space
    '
    ' Fix quotes.
    '
    strClean = VBA.Replace(strClean, "`", VBA.Chr(39))       ' replace back single quote with single quote
    strClean = VBA.Replace(strClean, VBA.Chr(145), VBA.Chr(39))  ' replace open single quote with single quote
    strClean = VBA.Replace(strClean, VBA.Chr(146), VBA.Chr(39))  ' replace close single quote with single quote
    strClean = VBA.Replace(strClean, VBA.Chr(147), VBA.Chr(34))  ' replace open double quote with double quote
    strClean = VBA.Replace(strClean, VBA.Chr(148), VBA.Chr(34))  ' replace close double quote with double quote

    While (VBA.InStr(strClean, "  ") > 0)
        strClean = VBA.Replace(strClean, "  ", " ")
    Wend
    
    mStringWhiteSpaceClean = strClean
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Function mStringEverythingClean _
( _
    ByRef strDirty As String _
) As String
    Dim strClean As String
    
    strClean = VBA.LCase(strDirty)
    
    '
    ' Fix white space.
    '
    strClean = VBA.Replace(strClean, VBA.Chr(9), " ")        ' replace tab with space
    strClean = VBA.Replace(strClean, VBA.Chr(11), " ")       ' replace line feed with space
    strClean = VBA.Replace(strClean, VBA.Chr(13), " ")       ' replace return with space
    '
    ' Fix quotes.
    '
    strClean = VBA.Replace(strClean, "`", VBA.Chr(39))       ' replace back single quote with single quote
    strClean = VBA.Replace(strClean, VBA.Chr(145), VBA.Chr(39))  ' replace open single quote with single quote
    strClean = VBA.Replace(strClean, VBA.Chr(146), VBA.Chr(39))  ' replace close single quote with single quote
    strClean = VBA.Replace(strClean, VBA.Chr(147), VBA.Chr(34))  ' replace open double quote with double quote
    strClean = VBA.Replace(strClean, VBA.Chr(148), VBA.Chr(34))  ' replace close double quote with double quote
    
    strClean = VBA.Replace(strClean, "'m", "m")          ' ignore ' in 'm
    strClean = VBA.Replace(strClean, "'s", "s")          ' ignore ' in 's
    strClean = VBA.Replace(strClean, "'t", "t")          ' ignore ' in 't
    strClean = VBA.Replace(strClean, "'ve", "ve")        ' ignore ' in 've
    
    '
    ' Remove basic punctuation.
    '
    strClean = VBA.Replace(strClean, "-", " ")
    strClean = VBA.Replace(strClean, ",", " ")
    strClean = VBA.Replace(strClean, ";", " ")
    strClean = VBA.Replace(strClean, ":", " ")
    strClean = VBA.Replace(strClean, ".", " ")
    strClean = VBA.Replace(strClean, "!", " ")
    strClean = VBA.Replace(strClean, "?", " ")
    strClean = VBA.Replace(strClean, "/", " ")
    strClean = VBA.Replace(strClean, VBA.Chr(39), " ")
    strClean = VBA.Replace(strClean, VBA.Chr(34), " ")
    
    While (VBA.InStr(strClean, "  ") > 0)
        strClean = VBA.Replace(strClean, "  ", " ")
    Wend
    
    '
    ' Replace commonly interchanged words.
    '
    strClean = " " & strClean & " "
    strClean = VBA.Replace(strClean, " oh ", " o ")
    strClean = VBA.Replace(strClean, " alleluia ", " hallelujah ")
    strClean = VBA.Replace(strClean, " allelujah ", " hallelujah ")
    strClean = VBA.Replace(strClean, " emanuel ", " immanuel ")
    strClean = VBA.Replace(strClean, " emmanuel ", " immanuel ")
    strClean = VBA.Replace(strClean, " imanuel ", " immanuel ")
    If (VBA.Len(strClean) >= 2) Then
        strClean = VBA.Left(strClean, VBA.Len(strClean) - 1)
        strClean = VBA.Right(strClean, VBA.Len(strClean) - 1)
    End If
    
    mStringEverythingClean = strClean
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mApplicationViewUpdate _
( _
    ByRef blnSave As Boolean _
)
    Static sblnWindowWindowStateMax As Boolean
    Static sintWindowState As PowerPoint.PpWindowState
    Static slngTop As Long
    Static slngLeft As Long
    Static slngHeight As Long
    Static slngWidth As Long
    
    Dim dwDocumentWindow As PowerPoint.DocumentWindow
    
    '
    ' Save current view.
    '
    If (blnSave = True) Then
        sblnWindowWindowStateMax = False
        For Each dwDocumentWindow In Application.Windows
            If (dwDocumentWindow.WindowState = PowerPoint.ppWindowMaximized) Then
                sblnWindowWindowStateMax = True
            End If
        Next
        sintWindowState = Application.WindowState
        Application.WindowState = PowerPoint.ppWindowNormal
        slngTop = Application.Top
        slngLeft = Application.Left
        slngHeight = Application.Height
        slngWidth = Application.Width
        Application.WindowState = sintWindowState
    '
    ' Restore old view.
    '
    Else
        If (sblnWindowWindowStateMax = True) Then
            Application.Windows(1).WindowState = PowerPoint.ppWindowMaximized
        End If
        Application.WindowState = PowerPoint.ppWindowNormal
        Application.Height = 0
        Application.Width = 0
        Application.Top = slngTop
        Application.Left = slngLeft
        Application.Height = slngHeight
        Application.Width = slngWidth
        Application.WindowState = sintWindowState
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mApplicationViewSet _
( _
)
    Dim lngTop As Long
    Dim lngLeft As Long
    Dim lngHeight As Long
    Dim lngWidth As Long
    
    Application.WindowState = PowerPoint.ppWindowMaximized
    lngTop = Application.Top + 3
    lngLeft = Application.Left + 3
    lngHeight = Application.Height - 6
    lngWidth = Application.Width - 6
    Application.WindowState = PowerPoint.ppWindowNormal
    Application.Height = 0
    Application.Width = 0
    Application.Top = lngTop
    Application.Left = lngLeft + Me.Width
    Application.Height = lngHeight
    Application.Width = lngWidth - Me.Width
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mPresentationViewLoad _
( _
    ByRef prePresentation As PowerPoint.Presentation _
)
    Dim intWindowState As PowerPoint.PpWindowState
    Dim intViewType As PowerPoint.PpViewType
    Dim intView_DisplaySlideMiniature As Office.MsoTriState
    Dim intView_ZoomToFit As Office.MsoTriState
    Dim dwDocumentWindow As PowerPoint.DocumentWindow
    Dim intSaved As Office.MsoTriState
    
    If (prePresentation.Windows.Count > 0) Then
        Set dwDocumentWindow = prePresentation.Windows(1)
        
        '
        ' Save current view.
        '
        intWindowState = dwDocumentWindow.WindowState
        intViewType = dwDocumentWindow.ViewType
        If (intViewType = PowerPoint.ppViewNormal) Then
            Select Case dwDocumentWindow.Panes(2).ViewType
                Case PowerPoint.ppViewSlide
                    intViewType = PowerPoint.ppViewNormal
                Case PowerPoint.ppViewSlideMaster
                    intViewType = PowerPoint.ppViewSlideMaster
            End Select
        End If
        intView_DisplaySlideMiniature = dwDocumentWindow.View.DisplaySlideMiniature
        intView_ZoomToFit = dwDocumentWindow.View.ZoomToFit
        
        intSaved = prePresentation.Saved
        prePresentation.Tags.Add "WorshipServiceAssistant_Window_WindowState", intWindowState
        prePresentation.Tags.Add "WorshipServiceAssistant_Window_ViewType", intViewType
        prePresentation.Tags.Add "WorshipServiceAssistant_Window_View_DisplaySlideMiniature", intView_DisplaySlideMiniature
        prePresentation.Tags.Add "WorshipServiceAssistant_Window_View_ZoomToFit", intView_ZoomToFit
        prePresentation.Saved = intSaved
        
        '
        ' Set view.
        '
        dwDocumentWindow.WindowState = PowerPoint.ppWindowMaximized
        dwDocumentWindow.ViewType = PowerPoint.ppViewSlide
        dwDocumentWindow.View.DisplaySlideMiniature = Office.msoFalse
        dwDocumentWindow.View.ZoomToFit = Office.msoTrue
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mPresentationViewUnload _
( _
    ByRef prePresentation As PowerPoint.Presentation _
)
    Dim intWindowState As PowerPoint.PpWindowState
    Dim intViewType As PowerPoint.PpViewType
    Dim intView_DisplaySlideMiniature As Office.MsoTriState
    Dim intView_ZoomToFit As Office.MsoTriState
    Dim dwDocumentWindow As PowerPoint.DocumentWindow
    
    If (prePresentation.Windows.Count > 0) Then
        Set dwDocumentWindow = prePresentation.Windows(1)
        
        '
        ' Unload current view.
        '
        intWindowState = prePresentation.Tags("WorshipServiceAssistant_Window_WindowState")
        intViewType = prePresentation.Tags("WorshipServiceAssistant_Window_ViewType")
        intView_DisplaySlideMiniature = prePresentation.Tags("WorshipServiceAssistant_Window_View_DisplaySlideMiniature")
        intView_ZoomToFit = prePresentation.Tags("WorshipServiceAssistant_Window_View_ZoomToFit")
        
        dwDocumentWindow.WindowState = intWindowState
        dwDocumentWindow.ViewType = intViewType
        dwDocumentWindow.View.DisplaySlideMiniature = intView_DisplaySlideMiniature
        dwDocumentWindow.View.ZoomToFit = intView_ZoomToFit
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mFormInitialize _
( _
)
    '
    ' Initialize sizes.
    '
    Const LngFormOverhead As Long = 22
    Const LngPageOverhead As Long = 6
    Const LngFrameOverhead As Long = 14
    
    Dim avarColor(6, 3) As Variant
    
    Me.Height = Application.Height
    Me.tabPages.Height = _
        Me.Height - _
        Me.tabPages.Top - _
        LngFormOverhead
    Me.fraSlideSelection.Height = _
        Me.tabPages.Height - _
        LngPageOverhead - _
        Me.fraSlideSelection.Top
    Me.lstSlideSelectionList.Height = _
        Me.fraSlideSelection.Height - _
        LngFrameOverhead - _
        Me.lstSlideSelectionList.Top
    
    '
    ' Initialize values.
    '
    avarColor(0, 0) = "Red":     avarColor(0, 1) = 255: avarColor(0, 2) = 0:   avarColor(0, 3) = 0
    avarColor(1, 0) = "Green":   avarColor(1, 1) = 0:   avarColor(1, 2) = 255: avarColor(1, 3) = 0
    avarColor(2, 0) = "Blue":    avarColor(2, 1) = 0:   avarColor(2, 2) = 0:   avarColor(2, 3) = 255
    avarColor(3, 0) = "Yellow":  avarColor(3, 1) = 255: avarColor(3, 2) = 255: avarColor(3, 3) = 0
    avarColor(4, 0) = "Magenta": avarColor(4, 1) = 255: avarColor(4, 2) = 0:   avarColor(4, 3) = 255
    avarColor(5, 0) = "Cyan":    avarColor(5, 1) = 0:   avarColor(5, 2) = 255: avarColor(5, 3) = 255
    avarColor(6, 0) = "White":   avarColor(6, 1) = 255: avarColor(6, 2) = 255: avarColor(6, 3) = 255
    
    With Me.lblBannerColorName
        .List = avarColor
        .ListIndex = 1
    End With
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mPresentationNameUpdate _
( _
    ByVal dwDocumentWindow As PowerPoint.DocumentWindow _
)
    Dim strName As String
    
    If (modActive.gblnActiveWindowExists = False) Then
        strName = ""
    Else
        strName = dwDocumentWindow.Presentation.Name
    End If
    If (VBA.Len(strName) >= 4) Then
        If (VBA.LCase(VBA.Right(strName, 4)) = ".ppt") Then
            strName = VBA.Left(strName, VBA.Len(strName) - 4)
        End If
    End If
    Me.lblPresentationSelectionName.Caption = strName
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mSlideListUpdate _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    Dim prePresentation As PowerPoint.Presentation
    Dim sldSlide As PowerPoint.Slide
    Dim blnFilterModeNumber As Boolean
    Dim blnFilterModeTitle As Boolean
    Dim blnFilterModeBody As Boolean
    Dim strFilterTextDirty As String
    Dim lngFilterTextDirtyLen As Long
    Dim strFilterTextClean As String
    Dim lngFilterTextCleanLen As Long
    Dim lngFilterNumber As Long
    Dim lngSelectedSlideIndex As Long
    Dim alngMatch() As Long
    Dim lngMatchCount As Long
    Dim alngList() As String
    Dim lngListCount As Long
    Dim lngListIndex As Long
    
    Set prePresentation = dwDocumentWindow.Presentation
    
    Me.lstSlideSelectionList.Clear
    
    If (modActive.gblnActiveWindowSlideExists(dwDocumentWindow) = False) Then
        Exit Sub
    End If
    
    If (modActive.gblnActiveSlideExists(dwDocumentWindow) = False) Then
        dwDocumentWindow.View.Slide = prePresentation.Slides(1)
    End If
    lngSelectedSlideIndex = modActive.gppActiveSlideGet(dwDocumentWindow).SlideIndex
    
    blnFilterModeNumber = False
    blnFilterModeTitle = False
    blnFilterModeBody = False
    Select Case Me.cmdSlideSelectionMode.Tag
        Case "0"
            blnFilterModeNumber = True
            blnFilterModeTitle = True
        Case "1"
            blnFilterModeNumber = True
            blnFilterModeTitle = True
            blnFilterModeBody = True
    End Select
    
    If (Me.lblSlideSelectionNumber <> "") Then
        strFilterTextDirty = Me.lblSlideSelectionNumber.Caption
        lngFilterTextDirtyLen = VBA.Len(strFilterTextDirty)
    Else
        strFilterTextDirty = Me.lblSlideSelectionText.Caption
        lngFilterTextDirtyLen = VBA.Len(strFilterTextDirty)
    End If
    strFilterTextClean = mStringEverythingClean(strFilterTextDirty)
    lngFilterTextCleanLen = VBA.Len(strFilterTextClean)
    
    lngListCount = 0
    lngListIndex = -1
    
    '
    ' Since the filter text is empty, there is no filter.
    ' Therefore, all slides are in the list.
    '
    If (lngFilterTextDirtyLen = 0) Then
        lngListCount = prePresentation.Slides.Count
        If (lngListCount > 0) Then
            ReDim alngList(lngListCount - 1, 1) As String
            lngListCount = 0
            For Each sldSlide In prePresentation.Slides
                alngList(lngListCount, 0) = sldSlide.SlideIndex
                alngList(lngListCount, 1) = sldSlide.Tags("WorshipServiceAssistant_TitleDisplay")
                If (sldSlide.SlideIndex = lngSelectedSlideIndex) Then
                    lngListIndex = lngListCount
                End If
                lngListCount = lngListCount + 1
            Next
        End If
    '
    ' Since the filter text is a number, assume the filter text is a slide number.
    ' Therefore, there is one slide in the list.
    '
    ElseIf (IsNumeric(strFilterTextDirty)) Then
        If (blnFilterModeNumber) Then
            lngFilterNumber = strFilterTextDirty
            If ((lngFilterNumber > 0) And (lngFilterNumber <= prePresentation.Slides.Count)) Then
                ReDim alngList(0, 1) As String
                lngListCount = 0
                Set sldSlide = prePresentation.Slides(lngFilterNumber)
                alngList(lngListCount, 0) = sldSlide.SlideIndex
                alngList(lngListCount, 1) = sldSlide.Tags("WorshipServiceAssistant_TitleDisplay")
                If (sldSlide.SlideIndex = lngSelectedSlideIndex) Then
                    lngListIndex = lngListCount
                End If
                lngListCount = lngListCount + 1
            End If
        End If
    '
    ' Since the filter text is non-empty and non-numeric, assume that it is slide text.
    '
    Else
        '
        ' Adding new elements to a ListBox list takes time.  As a result, it is
        ' faster to build an array with the elements, and assign the array to the
        ' ListBox list.
        '
        ReDim alngMatch(prePresentation.Slides.Count) As Long
        lngMatchCount = 0
        If ((blnFilterModeTitle = False) And (blnFilterModeBody = False)) Then
        ElseIf ((blnFilterModeTitle = True) And (blnFilterModeBody = False)) Then
            For Each sldSlide In prePresentation.Slides
                If ((VBA.Left(sldSlide.Tags("WorshipServiceAssistant_TitleMatch"), lngFilterTextCleanLen) = strFilterTextClean) Or _
                    (VBA.Left(sldSlide.Tags("WorshipServiceAssistant_TitleDisplay"), lngFilterTextCleanLen) = strFilterTextClean)) Then
                    alngMatch(lngMatchCount) = sldSlide.SlideIndex
                    lngMatchCount = lngMatchCount + 1
                End If
            Next
        ElseIf ((blnFilterModeTitle = False) And (blnFilterModeBody = True)) Then
            For Each sldSlide In prePresentation.Slides
                If (mblnBodyMatch(sldSlide, strFilterTextDirty)) Then
                    alngMatch(lngMatchCount) = sldSlide.SlideIndex
                    lngMatchCount = lngMatchCount + 1
                End If
            Next
        ElseIf ((blnFilterModeTitle = True) And (blnFilterModeBody = True)) Then
            For Each sldSlide In prePresentation.Slides
                If ((VBA.Left(sldSlide.Tags("WorshipServiceAssistant_TitleMatch"), lngFilterTextCleanLen) = strFilterTextClean) Or _
                    (VBA.Left(sldSlide.Tags("WorshipServiceAssistant_TitleDisplay"), lngFilterTextCleanLen) = strFilterTextClean) Or _
                    mblnBodyMatch(sldSlide, strFilterTextDirty)) Then
                    alngMatch(lngMatchCount) = sldSlide.SlideIndex
                    lngMatchCount = lngMatchCount + 1
                End If
            Next
        End If
        
        If (lngMatchCount > 0) Then
            ReDim alngList(lngMatchCount - 1, 1) As String
            For lngListCount = 0 To lngMatchCount - 1 Step 1
                alngList(lngListCount, 0) = alngMatch(lngListCount)
                alngList(lngListCount, 1) = prePresentation.Slides(alngMatch(lngListCount)).Tags("WorshipServiceAssistant_TitleDisplay")
                If (alngMatch(lngListCount) = lngSelectedSlideIndex) Then
                    lngListIndex = lngListCount
                End If
            Next
            lngListCount = lngMatchCount
        End If
    End If
    
    Me.lstSlideSelectionList.Clear
    If (lngListCount > 0) Then
        Me.lstSlideSelectionList.List() = alngList
    End If
    With Me.lstSlideSelectionList
        If (.ListCount > 0) Then
            If (lngListIndex >= 0) Then
                .ListIndex = lngListIndex
            Else
                .ListIndex = 0
            End If
            .TopIndex = .ListIndex
            dwDocumentWindow.View.Slide = prePresentation.Slides(Val(.List(.ListIndex, 0)))
        Else
            dwDocumentWindow.View.Slide = prePresentation.Slides(1)
        End If
    End With
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Function mblnBodyMatch _
( _
    ByRef sldCurrent As PowerPoint.Slide, _
    ByRef strCurrent As String _
) As Boolean
    Dim shpCurrent As PowerPoint.Shape
    
    mblnBodyMatch = False
    For Each shpCurrent In sldCurrent.Shapes
        If (shpCurrent.HasTextFrame) Then
            If (Not (shpCurrent.TextFrame.TextRange.Find(strCurrent, 0, Office.msoFalse, Office.msoFalse) Is Nothing)) Then
                mblnBodyMatch = True
                Exit For
            End If
        End If
    Next
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Function mblnPowerPointOptionGet _
( _
    ByRef strOptionName As String, _
    ByRef lngOptionValue As Long _
) As Boolean
    Dim lngResult As Long
    Dim lngKeyHandle As Long
    Dim strOptionPath As String
    Dim lngOptionType As Long
    Dim lngOptionBuffer As Long
    Dim lngOptionBufferSize As Long
    
    mblnPowerPointOptionGet = False
    
    strOptionPath = _
        "Software\Microsoft\Office" & _
        "\" & Application.VERSION & _
        "\" & "PowerPoint" & _
        "\" & "Options"
    
    lngResult = modWin32AdvAPI32.RegOpenKeyEx _
        (modWin32AdvAPI32.HKEY_CURRENT_USER, _
         strOptionPath, _
         0, _
         modWin32AdvAPI32.KEY_QUERY_VALUE, _
         lngKeyHandle)
    If (lngResult = modWin32AdvAPI32.ERROR_SUCCESS) Then
        lngResult = modWin32AdvAPI32.RegQueryValueEx _
            (lngKeyHandle, _
             strOptionName, _
             0&, _
             lngOptionType, _
             ByVal 0&, _
             lngOptionBufferSize)
        If (lngResult = modWin32AdvAPI32.ERROR_SUCCESS) Then
            If (lngOptionType = modWin32AdvAPI32.REG_DWORD) Then
                lngResult = modWin32AdvAPI32.RegQueryValueEx _
                    (lngKeyHandle, _
                     strOptionName, _
                     0&, _
                     lngOptionType, _
                     lngOptionBuffer, _
                     lngOptionBufferSize)
                If (lngResult = modWin32AdvAPI32.ERROR_SUCCESS) Then
                    mblnPowerPointOptionGet = True
                    lngOptionValue = lngOptionBuffer
                End If
            End If
        End If
        lngResult = modWin32AdvAPI32.RegCloseKey _
            (lngKeyHandle)
    End If
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mPowerPointOptionDialogTabSet _
( _
    lngIndex As Long _
)
    Dim blnResult As Boolean
    Dim lngValue As Long
    
    blnResult = mblnPowerPointOptionGet("Options dialog current tab", lngValue)
    If ((blnResult = True) And (lngValue <> lngIndex)) Then
        While ((blnResult = True) And (lngValue <> lngIndex))
            Application.CommandBars.FindControl(Id:=522).Execute
            VBA.SendKeys "^{TAB}", True
            VBA.SendKeys "{ENTER}", True
            blnResult = mblnPowerPointOptionGet("Options dialog current tab", lngValue)
        Wend
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mPowerPointOptionsConfigure _
( _
)
    Dim blnResult As Boolean
    Dim lngValue As Long
    
    blnResult = mblnPowerPointOptionGet("SSRightMouse", lngValue)
    If (blnResult = False) Then
        mPowerPointOptionDialogTabSet 0
        Application.CommandBars.FindControl(Id:=522).Execute
        VBA.SendKeys "%P", True
        VBA.SendKeys "{ENTER}", True
        blnResult = mblnPowerPointOptionGet("SSRightMouse", lngValue)
    End If
    If ((blnResult = True) And (lngValue <> 0)) Then
        mPowerPointOptionDialogTabSet 0
        Application.CommandBars.FindControl(Id:=522).Execute
        VBA.SendKeys "%P", True
        VBA.SendKeys "{ENTER}", True
    End If
    
    blnResult = mblnPowerPointOptionGet("SSMenuButton", lngValue)
    If (blnResult = False) Then
        mPowerPointOptionDialogTabSet 0
        Application.CommandBars.FindControl(Id:=522).Execute
        VBA.SendKeys "%S", True
        VBA.SendKeys "{ENTER}", True
        blnResult = mblnPowerPointOptionGet("SSMenuButton", lngValue)
    End If
    If ((blnResult = True) And (lngValue <> 0)) Then
        mPowerPointOptionDialogTabSet 0
        Application.CommandBars.FindControl(Id:=522).Execute
        VBA.SendKeys "%S", True
        VBA.SendKeys "{ENTER}", True
    End If
    
    blnResult = mblnPowerPointOptionGet("SSEndOnBlankSlide", lngValue)
    If (blnResult = False) Then
        mPowerPointOptionDialogTabSet 0
        Application.CommandBars.FindControl(Id:=522).Execute
        VBA.SendKeys "%E", True
        VBA.SendKeys "{ENTER}", True
        blnResult = mblnPowerPointOptionGet("SSEndOnBlankSlide", lngValue)
    End If
    If ((blnResult = True) And (lngValue <> 0)) Then
        mPowerPointOptionDialogTabSet 0
        Application.CommandBars.FindControl(Id:=522).Execute
        VBA.SendKeys "%E", True
        VBA.SendKeys "{ENTER}", True
    End If
    
    blnResult = mblnPowerPointOptionGet("SaveAutoRecoveryInfo", lngValue)
    If (blnResult = False) Then
        mPowerPointOptionDialogTabSet 4
        Application.CommandBars.FindControl(Id:=522).Execute
        VBA.SendKeys "%S", True
        VBA.SendKeys "{ENTER}", True
        blnResult = mblnPowerPointOptionGet("SaveAutoRecoveryInfo", lngValue)
    End If
    If ((blnResult = True) And (lngValue <> 0)) Then
        mPowerPointOptionDialogTabSet 4
        Application.CommandBars.FindControl(Id:=522).Execute
        VBA.SendKeys "%S", True
        VBA.SendKeys "{ENTER}", True
    End If
End Sub


'===============================================================================
' Private Event Handlers.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub UserForm_Initialize _
( _
)
    Dim prePresentation As PowerPoint.Presentation
    Dim dwDocumentWindow As PowerPoint.DocumentWindow
    Dim tlbTemp As Office.CommandBar
    
    mblnNavigatorFormLoaded = False
    
    If (modPresentation.gblnExists = False) Then
        VBA.Unload frmNavigator
        Exit Sub
    End If
    
    mblnNavigatorFormLoaded = True
    
    mblnNavigatorFormLocked = True
    
    '
    ' Configure the PowerPoint options to the most appropiate values
    ' for running the Navigator.
    '
    mPowerPointOptionsConfigure

    '
    ' Hide any floating or pop-up menus so that they do not interfere with
    ' the Navigator form.
    '
    For Each tlbTemp In Application.CommandBars
        If ((tlbTemp.Position = Office.msoBarFloating) Or _
            (tlbTemp.Position = Office.msoBarPopup)) Then
            If (tlbTemp.Visible = True) Then
                tlbTemp.Visible = False
            End If
        End If
    Next
    
    If (modPresentation.gblnIsPresentation(Application.ActiveWindow.Presentation) = False) Then
        If (modPresentation.gblnExists = True) Then
            For Each prePresentation In Application.Presentations
                If (modPresentation.gblnIsPresentation(prePresentation) = True) Then
                    If (prePresentation.Windows.Count > 0) Then
                        prePresentation.Windows(1).Activate
                    End If
                End If
            Next
        End If
    End If
    
    Set dwDocumentWindow = Application.ActiveWindow
    
    modBanner.gCreate
    
    For Each prePresentation In Application.Presentations
        If (modSlideShow.gblnIsSlideShow(prePresentation) = True) Then
            modSlideShow.gSetup prePresentation
        End If
    Next
    dwDocumentWindow.Activate
    For Each prePresentation In Application.Presentations
        If (modSlideShow.gblnIsSlideShow(prePresentation) = True) Then
            mSlideTitlesUpdate prePresentation
        End If
    Next
    dwDocumentWindow.Activate
    For Each prePresentation In Application.Presentations
        If (modSlideShow.gblnIsSlideShow(prePresentation) = True) Then
            mPresentationViewLoad prePresentation
        End If
    Next
    dwDocumentWindow.Activate
    
    mApplicationViewUpdate True
    
    mApplicationViewSet
    
    Me.tabPages.Value = Me.tabPages("tabPresentation").Index
    
    Me.cmdSlideSelectionMode.Tag = "0"
    Me.lblSlideSelectionNumber.Caption = ""
    Me.lblSlideSelectionText.Caption = ""
    Me.lblBannerSelectionText.Caption = ""
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left - Me.Width
    Me.Top = Application.Top
    
    mFormInitialize
    
    Me.gRefresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub UserForm_QueryClose _
( _
    ByRef intCancel As Integer, _
    ByRef intCloseMode As Integer _
)
    Dim dwDocumentWindow As PowerPoint.DocumentWindow
    Dim prePresentation As PowerPoint.Presentation
    Dim intResponse As VBA.VbMsgBoxResult
    
    '
    ' Exit without unloading form, because the form was never loaded.
    '
    If (mblnNavigatorFormLoaded = False) Then
        Exit Sub
    End If
    
    intResponse = VBA.MsgBox( _
        buttons:= _
            VBA.vbYesNo + VBA.vbDefaultButton2 + VBA.vbExclamation, _
        Title:= _
            modProject.GstrNamePretty, _
        Prompt:= _
            "Are you sure you want to exit the Navigator?")

    If (intResponse = VBA.vbYes) Then
        intCancel = 0
        
        Me.Hide
        
        Set dwDocumentWindow = Application.ActiveWindow
        For Each prePresentation In Application.Presentations
            If (modSlideShow.gblnIsSlideShow(prePresentation) = True) Then
                mPresentationViewUnload prePresentation
            End If
        Next
        dwDocumentWindow.Activate
        
        mApplicationViewUpdate False
        
        '
        ' Work around a PowerPoint 2002 bug that causes the
        ' user to be prompted to save the banner presentation,
        ' even though the banner presentation was marked
        ' as saved after all changes were made.
        '
        For Each prePresentation In Application.Presentations
            If (modBanner.gblnIsBanner(prePresentation)) Then
                prePresentation.Saved = Office.msoTrue
            End If
        Next
    Else
        intCancel = 1
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub UserForm_Activate _
( _
)
    mblnNavigatorFormLocked = True
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub tabPages_Change _
( _
)
    frmNavigator.gRefresh
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub cmdGeneralExit_Click _
( _
)
    VBA.Unload frmNavigator
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub cmdGeneralHelp_Click _
( _
)
    Dim strHelpFile As String
    
    strHelpFile = modHelp.gstrFileNameGet(True)
    
    If (strHelpFile = "") Then
        Exit Sub
     End If
        Call modHelp.HtmlHelp( _
        0&, _
        strHelpFile, _
        modHelp.HH_DISPLAY_TOPIC, _
        modHelp.GstrIDH_TopicPath_WSACommandBarNavigator)
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub cmdGeneralPageNext_Click _
( _
)
    mGeneralPageNext Application.ActiveWindow
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub cmdSlideShowLoad_Click _
( _
)
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If
    mSlideShowLoad Application.ActiveWindow
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub cmdSlideShowHide_Click _
( _
)
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If
    mSlideShowHide Application.ActiveWindow
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub cmdSlideShowRun_Click _
( _
)
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If
    mSlideShowRun Application.ActiveWindow
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub cmdSlideShowPause_Click _
( _
)
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If
    mSlideShowPause Application.ActiveWindow
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub cmdSlideShowEffectPrev_Click _
( _
)
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If
    mSlideShowEffectPrev Application.ActiveWindow
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub cmdSlideShowEffectNext_Click _
( _
)
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If
    mSlideShowEffectNext Application.ActiveWindow
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub cmdPresentationSelectionPrev_Click _
( _
)
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If
    mPresentationSelectionPrev Application.ActiveWindow
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub cmdPresentationSelectionNext_Click _
( _
)
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If
    mPresentationSelectionNext Application.ActiveWindow
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub cmdSlideSelectionMode_Click _
( _
)
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If
    mSlideSelectionMode Application.ActiveWindow
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub cmdSlideSelectionClear_Click _
( _
)
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If
    mSlideSelectionClear Application.ActiveWindow
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub lstSlideSelectionList_Click _
( _
)
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If
    mSlideSelectionUpdate Application.ActiveWindow
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub cmdBannerConfigurationBannerDisable_Click _
( _
)
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If
    mBannerConfigurationBannerDisable Application.ActiveWindow
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub cmdBannerShowLoad_Click _
( _
)
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If
    mBannerShowLoad Application.ActiveWindow
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub cmdBannerShowHide_Click _
( _
)
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If
    mBannerShowHide Application.ActiveWindow
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub cmdBannerColorPrev_Click _
( _
)
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If
    mBannerColorPrev Application.ActiveWindow
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub cmdBannerColorNext_Click _
( _
)
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If
    mBannerColorNext Application.ActiveWindow
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub cmdBannerSelectionClear_Click _
( _
)
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If
    mBannerSelectionClear Application.ActiveWindow
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'   This event handler forces the 'Empty' frame to hold the focus.
'-------------------------------------------------------------------------------
Private Sub fraEmpty_Exit _
( _
    ByVal Cancel As MSForms.ReturnBoolean _
)
    Cancel = mblnNavigatorFormLocked
    If (Cancel = True) Then
        mblnNavigatorValid
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub fraEmpty_KeyPress _
( _
    ByVal intKeyASCII As MSForms.ReturnInteger _
)
    Dim dwDocumentWindow As PowerPoint.DocumentWindow
    Dim strFilterText As String
    Dim strBannerText As String
    
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If
    
    Set dwDocumentWindow = Application.ActiveWindow
        
    If (Me.tabPages(Me.tabPages.Value).Name = "tabPresentation") Then
        If (Me.lblSlideSelectionNumber <> "") Then
            strFilterText = Me.lblSlideSelectionNumber.Caption
        Else
            strFilterText = Me.lblSlideSelectionText.Caption
        End If
        
        Select Case intKeyASCII
            Case 8:                   ' <backspace>
                If (VBA.Len(strFilterText) > 0) Then
                    strFilterText = VBA.Left(strFilterText, VBA.Len(strFilterText) - 1)
                End If
            Case 32 To 127:           ' <space> or printable character
                strFilterText = strFilterText & VBA.Chr(intKeyASCII)
            Case Else:
        End Select
                
        If (VBA.IsNumeric(strFilterText) = True) Then
            Me.lblSlideSelectionNumber.Caption = strFilterText
            Me.lblSlideSelectionText.Caption = ""
        Else
            Me.lblSlideSelectionNumber.Caption = ""
            Me.lblSlideSelectionText.Caption = strFilterText
        End If
        
        mSlideListUpdate dwDocumentWindow
    ElseIf (Me.tabPages(Me.tabPages.Value).Name = "tabBanner") Then
        If (modBanner.gblnEnabled = True) Then
            strBannerText = Me.lblBannerSelectionText.Caption
            Select Case intKeyASCII
                Case 8:                   ' <backspace>
                    If (VBA.Len(strBannerText) > 0) Then
                        strBannerText = VBA.Left(strBannerText, VBA.Len(strBannerText) - 1)
                    End If
                Case 32 To 127:           ' <space> or printable character
                    strBannerText = strBannerText & VBA.Chr(intKeyASCII)
                Case Else:
            End Select
            Me.lblBannerSelectionText.Caption = strBannerText
        End If
    End If
    
    mControlsUpdate dwDocumentWindow
    mblnNavigatorValid
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub fraEmpty_KeyDown _
( _
    ByVal intKeyCode As MSForms.ReturnInteger, _
    ByVal intKeyModifier As Integer _
)
    Dim dwDocumentWindow As PowerPoint.DocumentWindow
    Dim blnShift As Boolean
    Dim blnControl As Boolean
    Dim blnAlternate As Boolean
    
    If (mblnNavigatorValid = False) Then
        Exit Sub
    End If

    Set dwDocumentWindow = Application.ActiveWindow
    
    blnShift = intKeyModifier And 1
    blnControl = intKeyModifier And 2
    blnAlternate = intKeyModifier And 4
    
    If (Me.tabPages(Me.tabPages.Value).Name = "tabPresentation") Then
        If ((blnShift = False) And (blnControl = False) And (blnAlternate = False)) Then
            Select Case intKeyCode
                Case 13:                    ' RETURN
                    mSlideShowLoad dwDocumentWindow
                Case 46:                    ' DELETE
                    mSlideSelectionClear dwDocumentWindow
                Case 37:                    ' LEFT_ARROW
                    mPresentationSelectionPrev dwDocumentWindow
                Case 39:                    ' RIGHT_ARROW
                    mPresentationSelectionNext dwDocumentWindow
                Case 38:                    ' UP_ARROW
                    mSlideSelectionPrev dwDocumentWindow
                Case 40:                    ' DOWN_ARROW
                    mSlideSelectionNext dwDocumentWindow
            End Select
        ElseIf ((blnShift = False) And (blnControl = True) And (blnAlternate = False)) Then
            Select Case intKeyCode
                Case 72, 83:                ' "H", "S"
                    mSlideShowHide dwDocumentWindow
                Case 82:                    ' "R"
                    mSlideShowRun dwDocumentWindow
                Case 80:                    ' "P"
                    mSlideShowPause dwDocumentWindow
                Case 38:                    ' UP_ARROW
                    mSlideShowEffectPrev dwDocumentWindow
                Case 40:                    ' DOWN_ARROW
                    mSlideShowEffectNext dwDocumentWindow
                Case 77:                    ' "M"
                    mSlideSelectionMode dwDocumentWindow
            End Select
        End If
    ElseIf (Me.tabPages(Me.tabPages.Value).Name = "tabBanner") Then
        If ((blnShift = False) And (blnControl = False) And (blnAlternate = False)) Then
            Select Case intKeyCode
                Case 13:                    ' RETURN
                    mBannerShowLoad dwDocumentWindow
                Case 46:                    ' DELETE
                    mBannerSelectionClear dwDocumentWindow
                Case 37:                    ' LEFT_ARROW
                    mBannerColorPrev dwDocumentWindow
                Case 39:                    ' RIGHT_ARROW
                    mBannerColorNext dwDocumentWindow
            End Select
        ElseIf ((blnShift = False) And (blnControl = True) And (blnAlternate = False)) Then
            Select Case intKeyCode
                Case 72, 83:                ' "H", "S"
                    mBannerShowHide dwDocumentWindow
                Case 68, 69:                ' "D", "E"
                    mBannerConfigurationBannerDisable dwDocumentWindow
            End Select
        End If
    End If
    If ((blnShift = False) And (blnControl = False) And (blnAlternate = False)) Then
        Select Case intKeyCode
            Case 9:                     ' TAB
                mGeneralPageNext dwDocumentWindow
            Case 27:                    ' ESCAPE
                VBA.Unload frmNavigator
            Case 112:                   ' F1
                cmdGeneralHelp_Click
        End Select
    End If
    mblnNavigatorValid
End Sub

