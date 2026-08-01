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
' Warning:
'   To use the form, call the modNavigator.gRun routine.
'   Calls to the form's Load or Show methods will result in incorrect behavior.
'
' Description:
'   The form implements the Navigator.
'
'   The form is modal, allowing the form to take more control.
'
'   The form contains an empty frame (fraEmpty) to which the focus is locked.
'   The empty frame is not visible to the user because its height and width
'   are both 0.  The form handles all if the keyboard input processing using
'   its KeyPress and KeyDown event handlers.
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
'   However, there are PowerPoint actions (in particular, actions which
'   activate document windows or slide show windows) that do not behave
'   when there is a visible modal form. Therefore, this form will often
'   hide itself. It relies on the modNavigator.gRun routine to show
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
'     (1) Changed mNavigatorGeneralHelp_Action routine to use the
'         modHelp.gTopicShow routine.
'   1.04.0000:
'     (1) Reorganized the code and improved the comments.
'     (2) Changed updating routines (mNavigator*_Update) and input processing
'         routines (mNavigator*_Action, mNavigator*_KeyPress and
'         mNavigator*_KeyDown) to make better use of the Enabled property of
'         the controls.
'     (3) Added information to the Tag property of the
'         lblPresentationSelectionName and lstPresentationSlideSelectionList
'         controls so that they can now know whether or not they need
'         to be updated.
'     (4) Moved the PowerPoint and Application routines to the class module
'         WSAApplication.
'     (5) Moved the Presentation routines to the class module
'         WSAPresentation.
'     (6) Moved the TitleMatch and BodyMatch routines to the class module
'         WSASlide.
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
'     (2) Fixed banner checking bug in
'         mNavigatorPresentationSlideShowLoad_Action.
'   1.01.0003:
'     (1) Fixed errors in control tip text.
'     (2) Fixed errors in form re-sizing in UpdateFormSize.
'     (3) Changed LoadPresentationView and UnloadPresentationView
'         so that the slide miniture window will not be displayed and
'         so that the slides will zoom to fit.
'     (4) Changed form so that it does not resize based on page.
'     (5) Changed mNavigatorPresentationSlideShowLoad_Action so that
'         transitions between presentations would be more seamless.
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
' Indicates whether or not the Navigator form's focus is locked to the frmEmpty
' frame.  Normally the focus is locked to the frmEmpty frame. However, sometimes
' it is necessary to change the focus in order to force the Navigator form
' to be the active window.  During this time, the focus must be unlocked.
'
Private mblnNavigatorFocusLocked As Boolean

Private WSA As WorshipServiceAssistant.WSAApplication


'===============================================================================
' Public Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
'   Updates the form, so that the it reflects the active PowerPoint document
'   window.
'-------------------------------------------------------------------------------
Public Sub gUpdate _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    
    mNavigator_Update WSA.ActivePresentation
End Sub


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' mNavigator*_Initialize subroutines.
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

'-------------------------------------------------------------------------------
' Purpose:
'   Initializes the form.
'-------------------------------------------------------------------------------
Private Sub mNavigator_Initialize _
( _
)
    Const LngFormOverhead As Long = 22
    Const LngPageOverhead As Long = 6
    
    ' Initialize form size.
    Me.Height = Application.Height
    
    ' Initialize tabPages multipage size.
    Me.tabPages.Height = _
        Me.Height - _
        Me.tabPages.Top - _
        LngFormOverhead
    
    ' Initialize fraPresentationSlideSelection frame size.
    Me.fraPresentationSlideSelection.Height = _
        Me.tabPages.Height - _
        LngPageOverhead - _
        Me.fraPresentationSlideSelection.Top
    
    ' Initialize form position.
    Me.StartUpPosition = 0
    Me.Left = Application.Left - Me.Width
    Me.Top = Application.Top
    
    ' Initialize frames.
    mNavigatorGeneral_Initialize
    mNavigatorPresentationSlideShow_Initialize
    mNavigatorPresentationSelection_Initialize
    mNavigatorPresentationSlideSelection_Initialize
    mNavigatorBannerConfiguration_Initialize
    mNavigatorBannerShow_Initialize
    mNavigatorBannerColor_Initialize
    mNavigatorBannerSelection_Initialize
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Initializes the fraGeneral frame.
'-------------------------------------------------------------------------------
Private Sub mNavigatorGeneral_Initialize _
( _
)
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Initializes the fraPresentationSlideShow frame.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideShow_Initialize _
( _
)
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Initializes the fraPresentationSelection frame.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSelection_Initialize _
( _
)
    Me.lblPresentationSelectionName.Tag = ""
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Initializes the fraPresentationSlideSelection frame.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideSelection_Initialize _
( _
)
    Const LngFrameOverhead As Long = 14
    
    ' Initialize size.
    Me.lstPresentationSlideSelectionList.Height = _
        Me.fraPresentationSlideSelection.Height - _
        LngFrameOverhead - _
        Me.lstPresentationSlideSelectionList.Top
        
    Me.cmdPresentationSlideSelectionMode.Tag = "0"
    Me.lblPresentationSlideSelectionNumber.Caption = ""
    Me.lblPresentationSlideSelectionText.Caption = ""
    Me.lstPresentationSlideSelectionList.Tag = ""
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Initializes the fraBannerConfiguration frame.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerConfiguration_Initialize _
( _
)
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Initializes the fraBannerShow frame.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerShow_Initialize _
( _
)
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Initializes the fraBannerColor frame.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerColor_Initialize _
( _
)
    Dim avarColor(6, 3) As Variant
    
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
' Purpose:
'   Initializes the fraBannerSelection frame.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerSelection_Initialize _
( _
)
    Me.lblBannerSelectionText.Caption = ""
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' mNavigator*_Update subroutines.
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

'-------------------------------------------------------------------------------
' Purpose:
'   Updates the form to reflect the currently associated PowerPoint document
'   window (input dwCurrent).
' Inputs:
'    dwCurrent:
'      The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigator_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    mNavigatorBannerSelection_Update udtPresentation
    mNavigatorBannerColor_Update udtPresentation
    mNavigatorBannerShow_Update udtPresentation
    mNavigatorBannerConfiguration_Update udtPresentation
    mNavigatorPresentationSlideSelection_Update udtPresentation
    mNavigatorPresentationSelection_Update udtPresentation
    mNavigatorPresentationSlideShow_Update udtPresentation
    
    If (WSA.Presentations.Count <= 0) Then
        Me.tabPages("tabPresentation").Enabled = False
    End If
    
    ' Update the selected dependent information:
    '   (1) The form's caption.
    '   (2) The form's default control. This is only for the visual effect,
    '       effect as all key presses are intercepted and processed by,
    '       event handlers.
    '   (3) Disable hidden frames.
    Select Case Me.tabPages(Me.tabPages.Value).Name
        Case "tabPresentation"
            Me.Caption = "Navigator - Presentation"
            Me.cmdPresentationSlideShowLoad.Default = True
            Me.fraBannerConfiguration.Enabled = False
            Me.fraBannerShow.Enabled = False
            Me.fraBannerColor.Enabled = False
            Me.fraBannerSelection.Enabled = False
        Case "tabBanner"
            Me.Caption = "Navigator - Banner"
            Me.cmdBannerShowLoad.Default = True
            Me.fraPresentationSlideShow.Enabled = False
            Me.fraPresentationSelection.Enabled = False
            Me.fraPresentationSlideSelection.Enabled = False
    End Select
    
    ' Set the focus and attempt to make sure that the Navigator is
    ' the active window. Changing the focus seems to accomplish this.
    mblnNavigatorFocusLocked = False
    Me.fraGeneral.SetFocus
    Me.fraEmpty.SetFocus
    mblnNavigatorFocusLocked = True
    
    ' Set the mouse pointer.
    Me.MousePointer = fmMousePointerArrow
    
    Me.Repaint
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Updates the fraGeneral frame to reflect the currently associated PowerPoint
'   document window (input dwCurrent).
' Inputs:
'    dwCurrent:
'      The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorGeneral_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    ' Enable everything by default.
    Me.fraGeneral.Enabled = True
    Me.cmdGeneralPageNext.Enabled = True
    Me.cmdGeneralHelp.Enabled = True
    Me.cmdGeneralExit.Enabled = True
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Updates the fraPresentationSlideShow frame to reflect the currently associated
'   PowerPoint document window (input dwCurrent).
' Assumptions:
'   In setting the state of the fraPresentationSlideShow controls, this routine assumes
'   that the lstPresentationSlideSelectionList control is up-to-date. Therefore, routines
'   which update fraPresentationSlideShowList should be called before this routine.
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideShow_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    ' Enable everything by default.
    Me.fraPresentationSlideShow.Enabled = True
    Me.cmdPresentationSlideShowLoad.Enabled = True
    Me.cmdPresentationSlideShowHide.Enabled = True
    Me.cmdPresentationSlideShowRun.Enabled = True
    Me.cmdPresentationSlideShowPause.Enabled = True
    Me.cmdPresentationSlideShowEffectPrev.Enabled = True
    Me.cmdPresentationSlideShowEffectNext.Enabled = True
    
    If ((WSA.Presentations.Count <= 0) Or (udtPresentation Is Nothing)) Then
        Me.fraPresentationSlideShow.Enabled = False
        Me.cmdPresentationSlideShowLoad.Enabled = False
        Me.cmdPresentationSlideShowHide.Enabled = False
        Me.cmdPresentationSlideShowRun.Enabled = False
        Me.cmdPresentationSlideShowPause.Enabled = False
        Me.cmdPresentationSlideShowEffectPrev.Enabled = False
        Me.cmdPresentationSlideShowEffectNext.Enabled = False
        Exit Sub
    End If
    
    With udtPresentation.SlideShowWindow
        ' Since the current presentation does not have an associated slide show,
        ' disable the slide show controls, other than load control.
        If (.Exists = False) Then
            Me.cmdPresentationSlideShowHide.Caption = "Hide"
            Me.cmdPresentationSlideShowHide.Enabled = False
            Me.cmdPresentationSlideShowRun.Enabled = False
            Me.cmdPresentationSlideShowPause.Enabled = False
            Me.cmdPresentationSlideShowEffectPrev.Enabled = False
            Me.cmdPresentationSlideShowEffectNext.Enabled = False
        Else
            If (Me.lstPresentationSlideSelectionList.ListCount = 0) Then
                Me.cmdPresentationSlideShowLoad.Enabled = False
            End If
            ' Since the slide show is hidden,
            ' label the hide control as 'Show'.
            If (.Visible = False) Then
                Me.cmdPresentationSlideShowHide.Caption = "Show"
            End If
            ' Since the slide show is running,
            ' disable the run control.
            If (.Running = True) Then
                Me.cmdPresentationSlideShowRun.Enabled = False
            End If
            ' Since the slide show is paused,
            ' disable the pause control.
            If (.Paused = True) Then
                Me.cmdPresentationSlideShowPause.Enabled = False
            End If
        End If
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Updates the fraPresentationSelection frame to reflect the currently
'   associated PowerPoint document window (input dwCurrent).
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSelection_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    ' Enable everything by default.
    Me.fraPresentationSelection.Enabled = True
    Me.lblPresentationSelectionName.Enabled = True
    Me.cmdPresentationSelectionPrev.Enabled = True
    Me.cmdPresentationSelectionNext.Enabled = True
    
    ' Update the lstControlPresentationName control.
    mNavigatorPresentationSelectionName_Update udtPresentation
    
    If ((WSA.Presentations.Count <= 0) Or (udtPresentation Is Nothing)) Then
        Me.lblPresentationSelectionName.Enabled = False
        Me.cmdPresentationSelectionPrev.Enabled = False
        Me.cmdPresentationSelectionNext.Enabled = False
        Exit Sub
    End If
    
    If (WSA.Presentations.Count <= 1) Then
        Me.cmdPresentationSelectionPrev.Enabled = False
        Me.cmdPresentationSelectionNext.Enabled = False
    End If
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Updates the lblPresentationSelectionName control, so that its Caption
'   contains the name of the Presentation containing the currently associated
'   PowerPoint document window (input dwCurrent).
' Inputs:
'    dwCurrent:
'      The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSelectionName_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    Dim strName As String
    
    If ((WSA.Presentations.Count <= 0) Or (udtPresentation Is Nothing)) Then
        Me.lblPresentationSelectionName.Caption = ""
        Me.lblPresentationSelectionName.Tag = ""
        Exit Sub
    End If
    
    strName = udtPresentation.Name
    If (Me.lblPresentationSelectionName.Tag <> strName) Then
        Me.lblPresentationSelectionName.Tag = strName
        If (VBA.Len(strName) >= 4) Then
            If (VBA.LCase(VBA.Right(strName, 4)) = ".ppt") Then
                strName = VBA.Left(strName, VBA.Len(strName) - 4)
            End If
        End If
        Me.lblPresentationSelectionName.Caption = strName
    End If
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Updates the fraPresentationSlideSelection frame to reflect the currently associated
'   PowerPoint document window (input dwCurrent).
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideSelection_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    ' Enable everything by default.
    Me.fraPresentationSlideSelection.Enabled = True
    Me.cmdPresentationSlideSelectionMode.Enabled = True
    Me.lblPresentationSlideSelectionMode.Enabled = True
    Me.lblPresentationSlideSelectionNumber.Enabled = True
    Me.lblPresentationSlideSelectionText.Enabled = True
    Me.cmdPresentationSlideSelectionClear.Enabled = True
    Me.lstPresentationSlideSelectionList.Enabled = True
    
    ' Update the lstPresentationSlideSelectionList control.
    mNavigatorPresentationSlideSelectionList_Update udtPresentation
    
    If ((WSA.Presentations.Count <= 0) Or (udtPresentation Is Nothing)) Then
        Me.cmdPresentationSlideSelectionMode.Tag = ""
        Me.lblPresentationSlideSelectionMode.Caption = ""
        Me.lblPresentationSlideSelectionNumber.Caption = ""
        Me.lblPresentationSlideSelectionText.Caption = ""
        Me.fraPresentationSlideSelection.Enabled = False
        Me.cmdPresentationSlideSelectionMode.Enabled = False
        Me.lblPresentationSlideSelectionMode.Enabled = False
        Me.lblPresentationSlideSelectionNumber.Enabled = False
        Me.lblPresentationSlideSelectionText.Enabled = False
        Me.cmdPresentationSlideSelectionClear.Enabled = False
        Me.lstPresentationSlideSelectionList.Enabled = False
        Exit Sub
    End If
    
    ' Update the Slide Selection Mode dependent controls.
    Select Case Me.cmdPresentationSlideSelectionMode.Tag
        Case "0"
            Me.lblPresentationSlideSelectionMode.Caption = "Filter matching Number and Title"
            Me.fraPresentationSlideSelection.BorderColor = VBA.vbInactiveBorder
        Case "1"
            Me.lblPresentationSlideSelectionMode.Caption = "Filter matching Number and Title and Body"
            Me.fraPresentationSlideSelection.BorderColor = VBA.vbRed
        Case Else
            Me.lblPresentationSlideSelectionMode.Caption = ""
            Me.fraPresentationSlideSelection.BorderColor = VBA.vbInactiveBorder
    End Select
    
    ' Since the slide selection filter is clear,
    ' disable the slide selection filter clear control.
    If ((Me.lblPresentationSlideSelectionNumber.Caption = "") And _
        (Me.lblPresentationSlideSelectionText.Caption = "")) Then
        Me.cmdPresentationSlideSelectionClear.Enabled = False
    End If
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Updates the lstPresentationSlideSelectionList control, so that its List property
'   contains the numbers and titles of the slides in the currently associated
'   PowerPoint document window (input dwCurrent) which satisfy the filter
'   criteria specified  by the selection filter mode (form control
'   cmdPresentationSlideSelectionMode) and the slide selection filter (form controls
'   cmdSlideSelectionNumber and cmdSlideSelectionText).
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideSelectionList_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    Dim strFilterText As String
    Dim lngFilterNumber As Long
    Dim lngSelectedSlideIndex As Long
    Dim lngMatch() As Long
    Dim lngMatchCount As Long
    Dim lngList() As String
    Dim lngListCount As Long
    Dim lngListIndex As Long
    Dim strTag As String

    If ((WSA.Presentations.Count <= 0) Or (udtPresentation Is Nothing)) Then
        Me.lstPresentationSlideSelectionList.Tag = ""
        Me.lstPresentationSlideSelectionList.Clear
        Exit Sub
    End If
        
    If (udtPresentation.Windows.Item(1).HasActiveSlide = False) Then
        udtPresentation.Slides.Item(1).Activate
    End If
    lngSelectedSlideIndex = udtPresentation.Windows.Item(1).ActiveSlide.SlideIndex
    
    strTag = udtPresentation.Name & "?" & _
             udtPresentation.Windows.Item(1).ActiveSlide.SlideIndex & "?" & _
             Me.cmdPresentationSlideSelectionMode.Tag & "?" & _
             Me.lblPresentationSlideSelectionNumber.Caption & "?" & _
             Me.lblPresentationSlideSelectionText.Caption
             
    If (Me.lstPresentationSlideSelectionList.Tag = strTag) Then
        Exit Sub
    End If
    Me.lstPresentationSlideSelectionList.Tag = strTag
    
    strFilterText = _
        Me.lblPresentationSlideSelectionNumber.Caption & _
        Me.lblPresentationSlideSelectionText.Caption
    
    With udtPresentation.Slides
        lngMatchCount = 0
        ' Since the filter text is numeric, assume the it is a slide number.
        If (IsNumeric(strFilterText)) Then
            lngFilterNumber = strFilterText
            If ((lngFilterNumber > 0) And (lngFilterNumber <= .Count)) Then
                lngMatchCount = 1
                ReDim lngMatch(lngMatchCount - 1)
                With .Item(lngFilterNumber)
                    lngMatch(0) = .Index
                End With
            End If
        ' Since the filter text is non-numeric, assume that it is slide text.
        Else
            ' Selection mode is "Match Title"
            If (Me.cmdPresentationSlideSelectionMode.Tag = "0") Then
                lngMatchCount = .TitleMatches(lngMatch, strFilterText)
            ' Selection mode is "Match Title and Body"
            ElseIf (Me.cmdPresentationSlideSelectionMode.Tag = "1") Then
                lngMatchCount = .TitleOrBodyMatches(lngMatch, strFilterText)
            End If
        End If
        If (lngMatchCount > 0) Then
            ReDim lngList(lngMatchCount - 1, 1) As String
            For lngListCount = 0 To lngMatchCount - 1 Step 1
                With .Item(lngMatch(lngListCount))
                    lngList(lngListCount, 0) = .Index
                    lngList(lngListCount, 1) = .Title
                    If (.Index = lngSelectedSlideIndex) Then
                        lngListIndex = lngListCount
                    End If
                End With
            Next
            lngListCount = lngMatchCount
        End If
    End With
    
    With Me.lstPresentationSlideSelectionList
        .Clear
        If (lngListCount > 0) Then
            .List() = lngList
        End If
        
        If (.ListCount > 0) Then
            If (lngListIndex >= 0) Then
                .ListIndex = lngListIndex
            Else
                .ListIndex = 0
            End If
            .TopIndex = .ListIndex
            udtPresentation.Slides.Item(Val(.List(.ListIndex, 0))).Activate
        Else
            udtPresentation.Slides.Item(1).Activate
        End If
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Updates the fraBannerConfiguration frame to reflect the currently associated
'   PowerPoint document window (input dwCurrent).
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerConfiguration_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    ' Enable everything by default.
    Me.fraBannerConfiguration.Enabled = True
    Me.cmdBannerConfigurationBannerDisable.Enabled = True
    
    ' Set the default control captions.
    Me.cmdBannerConfigurationBannerDisable.Caption = "Banner Disable"
    
    ' Since the banner is disabled,
    ' label the banner disable control as 'Banner Enable'.
    If (WSA.Banner.Enabled = False) Then
        Me.cmdBannerConfigurationBannerDisable.Caption = "Banner Enable"
    End If
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Updates the fraBannerShow frame to reflect the currently associated
'   PowerPoint document window (input dwCurrent).
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerShow_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    ' Enable everything by default.
    Me.fraBannerShow.Enabled = True
    Me.cmdBannerShowLoad.Enabled = True
    Me.cmdBannerShowHide.Enabled = True
    
    ' Set the default control captions.
    Me.cmdBannerShowHide.Caption = "Hide"
    
    ' Since the banner is hidden,
    ' label the hide control as 'Show'.
    If (WSA.Banner.Visible = False) Then
        Me.cmdBannerShowHide.Caption = "Show"
    End If
    
    ' Since the banner is disabled,
    ' disable all the banner show controls.
    If (WSA.Banner.Enabled = False) Then
        Me.cmdBannerShowLoad.Enabled = False
        Me.cmdBannerShowHide.Enabled = False
        Me.fraBannerShow.Enabled = False
    End If
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Updates the fraBannerColor frame to reflect the currently associated
'   PowerPoint document window (input dwCurrent).
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerColor_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    ' Enable everything by default.
    Me.fraBannerColor.Enabled = True
    Me.lblBannerColorName.Enabled = True
    Me.cmdBannerColorPrev.Enabled = True
    Me.cmdBannerColorNext.Enabled = True
    
    ' Set banner color name color.
    If (WSA.Banner.Enabled = True) Then
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
    
    ' Since the banner is disabled,
    ' disable all the banner color controls.
    If (WSA.Banner.Enabled = False) Then
        Me.lblBannerColorName.Enabled = False
        Me.cmdBannerColorPrev.Enabled = False
        Me.cmdBannerColorNext.Enabled = False
        Me.fraBannerColor.Enabled = False
    End If
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Updates the fraBannerSelection frame to reflect the currently associated
'   PowerPoint document window (input dwCurrent).
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerSelection_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    ' Enable everything by default.
    Me.fraBannerSelection.Enabled = True
    Me.lblBannerSelectionText.Enabled = True
    Me.cmdBannerSelectionClear.Enabled = True
    
    ' Since the banner is disabled,
    ' disable all the banner selection controls.
    If (WSA.Banner.Enabled = False) Then
        Me.lblBannerSelectionText.Enabled = False
        Me.cmdBannerSelectionClear.Enabled = False
        Me.fraBannerSelection.Enabled = False
    End If
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' mNavigator*_KeyDown subroutines.
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

'-------------------------------------------------------------------------------
' Purpose:
'   Processes KeyDown events that effect the form.
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'   inKeyCode:
'     The KeyCode to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'   inKeyModifier:
'     The KeyModifier to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'-------------------------------------------------------------------------------
Private Sub mNavigator_KeyDown _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByRef intKeyCode As MSForms.ReturnInteger, _
    ByRef intKeyModifier As Integer _
)
    mNavigatorGeneral_KeyDown udtPresentation, intKeyCode, intKeyModifier
    mNavigatorPresentationSlideShow_KeyDown udtPresentation, intKeyCode, intKeyModifier
    mNavigatorPresentationSelection_KeyDown udtPresentation, intKeyCode, intKeyModifier
    mNavigatorPresentationSlideSelection_KeyDown udtPresentation, intKeyCode, intKeyModifier
    mNavigatorBannerConfiguration_KeyDown udtPresentation, intKeyCode, intKeyModifier
    mNavigatorBannerShow_KeyDown udtPresentation, intKeyCode, intKeyModifier
    mNavigatorBannerColor_KeyDown udtPresentation, intKeyCode, intKeyModifier
    mNavigatorBannerSelection_KeyDown udtPresentation, intKeyCode, intKeyModifier
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Processes KeyDown events that effect the fraGeneral frame.
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'   inKeyCode:
'     The KeyCode to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'   inKeyModifier:
'     The KeyModifier to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'-------------------------------------------------------------------------------
Private Sub mNavigatorGeneral_KeyDown _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByRef intKeyCode As MSForms.ReturnInteger, _
    ByRef intKeyModifier As Integer _
)
    If (Me.fraGeneral.Enabled = False) Then
        Exit Sub
    End If
    
    If (intKeyModifier = 0) Then        ' no shift + no control + no alternate
        Select Case intKeyCode
            Case 9:                     ' TAB
                mNavigatorGeneralPageNext_Action udtPresentation
            Case 112:                   ' F1
                mNavigatorGeneralHelp_Action udtPresentation
            Case 27:                    ' ESCAPE
                mNavigatorGeneralExit_Action
        End Select
    End If
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Processes KeyDown events that effect the fraPresentationSlideShow frame.
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'   inKeyCode:
'     The KeyCode to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'   inKeyModifier:
'     The KeyModifier to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideShow_KeyDown _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByRef intKeyCode As MSForms.ReturnInteger, _
    ByRef intKeyModifier As Integer _
)
    If (Me.fraPresentationSlideShow.Enabled = False) Then
        Exit Sub
    End If
    
    Select Case intKeyModifier
        Case 0:                             ' no shift + no control + no alternate
            Select Case intKeyCode
                Case 13:                    ' RETURN
                    mNavigatorPresentationSlideShowLoad_Action udtPresentation
                Case 46:                    ' DELETE
                    mNavigatorPresentationSlideSelectionClear_Action udtPresentation
            End Select
        Case 2:                             ' no shift + control + no alternate
            Select Case intKeyCode
                Case 72, 83:                ' "H", "S"
                    mNavigatorPresentationSlideShowHide_Action udtPresentation
                Case 82:                    ' "R"
                    mNavigatorPresentationSlideShowRun_Action udtPresentation
                Case 80:                    ' "P"
                    mNavigatorPresentationSlideShowPause_Action udtPresentation
                Case 38:                    ' UP_ARROW
                    mNavigatorPresentationSlideShowEffectPrev_Action udtPresentation
                Case 40:                    ' DOWN_ARROW
                    mNavigatorPresentationSlideShowEffectNext_Action udtPresentation
            End Select
    End Select
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Processes KeyDown events that effect the fraPresentationSelection frame.
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'   inKeyCode:
'     The KeyCode to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'   inKeyModifier:
'     The KeyModifier to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSelection_KeyDown _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByRef intKeyCode As MSForms.ReturnInteger, _
    ByRef intKeyModifier As Integer _
)
    If (Me.fraPresentationSelection.Enabled = False) Then
        Exit Sub
    End If
    
    Select Case intKeyModifier
        Case 0:                             ' no shift + no control + no alternate
            Select Case intKeyCode
                Case 37:                    ' LEFT_ARROW
                    mNavigatorPresentationSelectionPrev_Action udtPresentation
                Case 39:                    ' RIGHT_ARROW
                    mNavigatorPresentationSelectionNext_Action udtPresentation
            End Select
    End Select
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Processes KeyDown events that effect the fraPresentationSlideSelection frame.
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'   inKeyCode:
'     The KeyCode to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'   inKeyModifier:
'     The KeyModifier to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideSelection_KeyDown _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByRef intKeyCode As MSForms.ReturnInteger, _
    ByRef intKeyModifier As Integer _
)
    If (Me.fraPresentationSlideSelection.Enabled = False) Then
        Exit Sub
    End If
    
    Select Case intKeyModifier
        Case 0:                             ' no shift + no control + no alternate
            Select Case intKeyCode
                Case 38:                    ' UP_ARROW
                    mNavigatorPresentationSlideSelectionPrev_Action udtPresentation
                Case 40:                    ' DOWN_ARROW
                    mNavigatorPresentationSlideSelectionNext_Action udtPresentation
            End Select
        Case 2:                             ' no shift + control + no alternate
            Select Case intKeyCode
                Case 77:                    ' "M"
                    mNavigatorPresentationSlideSelectionMode_Action udtPresentation
            End Select
    End Select
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Processes KeyDown events that effect the fraBannerConfiguration frame.
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'   inKeyCode:
'     The KeyCode to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'   inKeyModifier:
'     The KeyModifier to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerConfiguration_KeyDown _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByRef intKeyCode As MSForms.ReturnInteger, _
    ByRef intKeyModifier As Integer _
)
    If (Me.fraBannerConfiguration.Enabled = False) Then
        Exit Sub
    End If
    
    Select Case intKeyModifier
        Case 2:                             ' no shift + control + no alternate
            Select Case intKeyCode
                Case 68, 69:                ' "D", "E"
                    mNavigatorBannerConfigurationBannerDisable_Action udtPresentation
            End Select
    End Select
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Processes KeyDown events that effect the fraBannerShow frame.
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'   inKeyCode:
'     The KeyCode to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'   inKeyModifier:
'     The KeyModifier to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerShow_KeyDown _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByRef intKeyCode As MSForms.ReturnInteger, _
    ByRef intKeyModifier As Integer _
)
    If (Me.fraBannerColor.Enabled = False) Then
        Exit Sub
    End If
    
    Select Case intKeyModifier
        Case 0:                             ' no shift + no control + no alternate
            Select Case intKeyCode
                Case 13:                    ' RETURN
                    mNavigatorBannerShowLoad_Action udtPresentation
            End Select
        Case 2:                             ' no shift + control + no alternate
            Select Case intKeyCode
                Case 72, 83:                ' "H", "S"
                    mNavigatorBannerShowHide_Action udtPresentation
            End Select
    End Select
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Processes KeyDown events that effect the fraBannerColor frame.
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'   inKeyCode:
'     The KeyCode to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'   inKeyModifier:
'     The KeyModifier to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerColor_KeyDown _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByRef intKeyCode As MSForms.ReturnInteger, _
    ByRef intKeyModifier As Integer _
)
    If (Me.fraBannerColor.Enabled = False) Then
        Exit Sub
    End If
    
    Select Case intKeyModifier
        Case 0:                             ' no shift + no control + no alternate
            Select Case intKeyCode
                Case 37:                    ' LEFT_ARROW
                    mNavigatorBannerColorPrev_Action udtPresentation
                Case 39:                    ' RIGHT_ARROW
                    mNavigatorBannerColorNext_Action udtPresentation
            End Select
    End Select
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Processes KeyDown events that effect the fraBannerSelection frame.
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'   inKeyCode:
'     The KeyCode to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'   inKeyModifier:
'     The KeyModifier to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerSelection_KeyDown _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByRef intKeyCode As MSForms.ReturnInteger, _
    ByRef intKeyModifier As Integer _
)
    If (Me.fraBannerSelection.Enabled = False) Then
        Exit Sub
    End If
    
    Select Case intKeyModifier
        Case 0:                             ' no shift + no control + no alternate
            Select Case intKeyCode
                Case 46:                    ' DELETE
                    mNavigatorBannerSelectionClear_Action udtPresentation
            End Select
    End Select
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' mNavigator*_KeyPress subroutines.
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

'-------------------------------------------------------------------------------
' Purpose:
'   Processes KeyPress events that effect the form.
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'   inKeyASCII:
'     The KeyASCII to be processed. A detailed description can be found in
'     the VBA help for the KeyPress event for UserForm controls.
'-------------------------------------------------------------------------------
Private Sub mNavigator_KeyPress _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByVal intKeyASCII As MSForms.ReturnInteger _
)
    mNavigatorGeneral_KeyPress udtPresentation, intKeyASCII
    mNavigatorPresentationSlideShow_KeyPress udtPresentation, intKeyASCII
    mNavigatorPresentationSelection_KeyPress udtPresentation, intKeyASCII
    mNavigatorPresentationSlideSelection_KeyPress udtPresentation, intKeyASCII
    mNavigatorBannerConfiguration_KeyPress udtPresentation, intKeyASCII
    mNavigatorBannerShow_KeyPress udtPresentation, intKeyASCII
    mNavigatorBannerColor_KeyPress udtPresentation, intKeyASCII
    mNavigatorBannerSelection_KeyPress udtPresentation, intKeyASCII
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Processes KeyPress events that effect the fraGeneral frame.
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'   inKeyASCII:
'     The KeyASCII to be processed. A detailed description can be found in
'     the VBA help for the KeyPress event for UserForm controls.
'-------------------------------------------------------------------------------
Private Sub mNavigatorGeneral_KeyPress _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByRef intKeyASCII As MSForms.ReturnInteger _
)
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Processes KeyPress events that effect the fraPresentationSlideShow frame.
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'   inKeyASCII:
'     The KeyASCII to be processed. A detailed description can be found in
'     the VBA help for the KeyPress event for UserForm controls.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideShow_KeyPress _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByRef intKeyASCII As MSForms.ReturnInteger _
)
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Processes KeyPress events that effect the fraPresentationSelection frame.
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'   inKeyASCII:
'     The KeyASCII to be processed. A detailed description can be found in
'     the VBA help for the KeyPress event for UserForm controls.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSelection_KeyPress _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByRef intKeyASCII As MSForms.ReturnInteger _
)
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Processes KeyPress events that effect the fraPresentationSlideSelection frame.
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'   inKeyASCII:
'     The KeyASCII to be processed. A detailed description can be found in
'     the VBA help for the KeyPress event for UserForm controls.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideSelection_KeyPress _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByRef intKeyASCII As MSForms.ReturnInteger _
)
    Dim strFilter As String
    
    If (Me.fraPresentationSlideSelection.Enabled = False) Then
        Exit Sub
    End If
    
    If (Me.lblPresentationSlideSelectionNumber <> "") Then
        strFilter = Me.lblPresentationSlideSelectionNumber.Caption
    Else
        strFilter = Me.lblPresentationSlideSelectionText.Caption
    End If
        
    Select Case intKeyASCII
        Case 8:                   ' <backspace>
            If (VBA.Len(strFilter) > 0) Then
                strFilter = VBA.Left(strFilter, VBA.Len(strFilter) - 1)
            End If
        Case 32 To 126:           ' <space> or printable character
            strFilter = strFilter & VBA.Chr(intKeyASCII)
        Case Else:
    End Select
        
    If (VBA.IsNumeric(strFilter) = True) Then
        Me.lblPresentationSlideSelectionNumber.Caption = strFilter
        Me.lblPresentationSlideSelectionText.Caption = ""
    Else
        Me.lblPresentationSlideSelectionNumber.Caption = ""
        Me.lblPresentationSlideSelectionText.Caption = strFilter
    End If
    
    mNavigator_Update udtPresentation
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Processes KeyPress events that effect the fraBannerConfiguration frame.
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'   inKeyASCII:
'     The KeyASCII to be processed. A detailed description can be found in
'     the VBA help for the KeyPress event for UserForm controls.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerConfiguration_KeyPress _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByRef intKeyASCII As MSForms.ReturnInteger _
)
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Processes KeyPress events that effect the fraBannerShow frame.
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'   inKeyASCII:
'     The KeyASCII to be processed. A detailed description can be found in
'     the VBA help for the KeyPress event for UserForm controls.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerShow_KeyPress _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByRef intKeyASCII As MSForms.ReturnInteger _
)
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Processes KeyPress events that effect the fraBannerColor frame.
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'   inKeyASCII:
'     The KeyASCII to be processed. A detailed description can be found in
'     the VBA help for the KeyPress event for UserForm controls.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerColor_KeyPress _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByRef intKeyASCII As MSForms.ReturnInteger _
)
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Processes KeyPress events that effect the fraBannerSelection frame.
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'   inKeyASCII:
'     The KeyASCII to be processed. A detailed description can be found in
'     the VBA help for the KeyPress event for UserForm controls.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerSelection_KeyPress _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByRef intKeyASCII As MSForms.ReturnInteger _
)
    Dim strBanner As String
    
    If (Me.fraBannerSelection.Enabled = False) Then
        Exit Sub
    End If
    
    strBanner = Me.lblBannerSelectionText.Caption
    
    Select Case intKeyASCII
        Case 8:                   ' <backspace>
            If (VBA.Len(strBanner) > 0) Then
                strBanner = VBA.Left(strBanner, VBA.Len(strBanner) - 1)
            End If
        Case 32 To 126:           ' <space> or printable character
            strBanner = strBanner & VBA.Chr(intKeyASCII)
        Case Else:
    End Select
    
    Me.lblBannerSelectionText.Caption = strBanner
    
    mNavigator_Update udtPresentation
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' mNavigator*_Action routines.
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorGeneralPageNext_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    Dim lngPage As Integer
    
    If (Me.cmdGeneralPageNext.Enabled = False) Then
        Exit Sub
    End If
    
    lngPage = Me.tabPages.Value
    lngPage = lngPage + 1
    If (lngPage >= Me.tabPages.Count) Then
        lngPage = 0
    End If
    
    mblnNavigatorFocusLocked = False
    Me.tabPages.Value = lngPage
    mblnNavigatorFocusLocked = True
    
    mNavigator_Update udtPresentation
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorGeneralHelp_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    If (Me.cmdGeneralHelp.Enabled = False) Then
        Exit Sub
    End If
    
    modHelp.gTopicShow WSACommandBarNavigator
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub mNavigatorGeneralExit_Action _
( _
)
    If (Me.cmdGeneralExit.Enabled = False) Then
        Exit Sub
    End If
    
    VBA.Unload frmNavigator
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideShowLoad_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    If (Me.cmdPresentationSlideShowLoad.Enabled = False) Then
        Exit Sub
    End If
    
    If (udtPresentation.Windows.Item(1).HasActiveSlide = False) Then
        Exit Sub
    End If

    Me.Hide
    
    udtPresentation.SlideShowWindow.Load Me.lstPresentationSlideSelectionList.Value
    
    udtPresentation.Activate
    
    Me.lblPresentationSlideSelectionNumber.Caption = ""
    Me.lblPresentationSlideSelectionText.Caption = ""
    
    mNavigator_Update udtPresentation
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideShowHide_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    If (Me.cmdPresentationSlideShowHide.Enabled = False) Then
        Exit Sub
    End If
    
    With udtPresentation.SlideShowWindow
        .Visible = Not .Visible
    End With
    
    mNavigator_Update udtPresentation
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideShowRun_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    If (Me.cmdPresentationSlideShowRun.Enabled = False) Then
        Exit Sub
    End If
    
    udtPresentation.SlideShowWindow.Run
    
    mNavigator_Update udtPresentation
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideShowPause_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    If (Me.cmdPresentationSlideShowPause.Enabled = False) Then
        Exit Sub
    End If
    
    udtPresentation.SlideShowWindow.Pause
    
    mNavigator_Update udtPresentation
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideShowEffectPrev_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    If (Me.cmdPresentationSlideShowEffectPrev.Enabled = False) Then
        Exit Sub
    End If
    
    udtPresentation.SlideShowWindow.EffectPrev
    
    Me.lblPresentationSlideSelectionNumber = ""
    Me.lblPresentationSlideSelectionText = ""
    
    mNavigator_Update udtPresentation
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideShowEffectNext_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    If (Me.cmdPresentationSlideShowEffectNext.Enabled = False) Then
        Exit Sub
    End If
    
    udtPresentation.SlideShowWindow.EffectNext
    
    Me.lblPresentationSlideSelectionNumber = ""
    Me.lblPresentationSlideSelectionText = ""
    
    mNavigator_Update udtPresentation
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSelectionPrev_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    Dim lngIndex As Long
    
    If (Me.cmdPresentationSelectionPrev.Enabled = False) Then
        Exit Sub
    End If
    
    With WSA.Presentations
        lngIndex = .Index(udtPresentation) - 1
        If (lngIndex < 1) Then
            lngIndex = .Count
        End If
        .Item(lngIndex).Activate
        mNavigator_Update WSA.ActivePresentation
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSelectionNext_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    Dim lngIndex As Long
    
    If (Me.cmdPresentationSelectionNext.Enabled = False) Then
        Exit Sub
    End If
    
    With WSA.Presentations
        lngIndex = .Index(udtPresentation) + 1
        If (lngIndex > .Count) Then
            lngIndex = 1
        End If
        .Item(lngIndex).Activate
        mNavigator_Update WSA.ActivePresentation
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideSelectionMode_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    If (Me.cmdPresentationSlideSelectionMode.Enabled = False) Then
        Exit Sub
    End If
    
    Select Case Me.cmdPresentationSlideSelectionMode.Tag
        Case "0"                        ' Title
            Me.cmdPresentationSlideSelectionMode.Tag = "1"
        Case "1"                        ' Title and Body
            Me.cmdPresentationSlideSelectionMode.Tag = "0"
        Case Else
            Me.cmdPresentationSlideSelectionMode.Tag = "0"
    End Select
    mNavigator_Update udtPresentation
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideSelectionClear_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    If (Me.cmdPresentationSlideSelectionClear.Enabled = False) Then
        Exit Sub
    End If
    
    Me.lblPresentationSlideSelectionNumber.Caption = ""
    Me.lblPresentationSlideSelectionText.Caption = ""
    mNavigator_Update udtPresentation
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideSelectionUpdate_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    Dim lngIndex As Long
    
    If (Me.lstPresentationSlideSelectionList.Enabled = False) Then
        Exit Sub
    End If
    
    '
    ' Set slide selected in the presentation to match the
    ' slide selected in the slide list control
    '
    With Me.lstPresentationSlideSelectionList
        If (.ListIndex >= 0) Then
            lngIndex = .Value
            udtPresentation.Slides.Item(lngIndex).Activate
        End If
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideSelectionPrev_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    If (Me.lstPresentationSlideSelectionList.Enabled = False) Then
        Exit Sub
    End If
    
    With Me.lstPresentationSlideSelectionList
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
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideSelectionNext_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    If (Me.lstPresentationSlideSelectionList.Enabled = False) Then
        Exit Sub
    End If
    
    With Me.lstPresentationSlideSelectionList
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
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerConfigurationBannerDisable_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    If (Me.cmdBannerConfigurationBannerDisable.Enabled = False) Then
        Exit Sub
    End If
    
    With WSA.Banner
        .Enabled = Not .Enabled
    End With
    mNavigator_Update udtPresentation
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerShowLoad_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    Dim intRed As Integer
    Dim intGreen As Integer
    Dim intBlue As Integer
    
    If (Me.cmdBannerShowLoad.Enabled = False) Then
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
    
    WSA.Banner.Load Me.lblBannerSelectionText.Caption, intRed, intGreen, intBlue
    mNavigatorBannerSelectionClear_Action udtPresentation
    mNavigator_Update udtPresentation
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerShowHide_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    If (Me.cmdBannerShowHide.Enabled = False) Then
        Exit Sub
    End If
    
    With WSA.Banner
        .Visible = Not .Visible
    End With
    mNavigator_Update udtPresentation
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerColorPrev_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    If (Me.cmdBannerColorPrev.Enabled = False) Then
        Exit Sub
    End If
    
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
    mNavigator_Update udtPresentation
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerColorNext_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    If (Me.cmdBannerColorNext.Enabled = False) Then
        Exit Sub
    End If
    
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
    mNavigator_Update udtPresentation
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   dwCurrent:
'     The current PowerPoint document window with which the form is associated.
'-------------------------------------------------------------------------------
Private Sub mNavigatorBannerSelectionClear_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    If (Me.cmdBannerSelectionClear.Enabled = False) Then
        Exit Sub
    End If
    
    If (WSA.Banner.Enabled = False) Then
        Exit Sub
    End If
    Me.lblBannerSelectionText.Caption = ""
    mNavigator_Update udtPresentation
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Support subroutines and functions.
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

'-------------------------------------------------------------------------------
' Purpose:
'   Checks whether or not it is valid for the form to be used. It is valid
'   for the form to be used if there is an active PowerPoint document window
'   with which the form can be associated.
' Effects:
'   Hides the form if it is not valid for the form to be used.
' Returns:
'   A boolean that is true only if it is valid for the form to be used.
'-------------------------------------------------------------------------------
Private Function mblnValid _
( _
) As Boolean
    mblnValid = True
    If (WSA.Presentations.Count >= 1) Then
        If (WSA.HasActivePresentation <> Office.msoTrue) Then
            mblnValid = False
            Me.Hide
        End If
    End If
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   intCancel:
'     The routine sets this value to a non-zero value if initialization
'     is was canceled.
'-------------------------------------------------------------------------------
Private Sub mInitialize _
( _
    ByRef intCancel As Integer _
)
    mblnNavigatorFormLoaded = False

    Set WSA = New WorshipServiceAssistant.WSAApplication
    WSA.Initialize Application
    
    If (WSA.HasActivePresentation <> Office.msoTrue) Then
        If (WSA.Presentations.Count >= 1) Then
            WSA.Presentations.Item(1).Activate
        End If
    End If
    
    Application.Left = Application.Left + Me.Width
    Application.Width = Application.Width - Me.Width
    
    mNavigator_Initialize
    
    mNavigator_Update WSA.ActivePresentation
    
    mblnNavigatorFormLoaded = True
    
    intCancel = 0
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub mTerminate _
( _
)
    ' Exit without unloading form, because the form was never loaded.
    If (mblnNavigatorFormLoaded = False) Then
        Exit Sub
    End If
    
    Me.Hide
    WSA.Terminate
    Set WSA = Nothing
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'   intCancel:
'     The routine sets this value to a non-zero value if initialization
'     is was canceled.
'-------------------------------------------------------------------------------
Private Sub mQueryClose _
( _
    ByRef intCancel As Integer, _
    ByRef intCloseMode As Integer _
)
    Dim intResponse As VBA.VbMsgBoxResult
    
    ' Exit without unloading form, because the form was never loaded.
    If (mblnNavigatorFormLoaded = False) Then
        mblnNavigatorFocusLocked = False
        intCancel = 0
        Exit Sub
    End If
    
    If ((intCloseMode = VBA.vbFormControlMenu) Or _
        (intCloseMode = VBA.vbFormCode)) Then
        intResponse = VBA.MsgBox( _
            buttons:= _
                VBA.vbYesNo + VBA.vbDefaultButton2 + VBA.vbExclamation, _
            Title:= _
                modProject.GstrNamePretty, _
            Prompt:= _
                "Are you sure you want to exit the Navigator?")
        If (intResponse = VBA.vbYes) Then
            mblnNavigatorFocusLocked = False
            intCancel = 0
        Else
            intCancel = 1
        End If
    End If
End Sub


'===============================================================================
' Private Event Handlers.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub UserForm_Initialize _
( _
)
    Dim intCancel As Integer
    
    intCancel = 0
    
    mInitialize intCancel
    
    If (intCancel <> 0) Then
        VBA.Unload frmNavigator
        Exit Sub
    End If
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub UserForm_Terminate _
( _
)
    mTerminate
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
'-------------------------------------------------------------------------------
Private Sub UserForm_QueryClose _
( _
    ByRef intCancel As Integer, _
    ByRef intCloseMode As Integer _
)
    mQueryClose intCancel, intCloseMode
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub UserForm_Activate _
( _
)
    mblnNavigatorFocusLocked = True
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub tabPages_Change _
( _
)
    mNavigator_Update WSA.ActivePresentation
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub cmdGeneralPageNext_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorGeneralPageNext_Action WSA.ActivePresentation
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub cmdGeneralHelp_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorGeneralHelp_Action WSA.ActivePresentation
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub cmdGeneralExit_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorGeneralExit_Action
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub cmdPresentationSlideShowLoad_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorPresentationSlideShowLoad_Action WSA.ActivePresentation
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub cmdPresentationSlideShowHide_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorPresentationSlideShowHide_Action WSA.ActivePresentation
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub cmdPresentationSlideShowRun_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorPresentationSlideShowRun_Action WSA.ActivePresentation
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub cmdPresentationSlideShowPause_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorPresentationSlideShowPause_Action WSA.ActivePresentation
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub cmdPresentationSlideShowEffectPrev_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorPresentationSlideShowEffectPrev_Action WSA.ActivePresentation
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub cmdPresentationSlideShowEffectNext_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorPresentationSlideShowEffectNext_Action WSA.ActivePresentation
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub cmdPresentationSelectionPrev_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorPresentationSelectionPrev_Action WSA.ActivePresentation
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub cmdPresentationSelectionNext_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorPresentationSelectionNext_Action WSA.ActivePresentation
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub cmdPresentationSlideSelectionMode_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorPresentationSlideSelectionMode_Action WSA.ActivePresentation
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub cmdPresentationSlideSelectionClear_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorPresentationSlideSelectionClear_Action WSA.ActivePresentation
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub lstPresentationSlideSelectionList_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorPresentationSlideSelectionUpdate_Action WSA.ActivePresentation
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub cmdBannerConfigurationBannerDisable_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorBannerConfigurationBannerDisable_Action WSA.ActivePresentation
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub cmdBannerShowLoad_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorBannerShowLoad_Action WSA.ActivePresentation
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub cmdBannerShowHide_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorBannerShowHide_Action WSA.ActivePresentation
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub cmdBannerColorPrev_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorBannerColorPrev_Action WSA.ActivePresentation
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub cmdBannerColorNext_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorBannerColorNext_Action WSA.ActivePresentation
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub cmdBannerSelectionClear_Click _
( _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigatorBannerSelectionClear_Action WSA.ActivePresentation
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Forces the fraEmpty frame to hold the focus if the form's focus is locked.
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub fraEmpty_Exit _
( _
    ByVal Cancel As MSForms.ReturnBoolean _
)
    Cancel = mblnNavigatorFocusLocked
    If (Cancel = True) Then
        mblnValid
    End If
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub fraEmpty_KeyPress _
( _
    ByVal intKeyASCII As MSForms.ReturnInteger _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigator_KeyPress WSA.ActivePresentation, intKeyASCII
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub fraEmpty_KeyDown _
( _
    ByVal intKeyCode As MSForms.ReturnInteger, _
    ByVal intKeyModifier As Integer _
)
    If (mblnValid = False) Then
        Exit Sub
    End If
    mNavigator_KeyDown WSA.ActivePresentation, intKeyCode, intKeyModifier
    mblnValid
End Sub
