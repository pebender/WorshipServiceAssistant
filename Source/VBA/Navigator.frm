VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Navigator 
   Caption         =   "Navigator"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   OleObjectBlob   =   "Navigator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Navigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================================
' Name:
'   WorshipServiceAssistant.Navigator
'
' Warning:
'   To use the form, use WorshipServiceAssistant.Application.ActivateNavigator.
'   Calls to the form's Load or Show methods will result in incorrect behavior.
'
' Description:
'   The form implements the Navigator.
'
'   The form is modal, allowing the form to take more control.
'
'   The form contains an empty frame (fraEmpty) to which the focus is locked.
'   The empty frame is not visible to the user because its height and width
'   are both 0. The form handles all if the keyboard input processing using
'   its KeyPress and KeyDown event handlers.
'
'   Since the form is modal, the user cannot change any of the presentations
'   while the form is running.  As a result, the form can perform all of the
'   configurations of the presentations when it is initialized.  This makes
'   the lauching of the form take more time. However, after that, the form
'   responds much more quickly.
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
' Public Variables.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================

' Indicates whether or not the Navigator form's focus is locked to the frmEmpty
' frame.  Normally the focus is locked to the frmEmpty frame. However, sometimes
' it is necessary to change the focus in order to force the Navigator form
' to be the active window.  During this time, the focus must be unlocked.
Private mblnNavigatorFocusLocked As Boolean


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
    If (Not mblnValid) Then
        Exit Sub
    End If
    
    mNavigator_Update WorshipServiceAssistant.Application.ActivePresentation
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
    Me.Height = WorshipServiceAssistant.Application.Doppelganger.Height
    
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
    Me.Left = WorshipServiceAssistant.Application.Doppelganger.Left - Me.Width
    Me.Top = WorshipServiceAssistant.Application.Doppelganger.Top
    
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
'
' Purpose:
'   Updates the form/frame/control to reflect the currently associated
'   presentation (input udtPresentation).
' Inputs:
'    udtPresentation:
'      The current presentation with which the form/frame/control is associated.
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Private Sub mNavigator_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    mNavigatorBannerSelection_Update udtPresentation
    mNavigatorBannerColor_Update udtPresentation
    mNavigatorBannerShow_Update udtPresentation
    mNavigatorBannerConfiguration_Update udtPresentation
    mNavigatorPresentationSlideSelection_Update udtPresentation
    mNavigatorPresentationSelection_Update udtPresentation
    mNavigatorPresentationSlideShow_Update udtPresentation
    
    If (WorshipServiceAssistant.Application.Presentations.Count <= 0) Then
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

Private Sub mNavigatorGeneral_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    ' Enable everything by default.
    Me.fraGeneral.Enabled = True
    Me.cmdGeneralPageNext.Enabled = True
    Me.cmdGeneralHelp.Enabled = True
    Me.cmdGeneralExit.Enabled = True
End Sub

'-------------------------------------------------------------------------------
' Assumptions:
'   In setting the state of the fraPresentationSlideShow controls, this routine
'   assumes that the lstPresentationSlideSelectionList control is up-to-date.
'   Therefore, routines which update fraPresentationSlideShowList control should
'   be called before this routine.
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideShow_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    ' Enable everything by default.
    Me.fraPresentationSlideShow.Enabled = True
    Me.cmdPresentationSlideShowLoad.Enabled = True
    Me.cmdPresentationSlideShowHide.Enabled = True
    Me.cmdPresentationSlideShowRun.Enabled = True
    Me.cmdPresentationSlideShowPause.Enabled = True
    Me.cmdPresentationSlideShowEffectPrev.Enabled = True
    Me.cmdPresentationSlideShowEffectNext.Enabled = True
    
    If ((WorshipServiceAssistant.Application.Presentations.Count <= 0) Or (udtPresentation Is Nothing)) Then
        Me.fraPresentationSlideShow.Enabled = False
        Me.cmdPresentationSlideShowLoad.Enabled = False
        Me.cmdPresentationSlideShowHide.Enabled = False
        Me.cmdPresentationSlideShowRun.Enabled = False
        Me.cmdPresentationSlideShowPause.Enabled = False
        Me.cmdPresentationSlideShowEffectPrev.Enabled = False
        Me.cmdPresentationSlideShowEffectNext.Enabled = False
        Exit Sub
    End If
    
    If (udtPresentation.SlideShowWindow.Doppelganger Is Nothing) Then
    ' Since the current presentation does not have an associated slide show,
    ' disable the slide show controls, other than the load control.
        Me.cmdPresentationSlideShowHide.Caption = "Hide"
        Me.cmdPresentationSlideShowHide.Enabled = False
        Me.cmdPresentationSlideShowRun.Enabled = False
        Me.cmdPresentationSlideShowPause.Enabled = False
        Me.cmdPresentationSlideShowEffectPrev.Enabled = False
        Me.cmdPresentationSlideShowEffectNext.Enabled = False
    Else
        With udtPresentation.SlideShowWindow
            If (Me.lstPresentationSlideSelectionList.ListCount = 0) Then
                Me.cmdPresentationSlideShowLoad.Enabled = False
            End If
            If (Not .Visible) Then
            ' Since the slide show is hidden,
            ' label the hide control as 'Show'.
                Me.cmdPresentationSlideShowHide.Caption = "Show"
            End If
            If (.Running) Then
            ' Since the slide show is running,
            ' disable the run control.
                Me.cmdPresentationSlideShowRun.Enabled = False
            End If
            If (.Paused) Then
            ' Since the slide show is paused,
            ' disable the pause control.
                Me.cmdPresentationSlideShowPause.Enabled = False
            End If
        End With
    End If
End Sub

Private Sub mNavigatorPresentationSelection_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    ' Enable everything by default.
    Me.fraPresentationSelection.Enabled = True
    Me.lblPresentationSelectionName.Enabled = True
    Me.cmdPresentationSelectionPrev.Enabled = True
    Me.cmdPresentationSelectionNext.Enabled = True
    
    ' Update the lstControlPresentationName control.
    mNavigatorPresentationSelectionName_Update udtPresentation
    
    If ((WorshipServiceAssistant.Application.Presentations.Count <= 0) Or (udtPresentation Is Nothing)) Then
        Me.lblPresentationSelectionName.Enabled = False
        Me.cmdPresentationSelectionPrev.Enabled = False
        Me.cmdPresentationSelectionNext.Enabled = False
        Exit Sub
    End If
    
    If (WorshipServiceAssistant.Application.Presentations.Count <= 1) Then
        Me.cmdPresentationSelectionPrev.Enabled = False
        Me.cmdPresentationSelectionNext.Enabled = False
    End If
End Sub

Private Sub mNavigatorPresentationSelectionName_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    Dim strName As String
    
    If ((WorshipServiceAssistant.Application.Presentations.Count <= 0) Or (udtPresentation Is Nothing)) Then
        Me.lblPresentationSelectionName.Caption = ""
        Me.lblPresentationSelectionName.Tag = ""
        Exit Sub
    End If
    
    strName = udtPresentation.Name
    If (Not Me.lblPresentationSelectionName.Tag = strName) Then
        Me.lblPresentationSelectionName.Tag = strName
        If (VBA.Len(strName) >= 4) Then
            If (VBA.LCase(VBA.Right(strName, 4)) = ".ppt") Then
                strName = VBA.Left(strName, VBA.Len(strName) - 4)
            End If
        End If
        Me.lblPresentationSelectionName.Caption = strName
    End If
End Sub

Private Sub mNavigatorPresentationSlideSelection_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
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
    
    If ((WorshipServiceAssistant.Application.Presentations.Count <= 0) Or (udtPresentation Is Nothing)) Then
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
'   Updates the lstPresentationSlideSelectionList control, so that its List
'   property contains the numbers and titles of the slides in the currently
'   associated presentation (input udtPresentation) which satisfy the filter
'   criteria specified by the selection filter mode (form control
'   cmdPresentationSlideSelectionMode) and the slide selection filter (form
'   controls cmdSlideSelectionNumber and cmdSlideSelectionText).
'-------------------------------------------------------------------------------
Private Sub mNavigatorPresentationSlideSelectionList_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
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

    If ((WorshipServiceAssistant.Application.Presentations.Count <= 0) Or (udtPresentation Is Nothing)) Then
        Me.lstPresentationSlideSelectionList.Tag = ""
        Me.lstPresentationSlideSelectionList.Clear
        Exit Sub
    End If

    If (udtPresentation.Windows.Item(1).ActiveSlide Is Nothing) Then
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
        If (IsNumeric(strFilterText)) Then
        ' Since the filter text is numeric, assume the it is a slide number.
            lngFilterNumber = strFilterText
            If ((lngFilterNumber > 0) And (lngFilterNumber <= .Count)) Then
                lngMatchCount = 1
                ReDim lngMatch(lngMatchCount - 1)
                With .Item(lngFilterNumber)
                    lngMatch(0) = .SlideIndex
                End With
            End If
        Else
        ' Since the filter text is non-numeric, assume that it is slide text.
            If (Me.cmdPresentationSlideSelectionMode.Tag = "0") Then
            ' Selection mode is "Match Title"
                lngMatchCount = .TitleMatches(lngMatch, strFilterText)
            ElseIf (Me.cmdPresentationSlideSelectionMode.Tag = "1") Then
            ' Selection mode is "Match Title and Body"
                lngMatchCount = .TitleOrBodyMatches(lngMatch, strFilterText)
            End If
        End If
        If (lngMatchCount > 0) Then
            ReDim lngList(lngMatchCount - 1, 1) As String
            For lngListCount = 0 To lngMatchCount - 1 Step 1
                With .Item(lngMatch(lngListCount))
                    lngList(lngListCount, 0) = .SlideIndex
                    lngList(lngListCount, 1) = .Title
                    If (.SlideIndex = lngSelectedSlideIndex) Then
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

Private Sub mNavigatorBannerConfiguration_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    ' Enable everything by default.
    Me.fraBannerConfiguration.Enabled = True
    Me.cmdBannerConfigurationBannerDisable.Enabled = True
    
    ' Set the default control captions.
    Me.cmdBannerConfigurationBannerDisable.Caption = "Banner Disable"
    
    ' Since the banner is disabled,
    ' label the banner disable control as 'Banner Enable'.
    If (Not WorshipServiceAssistant.Application.Banner.Enabled) Then
        Me.cmdBannerConfigurationBannerDisable.Caption = "Banner Enable"
    End If
End Sub

Private Sub mNavigatorBannerShow_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    ' Enable everything by default.
    Me.fraBannerShow.Enabled = True
    Me.cmdBannerShowLoad.Enabled = True
    Me.cmdBannerShowHide.Enabled = True
    
    ' Set the default control captions.
    Me.cmdBannerShowHide.Caption = "Hide"
    
    ' Since the banner is hidden,
    ' label the hide control as 'Show'.
    If (Not WorshipServiceAssistant.Application.Banner.Visible) Then
        Me.cmdBannerShowHide.Caption = "Show"
    End If
    
    ' Since the banner is disabled,
    ' disable all the banner show controls.
    If (Not WorshipServiceAssistant.Application.Banner.Enabled) Then
        Me.cmdBannerShowLoad.Enabled = False
        Me.cmdBannerShowHide.Enabled = False
        Me.fraBannerShow.Enabled = False
    End If
End Sub

Private Sub mNavigatorBannerColor_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    ' Enable everything by default.
    Me.fraBannerColor.Enabled = True
    Me.lblBannerColorName.Enabled = True
    Me.cmdBannerColorPrev.Enabled = True
    Me.cmdBannerColorNext.Enabled = True
    
    ' Set banner color name color.
    If (WorshipServiceAssistant.Application.Banner.Enabled = True) Then
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
    If (Not WorshipServiceAssistant.Application.Banner.Enabled) Then
        Me.lblBannerColorName.Enabled = False
        Me.cmdBannerColorPrev.Enabled = False
        Me.cmdBannerColorNext.Enabled = False
        Me.fraBannerColor.Enabled = False
    End If
End Sub

Private Sub mNavigatorBannerSelection_Update _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    ' Enable everything by default.
    Me.fraBannerSelection.Enabled = True
    Me.lblBannerSelectionText.Enabled = True
    Me.cmdBannerSelectionClear.Enabled = True
    
    ' Since the banner is disabled,
    ' disable all the banner selection controls.
    If (Not WorshipServiceAssistant.Application.Banner.Enabled) Then
        Me.lblBannerSelectionText.Enabled = False
        Me.cmdBannerSelectionClear.Enabled = False
        Me.fraBannerSelection.Enabled = False
    End If
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' mNavigator*_KeyDown subroutines.
'
' Purpose:
'   Processes KeyDown events that effect the form/frame/control.
' Inputs:
'    udtPresentation:
'      The current presentation with which the form/frame/control is associated.
'   inKeyCode:
'     The KeyCode to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'   inKeyModifier:
'     The KeyModifier to be processed. A detailed description can be found in
'     the VBA help for the KeyDown event for UserForm controls.
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Private Sub mNavigator_KeyDown _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation, _
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

Private Sub mNavigatorGeneral_KeyDown _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation, _
    ByRef intKeyCode As MSForms.ReturnInteger, _
    ByRef intKeyModifier As Integer _
)
    If (Not Me.fraGeneral.Enabled) Then
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

Private Sub mNavigatorPresentationSlideShow_KeyDown _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation, _
    ByRef intKeyCode As MSForms.ReturnInteger, _
    ByRef intKeyModifier As Integer _
)
    If (Not Me.fraPresentationSlideShow.Enabled) Then
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

Private Sub mNavigatorPresentationSelection_KeyDown _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation, _
    ByRef intKeyCode As MSForms.ReturnInteger, _
    ByRef intKeyModifier As Integer _
)
    If (Not Me.fraPresentationSelection.Enabled) Then
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

Private Sub mNavigatorPresentationSlideSelection_KeyDown _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation, _
    ByRef intKeyCode As MSForms.ReturnInteger, _
    ByRef intKeyModifier As Integer _
)
    If (Not Me.fraPresentationSlideSelection.Enabled) Then
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

Private Sub mNavigatorBannerConfiguration_KeyDown _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation, _
    ByRef intKeyCode As MSForms.ReturnInteger, _
    ByRef intKeyModifier As Integer _
)
    If (Not Me.fraBannerConfiguration.Enabled) Then
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

Private Sub mNavigatorBannerShow_KeyDown _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation, _
    ByRef intKeyCode As MSForms.ReturnInteger, _
    ByRef intKeyModifier As Integer _
)
    If (Not Me.fraBannerColor.Enabled) Then
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

Private Sub mNavigatorBannerColor_KeyDown _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation, _
    ByRef intKeyCode As MSForms.ReturnInteger, _
    ByRef intKeyModifier As Integer _
)
    If (Not Me.fraBannerColor.Enabled) Then
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

Private Sub mNavigatorBannerSelection_KeyDown _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation, _
    ByRef intKeyCode As MSForms.ReturnInteger, _
    ByRef intKeyModifier As Integer _
)
    If (Not Me.fraBannerSelection.Enabled) Then
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
'
' Purpose:
'   Processes KeyPress events that effect the form/frame/control.
' Inputs:
'    udtPresentation:
'      The current presentation with which the form/frame/control is associated.
'   inKeyASCII:
'     The KeyASCII to be processed. A detailed description can be found in
'     the VBA help for the KeyPress event for UserForm controls.
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Private Sub mNavigator_KeyPress _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation, _
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

Private Sub mNavigatorGeneral_KeyPress _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation, _
    ByRef intKeyASCII As MSForms.ReturnInteger _
)
End Sub

Private Sub mNavigatorPresentationSlideShow_KeyPress _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation, _
    ByRef intKeyASCII As MSForms.ReturnInteger _
)
End Sub

Private Sub mNavigatorPresentationSelection_KeyPress _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation, _
    ByRef intKeyASCII As MSForms.ReturnInteger _
)
End Sub

Private Sub mNavigatorPresentationSlideSelection_KeyPress _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation, _
    ByRef intKeyASCII As MSForms.ReturnInteger _
)
    Dim strFilter As String
    
    If (Not Me.fraPresentationSlideSelection.Enabled) Then
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

Private Sub mNavigatorBannerConfiguration_KeyPress _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation, _
    ByRef intKeyASCII As MSForms.ReturnInteger _
)
End Sub

Private Sub mNavigatorBannerShow_KeyPress _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation, _
    ByRef intKeyASCII As MSForms.ReturnInteger _
)
End Sub

Private Sub mNavigatorBannerColor_KeyPress _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation, _
    ByRef intKeyASCII As MSForms.ReturnInteger _
)
End Sub

Private Sub mNavigatorBannerSelection_KeyPress _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation, _
    ByRef intKeyASCII As MSForms.ReturnInteger _
)
    Dim strBanner As String
    
    If (Not Me.fraBannerSelection.Enabled) Then
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
'
' Purpose:
'   Perform the action for the form/frame/control.
' Inputs:
'    udtPresentation:
'      The current presentation with which the form/frame/control is associated.
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Private Sub mNavigatorGeneralPageNext_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    Dim lngPage As Long
    
    If (Not Me.cmdGeneralPageNext.Enabled) Then
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

Private Sub mNavigatorGeneralHelp_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    If (Not Me.cmdGeneralHelp.Enabled) Then
        Exit Sub
    End If
    
    WorshipServiceAssistant.Application.Help.ShowTopic WorshipServiceAssistant.wsaHelpTopicToolbarNavigator
End Sub

Private Sub mNavigatorGeneralExit_Action _
( _
)
    If (Not Me.cmdGeneralExit.Enabled) Then
        Exit Sub
    End If
    
    VBA.Unload Navigator
End Sub

Private Sub mNavigatorPresentationSlideShowLoad_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    If (Not Me.cmdPresentationSlideShowLoad.Enabled) Then
        Exit Sub
    End If
    
    If (udtPresentation.Windows.Item(1).ActiveSlide Is Nothing) Then
        Exit Sub
    End If
    
    Me.Hide

    udtPresentation.SlideShowWindow.Load Me.lstPresentationSlideSelectionList.Value
    
    udtPresentation.Activate
    
    Me.lblPresentationSlideSelectionNumber.Caption = ""
    Me.lblPresentationSlideSelectionText.Caption = ""
    
    mNavigator_Update udtPresentation
End Sub

Private Sub mNavigatorPresentationSlideShowHide_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    If (Not Me.cmdPresentationSlideShowHide.Enabled) Then
        Exit Sub
    End If
    
    With udtPresentation.SlideShowWindow
        .Visible = Not .Visible
    End With
    
    mNavigator_Update udtPresentation
End Sub

Private Sub mNavigatorPresentationSlideShowRun_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    If (Not Me.cmdPresentationSlideShowRun.Enabled) Then
        Exit Sub
    End If
    
    udtPresentation.SlideShowWindow.Run
    
    mNavigator_Update udtPresentation
End Sub

Private Sub mNavigatorPresentationSlideShowPause_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    If (Not Me.cmdPresentationSlideShowPause.Enabled) Then
        Exit Sub
    End If
    
    udtPresentation.SlideShowWindow.Pause
    
    mNavigator_Update udtPresentation
End Sub

Private Sub mNavigatorPresentationSlideShowEffectPrev_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    If (Not Me.cmdPresentationSlideShowEffectPrev.Enabled) Then
        Exit Sub
    End If
    
    udtPresentation.SlideShowWindow.EffectPrev
    
    Me.lblPresentationSlideSelectionNumber = ""
    Me.lblPresentationSlideSelectionText = ""
    
    mNavigator_Update udtPresentation
End Sub

Private Sub mNavigatorPresentationSlideShowEffectNext_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    If (Not Me.cmdPresentationSlideShowEffectNext.Enabled) Then
        Exit Sub
    End If
    
    udtPresentation.SlideShowWindow.EffectNext
    
    Me.lblPresentationSlideSelectionNumber = ""
    Me.lblPresentationSlideSelectionText = ""
    
    mNavigator_Update udtPresentation
End Sub

Private Sub mNavigatorPresentationSelectionPrev_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    Dim lngIndex As Long
    
    If (Not Me.cmdPresentationSelectionPrev.Enabled) Then
        Exit Sub
    End If
    
    With WorshipServiceAssistant.Application.Presentations
        lngIndex = .Index(udtPresentation) - 1
        If (lngIndex < 1) Then
            lngIndex = .Count
        End If
        .Item(lngIndex).Activate
        mNavigator_Update WorshipServiceAssistant.Application.ActivePresentation
    End With
End Sub

Private Sub mNavigatorPresentationSelectionNext_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    Dim lngIndex As Long
    
    If (Not Me.cmdPresentationSelectionNext.Enabled) Then
        Exit Sub
    End If
    
    With WorshipServiceAssistant.Application.Presentations
        lngIndex = .Index(udtPresentation) + 1
        If (lngIndex > .Count) Then
            lngIndex = 1
        End If
        .Item(lngIndex).Activate
        mNavigator_Update WorshipServiceAssistant.Application.ActivePresentation
    End With
End Sub

Private Sub mNavigatorPresentationSlideSelectionMode_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    If (Not Me.cmdPresentationSlideSelectionMode.Enabled) Then
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

Private Sub mNavigatorPresentationSlideSelectionClear_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    If (Not Me.cmdPresentationSlideSelectionClear.Enabled) Then
        Exit Sub
    End If
    
    Me.lblPresentationSlideSelectionNumber.Caption = ""
    Me.lblPresentationSlideSelectionText.Caption = ""
    mNavigator_Update udtPresentation
End Sub

Private Sub mNavigatorPresentationSlideSelectionUpdate_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    Dim lngIndex As Long
    
    If (Not Me.lstPresentationSlideSelectionList.Enabled) Then
        Exit Sub
    End If
    
    ' Set slide selected in the presentation to match the
    ' slide selected in the slide list control
    With Me.lstPresentationSlideSelectionList
        If (.ListIndex >= 0) Then
            lngIndex = .Value
            udtPresentation.Slides.Item(lngIndex).Activate
        End If
    End With
End Sub

Private Sub mNavigatorPresentationSlideSelectionPrev_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    If (Not Me.lstPresentationSlideSelectionList.Enabled) Then
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

Private Sub mNavigatorPresentationSlideSelectionNext_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    If (Not Me.lstPresentationSlideSelectionList.Enabled) Then
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

Private Sub mNavigatorBannerConfigurationBannerDisable_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    If (Not Me.cmdBannerConfigurationBannerDisable.Enabled) Then
        Exit Sub
    End If
    
    With WorshipServiceAssistant.Application.Banner
        .Enabled = Not .Enabled
    End With
    mNavigator_Update udtPresentation
End Sub

Private Sub mNavigatorBannerShowLoad_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    Dim lngRed As Long
    Dim lngGreen As Long
    Dim lngBlue As Long
    
    If (Not Me.cmdBannerShowLoad.Enabled) Then
        Exit Sub
    End If
    
    With Me.lblBannerColorName
        If (.ListIndex >= 0) Then
            lngRed = VBA.CLng(.List(.ListIndex, 1))
            lngGreen = VBA.CLng(.List(.ListIndex, 2))
            lngBlue = VBA.CLng(.List(.ListIndex, 3))
        Else
            lngRed = 0
            lngGreen = 0
            lngBlue = 0
        End If
    End With
    
    WorshipServiceAssistant.Application.Banner.Load Me.lblBannerSelectionText.Caption, lngRed, lngGreen, lngBlue
    mNavigatorBannerSelectionClear_Action udtPresentation
    mNavigator_Update udtPresentation
End Sub

Private Sub mNavigatorBannerShowHide_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    If (Not Me.cmdBannerShowHide.Enabled) Then
        Exit Sub
    End If
    
    With WorshipServiceAssistant.Application.Banner
        .Visible = Not .Visible
    End With
    mNavigator_Update udtPresentation
End Sub

Private Sub mNavigatorBannerColorPrev_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    If (Not Me.cmdBannerColorPrev.Enabled) Then
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

Private Sub mNavigatorBannerColorNext_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    If (Not Me.cmdBannerColorNext.Enabled) Then
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

Private Sub mNavigatorBannerSelectionClear_Action _
( _
    ByRef udtPresentation As WorshipServiceAssistant.Presentation _
)
    If (Not Me.cmdBannerSelectionClear.Enabled) Then
        Exit Sub
    End If
    
    If (Not WorshipServiceAssistant.Application.Banner.Enabled) Then
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
    If (WorshipServiceAssistant.Application.Presentations.Count > 0) Then
        If (WorshipServiceAssistant.Application.ActivePresentation Is Nothing) Then
            mblnValid = False
            Me.Hide
        End If
    End If
End Function

'-------------------------------------------------------------------------------
' Purpose:
'-------------------------------------------------------------------------------
Private Sub mInitialize _
( _
)
    WorshipServiceAssistant.Application.Invariant = True
    
    If (WorshipServiceAssistant.Application.Presentations.Count > 0) Then
        Do While (WorshipServiceAssistant.Application.ActivePresentation Is Nothing)
            WorshipServiceAssistant.Application.Presentations.Item(1).Activate
        Loop
    End If
    
    WorshipServiceAssistant.Application.Doppelganger.Left = WorshipServiceAssistant.Application.Doppelganger.Left + Me.Width
    WorshipServiceAssistant.Application.Doppelganger.Width = WorshipServiceAssistant.Application.Doppelganger.Width - Me.Width
    
    mNavigator_Initialize
    
    mNavigator_Update WorshipServiceAssistant.Application.ActivePresentation
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'-------------------------------------------------------------------------------
Private Sub mTerminate _
( _
)
    WorshipServiceAssistant.Application.Invariant = False
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Inputs:
'-------------------------------------------------------------------------------
Private Sub mQueryClose _
( _
    ByRef intCancel As Integer, _
    ByRef intCloseMode As Integer _
)
    Dim intResponse As VBA.VbMsgBoxResult
    
    If ((intCloseMode = VBA.vbFormControlMenu) Or _
        (intCloseMode = VBA.vbFormCode)) Then
        intResponse = VBA.MsgBox( _
            Buttons:= _
                VBA.vbYesNo + VBA.vbDefaultButton2 + VBA.vbExclamation, _
            Title:= _
                WorshipServiceAssistant.wsaApplicationNamePretty, _
            Prompt:= _
                "Are you sure you want to exit the Navigator?")
        intCancel = VBA.IIf(intResponse = VBA.vbYes, 0, 1)
    End If
End Sub


'===============================================================================
' Private Event Handlers.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
'-------------------------------------------------------------------------------
Private Sub UserForm_Initialize _
( _
)
    mInitialize
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'-------------------------------------------------------------------------------
Private Sub UserForm_Terminate _
( _
)
    mTerminate
End Sub

'-------------------------------------------------------------------------------
' Purpose:
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
'-------------------------------------------------------------------------------
Private Sub UserForm_Activate _
( _
)
    mblnNavigatorFocusLocked = True
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'-------------------------------------------------------------------------------
Private Sub tabPages_Change _
( _
)
    mNavigator_Update WorshipServiceAssistant.Application.ActivePresentation
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Inputs:
'-------------------------------------------------------------------------------
Private Sub fraEmpty_Exit _
( _
    ByVal Cancel As MSForms.ReturnBoolean _
)
    Cancel = mblnNavigatorFocusLocked
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'-------------------------------------------------------------------------------
Private Sub fraEmpty_KeyPress _
( _
    ByVal intKeyASCII As MSForms.ReturnInteger _
)
    If (Not mblnValid) Then Exit Sub
    mNavigator_KeyPress WorshipServiceAssistant.Application.ActivePresentation, intKeyASCII
    mblnValid
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'-------------------------------------------------------------------------------
Private Sub fraEmpty_KeyDown _
( _
    ByVal intKeyCode As MSForms.ReturnInteger, _
    ByVal intKeyModifier As Integer _
)
    If (Not mblnValid) Then Exit Sub
    mNavigator_KeyDown WorshipServiceAssistant.Application.ActivePresentation, intKeyCode, intKeyModifier
    mblnValid
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' *_Click event handlers.
'
' Purpose:
'   Perform the click action for the control.
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Private Sub cmdGeneralPageNext_Click _
( _
)
    If (Not mblnValid) Then Exit Sub
    mNavigatorGeneralPageNext_Action WorshipServiceAssistant.Application.ActivePresentation
    mblnValid
End Sub

Private Sub cmdGeneralHelp_Click _
( _
)
    If (Not mblnValid) Then Exit Sub
    mNavigatorGeneralHelp_Action WorshipServiceAssistant.Application.ActivePresentation
    mblnValid
End Sub

Private Sub cmdGeneralExit_Click _
( _
)
    mNavigatorGeneralExit_Action
End Sub

Private Sub cmdPresentationSlideShowLoad_Click _
( _
)
    If (Not mblnValid) Then Exit Sub
    mNavigatorPresentationSlideShowLoad_Action WorshipServiceAssistant.Application.ActivePresentation
    mblnValid
End Sub

Private Sub cmdPresentationSlideShowHide_Click _
( _
)
    If (Not mblnValid) Then Exit Sub
    mNavigatorPresentationSlideShowHide_Action WorshipServiceAssistant.Application.ActivePresentation
    mblnValid
End Sub

Private Sub cmdPresentationSlideShowRun_Click _
( _
)
    If (Not mblnValid) Then Exit Sub
    mNavigatorPresentationSlideShowRun_Action WorshipServiceAssistant.Application.ActivePresentation
    mblnValid
End Sub

Private Sub cmdPresentationSlideShowPause_Click _
( _
)
    If (Not mblnValid) Then Exit Sub
    mNavigatorPresentationSlideShowPause_Action WorshipServiceAssistant.Application.ActivePresentation
    mblnValid
End Sub

Private Sub cmdPresentationSlideShowEffectPrev_Click _
( _
)
    If (Not mblnValid) Then Exit Sub
    mNavigatorPresentationSlideShowEffectPrev_Action WorshipServiceAssistant.Application.ActivePresentation
    mblnValid
End Sub

Private Sub cmdPresentationSlideShowEffectNext_Click _
( _
)
    If (Not mblnValid) Then Exit Sub
    mNavigatorPresentationSlideShowEffectNext_Action WorshipServiceAssistant.Application.ActivePresentation
    mblnValid
End Sub

Private Sub cmdPresentationSelectionPrev_Click _
( _
)
    If (Not mblnValid) Then Exit Sub
    mNavigatorPresentationSelectionPrev_Action WorshipServiceAssistant.Application.ActivePresentation
    mblnValid
End Sub

Private Sub cmdPresentationSelectionNext_Click _
( _
)
    If (Not mblnValid) Then Exit Sub
    mNavigatorPresentationSelectionNext_Action WorshipServiceAssistant.Application.ActivePresentation
    mblnValid
End Sub

Private Sub cmdPresentationSlideSelectionMode_Click _
( _
)
    If (Not mblnValid) Then Exit Sub
    mNavigatorPresentationSlideSelectionMode_Action WorshipServiceAssistant.Application.ActivePresentation
    mblnValid
End Sub

Private Sub cmdPresentationSlideSelectionClear_Click _
( _
)
    If (Not mblnValid) Then Exit Sub
    mNavigatorPresentationSlideSelectionClear_Action WorshipServiceAssistant.Application.ActivePresentation
    mblnValid
End Sub

Private Sub lstPresentationSlideSelectionList_Click _
( _
)
    If (Not mblnValid) Then Exit Sub
    mNavigatorPresentationSlideSelectionUpdate_Action WorshipServiceAssistant.Application.ActivePresentation
    mblnValid
End Sub

Private Sub cmdBannerConfigurationBannerDisable_Click _
( _
)
    If (Not mblnValid) Then Exit Sub
    mNavigatorBannerConfigurationBannerDisable_Action WorshipServiceAssistant.Application.ActivePresentation
    mblnValid
End Sub

Private Sub cmdBannerShowLoad_Click _
( _
)
    If (Not mblnValid) Then Exit Sub
    mNavigatorBannerShowLoad_Action WorshipServiceAssistant.Application.ActivePresentation
    mblnValid
End Sub

Private Sub cmdBannerShowHide_Click _
( _
)
    If (Not mblnValid) Then Exit Sub
    mNavigatorBannerShowHide_Action WorshipServiceAssistant.Application.ActivePresentation
    mblnValid
End Sub

Private Sub cmdBannerColorPrev_Click _
( _
)
    If (Not mblnValid) Then Exit Sub
    mNavigatorBannerColorPrev_Action WorshipServiceAssistant.Application.ActivePresentation
    mblnValid
End Sub

Private Sub cmdBannerColorNext_Click _
( _
)
    If (Not mblnValid) Then Exit Sub
    mNavigatorBannerColorNext_Action WorshipServiceAssistant.Application.ActivePresentation
    mblnValid
End Sub

Private Sub cmdBannerSelectionClear_Click _
( _
)
    If (Not mblnValid) Then Exit Sub
    mNavigatorBannerSelectionClear_Action WorshipServiceAssistant.Application.ActivePresentation
    mblnValid
End Sub
