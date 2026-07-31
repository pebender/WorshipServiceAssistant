Attribute VB_Name = "modSlideShow"
'===============================================================================
' Name:
'   WorshipServiceAssistant.modSlideShow
'
' Description:
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
'     (1) Made changes to the source code so that it follows Microsoft's
'         Visual Basic coding conventions.
'   1.01.0007:
'     (1) Fixed bug that would cause SlideShow_modSlideShow.gblnIsSlideShow to crash
'         when the presentation has no document windows.
'     (2) Made changes to the SlideShow_Prev and SlideShow_Next routines
'         so that they would work under PowerPoint 2002.
'   1.01.0000:
'     (1) Made room for the text banner above the slide show.
'     (2) Moved slide show scaling into a separate routine so that it could
'         be used by the Banner as well.
'   1.00.0001:
'     (1) Changed SlideShow_Setup routine so that the "Set Up Show" dialog
'         is activated by using the control ID rather than the command bar
'         and control names.
'     (2) Eliminated SlideShow_WindowDisplay and SlideShow_WindowSize constants.
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
' Public Variables.
'===============================================================================


'===============================================================================
' Private Constants.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================
Private mlngSlideShowWindowDisplay As Long
Private mlngSlideShowWindowSize As Long


'===============================================================================
' Public Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Function gblnIsSlideShow _
( _
    ByRef prePresentation As PowerPoint.Presentation _
) As Boolean
    gblnIsSlideShow = True
    If (gblnIsSlideShow) Then
        If (prePresentation.Windows.Count = 0) Then
            gblnIsSlideShow = False
        End If
    End If
    If (gblnIsSlideShow) Then
        If (modActive.gblnActiveWindowSlideExists(prePresentation.Windows(1)) = False) Then
            gblnIsSlideShow = False
        End If
    End If
    If (gblnIsSlideShow) Then
        If (modBanner.gblnIsBanner(prePresentation) = True) Then
            gblnIsSlideShow = False
        End If
    End If
    
    gblnIsSlideShow = gblnIsSlideShow
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub gInitialize _
( _
)
    mlngSlideShowWindowDisplay = 0
    mlngSlideShowWindowSize = 0
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub gWindowDisplaySet _
( _
    ByRef lngWindowDisplay As Long _
)
    mlngSlideShowWindowDisplay = lngWindowDisplay
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub gWindowSizeSet _
( _
    ByRef lngWindowSize As Long _
)
    mlngSlideShowWindowSize = lngWindowSize
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Function glngWindowDisplayGet _
( _
) As Long
    glngWindowDisplayGet = mlngSlideShowWindowDisplay
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Function glngWindowSizeGet _
( _
) As Long
    glngWindowSizeGet = mlngSlideShowWindowSize
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub gSetup _
( _
    ByRef prePresentation As PowerPoint.Presentation _
)
    Dim blnPresentationSaved As Boolean
    Dim lngIndex As Long
    
    '
    ' Cannot setup a slide show if the presentation has an active slide show.
    '
    If (modActive.gblnActiveSlideShowExists(prePresentation) = True) Then
        Exit Sub
    End If
    '
    ' Cannot setup a slide show if the presentation has no slides.
    '
'    If (modActive.gblnActiveWindowSlideExists(dwDocumentWindow) = False) Then
'        Exit Sub
'    End If
    
    blnPresentationSaved = prePresentation.Saved

    '
    ' Configure slide show settings.  Unfortunately, the Slide Show
    ' display monitor is not part of the PowerPoint 9.0 object hierarchy.
    '
    With prePresentation.SlideShowSettings
        If (.ShowType <> PowerPoint.ppShowTypeSpeaker) Then
            .ShowType = PowerPoint.ppShowTypeSpeaker
        End If
        If (.RangeType <> PowerPoint.ppShowAll) Then
            .RangeType = PowerPoint.ppShowAll
        End If
        If (.AdvanceMode = PowerPoint.ppSlideShowManualAdvance) Then
            .AdvanceMode = PowerPoint.ppSlideShowManualAdvance
        End If
        If (.LoopUntilStopped <> Office.msoTrue) Then
            .LoopUntilStopped = Office.msoTrue
        End If
        If (.ShowWithAnimation <> Office.msoTrue) Then
            .ShowWithAnimation = Office.msoTrue
        End If
        If (.ShowWithNarration <> Office.msoTrue) Then
            .ShowWithNarration = Office.msoTrue
        End If
    End With
    
    '
    ' Hack to workaround the fact that Slide Show display monitor is
    ' not part of the PowerPoint 9.0 object hierarchy.
    ' First, the presentation is activiated.
    ' Second, the "Set Up Show" dialog box is activated using
    '   its command bar control identifier.
    ' Third, the "Show On" dropdown is selected using
    '   its keyboard shortcut.
    ' Fourth, the desired monitor is selected using keyboard
    '   shortcuts.
    ' Fifth, the "Set Up Show" dialog box is closed using
    '   the ENTER key.
    '
    prePresentation.Windows(1).Activate
    Application.CommandBars.FindControl(Id:=2744).Execute
    VBA.SendKeys "%o", True
    If (mlngSlideShowWindowDisplay > 0) Then
        VBA.SendKeys "{PGUP}", True
        For lngIndex = 1 To mlngSlideShowWindowDisplay Step 1
            VBA.SendKeys "{DOWN}", True
        Next
        VBA.SendKeys "{ENTER}", True
    Else
        VBA.SendKeys "{PGDN}", True
    End If
    VBA.SendKeys "{ENTER}", True
    
    prePresentation.Saved = blnPresentationSaved
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub gQuit _
( _
)
    Dim lngIndex As Long
    
    '
    ' Exit all slide shows.  For the sake of appearance, all slide shows
    ' are blacked before any slide shows are exited.
    '
    For lngIndex = Application.SlideShowWindows.Count To 1 Step -1
        If (modPresentation.gblnIsPresentation(Application.SlideShowWindows(lngIndex).Presentation) = True) Then
            Application.SlideShowWindows(lngIndex).View.State = PowerPoint.ppSlideShowBlackScreen
        End If
    Next
    For lngIndex = Application.SlideShowWindows.Count To 1 Step -1
        If (modPresentation.gblnIsPresentation(Application.SlideShowWindows(lngIndex).Presentation) = True) Then
            Application.SlideShowWindows(lngIndex).View.Exit
        End If
    Next
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub gBegin _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    Dim prePresentation As PowerPoint.Presentation
    Dim blnPresentationSaved As Boolean
    
    Set prePresentation = dwDocumentWindow.Presentation
    blnPresentationSaved = prePresentation.Saved
    
    '
    ' Unfortunately, there appears to be no way to start the slide show
    ' with a black screen.  However, the first slide will flash for a
    ' moment.
    '
    If (modActive.gblnActiveSlideExists(dwDocumentWindow) = True) Then
        prePresentation.SlideShowSettings.Run.View.GotoSlide modActive.gppActiveSlideGet(dwDocumentWindow).SlideIndex, Office.msoTrue
    Else
        prePresentation.SlideShowSettings.Run
    End If
    
    '
    ' Set the slide show window size.
    '
    modSlideShow.gSizeSet dwDocumentWindow.Presentation
        
    prePresentation.Saved = blnPresentationSaved
    
    dwDocumentWindow.Presentation.SlideShowWindow.View.State = PowerPoint.ppSlideShowPaused
    prePresentation.SlideShowWindow.View.PointerType = PowerPoint.ppSlideShowPointerArrow
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub gSizeSet _
( _
    ByRef prePresentation As PowerPoint.Presentation _
)
    Dim blnPresentationSaved As Boolean
    Dim lngHeight As Long
    Dim lngWidth As Long
    Dim lngIndex As Long
    
    blnPresentationSaved = prePresentation.Saved
    
    '
    ' Set the slide show size.
    '
    lngHeight = prePresentation.SlideShowWindow.Height
    lngWidth = prePresentation.SlideShowWindow.Width
    For lngIndex = 1 To mlngSlideShowWindowSize - 1 Step 1
        lngHeight = lngHeight / 2
        lngWidth = lngWidth / 2
    Next
    prePresentation.SlideShowWindow.Height = lngHeight
    prePresentation.SlideShowWindow.Width = lngWidth
        
    prePresentation.Saved = blnPresentationSaved
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Load the currently active slide in the presentation's window into the
'   presentation's slide show window.
'-------------------------------------------------------------------------------
Public Sub gLoad _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    Dim lngIndex As Long
    
    '
    ' Abort if no active slide show exists.
    '
    If (modActive.gblnActiveSlideShowExists(dwDocumentWindow.Presentation) = False) Then
        Exit Sub
    End If
    '
    ' Abort if no active slide exists.
    '
    If (modActive.gblnActiveSlideExists(dwDocumentWindow) = False) Then
        Exit Sub
    End If
    
    lngIndex = modActive.gppActiveSlideGet(dwDocumentWindow).SlideIndex
    dwDocumentWindow.Presentation.SlideShowWindow.View.GotoSlide lngIndex, Office.msoFalse
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Toggles the slide show display between hidden and shown.
'-------------------------------------------------------------------------------
Public Sub gHide _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    '
    ' Abort if no active slide show exists.
    '
    If (modActive.gblnActiveSlideShowExists(dwDocumentWindow.Presentation) = False) Then
        Exit Sub
    End If
    '
    ' Abort if no active slide exists.
    '
    If (modActive.gblnActiveSlideExists(dwDocumentWindow) = False) Then
        Exit Sub
    End If
    
    If (dwDocumentWindow.Presentation.SlideShowWindow.View.State = PowerPoint.ppSlideShowBlackScreen) Then
        dwDocumentWindow.Presentation.SlideShowWindow.View.State = PowerPoint.ppSlideShowRunning
        dwDocumentWindow.Presentation.SlideShowWindow.View.State = PowerPoint.ppSlideShowPaused
        dwDocumentWindow.Presentation.SlideShowWindow.View.PointerType = PowerPoint.ppSlideShowPointerArrow
    Else
        dwDocumentWindow.Presentation.SlideShowWindow.View.State = PowerPoint.ppSlideShowBlackScreen
        dwDocumentWindow.Presentation.SlideShowWindow.View.PointerType = PowerPoint.ppSlideShowPointerArrow
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Run the slide show associated with the presentation window.
'-------------------------------------------------------------------------------
Public Sub gRun _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    Dim lngIndex As Long
    
    '
    ' Abort if no active slide show exists.
    '
    If (modActive.gblnActiveSlideShowExists(dwDocumentWindow.Presentation) = False) Then
        Exit Sub
    End If
    '
    ' Abort if no active slide exists.
    '
    If (modActive.gblnActiveSlideExists(dwDocumentWindow) = False) Then
        Exit Sub
    End If
    
    '
    ' Abort if the slide show is already running.
    '
    If (dwDocumentWindow.Presentation.SlideShowWindow.View.State = PowerPoint.ppSlideShowRunning) Then
        Exit Sub
    End If
    
    lngIndex = dwDocumentWindow.Presentation.SlideShowWindow.View.Slide.SlideIndex
    dwDocumentWindow.Presentation.SlideShowWindow.View.GotoSlide lngIndex, Office.msoFalse
    dwDocumentWindow.Presentation.SlideShowWindow.View.State = PowerPoint.ppSlideShowRunning
    dwDocumentWindow.Presentation.SlideShowWindow.View.PointerType = PowerPoint.ppSlideShowPointerArrow
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Pause the slide show associated with the presentation window.
'-------------------------------------------------------------------------------
Public Sub gPause _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    '
    ' Abort if no active slide show exists.
    '
    If (modActive.gblnActiveSlideShowExists(dwDocumentWindow.Presentation) = False) Then
        Exit Sub
    End If
    '
    ' Abort if no active slide exists.
    '
    If (modActive.gblnActiveSlideExists(dwDocumentWindow) = False) Then
        Exit Sub
    End If
    
    dwDocumentWindow.Presentation.SlideShowWindow.View.State = PowerPoint.ppSlideShowRunning
    dwDocumentWindow.Presentation.SlideShowWindow.View.State = PowerPoint.ppSlideShowPaused
    dwDocumentWindow.Presentation.SlideShowWindow.View.PointerType = PowerPoint.ppSlideShowPointerArrow
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Move forward in the windows's slide show.
'-------------------------------------------------------------------------------
Public Sub gEffectNext _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    Dim lngIndex As Long
    
    '
    ' Abort if no active slide show exists.
    '
    If (modActive.gblnActiveSlideShowExists(dwDocumentWindow.Presentation) = False) Then
        Exit Sub
    End If
    '
    ' Abort if no active slide exists.
    '
    If (modActive.gblnActiveSlideExists(dwDocumentWindow) = False) Then
        Exit Sub
    End If
    
    '
    ' PowerPoint 2002 effects may not advance correctly
    ' if the slide show is not running.
    ' Therefore, start the slide show running if it is not running.
    '
    If (dwDocumentWindow.Presentation.SlideShowWindow.View.State <> PowerPoint.ppSlideShowRunning) Then
        modSlideShow.gRun dwDocumentWindow
    End If
    
    dwDocumentWindow.Presentation.SlideShowWindow.View.Next
    dwDocumentWindow.View.Slide = dwDocumentWindow.Presentation.SlideShowWindow.View.Slide
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Move backward in the window's slide show.
'-------------------------------------------------------------------------------
Public Sub gEffectPrev _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    Dim lngIndex As Long
    
    '
    ' Abort if no active slide show exists.
    '
    If (modActive.gblnActiveSlideShowExists(dwDocumentWindow.Presentation) = False) Then
        Exit Sub
    End If
    '
    ' Abort if no active slide exists.
    '
    If (modActive.gblnActiveSlideExists(dwDocumentWindow) = False) Then
        Exit Sub
    End If
    
    '
    ' PowerPoint 2002 effects may not advance correctly
    ' if the slide show is not running.
    ' Therefore, start the slide show running if it is not running.
    '
    If (dwDocumentWindow.Presentation.SlideShowWindow.View.State <> PowerPoint.ppSlideShowRunning) Then
        modSlideShow.gRun dwDocumentWindow
    End If
    
    dwDocumentWindow.Presentation.SlideShowWindow.View.Previous
    dwDocumentWindow.View.Slide = dwDocumentWindow.Presentation.SlideShowWindow.View.Slide
End Sub

'===============================================================================
' Private Subroutines and Functions.
'===============================================================================
