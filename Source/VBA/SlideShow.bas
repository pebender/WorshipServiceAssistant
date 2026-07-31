Attribute VB_Name = "SlideShow"
'===============================================================================
' Name:
'   WorshipServiceAssistant.SlideShow
'
' Description:
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
Private SlideShowWindowDisplay As Integer
Private SlideShowWindowSize As Integer


'===============================================================================
' Public Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Function SlideShow_IsSlideShow(ByVal P As PowerPoint.Presentation) As Boolean
    SlideShow_IsSlideShow = _
        (ActiveWindowSlideExists(P.Windows(1)) = True) And _
        (Banner.IsBanner(P) = False)
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub SlideShow_Initialize()
    SlideShowWindowDisplay = 0
    SlideShowWindowSize = 0
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub SlideShow_SetWindowDisplay(ByVal WindowDisplay As Integer)
    SlideShowWindowDisplay = WindowDisplay
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub SlideShow_SetWindowSize(ByVal WindowSize As Integer)
    SlideShowWindowSize = WindowSize
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Function SlideShow_GetWindowDisplay() As Integer
    SlideShow_GetWindowDisplay = SlideShowWindowDisplay
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Function SlideShow_GetWindowSize() As Integer
    SlideShow_GetWindowSize = SlideShowWindowSize
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub SlideShow_Setup(ByVal P As PowerPoint.Presentation)
    Dim PSaved As Boolean
    Dim Index As Integer
    
    '
    ' Cannot setup a slide show if the presentation has an active slide show.
    '
    If (ActiveSlideShowExists(P) = True) Then
        Exit Sub
    End If
    '
    ' Cannot setup a slide show if the presentation has no slides.
    '
'    If (ActiveWindowSlideExists(W) = False) Then
'        Exit Sub
'    End If
    
    PSaved = P.Saved

    '
    ' Configure slide show settings.  Unfortunately, the Slide Show
    ' display monitor is not part of the PowerPoint 9.0 object hierarchy.
    '
    With P.SlideShowSettings
        If (.ShowType <> ppShowTypeSpeaker) Then
            .ShowType = ppShowTypeSpeaker
        End If
        If (.RangeType <> ppShowAll) Then
            .RangeType = ppShowAll
        End If
        If (.AdvanceMode = ppSlideShowManualAdvance) Then
            .AdvanceMode = ppSlideShowManualAdvance
        End If
        If (.LoopUntilStopped <> msoTrue) Then
            .LoopUntilStopped = msoTrue
        End If
        If (.ShowWithAnimation <> msoTrue) Then
            .ShowWithAnimation = msoTrue
        End If
        If (.ShowWithNarration <> msoTrue) Then
            .ShowWithNarration = msoTrue
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
    P.Windows(1).Activate
    Application.CommandBars.FindControl(Id:=2744).Execute
    SendKeys "%o", True
    If (SlideShowWindowDisplay > 0) Then
        SendKeys "{PGUP}", True
        For Index = 1 To SlideShowWindowDisplay Step 1
            SendKeys "{DOWN}", True
        Next
        SendKeys "{ENTER}", True
    Else
        SendKeys "{PGDN}", True
    End If
    SendKeys "{ENTER}", True
    
    P.Saved = PSaved
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub SlideShow_End()
    Dim i As Long
    
    '
    ' Exit all slide shows.  For the sake of appearance, all slide shows
    ' are blacked before any slide shows are exited.
    '
    For i = Application.SlideShowWindows.Count To 1 Step -1
        If (Presentation.IsPresentation(Application.SlideShowWindows(i).Presentation) = True) Then
            Application.SlideShowWindows(i).View.State = ppSlideShowBlackScreen
        End If
    Next
    For i = Application.SlideShowWindows.Count To 1 Step -1
        If (Presentation.IsPresentation(Application.SlideShowWindows(i).Presentation) = True) Then
            Application.SlideShowWindows(i).View.Exit
        End If
    Next
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub SlideShow_Begin(ByVal W As PowerPoint.DocumentWindow)
    Dim P As PowerPoint.Presentation
    Dim PSaved As Boolean
    
    Set P = W.Presentation
    PSaved = P.Saved
    
    '
    ' Unfortunately, there appears to be no way to start the slide show
    ' with a black screen.  However, the first slide will flash for a
    ' moment.
    '
    If (ActiveSlideExists(W) = True) Then
        P.SlideShowSettings.Run.View.GotoSlide ActiveSlide(W).SlideIndex, msoTrue
    Else
        P.SlideShowSettings.Run
    End If
    
    '
    ' Set the slide show window size.
    '
    SlideShow_Scale W.Presentation
        
    P.Saved = PSaved
    
    W.Presentation.SlideShowWindow.View.State = ppSlideShowPaused
    P.SlideShowWindow.View.PointerType = ppSlideShowPointerArrow
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub SlideShow_Scale(ByVal P As PowerPoint.Presentation)
    Dim PSaved As Boolean
    Dim Top As Long
    Dim Left As Long
    Dim Height As Long
    Dim Width As Long
    Dim Index As Long
    
    PSaved = P.Saved
    
    '
    ' Scale the slide show size.
    '
    Height = P.SlideShowWindow.Height
    Width = P.SlideShowWindow.Width
    For Index = 1 To SlideShowWindowSize - 1 Step 1
        Height = Height / 2
        Width = Width / 2
    Next
    P.SlideShowWindow.Height = Height
    P.SlideShowWindow.Width = Width
        
    P.Saved = PSaved
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Load the currently active slide in the presentation's window into the
'   presentation's slide show window.
'-------------------------------------------------------------------------------
Public Sub SlideShow_Load(ByVal W As PowerPoint.DocumentWindow)
    '
    ' Abort if no active slide show exists.
    '
    If (ActiveSlideShowExists(W.Presentation) = False) Then
        Exit Sub
    End If
    '
    ' Abort if no active slide exists.
    '
    If (ActiveSlideExists(W) = False) Then
        Exit Sub
    End If
    
    Dim Index As Long
    
    Index = ActiveSlide(W).SlideIndex
    W.Presentation.SlideShowWindow.View.GotoSlide Index, msoFalse
    W.Presentation.SlideShowWindow.View.State = ppSlideShowPaused
    W.Presentation.SlideShowWindow.View.PointerType = ppSlideShowPointerArrow
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Toggles the slide show display between hidden and shown.
'-------------------------------------------------------------------------------
Public Sub SlideShow_Hide(ByVal W As PowerPoint.DocumentWindow)
    '
    ' Abort if no active slide show exists.
    '
    If (ActiveSlideShowExists(W.Presentation) = False) Then
        Exit Sub
    End If
    '
    ' Abort if no active slide exists.
    '
    If (ActiveSlideExists(W) = False) Then
        Exit Sub
    End If
    
    If (W.Presentation.SlideShowWindow.View.State = ppSlideShowBlackScreen) Then
        W.Presentation.SlideShowWindow.View.State = ppSlideShowRunning
        W.Presentation.SlideShowWindow.View.State = ppSlideShowPaused
        W.Presentation.SlideShowWindow.View.PointerType = ppSlideShowPointerArrow
    Else
        W.Presentation.SlideShowWindow.View.State = ppSlideShowBlackScreen
        W.Presentation.SlideShowWindow.View.PointerType = ppSlideShowPointerArrow
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Run the slide show associated with the presentation window.
'-------------------------------------------------------------------------------
Public Sub SlideShow_Run(ByVal W As PowerPoint.DocumentWindow)
    '
    ' Abort if no active slide show exists.
    '
    If (ActiveSlideShowExists(W.Presentation) = False) Then
        Exit Sub
    End If
    '
    ' Abort if no active slide exists.
    '
    If (ActiveSlideExists(W) = False) Then
        Exit Sub
    End If
    
    '
    ' Abort if the slide show is already running.
    '
    If (W.Presentation.SlideShowWindow.View.State = ppSlideShowRunning) Then
        Exit Sub
    End If
    
    Dim Index As Long
    
    Index = W.Presentation.SlideShowWindow.View.Slide.SlideIndex
    W.Presentation.SlideShowWindow.View.GotoSlide Index, msoFalse
    W.Presentation.SlideShowWindow.View.State = ppSlideShowRunning
    W.Presentation.SlideShowWindow.View.PointerType = ppSlideShowPointerArrow
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Pause the slide show associated with the presentation window.
'-------------------------------------------------------------------------------
Public Sub SlideShow_Pause(ByVal W As PowerPoint.DocumentWindow)
    '
    ' Abort if no active slide show exists.
    '
    If (ActiveSlideShowExists(W.Presentation) = False) Then
        Exit Sub
    End If
    '
    ' Abort if no active slide exists.
    '
    If (ActiveSlideExists(W) = False) Then
        Exit Sub
    End If
    
    W.Presentation.SlideShowWindow.View.State = ppSlideShowRunning
    W.Presentation.SlideShowWindow.View.State = ppSlideShowPaused
    W.Presentation.SlideShowWindow.View.PointerType = ppSlideShowPointerArrow
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Move forward in the windows's slide show.
'-------------------------------------------------------------------------------
Public Sub SlideShow_Next(ByVal W As PowerPoint.DocumentWindow)
    '
    ' Abort if no active slide show exists.
    '
    If (ActiveSlideShowExists(W.Presentation) = False) Then
        Exit Sub
    End If
    '
    ' Abort if no active slide exists.
    '
    If (ActiveSlideExists(W) = False) Then
        Exit Sub
    End If
    
    W.Presentation.SlideShowWindow.View.Next
    W.View.Slide = W.Presentation.SlideShowWindow.View.Slide
    W.Presentation.SlideShowWindow.View.State = ppSlideShowPaused
    W.Presentation.SlideShowWindow.View.PointerType = ppSlideShowPointerArrow
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Move backward in the window's slide show.
'-------------------------------------------------------------------------------
Public Sub SlideShow_Prev(ByVal W As PowerPoint.DocumentWindow)
    '
    ' Abort if no active slide show exists.
    '
    If (ActiveSlideShowExists(W.Presentation) = False) Then
        Exit Sub
    End If
    '
    ' Abort if no active slide exists.
    '
    If (ActiveSlideExists(W) = False) Then
        Exit Sub
    End If
    
    W.Presentation.SlideShowWindow.View.Previous
    W.View.Slide = W.Presentation.SlideShowWindow.View.Slide
    W.Presentation.SlideShowWindow.View.State = ppSlideShowPaused
    W.Presentation.SlideShowWindow.View.PointerType = ppSlideShowPointerArrow
End Sub

'===============================================================================
' Private Subroutines and Functions.
'===============================================================================
