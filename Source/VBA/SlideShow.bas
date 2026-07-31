Attribute VB_Name = "SlideShow"
'===============================================================================
' Name:
'   WorshipServiceAssistant.SlideShow
'
' Description:
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
Enum SlideShow_WindowDisplay
    Last = 0
    First = 1
    Display01 = 2
    Display02 = 3
    Display03 = 4
    Display04 = 5
End Enum

Enum SlideShow_WindowSize
    Full = 0
    Half = 1
    Quarter = 2
End Enum


'===============================================================================
' Public Variables.
'===============================================================================


'===============================================================================
' Private Constants.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================
Private SlideShowWindowDisplay As SlideShow_WindowDisplay
Private SlideShowWindowSize As SlideShow_WindowSize


'===============================================================================
' Public Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub SlideShow_Initialize()
    SlideShowWindowDisplay = SlideShow_WindowDisplay.Last
    SlideShowWindowSize = SlideShow_WindowSize.Full
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub SlideShow_SetWindowDisplay(ByVal WindowDisplay As SlideShow_WindowDisplay)
    SlideShowWindowDisplay = WindowDisplay
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub SlideShow_SetWindowSize(ByVal WindowSize As SlideShow_WindowSize)
    SlideShowWindowSize = WindowSize
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Function SlideShow_GetWindowDisplay() As SlideShow_WindowDisplay
    SlideShow_GetWindowDisplay = SlideShowWindowDisplay
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Function SlideShow_GetWindowSize() As SlideShow_WindowSize
    SlideShow_GetWindowSize = SlideShowWindowSize
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub SlideShow_Setup(ByVal P As Presentation)
    Dim PSaved As Boolean
    
    '
    ' Cannot setup a slide show if the presentation has an active slide show.
    '
    If (ActiveSlideShowExists(P.Windows(1)) = True) Then
        Exit Sub
    End If
    '
    ' Cannot setup a slide show if the presentation has no slides.
    '
    If (ActiveWindowSlideExists(P.Windows(1)) = False) Then
        Exit Sub
    End If
    
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
    ' not part of the PowerPoint 9.0 object hierarchy.  This hack starts
    ' the slide show on the last monitor in the "Show On" list.
    '
    
    P.Windows(1).Activate
    '
    ' Set the slide show window display.  By default, the slide show uses
    ' the last display.
    '
    Select Case SlideShowWindowDisplay
        Case SlideShow_WindowDisplay.Last:
            SendKeys "%dS%o{PGDN}{ENTER}", True
        Case SlideShow_WindowDisplay.First:
            SendKeys "%dS%o{PGUP}{ENTER}", True
        Case SlideShow_WindowDisplay.Display01:
            SendKeys "%dS%o{PGUP}{ENTER}", True
        Case SlideShow_WindowDisplay.Display02:
            SendKeys "%dS%o{PGUP}{DOWN}{DOWN}{ENTER}{ENTER}", True
        Case SlideShow_WindowDisplay.Display03:
            SendKeys "%dS%o{PGUP}{DOWN}{DOWN}{DOWN}{ENTER}{ENTER}", True
        Case SlideShow_WindowDisplay.Display04:
            SendKeys "%dS%o{PGUP}{DOWN}{DOWN}{DOWN}{DOWN}{ENTER}{ENTER}", True
        Case Else:
            SendKeys "%dS%o{PGDN}{ENTER}", True
    End Select
    
    P.Saved = PSaved
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub SlideShow_End()
    Dim I As Long
    
    '
    ' Exit all slide shows.  For the sake of appearance, all slide shows
    ' are blacked before any slide shows are exited.
    '
    For I = Application.SlideShowWindows.Count To 1 Step -1
        Application.SlideShowWindows(I).View.State = ppSlideShowBlackScreen
    Next
    For I = Application.SlideShowWindows.Count To 1 Step -1
        Application.SlideShowWindows(I).View.Exit
    Next
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub SlideShow_Begin(ByVal W As DocumentWindow)
    Dim P As Presentation
    Dim PSaved As Boolean
    
    Set P = W.Presentation
    PSaved = P.Saved
    
    '
    ' Start a slide show window for the presentation currentPresentation.
    ' If a slide show window already exists, a new window will not be
    ' started.  This will also activate the slide show window.
    '
    P.SlideShowSettings.Run

    '
    ' Set the slide show window size.  By default, the slide show uses
    ' a full size window.
    '
    Select Case SlideShowWindowSize
        Case SlideShow_WindowSize.Full:
        Case SlideShow_WindowSize.Half:
            P.SlideShowWindow.Height = P.SlideShowWindow.Height / 2
            P.SlideShowWindow.Width = P.SlideShowWindow.Width / 2
        Case SlideShow_WindowSize.Quarter:
            P.SlideShowWindow.Height = P.SlideShowWindow.Height / 4
            P.SlideShowWindow.Width = P.SlideShowWindow.Width / 4
        Case Else:
    End Select
        
    P.Saved = PSaved
    
    W.Presentation.SlideShowWindow.View.State = ppSlideShowPaused
    P.SlideShowWindow.View.PointerType = ppSlideShowPointerArrow
End Sub

'-------------------------------------------------------------------------------
' Description:
'   Load the currently active slide in the presentation's window into the
'   presentation's slide show window.
'-------------------------------------------------------------------------------
Public Sub SlideShow_Load(ByVal W As DocumentWindow)
    '
    ' Abort if no active slide show exists.
    '
    If (ActiveSlideShowExists(W) = False) Then
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
Public Sub SlideShow_Hide(ByVal W As DocumentWindow)
    '
    ' Abort if no active slide show exists.
    '
    If (ActiveSlideShowExists(W) = False) Then
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
Public Sub SlideShow_Run(ByVal W As DocumentWindow)
    '
    ' Abort if no active slide show exists.
    '
    If (ActiveSlideShowExists(W) = False) Then
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
Public Sub SlideShow_Pause(ByVal W As DocumentWindow)
    '
    ' Abort if no active slide show exists.
    '
    If (ActiveSlideShowExists(W) = False) Then
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
Public Sub SlideShow_Next(ByVal W As DocumentWindow)
    '
    ' Abort if no active slide show exists.
    '
    If (ActiveSlideShowExists(W) = False) Then
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
Public Sub SlideShow_Prev(ByVal W As DocumentWindow)
    '
    ' Abort if no active slide show exists.
    '
    If (ActiveSlideShowExists(W) = False) Then
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
