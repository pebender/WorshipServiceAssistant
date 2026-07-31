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
    ' Start a slide show window for the presentation currentPresentation.
    ' If a slide show window already exists, a new window will not be
    ' started.  This will also activate the slide show window.
    '
    P.SlideShowSettings.Run

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
