VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WSASlideShowWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.WSASlideShowWindow
'
' Description:
'
' Author:
'   Paul Bender <pbender@alumni.ucsd.edu>
'
' Copyright:
'   Copyright (c) 2002 Paul Bender
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
'   1.04.0002:
'     (1) Removed code that set the slide show window monitor.
'         Setting of the slide show window monitor is now being performed
'         by WSAApplication.
'   1.04.0000:
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
Private Const mlngPpControlIdSlideShowSetUpShow As Long = 2744


'===============================================================================
' Private Variables.
'===============================================================================
Private mudtApplication As WorshipServiceAssistant.WSAApplication
Private mudtPresentation As WorshipServiceAssistant.WSAPresentation


'===============================================================================
' Public Properties and Methods.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Exists _
( _
) As Boolean
    On Error GoTo Exists_False
    
    Exists = True
    
    If (VBA.IsNull(mudtPresentation.Doppelganger.SlideShowWindow) = True) Then
        GoTo Exists_False
    End If
    If (VBA.IsEmpty(mudtPresentation.Doppelganger.SlideShowWindow) = True) Then
        GoTo Exists_False
    End If
Exit Property

Exists_False:
    Exists = False
End Property

'-------------------------------------------------------------------------------
' Purpose:
'   Load the currently active slide in the presentation's window into the
'   presentation's slide show window.
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Load _
( _
    ByRef lngSlideIndex As Long _
)
    Dim lngIndex As Long
    
    ' Begin a the slide show if no active slide show exists.
    If (Me.Exists = False) Then
        mCreate lngSlideIndex
    End If
    
    With mudtPresentation.Doppelganger
        ' Abort if the slide index is out of range.
        If ((lngSlideIndex < 1) Or (lngSlideIndex > .Slides.Count)) Then
            Exit Sub
        End If
        
        .SlideShowWindow.View.GotoSlide lngSlideIndex, Office.msoFalse
        .SlideShowWindow.Activate
    End With
    
    ' Black other slide shows.
    ' This will stop the slide shows from running.
    With mudtApplication.Presentations
        For lngIndex = 1 To .Count Step 1
            If ((.Item(lngIndex) Is mudtPresentation) = False) Then
                .Item(lngIndex).SlideShowWindow.Visible = False
            End If
        Next
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Visible _
( _
) As Boolean
    ' Abort if no active slide show exists.
    If (Me.Exists = False) Then
        Visible = False
        Exit Property
    End If
    
    With mudtPresentation.Doppelganger.SlideShowWindow.View
        If (.State = PowerPoint.ppSlideShowBlackScreen) Then
            Visible = False
        Else
            Visible = True
        End If
    End With
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Let Visible _
( _
    ByRef blnVisible As Boolean _
)
    ' Abort if no active slide show exists.
    If (Me.Exists = False) Then
        Exit Property
    End If
    
    With mudtPresentation.Doppelganger.SlideShowWindow.View
        If (blnVisible = True) Then
            .State = PowerPoint.ppSlideShowRunning
            .State = PowerPoint.ppSlideShowPaused
            .PointerType = PowerPoint.ppSlideShowPointerArrow
        Else
            .State = PowerPoint.ppSlideShowBlackScreen
            .PointerType = PowerPoint.ppSlideShowPointerArrow
        End If
    End With
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Running _
( _
) As Boolean
    ' Abort if no active slide show exists.
    If (Me.Exists = False) Then
        Running = False
        Exit Property
    End If
    
    With mudtPresentation.Doppelganger.SlideShowWindow.View
        If (.State = PowerPoint.ppSlideShowRunning) Then
            Running = True
        Else
            Running = False
        End If
    End With
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Run _
( _
)
    ' Abort if no active slide show exists.
    If (Me.Exists = False) Then
        Exit Sub
    End If
    
    With mudtPresentation.Doppelganger.SlideShowWindow.View
        ' The GotoSlide "kicks" the slide show into running the first time.
        ' Without it, the slide show will not start running the first time.
        .GotoSlide .Slide.SlideIndex, Office.msoFalse
        .State = PowerPoint.ppSlideShowRunning
        .PointerType = PowerPoint.ppSlideShowPointerArrow
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Paused _
( _
) As Boolean
    ' Abort if no active slide show exists.
    If (Me.Exists = False) Then
        Paused = False
        Exit Property
    End If
    
    With mudtPresentation.Doppelganger.SlideShowWindow.View
        If (.State = PowerPoint.ppSlideShowPaused) Then
            Paused = True
        Else
            Paused = False
        End If
    End With
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Pause _
( _
)
    ' Abort if no active slide show exists.
    If (Me.Exists = False) Then
        Exit Sub
    End If
    
    With mudtPresentation.Doppelganger.SlideShowWindow.View
        .State = PowerPoint.ppSlideShowRunning
        .State = PowerPoint.ppSlideShowPaused
        .PointerType = PowerPoint.ppSlideShowPointerArrow
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Move forward in the slide show.
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub EffectNext _
( _
)
    Dim lngIndex As Long
    
    ' Abort if no active slide show exists.
    If (Me.Exists = False) Then
        Exit Sub
    End If
    
    With mudtPresentation.Doppelganger
        ' PowerPoint 2002 effects may not advance correctly
        ' if the slide show is not running.
        ' Therefore, start the slide show running if it is not running.
        If (Me.Running = False) Then
            Me.Run
        End If
    
        .SlideShowWindow.View.Next
        If (.Windows.Count >= 1) Then
            For lngIndex = 1 To .Windows.Count Step 1
                .Windows(lngIndex).View.Slide = .SlideShowWindow.View.Slide
            Next
        End If
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
'   Move backward in the slide show.
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub EffectPrev _
( _
)
    Dim lngIndex As Long
    
    ' Abort if no active slide show exists.
    If (Me.Exists = False) Then
        Exit Sub
    End If
    
    With mudtPresentation.Doppelganger
        ' PowerPoint 2002 effects may not advance correctly
        ' if the slide show is not running.
        ' Therefore, start the slide show running if it is not running.
        If (Me.Running = False) Then
            Me.Run
        End If
    
        .SlideShowWindow.View.Previous
        If (.Windows.Count >= 1) Then
            For lngIndex = 1 To .Windows.Count Step 1
                .Windows(lngIndex).View.Slide = .SlideShowWindow.View.Slide
            Next
        End If
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Initialize _
( _
    ByRef udtApplication As WorshipServiceAssistant.WSAApplication, _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation _
)
    Set mudtApplication = udtApplication
    Set mudtPresentation = udtPresentation
    
    mConfigurationInitialize
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Terminate _
( _
)
    Set mudtPresentation = Nothing
    Set mudtApplication = Nothing
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Doppelganger _
( _
) As PowerPoint.SlideShowWindow
    If (Me.Exists = True) Then
        Set Doppelganger = mudtPresentation.Doppelganger.SlideShowWindow
    Else
        Set Doppelganger = mudtPresentation.Doppelganger.SlideShowWindow
    End If
End Property


'===============================================================================
' Private Properties and Methods.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub mConfigurationInitialize _
( _
)
    Dim blnPresentationSaved As Boolean
    Dim lngIndex As Long
    
    ' Cannot setup a slide show if the presentation has an active slide show.
    If (Me.Exists = True) Then
        Exit Sub
    End If
    
    blnPresentationSaved = mudtPresentation.Doppelganger.Saved

    ' Configure slide show settings.
    With mudtPresentation.Doppelganger.SlideShowSettings
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
    
    mudtPresentation.Doppelganger.Saved = blnPresentationSaved
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub mCreate _
( _
    lngSlideIndex As Long _
)
    Dim blnPresentationSaved As Boolean
    Dim lngHeight As Long
    Dim lngWidth As Long
    Dim lngIndex As Long
    
    With mudtPresentation.Doppelganger
        ' Abort if the slide index is out of range.
        If ((lngSlideIndex < 1) Or (lngSlideIndex > .Slides.Count)) Then
            Exit Sub
        End If
    
        blnPresentationSaved = .Saved
        
        ' Unfortunately, there appears to be no way to start the slide show
        ' with a black screen.  However, the first slide will flash for a
        ' moment.
        .SlideShowSettings.Run.View.GotoSlide lngSlideIndex, Office.msoTrue
        
        .SlideShowWindow.View.State = PowerPoint.ppSlideShowPaused
        .SlideShowWindow.View.PointerType = PowerPoint.ppSlideShowPointerArrow
        
        ' Set the slide show window size.
        lngHeight = .SlideShowWindow.Height
        lngWidth = .SlideShowWindow.Width
        For lngIndex = 1 To mudtApplication.SlideShowWindowSize - 1 Step 1
            lngHeight = lngHeight / 2
            lngWidth = lngWidth / 2
        Next
        .SlideShowWindow.Height = lngHeight
        .SlideShowWindow.Width = lngWidth
            
        mudtApplication.Banner.AdjustSlideShowWindow Me
        
        .Saved = blnPresentationSaved
    End With
End Sub


'===============================================================================
' Private Event Handlers.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub Class_Initialize _
( _
)
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub Class_Terminate _
( _
)
End Sub

