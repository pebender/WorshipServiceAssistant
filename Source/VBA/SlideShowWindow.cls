VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SlideShowWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.SlideShowWindow
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
' Public Constants.
'===============================================================================


'===============================================================================
' Private Constants.
'===============================================================================
Private Const mlngPpControlIdSlideShowSetUpShow As Long = 2744


'===============================================================================
' Public Variables.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================
Private mblnInvariant As Boolean
Private mudtApplication As WorshipServiceAssistant.Application
Private mudtParent As WorshipServiceAssistant.Presentation
Private mprePresentation As PowerPoint.Presentation


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
Public Sub Initialize _
( _
    ByRef udtApplication As WorshipServiceAssistant.Application, _
    ByRef udtParent As WorshipServiceAssistant.Presentation _
)
    mblnInvariant = False
    
    Set mudtApplication = udtApplication
    Set mudtParent = udtParent
    Set mprePresentation = mudtParent.Doppelganger
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
    Set mprePresentation = Nothing
    Set mudtParent = Nothing
    Set mudtApplication = Nothing
    
    mblnInvariant = False
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Invariant _
( _
) As Boolean
    Invariant = mblnInvariant
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Let Invariant _
( _
    ByRef blnInvariant As Boolean _
)
    If (Not blnInvariant = mblnInvariant) Then
        If (blnInvariant = True) Then
            mblnInvariant = blnInvariant
            mConfigurationInitialize
        End If
        If (blnInvariant = False) Then
            mblnInvariant = blnInvariant
        End If
    End If
End Property

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
    Set Doppelganger = Nothing
    On Error Resume Next
    Set Doppelganger = mprePresentation.SlideShowWindow
    On Error GoTo 0
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Application _
( _
) As WorshipServiceAssistant.Application
    Set Application = mudtApplication
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Parent _
( _
) As WorshipServiceAssistant.Presentation
    Set Parent = mudtParent
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Presentation _
( _
) As WorshipServiceAssistant.Presentation
    Set Presentation = mudtParent
End Property

'-------------------------------------------------------------------------------
' Purpose:
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
    If (Me.Doppelganger Is Nothing) Then mCreate lngSlideIndex
    
    With mprePresentation
        ' Abort if the slide index is out of range.
        If ((lngSlideIndex < 1) Or (lngSlideIndex > .Slides.Count)) Then Exit Sub
        
        .SlideShowWindow.View.GotoSlide lngSlideIndex, Office.msoFalse
        .SlideShowWindow.Activate
    End With
    
    ' Black other slide shows.
    ' This will stop the slide shows from running.
    With mudtApplication.Presentations
        For lngIndex = 1 To .Count Step 1
            If ((.Item(lngIndex).Doppelganger Is mprePresentation) = False) Then
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
    Visible = False
    
    ' Abort if no active slide show exists.
    If (Me.Doppelganger Is Nothing) Then Exit Property
    
    With mprePresentation.SlideShowWindow.View
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
    If (Me.Doppelganger Is Nothing) Then Exit Property
    
    With mprePresentation.SlideShowWindow.View
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
    Running = False
    
    ' Abort if no active slide show exists.
    If (Me.Doppelganger Is Nothing) Then Exit Property
    
    With mprePresentation.SlideShowWindow.View
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
    If (Me.Doppelganger Is Nothing) Then Exit Sub
    
    With mprePresentation.SlideShowWindow.View
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
    Paused = False
    
    ' Abort if no active slide show exists.
    If (Me.Doppelganger Is Nothing) Then Exit Property
    
    With mprePresentation.SlideShowWindow.View
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
    If (Me.Doppelganger Is Nothing) Then Exit Sub
    
    With mprePresentation.SlideShowWindow.View
        .State = PowerPoint.ppSlideShowRunning
        .State = PowerPoint.ppSlideShowPaused
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
Public Sub EffectNext _
( _
)
    Dim lngIndex As Long
    
    ' Abort if no active slide show exists.
    If (Me.Doppelganger Is Nothing) Then Exit Sub
    
    With mprePresentation
        ' PowerPoint 10.0 effects may not advance correctly
        ' if the slide show is not running.
        ' Therefore, start the slide show running if it is not running.
        If (Not Me.Running) Then Me.Run
    
        .SlideShowWindow.View.Next
        For lngIndex = 1 To .Windows.Count Step 1
            .Windows.Item(lngIndex).View.Slide = .SlideShowWindow.View.Slide
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
Public Sub EffectPrev _
( _
)
    Dim lngIndex As Long
    
    ' Abort if no active slide show exists.
    If (Me.Doppelganger Is Nothing) Then Exit Sub
    
    With mprePresentation
        ' PowerPoint 10.0 effects may not advance correctly
        ' if the slide show is not running.
        ' Therefore, start the slide show running if it is not running.
        If (Not Me.Running) Then Me.Run
    
        .SlideShowWindow.View.previous
        For lngIndex = 1 To .Windows.Count Step 1
            .Windows.Item(lngIndex).View.Slide = .SlideShowWindow.View.Slide
        Next
    End With
End Sub


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
    
    ' Cannot setup a slide show if the presentation has an active slide show.
    If (Not Me.Doppelganger Is Nothing) Then Exit Sub
    
    blnPresentationSaved = mprePresentation.Saved

    ' Configure slide show settings.
    With mprePresentation.SlideShowSettings
        .ShowType = PowerPoint.ppShowTypeSpeaker
        .RangeType = PowerPoint.ppShowAll
        .AdvanceMode = PowerPoint.ppSlideShowUseSlideTimings
        .LoopUntilStopped = Office.msoTrue
        .ShowWithAnimation = Office.msoTrue
        .ShowWithNarration = Office.msoTrue
    End With
    
    mprePresentation.Saved = blnPresentationSaved
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
    
    With mprePresentation
        ' Abort if the slide index is out of range.
        If ((lngSlideIndex < 1) Or (lngSlideIndex > .Slides.Count)) Then
            Exit Sub
        End If
    
        blnPresentationSaved = .Saved
        
        With .SlideShowSettings
            ' Unfortunately, there appears to be no way to start the slide show
            ' with a black screen or to start a slide show on other than the
            ' first slide. Therefore, the first slide will flash for a moment.
            .Run.View.GotoSlide lngSlideIndex, Office.msoTrue
        End With
        
        With .SlideShowWindow
            .View.State = PowerPoint.ppSlideShowPaused
            .View.PointerType = PowerPoint.ppSlideShowPointerArrow
            
            ' Set the slide show window size.
            .Height = .Height / mudtApplication.SlideShowWindowSize
            .Width = .Width / mudtApplication.SlideShowWindowSize
        End With
        
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

