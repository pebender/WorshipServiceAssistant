VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.DocumentWindow
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


'===============================================================================
' Public Variables.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================
Private mblnInvariant As Boolean
Private mudtApplication As WorshipServiceAssistant.Application
Private mudtParent As WorshipServiceAssistant.Presentation
Private mdwDoppelganger As PowerPoint.DocumentWindow
Private mlngWindowState As PowerPoint.PpWindowState
Private mlngViewType As PowerPoint.ppViewType
Private mlngView_DisplaySlideMiniature As Office.MsoTriState
Private mlngView_ZoomToFit As Office.MsoTriState


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
    ByRef udtParent As WorshipServiceAssistant.Presentation, _
    ByRef dwDoppelganger As PowerPoint.DocumentWindow _
)
    mblnInvariant = False
    
    Set mudtApplication = udtApplication
    Set mudtParent = udtParent
    Set mdwDoppelganger = dwDoppelganger
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
    Set mdwDoppelganger = Nothing
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
            mConfigurationSave
            mConfigurationInitialize
            mblnInvariant = blnInvariant
        End If
        If (blnInvariant = False) Then
            mblnInvariant = blnInvariant
           mConfigurationRestore
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
) As PowerPoint.DocumentWindow
    Set Doppelganger = mdwDoppelganger
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
'   This property returns the active slide, if there is an active slide.
'   Otherwise it returns Nothing.
' Assumptions:
' Effects:
' Returns:
'-------------------------------------------------------------------------------
Public Function ActiveSlide _
( _
) As WorshipServiceAssistant.Slide
    Dim udtSlideSelection As WorshipServiceAssistant.SlideRange
    
    Set udtSlideSelection = Me.ActiveSlideSelection
    
    If (udtSlideSelection.Count = 0) Then
        Set ActiveSlide = Nothing
    Else
        Set ActiveSlide = udtSlideSelection.Item(1)
    End If
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Function ActiveSlideSelection _
( _
) As WorshipServiceAssistant.SlideRange
    Dim sldSlideSelection As PowerPoint.SlideRange
    Dim lngViewTypeCurrent As PowerPoint.ppViewType
    Dim lngIndex As Long
    
    Set ActiveSlideSelection = New WorshipServiceAssistant.SlideRange
    
    With mdwDoppelganger
        With .Panes
            If (.Count = 1) Then
                lngViewTypeCurrent = .Item(1).ViewType
            Else
                lngViewTypeCurrent = .Item(2).ViewType
           End If
        End With
        
        Set sldSlideSelection = Nothing
        If ((lngViewTypeCurrent = PowerPoint.ppViewSlide) Or _
            (lngViewTypeCurrent = PowerPoint.ppViewSlideSorter)) Then
            Select Case .Selection.Type
                Case ppSelectionNone
                    On Error Resume Next
                    Set sldSlideSelection = .Parent.Slides.Range(.View.Slide.SlideIndex)
                    On Error GoTo 0
                Case ppSelectionSlides
                    Set sldSlideSelection = .Selection.SlideRange
                Case ppSelectionShapes
                    Set sldSlideSelection = .Selection.SlideRange
                Case ppSelectionText
                    Set sldSlideSelection = .Selection.SlideRange
            End Select
        End If
    End With
    
    If (Not sldSlideSelection Is Nothing) Then
        ActiveSlideSelection.Initialize mudtApplication, Me, sldSlideSelection
    End If
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Active _
( _
) As Office.MsoTriState
    Active = mdwDoppelganger.Active
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Activate _
( _
)
    mdwDoppelganger.Activate
End Sub


'===============================================================================
' Private Properties and Methods.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub mConfigurationSave _
( _
)
    With mdwDoppelganger
        mlngWindowState = .WindowState
        mlngViewType = .ViewType
        If (mlngViewType = PowerPoint.ppViewNormal) Then
            Select Case .Panes(2).ViewType
                Case PowerPoint.ppViewSlide
                    mlngViewType = PowerPoint.ppViewNormal
                Case PowerPoint.ppViewSlideMaster
                    mlngViewType = PowerPoint.ppViewSlideMaster
            End Select
        End If
        mlngView_DisplaySlideMiniature = .View.DisplaySlideMiniature
        mlngView_ZoomToFit = .View.ZoomToFit
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub mConfigurationRestore _
( _
)
    With mdwDoppelganger
        .WindowState = mlngWindowState
        .ViewType = mlngViewType
        .View.DisplaySlideMiniature = mlngView_DisplaySlideMiniature
        .View.ZoomToFit = mlngView_ZoomToFit
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
'-------------------------------------------------------------------------------
Private Sub mConfigurationInitialize _
( _
)
    With mdwDoppelganger
        .WindowState = PowerPoint.ppWindowMaximized
        .ViewType = PowerPoint.ppViewSlide
        .View.DisplaySlideMiniature = Office.msoFalse
        .View.ZoomToFit = Office.msoTrue
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

