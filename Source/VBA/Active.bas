Attribute VB_Name = "Active"
'===============================================================================
' Name:
'   WorshipServiceAssistant.Active
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
'   1.03.0001:
'     (1) Fixed a bug that caused ActiveSlideExists and ActiveSlide to
'         incorrectly believe that a slide is active when no slide was selected
'         in the slide window of the outline pane in the normal view.
'         This bug resulted from the addition of the Thumbnails view
'         to the the Normal view in PowerPoint 2002 and some bad programmg.
'     (2) Modified ActiveWindowExists and ActiveSlideExists to work around
'         the PowerPoint 2002 VBA API bug described in Microsoft knowledge base
'         article Q285436.
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


'===============================================================================
' Public Variables.
'===============================================================================


'===============================================================================
' Private Constants.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================


'===============================================================================
' Public Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'   This function returns TRUE if there is an active document window and
'   the active document window is one of the following types: normal, outline,
'   notes, slide, or slide sorter.  Otherwise, it returns FALSE.
'-------------------------------------------------------------------------------
Public Function ActiveWindowExists() As Boolean
    On Error GoTo ActiveWindowExists_False
    
    Dim ActiveWindowViewType As PowerPoint.PpViewType
    
    With Application.ActiveWindow
        ActiveWindowViewType = .ViewType
        If (.ViewType = PowerPoint.ppViewNormal) Then
            Select Case .Panes(2).ViewType
                Case PowerPoint.ppViewSlide
                    ActiveWindowViewType = PowerPoint.ppViewNormal
                Case PowerPoint.ppViewSlideMaster
                    ActiveWindowViewType = PowerPoint.ppViewSlideMaster
            End Select
        End If
    End With
    
    ActiveWindowExists = True
    
    If ((ActiveWindowViewType <> PowerPoint.ppViewSlide) And _
        (ActiveWindowViewType <> PowerPoint.ppViewNormal) And _
        (ActiveWindowViewType <> PowerPoint.ppViewSlideSorter)) Then
        GoTo ActiveWindowExists_False
    End If
Exit Function

ActiveWindowExists_False:
    ActiveWindowExists = False
    Exit Function
End Function

'-------------------------------------------------------------------------------
' Description:
'    This function returns TRUE if the presentation, P, has an associated
'    active slide show window.  Otherwise, it returns FALSE.
'-------------------------------------------------------------------------------
Public Function ActiveSlideShowExists(ByVal P As PowerPoint.Presentation) As Boolean
    On Error GoTo ActiveSlideShowExists_False
    
    ActiveSlideShowExists = True
    
    If (IsNull(P.SlideShowWindow) = True) Then
        GoTo ActiveSlideShowExists_False
    End If
    If (IsEmpty(P.SlideShowWindow) = True) Then
        GoTo ActiveSlideShowExists_False
    End If
Exit Function

ActiveSlideShowExists_False:
    ActiveSlideShowExists = False
End Function

'-------------------------------------------------------------------------------
' Description:
'   This function returns TRUE if the document window, W, has at least
'   one slide.  Otherwise, it returns FALSE.
'-------------------------------------------------------------------------------
Public Function ActiveWindowSlideExists(ByVal W As PowerPoint.DocumentWindow) As Boolean
    ActiveWindowSlideExists = False
    If (W.Presentation.Slides.Count = 0) Then
        Exit Function
    End If
    ActiveWindowSlideExists = True
End Function

'-------------------------------------------------------------------------------
' Description:
'    This function returns TRUE if the document window, W, has a selection.
'    Otherwise, it returns FALSE.
'-------------------------------------------------------------------------------
Public Function ActiveSelectionExists(ByVal W As PowerPoint.DocumentWindow) As Boolean
    ActiveSelectionExists = False
    If (W.Selection.Type = PowerPoint.ppSelectionNone) Then
        Exit Function
    End If
    ActiveSelectionExists = True
End Function

'-------------------------------------------------------------------------------
' Description:
'    This function returns TRUE if the document window, W, has an active
'    slide.  Otherwise, it returns FALSE.
'-------------------------------------------------------------------------------
Public Function ActiveSlideExists(ByVal W As PowerPoint.DocumentWindow) As Boolean
    On Error GoTo ActiveSlideExists_Exit
    
    ActiveSlideExists = False
    If (ActiveWindowExists = True) Then
        If (ActiveSelectionExists(W) = True) Then
            ActiveSlideExists = True
        End If
        If (IsNull(W.View.Slide) = False) Then
            ActiveSlideExists = True
        End If
    End If

ActiveSlideExists_Exit:
End Function

'-------------------------------------------------------------------------------
' Description:
'   This function returns the active slide associated with the document
'   window, W.  This function does not test whether or not the document
'   window has an active slide.  If no active slide exists, then this
'   function will generate an error.
'-------------------------------------------------------------------------------
Public Function ActiveSlide(ByVal W As PowerPoint.DocumentWindow) As PowerPoint.Slide
    If (ActiveSelectionExists(W) = True) Then
        Set ActiveSlide = W.Selection.SlideRange(1)
    Else
        Set ActiveSlide = W.View.Slide
    End If
End Function


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================
