Attribute VB_Name = "modActive"
'===============================================================================
' Name:
'   WorshipServiceAssistant.modActive
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
'   1.03.0002:
'     (1) Made changes to the source code so that it follows Microsoft's
'         Visual Basic coding conventions.
'   1.03.0001:
'     (1) Fixed a bug that caused gblnActiveSlideExists and ActiveSlide to
'         incorrectly believe that a slide is active when no slide was selected
'         in the slide window of the outline pane in the normal view.
'         This bug resulted from the addition of the Thumbnails view
'         to the the Normal view in PowerPoint 2002 and some bad programmg.
'     (2) Modified gblnActiveWindowExists and gblnActiveSlideExists to work around
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
Public Function gblnActiveWindowExists _
( _
) As Boolean
    Dim ppActiveWindowViewType As PowerPoint.PpViewType
    
    On Error GoTo gblnActiveWindowExists_False
    
    With Application.ActiveWindow
        ppActiveWindowViewType = .ViewType
        If (.ViewType = PowerPoint.ppViewNormal) Then
            Select Case .Panes(2).ViewType
                Case PowerPoint.ppViewSlide
                    ppActiveWindowViewType = PowerPoint.ppViewNormal
                Case PowerPoint.ppViewSlideMaster
                    ppActiveWindowViewType = PowerPoint.ppViewSlideMaster
            End Select
        End If
    End With
    
    gblnActiveWindowExists = True
    
    If ((ppActiveWindowViewType <> PowerPoint.ppViewSlide) And _
        (ppActiveWindowViewType <> PowerPoint.ppViewNormal) And _
        (ppActiveWindowViewType <> PowerPoint.ppViewSlideSorter)) Then
        GoTo gblnActiveWindowExists_False
    End If
Exit Function

gblnActiveWindowExists_False:
    gblnActiveWindowExists = False
    Exit Function
End Function

'-------------------------------------------------------------------------------
' Description:
'    This function returns TRUE if the presentation, prePresentation, has an
'    associated active slide show window.  Otherwise, it returns FALSE.
'-------------------------------------------------------------------------------
Public Function gblnActiveSlideShowExists _
( _
    ByRef prePresentation As PowerPoint.Presentation _
) As Boolean
    On Error GoTo gblnActiveSlideShowExists_False
    
    gblnActiveSlideShowExists = True
    
    If (VBA.IsNull(prePresentation.SlideShowWindow) = True) Then
        GoTo gblnActiveSlideShowExists_False
    End If
    If (VBA.IsEmpty(prePresentation.SlideShowWindow) = True) Then
        GoTo gblnActiveSlideShowExists_False
    End If
Exit Function

gblnActiveSlideShowExists_False:
    gblnActiveSlideShowExists = False
End Function

'-------------------------------------------------------------------------------
' Description:
'   This function returns TRUE if the document window, dwDocumentWindow, has at
'   least one slide.  Otherwise, it returns FALSE.
'-------------------------------------------------------------------------------
Public Function gblnActiveWindowSlideExists _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
) As Boolean
    gblnActiveWindowSlideExists = False
    If (dwDocumentWindow.Presentation.Slides.Count = 0) Then
        Exit Function
    End If
    gblnActiveWindowSlideExists = True
End Function

'-------------------------------------------------------------------------------
' Description:
'    This function returns TRUE if the document window, dwDocumentWindow, has a
'     selection.  Otherwise, it returns FALSE.
'-------------------------------------------------------------------------------
Public Function gblnActiveSelectionExists _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
) As Boolean
    gblnActiveSelectionExists = False
    If (dwDocumentWindow.Selection.Type = PowerPoint.ppSelectionNone) Then
        Exit Function
    End If
    gblnActiveSelectionExists = True
End Function

'-------------------------------------------------------------------------------
' Description:
'    This function returns TRUE if the document window, dwDocumentWindow, has an
'    active slide.  Otherwise, it returns FALSE.
'-------------------------------------------------------------------------------
Public Function gblnActiveSlideExists _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
) As Boolean
    On Error GoTo gblnActiveSlideExists_Exit
    
    gblnActiveSlideExists = False
    If (gblnActiveWindowExists = True) Then
        If (modActive.gblnActiveSelectionExists(dwDocumentWindow) = True) Then
            gblnActiveSlideExists = True
        End If
        If (VBA.IsNull(dwDocumentWindow.View.Slide) = False) Then
            gblnActiveSlideExists = True
        End If
    End If

gblnActiveSlideExists_Exit:
End Function

'-------------------------------------------------------------------------------
' Description:
'   This function returns the active slide associated with the document
'   window, dwDocumentWindow.  This function does not test whether or not the
'   documentwindow has an active slide.  If no active slide exists, then this
'   function will generate an error.
'-------------------------------------------------------------------------------
Public Function gppActiveSlideGet _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
) As PowerPoint.Slide
    If (modActive.gblnActiveSelectionExists(dwDocumentWindow) = True) Then
        Set gppActiveSlideGet = dwDocumentWindow.Selection.SlideRange(1)
    Else
        Set gppActiveSlideGet = dwDocumentWindow.View.Slide
    End If
End Function


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================
