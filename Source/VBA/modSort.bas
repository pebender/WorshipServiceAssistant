Attribute VB_Name = "modSort"
'===============================================================================
' Name:
'   WorshipServiceAssistant.modSort
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
'   1.04.0000:
'     (1) Deleted the unused mSortSlidesByCategory routine.
'     (2) Modified mSortSlidesByCategory routine to use the Slide object
'         MoveTo method added in PowerPoint 2002.
'     (3) Modified the gRun routine to remove the code that was added
'         in 1.00.0001 to collapse the pasted slides, since the use of the
'         MoveTo method made it unnecessary.
'     (4) Cleaned up the mSortSlidesByTitle routine.
'   1.03.0002:
'     (1) Made changes to the source code so that it follows Microsoft's
'         Visual Basic coding conventions.
'   1.00.0001:
'     (1) Modified the gRun routine to collapse the pasted slides in the
'         outline pane.
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
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub gRun _
( _
    ByRef dwCurrent As PowerPoint.DocumentWindow _
)
    If ((dwCurrent Is Nothing) = False) Then
        mSlidesSortByTitle _
            prePresentation:=dwCurrent.Presentation, _
            lngLowerIndex:=1, _
            lngUpperIndex:=dwCurrent.Presentation.Slides.Count
    End If
End Sub


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub mSlidesSortByTitle _
( _
    ByRef prePresentation As PowerPoint.Presentation, _
    ByRef lngLowerIndex As Long, _
    ByRef lngUpperIndex As Long _
)
    Dim strTitle As String
    Dim lngDelta As Long
    Dim lngMatch As Long
    Dim lngIndex As Long
    
    With prePresentation.Slides
        If ((lngLowerIndex < 1) Or _
            (lngUpperIndex > .Count) Or _
            (lngLowerIndex >= lngUpperIndex)) Then
            Exit Sub
        End If
        
        For lngIndex = lngLowerIndex + 1 To lngUpperIndex Step 1
            strTitle = mstrTitle(.Item(lngIndex))
            lngDelta = 1
            While (lngDelta < (lngIndex - lngLowerIndex + 1))
                lngDelta = lngDelta * 2
            Wend
            lngMatch = lngLowerIndex - 1 + lngDelta / 2
            While lngDelta > 1
                lngDelta = lngDelta / 2
                If (lngMatch - lngDelta >= 1) Then
                    If (strTitle < mstrTitle(.Item(lngMatch - lngDelta))) Then
                        lngMatch = lngMatch - lngDelta
                    End If
                End If
                If (lngMatch + lngDelta - 1 < lngIndex) Then
                    If (strTitle >= mstrTitle(.Item(lngMatch + lngDelta - 1))) Then
                        lngMatch = lngMatch + lngDelta
                    End If
                End If
            Wend
            If (lngIndex <> lngMatch) Then
                .Item(lngIndex).MoveTo lngMatch
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
Private Function mstrTitle _
( _
    ByRef sldSlide As PowerPoint.Slide _
) As String
    mstrTitle = ""
    With sldSlide.Shapes
        If (.HasTitle = Office.msoTrue) Then
            With .Title
                If (.HasTextFrame = Office.msoTrue) Then
                    With .TextFrame
                        If (.HasText = Office.msoTrue) Then
                            mstrTitle = .TextRange.Text
                        End If
                    End With
                End If
            End With
        End If
    End With
End Function

