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
'   1.00.0001:
'     (1) Modified the Sort_Run routine to collapse the pasted slides in the
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
' Description:
'-------------------------------------------------------------------------------
Public Sub gRun _
( _
    ByRef dwDocumentWindow As PowerPoint.DocumentWindow _
)
    mSlidesSortByTitle _
        prePresentation:=dwDocumentWindow.Presentation, _
        lngLowerIndex:=1, _
        lngUpperIndex:=dwDocumentWindow.Presentation.Slides.Count
        '
        ' Collapse the pasted slide by selecting the slides, and using the
        ' control identifier to find and execute the "Collapse" command.
        '
        dwDocumentWindow.Presentation.Slides.Range.Select
        Application.CommandBars.FindControl(Id:=138).Execute
End Sub


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mSlidesSortByTitle _
( _
    ByRef prePresentation As PowerPoint.Presentation, _
    ByRef lngLowerIndex As Long, _
    ByRef lngUpperIndex As Long _
)
    Dim lngJDelta As Long
    Dim lngIndex As Long
    Dim lngI As Long
    
    If ((lngLowerIndex < 1) Or _
        (lngUpperIndex > prePresentation.Slides.Count) Or _
        (lngLowerIndex >= lngUpperIndex)) Then
        Exit Sub
    End If
    
    For lngI = lngLowerIndex + 1 To lngUpperIndex Step 1
        If (prePresentation.Slides(lngI).Shapes.HasTitle = Office.msoTrue) Then
            lngJDelta = 1
            While (lngJDelta < (lngI - lngLowerIndex + 1))
                lngJDelta = lngJDelta * 2
            Wend
            lngIndex = lngLowerIndex - 1 + lngJDelta / 2
            While lngJDelta > 1
                lngJDelta = lngJDelta / 2
                If (lngIndex - lngJDelta >= 1) Then
                    If (prePresentation.Slides(lngIndex - lngJDelta).Shapes.HasTitle = Office.msoTrue) Then
                        If (prePresentation.Slides(lngI).Shapes.Title.TextFrame.TextRange.Text < _
                            prePresentation.Slides(lngIndex - lngJDelta).Shapes.Title.TextFrame.TextRange.Text) Then
                            lngIndex = lngIndex - lngJDelta
                        End If
                    End If
                End If
                If (lngIndex + lngJDelta - 1 < lngI) Then
                    If (prePresentation.Slides(lngIndex + lngJDelta - 1).Shapes.HasTitle = Office.msoTrue) Then
                        If (prePresentation.Slides(lngI).Shapes.Title.TextFrame.TextRange.Text >= _
                            prePresentation.Slides(lngIndex + lngJDelta - 1).Shapes.Title.TextFrame.TextRange.Text) Then
                            lngIndex = lngIndex + lngJDelta
                        End If
                    End If
                End If
            Wend
            If (lngIndex <> lngI) Then
                prePresentation.Slides(lngI).Copy
                prePresentation.Slides(lngI).Delete
                prePresentation.Slides.Paste lngIndex
            End If
        End If
    Next
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub mSlidesSortByCategory _
( _
    ByRef prePresentation As PowerPoint.Presentation, _
    ByRef lngLowerIndex As Long, _
    ByRef lngUpperIndex As Long _
)
    Dim lngJDelta As Long
    Dim lngIndex
    Dim lngI As Long
    
    If ((lngLowerIndex < 1) Or _
        (lngUpperIndex > prePresentation.Slides.Count) Or _
        (lngLowerIndex >= lngUpperIndex)) Then
        Exit Sub
    End If
    
    For lngI = lngLowerIndex + 1 To lngUpperIndex Step 1
        lngJDelta = 1
        While (lngJDelta < (lngI - lngLowerIndex + 1))
            lngJDelta = lngJDelta * 2
        Wend
        lngIndex = lngLowerIndex - 1 + lngJDelta / 2
        While lngJDelta > 1
            lngJDelta = lngJDelta / 2
            If (lngIndex - lngJDelta >= 1) Then
                If (prePresentation.Slides(lngI).Tags("CategoryIndex") < _
                    prePresentation.Slides(lngIndex - lngJDelta).Tags("CategoryIndex")) Then
                    lngIndex = lngIndex - lngJDelta
                End If
            End If
            If (lngIndex + lngJDelta - 1 < lngI) Then
                If (prePresentation.Slides(lngI).Tags("CategoryIndex") >= _
                    prePresentation.Slides(lngIndex + lngJDelta - 1).Tags("CategoryIndex")) Then
                    lngIndex = lngIndex + lngJDelta
                End If
            End If
        Wend
        If (lngIndex <> lngI) Then
            prePresentation.Slides(lngI).Copy
            prePresentation.Slides(lngI).Delete
            prePresentation.Slides.Paste lngIndex
        End If
    Next
End Sub
