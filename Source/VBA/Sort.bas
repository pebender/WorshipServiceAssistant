Attribute VB_Name = "Sort"
'===============================================================================
' Name:
'   WorshipServiceAssistant.Sort
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
Public Sub Sort_Run(ByVal W As PowerPoint.DocumentWindow)
    SortSlidesByTitle _
        P:=W.Presentation, _
        LowerIndex:=1, _
        UpperIndex:=W.Presentation.slides.Count
        '
        ' Collapse the pasted slide by selecting the slides, and using the
        ' control identifier to find and execute the "Collapse" command.
        '
        W.Presentation.slides.Range.Select
        Application.CommandBars.FindControl(Id:=138).Execute
End Sub


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Sub SortSlidesByTitle(ByVal P As PowerPoint.Presentation, ByVal LowerIndex As Long, ByVal UpperIndex As Long)
    Dim JDelta As Long
    Dim Index
    Dim i As Long
    
    If ((LowerIndex < 1) Or _
        (UpperIndex > P.slides.Count) Or _
        (LowerIndex >= UpperIndex)) Then
        Exit Sub
    End If
    
    For i = LowerIndex + 1 To UpperIndex Step 1
        If (P.slides(i).Shapes.HasTitle = msoTrue) Then
            JDelta = 1
            While (JDelta < (i - LowerIndex + 1))
                JDelta = JDelta * 2
            Wend
            Index = LowerIndex - 1 + JDelta / 2
            While JDelta > 1
                JDelta = JDelta / 2
                If (Index - JDelta >= 1) Then
                    If (P.slides(Index - JDelta).Shapes.HasTitle = msoTrue) Then
                        If (P.slides(i).Shapes.Title.TextFrame.TextRange.Text < _
                            P.slides(Index - JDelta).Shapes.Title.TextFrame.TextRange.Text) Then
                            Index = Index - JDelta
                        End If
                    End If
                End If
                If (Index + JDelta - 1 < i) Then
                    If (P.slides(Index + JDelta - 1).Shapes.HasTitle = msoTrue) Then
                        If (P.slides(i).Shapes.Title.TextFrame.TextRange.Text >= _
                            P.slides(Index + JDelta - 1).Shapes.Title.TextFrame.TextRange.Text) Then
                            Index = Index + JDelta
                        End If
                    End If
                End If
            Wend
            If (Index <> i) Then
                P.slides(i).Copy
                P.slides(i).Delete
                P.slides.Paste Index
            End If
        End If
    Next
End Sub

Private Sub SortSlidesByCategory(ByVal P As PowerPoint.Presentation, ByVal LowerIndex As Long, ByVal UpperIndex As Long)
    Dim JDelta As Long
    Dim Index
    Dim i As Long
    
    If ((LowerIndex < 1) Or _
        (UpperIndex > P.slides.Count) Or _
        (LowerIndex >= UpperIndex)) Then
        Exit Sub
    End If
    
    For i = LowerIndex + 1 To UpperIndex Step 1
        JDelta = 1
        While (JDelta < (i - LowerIndex + 1))
            JDelta = JDelta * 2
        Wend
        Index = LowerIndex - 1 + JDelta / 2
        While JDelta > 1
            JDelta = JDelta / 2
            If (Index - JDelta >= 1) Then
                If (P.slides(i).Tags("CategoryIndex") < _
                    P.slides(Index - JDelta).Tags("CategoryIndex")) Then
                    Index = Index - JDelta
                End If
            End If
            If (Index + JDelta - 1 < i) Then
                If (P.slides(i).Tags("CategoryIndex") >= _
                    P.slides(Index + JDelta - 1).Tags("CategoryIndex")) Then
                    Index = Index + JDelta
                End If
            End If
        Wend
        If (Index <> i) Then
            P.slides(i).Copy
            P.slides(i).Delete
            P.slides.Paste Index
        End If
    Next
End Sub
