Attribute VB_Name = "Sort"
'===============================================================================
' Name:
'   WorshipServiceAssistant.Sort
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
Public Sub Sort_Run(ByVal W As DocumentWindow)
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
Private Sub SortSlidesByTitle(ByVal P As Presentation, ByVal LowerIndex As Long, ByVal UpperIndex As Long)
    Dim JDelta As Long
    Dim Index
    Dim I As Long
    
    If ((LowerIndex < 1) Or _
        (UpperIndex > P.slides.Count) Or _
        (LowerIndex >= UpperIndex)) Then
        Exit Sub
    End If
    
    For I = LowerIndex + 1 To UpperIndex Step 1
        If (P.slides(I).Shapes.HasTitle = msoTrue) Then
            JDelta = 1
            While (JDelta < (I - LowerIndex + 1))
                JDelta = JDelta * 2
            Wend
            Index = LowerIndex - 1 + JDelta / 2
            While JDelta > 1
                JDelta = JDelta / 2
                If (Index - JDelta >= 1) Then
                    If (P.slides(Index - JDelta).Shapes.HasTitle = msoTrue) Then
                        If (P.slides(I).Shapes.Title.TextFrame.TextRange.Text < _
                            P.slides(Index - JDelta).Shapes.Title.TextFrame.TextRange.Text) Then
                            Index = Index - JDelta
                        End If
                    End If
                End If
                If (Index + JDelta - 1 < I) Then
                    If (P.slides(Index + JDelta - 1).Shapes.HasTitle = msoTrue) Then
                        If (P.slides(I).Shapes.Title.TextFrame.TextRange.Text >= _
                            P.slides(Index + JDelta - 1).Shapes.Title.TextFrame.TextRange.Text) Then
                            Index = Index + JDelta
                        End If
                    End If
                End If
            Wend
            If (Index <> I) Then
                P.slides(I).Copy
                P.slides(I).Delete
                P.slides.Paste Index
            End If
        End If
    Next
End Sub

Private Sub SortSlidesByCategory(ByVal P As Presentation, ByVal LowerIndex As Long, ByVal UpperIndex As Long)
    Dim JDelta As Long
    Dim Index
    Dim I As Long
    
    If ((LowerIndex < 1) Or _
        (UpperIndex > P.slides.Count) Or _
        (LowerIndex >= UpperIndex)) Then
        Exit Sub
    End If
    
    For I = LowerIndex + 1 To UpperIndex Step 1
        JDelta = 1
        While (JDelta < (I - LowerIndex + 1))
            JDelta = JDelta * 2
        Wend
        Index = LowerIndex - 1 + JDelta / 2
        While JDelta > 1
            JDelta = JDelta / 2
            If (Index - JDelta >= 1) Then
                If (P.slides(I).Tags("CategoryIndex") < _
                    P.slides(Index - JDelta).Tags("CategoryIndex")) Then
                    Index = Index - JDelta
                End If
            End If
            If (Index + JDelta - 1 < I) Then
                If (P.slides(I).Tags("CategoryIndex") >= _
                    P.slides(Index + JDelta - 1).Tags("CategoryIndex")) Then
                    Index = Index + JDelta
                End If
            End If
        Wend
        If (Index <> I) Then
            P.slides(I).Copy
            P.slides(I).Delete
            P.slides.Paste Index
        End If
    Next
End Sub
