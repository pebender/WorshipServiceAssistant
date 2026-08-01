VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Slides"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.Slides
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
Private msldDoppelganger As PowerPoint.Slides
Private mcudtSlides As New VBA.Collection


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
    ByRef sldDoppelganger As PowerPoint.Slides _
)
    mblnInvariant = False
    
    Set mudtApplication = udtApplication
    Set mudtParent = udtParent
    Set msldDoppelganger = sldDoppelganger
    
    With mcudtSlides
        Do While (.Count > 0)
            .Remove .Count
        Loop
    End With
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
    With mcudtSlides
        Do While (.Count > 0)
            .Item(.Count).Terminate
            .Remove .Count
        Loop
    End With
    
    Set msldDoppelganger = Nothing
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
    Dim lng As Long
    
    If (Not blnInvariant = mblnInvariant) Then
        If (blnInvariant = True) Then
            Set mcudtSlides = New VBA.Collection
            With msldDoppelganger
                For lng = 1 To .Count Step 1
                    mcudtSlides.Add New WorshipServiceAssistant.Slide
                    mcudtSlides.Item(lng).Initialize mudtApplication, mudtParent, msldDoppelganger.Item(lng)
                Next
            End With
            mblnInvariant = blnInvariant
        End If
        With mcudtSlides
            For lng = 1 To .Count Step 1
                .Item(lng).Invariant = blnInvariant
            Next
        End With
        If (blnInvariant = False) Then
            mblnInvariant = blnInvariant
            Set mcudtSlides = Nothing
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
) As PowerPoint.Slides
    Set Doppelganger = msldDoppelganger
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
Public Property Get Count _
( _
)
    If (mblnInvariant) Then
        Count = mcudtSlides.Count
    Else
        Count = msldDoppelganger.Count
    End If
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Function Item _
( _
    ByRef lngIndex As Long _
) As WorshipServiceAssistant.Slide
Attribute Item.VB_UserMemId = 0
    If (mblnInvariant) Then
        Set Item = mcudtSlides.Item(lngIndex)
    Else
        Set Item = New WorshipServiceAssistant.Slide
        Item.Initialize mudtApplication, mudtParent, msldDoppelganger.Item(lngIndex)
    End If
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Function TitleMatches _
( _
    ByRef lngIndexes() As Long, _
    ByRef strMatch As String _
) As Long
    Dim lngIndex As Long
    Dim lngCount As Long
    Dim strTitleMatch As String
    Dim lngTitleMatchLen As Long
    
    ReDim lngIndexes(0 To Me.Count - 1)
    lngCount = 0
    With Me
        If (.Count > 0) Then
            strTitleMatch = .Item(1).StringForTitleMatch(strMatch)
            lngTitleMatchLen = VBA.Len(strTitleMatch)
            For lngIndex = 1 To .Count Step 1
                If (.Item(lngIndex).TitleMatches(strTitleMatch, lngTitleMatchLen) = True) Then
                    lngIndexes(lngCount) = lngIndex
                    lngCount = lngCount + 1
                End If
            Next
        End If
    End With
    TitleMatches = lngCount
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Function TitleOrBodyMatches _
( _
    ByRef lngIndexes() As Long, _
    ByRef strMatch As String _
) As Long
    Dim lngIndex As Long
    Dim lngCount As Long
    Dim strTitleOrBodyMatch As String
    Dim lngTitleOrBodyMatchLen As Long
    
    ReDim lngIndexes(0 To Me.Count - 1)
    lngCount = 0
    With Me
        If (.Count > 0) Then
            strTitleOrBodyMatch = .Item(1).StringForTitleMatch(strMatch)
            lngTitleOrBodyMatchLen = VBA.Len(strTitleOrBodyMatch)
            
            For lngIndex = 1 To .Count Step 1
                If (.Item(lngIndex).TitleOrBodyMatches(strTitleOrBodyMatch, lngTitleOrBodyMatchLen) = True) Then
                    lngIndexes(lngCount) = lngIndex
                    lngCount = lngCount + 1
                End If
            Next
        End If
    End With
    TitleOrBodyMatches = lngCount
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub SortByTitle _
( _
)
    Dim strTitle As String
    Dim lngDelta As Long
    Dim lngMatch As Long
    Dim lngIndex As Long
    Dim udtSlide As WorshipServiceAssistant.Slide
    
    With Me
        For lngIndex = 2 To .Count Step 1
            strTitle = .Item(lngIndex).Title
            lngDelta = 1
            Do While (lngDelta < lngIndex)
                lngDelta = lngDelta * 2
            Loop
            lngMatch = lngDelta / 2
            Do While (lngDelta > 1)
                lngDelta = lngDelta / 2
                If (lngMatch - lngDelta > 0) Then
                    If (strTitle < .Item(lngMatch - lngDelta).Title) Then
                        lngMatch = lngMatch - lngDelta
                    End If
                End If
                If (lngMatch + lngDelta - 1 < lngIndex) Then
                    If (strTitle >= .Item(lngMatch + lngDelta - 1).Title) Then
                        lngMatch = lngMatch + lngDelta
                    End If
                End If
            Loop
            If (Not lngIndex = lngMatch) Then
                If (mblnInvariant) Then
                    With mcudtSlides
                        .Remove lngIndex
                        If (lngMatch < lngIndex) Then
                            .Add Item:=udtSlide, Before:=lngMatch
                        Else
                            .Add Item:=udtSlide, Before:=lngMatch - 1
                        End If
                    End With
                End If
                msldDoppelganger.Item(lngIndex).MoveTo lngMatch
            End If
        Next
    End With
End Sub


'===============================================================================
' Private Properties and Methods.
'===============================================================================


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
