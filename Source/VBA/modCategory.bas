Attribute VB_Name = "modCategory"
'===============================================================================
' Name:
'   WorshipServiceAssistant.modCategory
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
'   1.04.0003:
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
' Private Constants.
'===============================================================================


'===============================================================================
' Public Variables.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================
Private mblnInitialized As Boolean
Private mastrCategoryList() As String


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
Public Function Item _
( _
    ByRef lngIndex As Long _
) As String
    If (mblnInitialized = False) Then
        mInitialize
    End If
    
    Item = ""
    
    If ((lngIndex >= LBound(mastrCategoryList)) And lngIndex <= UBound(mastrCategoryList)) Then
        Item = mastrCategoryList(lngIndex)
    End If
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Function Count _
( _
) As Long
    If (mblnInitialized = False) Then
        mInitialize
    End If
    
    Count = UBound(mastrCategoryList) - LBound(mastrCategoryList) + 1
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub CategorySet _
( _
    ByRef sldRange As PowerPoint.SlideRange, _
    ByRef strCategory As String _
)
    Dim sldSlide As PowerPoint.Slide
    
    If ((sldRange Is Nothing) = False) Then
        If (sldRange.Count >= 1) Then
            For Each sldSlide In sldRange
                If ((strCategory = "") Or (strCategory = "<none>")) Then
                    sldSlide.Tags.Delete "Category"
                Else
                    sldSlide.Tags.Add "Category", strCategory
                End If
            Next
        End If
    End If
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Function CategoryGet _
( _
    ByRef sldRange As PowerPoint.SlideRange _
) As String
    Dim sldSlide As PowerPoint.Slide
    Dim strCategory As String
    
    strCategory = ""
    
    If ((sldRange Is Nothing) = False) Then
        If (sldRange.Count >= 1) Then
            strCategory = sldRange(1).Tags("Category")
            For Each sldSlide In sldRange
                If (sldSlide.Tags("Category") <> strCategory) Then
                        strCategory = ""
                    Exit For
                End If
            Next
        End If
    End If
    
    CategoryGet = strCategory
End Function


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
Public Sub mInitialize _
( _
)
    ReDim mastrCategoryList(1 To 7)
    mastrCategoryList(1) = "<none>"
    mastrCategoryList(2) = "Worship"
    mastrCategoryList(3) = "Choir"
    mastrCategoryList(4) = "Hymn"
    mastrCategoryList(5) = "Carol"
    mastrCategoryList(6) = "Children"
    mastrCategoryList(7) = "Liturgy"
    mblnInitialized = True
End Sub

