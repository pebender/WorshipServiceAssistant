VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WSASlide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.WSASlide
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
'   1.04.0000:
'     Initial revision.
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
' Public Variables.
'===============================================================================


'===============================================================================
' Private Constants.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================
Private mudtApplication As WorshipServiceAssistant.WSAApplication
Private mudtPresentation As WorshipServiceAssistant.WSAPresentation
Private msldDoppelganger As PowerPoint.Slide

Private mlngIndex As Long
Private mstrTitle As String
Private mstrTitleMatch1 As String
Private mstrTitleMatch2 As String
Private mstrTitleOrBodyMatch1 As String
Private mstrTitleOrBodyMatch2 As String


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
Public Property Get Index _
( _
) As Long
    Index = mlngIndex
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Title _
( _
) As String
    Title = mstrTitle
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Category _
( _
) As String
    Category = msldDoppelganger.Tags("Category")
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Let Category _
( _
    ByRef strCategory As String _
)
    msldDoppelganger.Tags.Add "Category", strCategory
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
    Dim lngIndex As Long

    With msldDoppelganger.Parent.Windows
        If (.Count >= 1) Then
            For lngIndex = 1 To .Count Step 1
                .Item(lngIndex).View.Slide = msldDoppelganger
            Next
        End If
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Function TitleMatches _
( _
    ByRef strMatch As String, _
    ByRef lngMatchLen As Long _
) As Boolean
    If ((VBA.Left(mstrTitleMatch1, lngMatchLen) = strMatch) Or _
        (VBA.Left(mstrTitleMatch2, lngMatchLen) = strMatch)) Then
        TitleMatches = True
    Else
        TitleMatches = False
    End If
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Function StringForTitleMatch _
( _
    ByRef str As String _
) As String
    StringForTitleMatch = mstrEverythingClean(str)
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
    ByRef strMatch As String, _
    ByRef lngMatchLen As Long _
) As Boolean
    If ((InStr(1, mstrTitleOrBodyMatch1, strMatch, VBA.vbTextCompare) > 0) Or _
        (InStr(1, mstrTitleOrBodyMatch2, strMatch, VBA.vbTextCompare) > 0)) Then
        TitleOrBodyMatches = True
    Else
        TitleOrBodyMatches = False
    End If
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Function StringForTitleOrBodyMatch _
( _
    ByRef str As String _
) As String
    StringForTitleOrBodyMatch = " " & mstrEverythingClean(str)
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Initialize _
( _
    ByRef udtApplication As WorshipServiceAssistant.WSAApplication, _
    ByRef udtPresentation As WorshipServiceAssistant.WSAPresentation, _
    ByRef sldDoppelganger As PowerPoint.Slide _
)
    Dim strTitle As String
    Dim strTitleOrBody As String
    
    Set mudtApplication = udtApplication
    Set mudtPresentation = udtPresentation
    Set msldDoppelganger = sldDoppelganger
        
    With msldDoppelganger
        mlngIndex = .SlideIndex
        
        strTitle = mstrTitleGet
        mstrTitle = mstrWhiteSpaceClean(strTitle)
        mstrTitleMatch1 = mstrWhiteSpaceClean(strTitle)
        mstrTitleMatch2 = mstrEverythingClean(strTitle)

        strTitleOrBody = mstrTitleOrBodyGet
        mstrTitleOrBodyMatch1 = mstrWhiteSpaceClean(strTitleOrBody)
        mstrTitleOrBodyMatch2 = mstrEverythingClean(strTitleOrBody)
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
    mstrTitleOrBodyMatch2 = ""
    mstrTitleOrBodyMatch1 = ""
    mstrTitleMatch2 = ""
    mstrTitleMatch1 = ""
    mstrTitle = ""
    mlngIndex = 0
    
    Set msldDoppelganger = Nothing
    Set mudtPresentation = Nothing
    Set mudtApplication = Nothing
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Doppelganger _
( _
) As PowerPoint.Slide
    Set Doppelganger = msldDoppelganger
End Property


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
Private Function mstrTitleGet _
( _
) As String
    mstrTitleGet = ""
    With msldDoppelganger
        If (.Shapes.HasTitle = Office.msoTrue) Then
            If (.Shapes.Title.HasTextFrame = Office.msoTrue) Then
                If (.Shapes.Title.TextFrame.HasText = Office.msoTrue) Then
                    mstrTitleGet = .Shapes.Title.TextFrame.TextRange.Text
                End If
            End If
        End If
    End With
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mstrTitleOrBodyGet _
( _
) As String
    Dim lngIndex As Long
    
    mstrTitleOrBodyGet = VBA.vbCrLf & mstrTitleGet
    With msldDoppelganger.Shapes
        For lngIndex = 1 To .Count Step 1
            With .Item(lngIndex)
                If (.HasTextFrame = Office.msoTrue) Then
                    If (.TextFrame.HasText = Office.msoTrue) Then
                        mstrTitleOrBodyGet = _
                            mstrTitleOrBodyGet & _
                            VBA.vbCrLf & _
                            .TextFrame.TextRange.Text
                    End If
                End If
            End With
        Next
    End With
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mstrWhiteSpaceClean _
( _
    ByRef strDirty As String _
) As String
    Dim strClean As String
    
    strClean = strDirty
    
    '
    ' Fix white space.
    '
    strClean = VBA.Replace(strClean, VBA.Chr(9), " ")        ' replace tab with space
    strClean = VBA.Replace(strClean, VBA.Chr(11), " ")       ' replace line feed with space
    strClean = VBA.Replace(strClean, VBA.Chr(13), " ")       ' replace return with space
    '
    ' Fix quotes.
    '
    strClean = VBA.Replace(strClean, "`", VBA.Chr(39))       ' replace back single quote with single quote
    strClean = VBA.Replace(strClean, VBA.Chr(145), VBA.Chr(39))  ' replace open single quote with single quote
    strClean = VBA.Replace(strClean, VBA.Chr(146), VBA.Chr(39))  ' replace close single quote with single quote
    strClean = VBA.Replace(strClean, VBA.Chr(147), VBA.Chr(34))  ' replace open double quote with double quote
    strClean = VBA.Replace(strClean, VBA.Chr(148), VBA.Chr(34))  ' replace close double quote with double quote

    While (VBA.InStr(strClean, "  ") > 0)
        strClean = VBA.Replace(strClean, "  ", " ")
    Wend
    
    mstrWhiteSpaceClean = strClean
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mstrEverythingClean _
( _
    ByRef strDirty As String _
) As String
    Dim strClean As String
    
    strClean = VBA.LCase(strDirty)
    
    '
    ' Fix white space.
    '
    strClean = VBA.Replace(strClean, VBA.Chr(9), " ")        ' replace tab with space
    strClean = VBA.Replace(strClean, VBA.Chr(11), " ")       ' replace line feed with space
    strClean = VBA.Replace(strClean, VBA.Chr(13), " ")       ' replace return with space
    '
    ' Fix quotes.
    '
    strClean = VBA.Replace(strClean, "`", VBA.Chr(39))       ' replace back single quote with single quote
    strClean = VBA.Replace(strClean, VBA.Chr(145), VBA.Chr(39))  ' replace open single quote with single quote
    strClean = VBA.Replace(strClean, VBA.Chr(146), VBA.Chr(39))  ' replace close single quote with single quote
    strClean = VBA.Replace(strClean, VBA.Chr(147), VBA.Chr(34))  ' replace open double quote with double quote
    strClean = VBA.Replace(strClean, VBA.Chr(148), VBA.Chr(34))  ' replace close double quote with double quote
    
    strClean = VBA.Replace(strClean, "'m", "m")          ' ignore ' in 'm
    strClean = VBA.Replace(strClean, "'s", "s")          ' ignore ' in 's
    strClean = VBA.Replace(strClean, "'t", "t")          ' ignore ' in 't
    strClean = VBA.Replace(strClean, "'ve", "ve")        ' ignore ' in 've
    
    '
    ' Remove basic punctuation.
    '
    strClean = VBA.Replace(strClean, "-", " ")
    strClean = VBA.Replace(strClean, ",", " ")
    strClean = VBA.Replace(strClean, ";", " ")
    strClean = VBA.Replace(strClean, ":", " ")
    strClean = VBA.Replace(strClean, ".", " ")
    strClean = VBA.Replace(strClean, "!", " ")
    strClean = VBA.Replace(strClean, "?", " ")
    strClean = VBA.Replace(strClean, "/", " ")
    strClean = VBA.Replace(strClean, VBA.Chr(39), " ")
    strClean = VBA.Replace(strClean, VBA.Chr(34), " ")
    
    While (VBA.InStr(strClean, "  ") > 0)
        strClean = VBA.Replace(strClean, "  ", " ")
    Wend
    
    '
    ' Replace commonly interchanged words.
    '
    strClean = " " & strClean & " "
    strClean = VBA.Replace(strClean, " oh ", " o ")
    strClean = VBA.Replace(strClean, " alleluia ", " hallelujah ")
    strClean = VBA.Replace(strClean, " allelujah ", " hallelujah ")
    strClean = VBA.Replace(strClean, " emanuel ", " immanuel ")
    strClean = VBA.Replace(strClean, " emmanuel ", " immanuel ")
    strClean = VBA.Replace(strClean, " imanuel ", " immanuel ")
    If (VBA.Len(strClean) >= 2) Then
        strClean = VBA.Left(strClean, VBA.Len(strClean) - 1)
        strClean = VBA.Right(strClean, VBA.Len(strClean) - 1)
    End If
    
    mstrEverythingClean = strClean
End Function


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
