VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Slide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.Slide
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
Private msldDoppelganger As PowerPoint.Slide
Private strTitle As String
Private strTitleMatch1 As String
Private strTitleMatch2 As String
Private strTitleOrBodyMatch1 As String
Private strTitleOrBodyMatch2 As String


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
    ByRef sldDoppelganger As PowerPoint.Slide _
)
    mblnInvariant = False
    
    Set mudtApplication = udtApplication
    Set mudtParent = udtParent
    Set msldDoppelganger = sldDoppelganger
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
    If (Not blnInvariant = mblnInvariant) Then
        If (blnInvariant = True) Then
            strTitle = mstrLevel1Clean(strTitleGet)
            strTitleMatch1 = mstrLevel1Clean(strTitleGet)
            strTitleMatch2 = mstrLevel2Clean(strTitleMatch1)
            strTitleOrBodyMatch1 = mstrLevel1Clean(strTitleOrBodyGet)
            strTitleOrBodyMatch2 = mstrLevel2Clean(strTitleOrBodyMatch1)
            mblnInvariant = blnInvariant
        End If
        If (blnInvariant = False) Then
            mblnInvariant = blnInvariant
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
) As PowerPoint.Slide
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
Public Property Get SlideIndex _
( _
) As Long
    SlideIndex = msldDoppelganger.SlideIndex
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get SlideID _
( _
) As Long
    SlideID = msldDoppelganger.SlideID
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
    If (Not mblnInvariant) Then
        strTitle = mstrLevel1Clean(strTitleGet)
    End If
    Title = strTitle
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
    If (Category = "") Then Category = "<none>"
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
    If ((strCategory = "") Or (strCategory = "<none>")) Then
        msldDoppelganger.Tags.Delete "Category"
    Else
        msldDoppelganger.Tags.Add "Category", strCategory
    End If
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
    Dim lng As Long

    With msldDoppelganger.Parent.Windows
        For lng = 1 To .Count Step 1
            .Item(lng).View.Slide = msldDoppelganger
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
Public Function TitleMatches _
( _
    ByRef strMatch As String, _
    ByRef lngMatchLen As Long _
) As Boolean
    If (Not mblnInvariant) Then
        strTitleMatch1 = mstrLevel1Clean(strTitleGet)
        strTitleMatch2 = mstrLevel2Clean(strTitleMatch1)
    End If
    
    If ((VBA.Left(strTitleMatch1, lngMatchLen) = strMatch) Or _
        (VBA.Left(strTitleMatch2, lngMatchLen) = strMatch)) Then
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
    StringForTitleMatch = mstrLevel2Clean(str)
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
    If (Not mblnInvariant) Then
        strTitleOrBodyMatch1 = mstrLevel1Clean(strTitleOrBodyGet)
        strTitleOrBodyMatch2 = mstrLevel2Clean(strTitleOrBodyMatch1)
    End If
    
    If ((InStr(1, strTitleOrBodyMatch1, strMatch, VBA.vbTextCompare) > 0) Or _
        (InStr(1, strTitleOrBodyMatch2, strMatch, VBA.vbTextCompare) > 0)) Then
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
    StringForTitleOrBodyMatch = " " & mstrLevel2Clean(str)
End Function


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
Private Function strTitleGet _
( _
) As String
    strTitleGet = ""
    With msldDoppelganger
        If (.Shapes.HasTitle = Office.msoTrue) Then
            If (.Shapes.Title.HasTextFrame = Office.msoTrue) Then
                If (.Shapes.Title.TextFrame.HasText = Office.msoTrue) Then
                    strTitleGet = .Shapes.Title.TextFrame.TextRange.Text
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
Private Function strTitleOrBodyGet _
( _
) As String
    Dim lngIndex As Long
    
    strTitleOrBodyGet = VBA.vbNewLine & strTitleGet
    With msldDoppelganger.Shapes
        For lngIndex = 1 To .Count Step 1
            With .Item(lngIndex)
                If (.HasTextFrame = Office.msoTrue) Then
                    If (.TextFrame.HasText = Office.msoTrue) Then
                        strTitleOrBodyGet = _
                            strTitleOrBodyGet & _
                            VBA.vbNewLine & _
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
Private Function mstrLevel1Clean _
( _
    ByRef strDirty As String _
) As String
    Dim strClean As String
    
    strClean = strDirty
    
    ' Fix white space.
    strClean = VBA.Replace(strClean, VBA.vbTab, " ")             ' replace tab with space
    strClean = VBA.Replace(strClean, VBA.vbLf, " ")              ' replace line feed with space
    strClean = VBA.Replace(strClean, VBA.vbCr, " ")              ' replace return with space
    
    ' Fix quotes.
    strClean = VBA.Replace(strClean, VBA.Chr(96), VBA.Chr(39))   ' replace back single quote with single quote
    strClean = VBA.Replace(strClean, VBA.Chr(145), VBA.Chr(39))  ' replace open single quote with single quote
    strClean = VBA.Replace(strClean, VBA.Chr(146), VBA.Chr(39))  ' replace close single quote with single quote
    strClean = VBA.Replace(strClean, VBA.Chr(147), VBA.Chr(34))  ' replace open double quote with double quote
    strClean = VBA.Replace(strClean, VBA.Chr(148), VBA.Chr(34))  ' replace close double quote with double quote

    ' Collapse multiple spaces into a single space.
    Do While (VBA.InStr(strClean, "  ") > 0)
        strClean = VBA.Replace(strClean, "  ", " ")
    Loop
    
    ' Remove leading and trailing spaces.
    If (VBA.Left(strClean, 1) = " ") Then
        strClean = VBA.Right(strClean, VBA.Len(strClean) - 1)
    End If
    If (VBA.Right(strClean, 1) = " ") Then
        strClean = VBA.Left(strClean, VBA.Len(strClean) - 1)
    End If
    
    mstrLevel1Clean = strClean
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
'   This function assumes that the input strDirty is the output of the function
'   mstrLevel1Clean.
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mstrLevel2Clean _
( _
    ByRef strDirty As String _
) As String
    Dim strClean As String
    
    strClean = VBA.LCase(strDirty)
    
    ' Remove basic punctuation.
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
    
    ' Replace commonly interchanged words.
    strClean = " " & strClean & " "
    strClean = VBA.Replace(strClean, " oh ", " o ")
    strClean = VBA.Replace(strClean, " alleluia ", " hallelujah ")
    strClean = VBA.Replace(strClean, " allelujah ", " hallelujah ")
    strClean = VBA.Replace(strClean, " emanuel ", " immanuel ")
    strClean = VBA.Replace(strClean, " emmanuel ", " immanuel ")
    strClean = VBA.Replace(strClean, " imanuel ", " immanuel ")
    
    ' Replace multiple spaces with a single space.
    Do While (VBA.InStr(strClean, "  ") > 0)
        strClean = VBA.Replace(strClean, "  ", " ")
    Loop
    
    ' Remove leading and trailing spaces.
    If (VBA.Left(strClean, 1) = " ") Then
        strClean = VBA.Right(strClean, VBA.Len(strClean) - 1)
    End If
    If (VBA.Right(strClean, 1) = " ") Then
        strClean = VBA.Left(strClean, VBA.Len(strClean) - 1)
    End If
    
    mstrLevel2Clean = strClean
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

