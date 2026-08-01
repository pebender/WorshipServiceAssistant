VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Banner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.Banner
'
' Description:
'
' Author:
'   Paul Bender <pbender@alumni.ucsd.edu>
'
' Copyright:
'   Copyright (c) 2001,2002 Paul Bender
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
Private Const MstrBannerTag        As String = "WorshipServiceAssistant_Banner"
Private Const MstrBannerTagEnabled As String = "WorshipServiceAssistant_Banner_Enabled"


'===============================================================================
' Public Variables.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================
Private mblnInvariant As Boolean
Private mudtApplication As WorshipServiceAssistant.Application
Private mudtParent As WorshipServiceAssistant.Application
Private mudtPresentation As New WorshipServiceAssistant.Presentation
    

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
    ByRef udtParent As WorshipServiceAssistant.Application _
)
    mblnInvariant = False
    
    Set mudtApplication = udtApplication
    Set mudtParent = udtParent
    
    mUpdate
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
    ' This should not be needed.
    If (Not mudtPresentation.Doppelganger Is Nothing) Then
        mudtPresentation.Doppelganger.Saved = Office.msoTrue
    End If
    
    mudtPresentation.Terminate
    
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
            mUpdate
            mblnInvariant = blnInvariant
        End If
        mudtPresentation.Invariant = blnInvariant
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
Public Function Doppelganger _
( _
) As PowerPoint.Presentation
    If (Not mblnInvariant) Then mUpdate
    
    Set Doppelganger = mudtPresentation.Doppelganger
End Function

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
) As WorshipServiceAssistant.Application
    Set Parent = mudtParent
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub Load _
( _
    ByRef strBanner As String, _
    ByRef lngRed As Long, _
    ByRef lngGreen As Long, _
    ByRef lngBlue As Long _
)
    If (Not mblnInvariant) Then mUpdate
    
    With mudtPresentation.Doppelganger
        .SlideMaster.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = VBA.RGB(0, 0, 0)
        .Slides.Item(1).Shapes.Title.TextFrame.TextRange.Text = strBanner
        .SlideMaster.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = VBA.RGB(lngRed, lngGreen, lngBlue)
        Me.Visible = True
        .Saved = Office.msoTrue
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Visible _
( _
) As Boolean
    If (Not mblnInvariant) Then mUpdate
    
    Visible = mudtPresentation.SlideShowWindow.Visible
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Let Visible _
( _
    ByRef blnVisible As Boolean _
)
    If (Not mblnInvariant) Then mUpdate
    
    mudtPresentation.SlideShowWindow.Visible = blnVisible
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Enabled _
( _
) As Boolean
    If (Not mblnInvariant) Then mUpdate
    
    With mudtPresentation.Doppelganger
        If (.Tags(MstrBannerTagEnabled) = "True") Then
            Enabled = True
        Else
            Enabled = False
        End If
        .Saved = Office.msoTrue
    End With
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Let Enabled _
( _
    ByRef blnEnabled As Boolean _
)
    Dim lng As Long
    
    If (Not mblnInvariant) Then mUpdate
    
    With mudtPresentation
        With .Doppelganger
            If (blnEnabled) Then
                .Tags.Add MstrBannerTagEnabled, "True"
            Else
                Me.Visible = False
                .Tags.Add MstrBannerTagEnabled, "False"
            End If
            .Saved = Office.msoTrue
        End With
        If (blnEnabled) Then
            If (.SlideShowWindow.Doppelganger Is Nothing) Then
                .SlideShowWindow.Load 1
            End If
        End If
    End With
    
    With mudtApplication.Presentations
        For lng = 1 To .Count Step 1
            Me.AdjustSlideShowWindow .Item(lng).SlideShowWindow
        Next
    End With
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub AdjustSlideShowWindow _
( _
    ByRef udtSlideShowWindow As WorshipServiceAssistant.SlideShowWindow _
)
    Dim sngTop As Single
    Dim sngHeight As Single
    Dim sngLeft As Single
    Dim sngWidth As Single
    Dim sngDelta As Single
    
    If (Not mblnInvariant) Then mUpdate
    
    If (mudtPresentation.SlideShowWindow.Doppelganger Is Nothing) Then Exit Sub
    If (udtSlideShowWindow.Doppelganger Is Nothing) Then Exit Sub
    If (udtSlideShowWindow Is mudtPresentation.SlideShowWindow) Then Exit Sub
    
    With mudtPresentation.Doppelganger
        If (Me.Enabled) Then
            sngDelta = _
                (.SlideMaster.Shapes.Title.Height * _
                 .SlideShowWindow.Height) / _
                 .PageSetup.SlideHeight
        Else
            sngDelta = 0
        End If
        With .SlideShowWindow
            sngTop = .Top
            sngHeight = .Height
            sngLeft = .Left
            sngWidth = .Width
        End With
    End With
    With udtSlideShowWindow.Doppelganger
        If (.Height > sngHeight - sngDelta) Then
            .Height = sngHeight - sngDelta
            .Top = sngTop + sngDelta
            .Width = sngWidth
            .Left = sngLeft
        Else
            .Top = sngTop + sngDelta
            .Height = sngHeight - sngDelta
            .Left = sngLeft
            .Width = sngWidth
        End If
    End With
End Sub


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
Private Sub mUpdate _
( _
)
    Dim app As PowerPoint.Application
    
    Set app = Nothing
    On Error Resume Next
    Set app = mudtPresentation.Doppelganger.Application
    On Error GoTo 0
    
    If (app Is Nothing) Then
        mudtPresentation.Initialize mudtApplication, mudtParent, mprePresentationGet
    End If
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mprePresentationGet _
( _
)
    Dim pre As PowerPoint.Presentation
    
    For Each pre In PowerPoint.Presentations
        If (pre.Tags(MstrBannerTag) = "True") Then
            Exit For
        End If
    Next
    If (pre Is Nothing) Then
        Set pre = mprePresentationCreate
        pre.Tags.Add MstrBannerTag, "True"
    End If
    
    Set mprePresentationGet = pre
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mprePresentationCreate _
( _
) As PowerPoint.Presentation
    Dim prePresentation As PowerPoint.Presentation
    Dim lngHeight As Long
    Dim lngWidth As Long
    Dim sswActive As PowerPoint.SlideShowWindow
    Dim lngIndex As Long
    
    Set prePresentation = mudtApplication.Doppelganger.Presentations.Add(withwindow:=Office.msoFalse)
    
    With prePresentation
        With .PageSetup
            .SlideSize = PowerPoint.ppSlideSizeOnScreen
            .SlideOrientation = Office.msoOrientationHorizontal
        End With
        With .Slides
            For lngIndex = .Count To 1 Step -1
                .Item(lngIndex).Delete
            Next
        End With
        With .SlideMaster
            With .Background
                With .Fill
                    .Solid
                    .BackColor.RGB = VBA.RGB(0, 0, 0)
                    .ForeColor.RGB = VBA.RGB(0, 0, 0)
                End With
            End With
            For lngIndex = .Shapes.Count To 1 Step -1
                .Shapes(lngIndex).Delete
            Next
            .Shapes.AddTitle
            With .Shapes.Title
                .Top = 0
                .Left = 0
                .Width = prePresentation.PageSetup.SlideWidth
                .Height = 0.5 * 72
            End With
            With .Shapes.Title.TextFrame
                .HorizontalAnchor = Office.msoAnchorCenter
                .VerticalAnchor = Office.msoAnchorTop
                .Orientation = Office.msoTextOrientationHorizontal
                .WordWrap = Office.msoFalse
                .MarginTop = 4
                .MarginRight = 4
                .MarginBottom = 4
                .MarginLeft = 4
                With .TextRange
                    With .ParagraphFormat
                        .Alignment = PowerPoint.ppAlignCenter
                        .Bullet = Office.msoFalse
                        .WordWrap = Office.msoFalse
                    End With
                    With .Font
                        .NameAscii = "Arial"
                        .Size = 28
                        .Bold = Office.msoFalse
                        .Italic = Office.msoFalse
                        .Underline = Office.msoFalse
                        .Subscript = Office.msoFalse
                        .Superscript = Office.msoFalse
                        .Emboss = Office.msoFalse
                        .Shadow = Office.msoFalse
                        .Color.RGB = VBA.RGB(0, 0, 0)
                    End With
                End With
            End With
        End With
        .Slides.Add 1, PowerPoint.ppLayoutTitleOnly
        
        '
        ' There should be no need to fill in the title text.
        ' However, it appears that under PowerPoint 10.0,
        ' if there is no title text when the slide show is started,
        ' then no title text can be added later.
        '
        .SlideMaster.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = VBA.RGB(0, 0, 0)
        .Slides.Item(1).Shapes.Title.TextFrame.TextRange.Text = "Welcome"
        .SlideMaster.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = VBA.RGB(0, 0, 0)
    
        .Saved = Office.msoTrue
    End With
    
    Set mprePresentationCreate = prePresentation
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

