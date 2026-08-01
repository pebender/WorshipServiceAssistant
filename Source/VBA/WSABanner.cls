VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WSABanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.WSABanner
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
'   1.04.0001:
'     (1) Fixed a bug that caused the banner to display on the wrong
'         monitor. As a result of this fix, the banner presentation window
'         is no longer hidden.
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
Private mudtApplication As WorshipServiceAssistant.WSAApplication

Private mudtPresentation As New WorshipServiceAssistant.WSAPresentation


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
Public Property Get Presentation _
( _
) As WorshipServiceAssistant.WSAPresentation
    Set Presentation = mudtPresentation
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
    ByVal strBanner As String, _
    ByVal intRed As Integer, _
    ByVal intGreen As Integer, _
    ByVal intBlue As Integer _
)
    If (Me.Enabled = False) Then
        Exit Sub
    End If
    
    With mudtPresentation.Doppelganger
        .SlideMaster.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = VBA.RGB(0, 0, 0)
        .Slides(1).Shapes.Title.TextFrame.TextRange.Text = strBanner
        .SlideMaster.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = VBA.RGB(intRed, intGreen, intBlue)
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
    ByVal blnVisible As Boolean _
)
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
    ByVal blnEnabled As Boolean _
)
    Dim lngIndex As Long
    
    With mudtPresentation.Doppelganger
        If (blnEnabled = True) Then
            .Tags.Add MstrBannerTagEnabled, "True"
        Else
            Me.Visible = False
            .Tags.Add MstrBannerTagEnabled, "False"
        End If
        .Saved = Office.msoTrue
    End With
    
    With mudtApplication.Presentations
        For lngIndex = 1 To .Count Step 1
            Me.AdjustSlideShowWindow .Item(lngIndex).SlideShowWindow
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
Public Sub Initialize _
( _
    ByRef udtApplication As WorshipServiceAssistant.WSAApplication _
)
    Dim prePresentation As PowerPoint.Presentation
    Dim lngIndex As Long
    
    Set mudtApplication = udtApplication
    
    With mudtApplication.Doppelganger.Presentations
        For lngIndex = 1 To .Count Step 1
            If (.Item(lngIndex).Tags(MstrBannerTag) = "True") Then
                Set prePresentation = .Item(lngIndex)
                Exit For
            End If
        Next
        If (lngIndex > .Count) Then
            Set prePresentation = mpreCreate
            prePresentation.Tags.Add MstrBannerTag, "True"
        End If
    End With
    
    mudtPresentation.Initialize mudtApplication, prePresentation
    
    If (mudtPresentation.SlideShowWindow.Exists = False) Then
        mudtPresentation.SlideShowWindow.Load 1
        Me.Enabled = False
    End If
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
    mudtPresentation.Doppelganger.Saved = Office.msoTrue
    
    mudtPresentation.Terminate
    
    Set mudtApplication = Nothing
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Sub AdjustSlideShowWindow _
( _
    ByRef udtSlideShowWindow As WorshipServiceAssistant.WSASlideShowWindow _
)
    Dim sngTop As Single
    Dim sngHeight As Single
    Dim sngLeft As Single
    Dim sngWidth As Single
    Dim sngDelta As Single
    
    If ((udtSlideShowWindow Is mudtPresentation.SlideShowWindow) = True) Then
        Exit Sub
    End If
    
    If (udtSlideShowWindow.Exists = False) Then
        Exit Sub
    End If
    
    With mudtPresentation.Doppelganger
        If (Me.Enabled = False) Then
            sngDelta = 0
        Else
            sngDelta = _
                (.SlideMaster.Shapes.Title.Height * _
                 .SlideShowWindow.Height) / _
                 .PageSetup.SlideHeight
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
Private Function mpreCreate _
( _
) As PowerPoint.Presentation
    Dim prePresentation As PowerPoint.Presentation
    Dim lngHeight As Long
    Dim lngWidth As Long
    Dim sswActive As PowerPoint.SlideShowWindow
    Dim lngIndex As Long
    
    Set prePresentation = mudtApplication.Doppelganger.Presentations.Add
    
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
        ' However, it appears that under PowerPoint 2002,
        ' if there is no title text when the slide show is started,
        ' then no title text can be added later.
        '
        .SlideMaster.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = VBA.RGB(0, 0, 0)
        .Slides(1).Shapes.Title.TextFrame.TextRange.Text = "Welcome"
        .SlideMaster.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = VBA.RGB(0, 0, 0)
    
        .Saved = Office.msoTrue
    End With
    
    Set mpreCreate = prePresentation
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

