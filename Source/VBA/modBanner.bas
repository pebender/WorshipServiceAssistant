Attribute VB_Name = "modBanner"
'===============================================================================
' Name:
'   WorshipServiceAssistant.modBanner
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
'   1.03.0002:
'     (1) Made changes to the source code so that it follows Microsoft's
'         Visual Basic coding conventions.
'   1.00.1005:
'     (1) Worked around problem under PowerPoint 2002 that would cause
'         the banner to fail to display.
'   1.00.0002:
'     (1) Added support for abitrary banner color.
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
Private mpreBanner As PowerPoint.Presentation


'===============================================================================
' Public Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Function gblnIsBanner _
( _
    ByRef prePresentation As PowerPoint.Presentation _
) As Boolean
    If (prePresentation.Tags("WorshipServiceAssistantType") = "Banner") Then
        gblnIsBanner = True
    Else
        gblnIsBanner = False
    End If
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub gApply _
( _
    ByVal sswSlideShowWindow As PowerPoint.SlideShowWindow _
)
    Dim lngHeight As Long
    
    If (modBanner.gblnExists = False) Then
        Exit Sub
    End If
    If (modBanner.gblnIsBanner(sswSlideShowWindow.Presentation) = True) Then
        Exit Sub
    End If
    
    sswSlideShowWindow.Left = mpreBanner.SlideShowWindow.Left
    sswSlideShowWindow.Width = mpreBanner.SlideShowWindow.Width
    If (modBanner.gblnEnabled = False) Then
        lngHeight = 0
        sswSlideShowWindow.Height = mpreBanner.SlideShowWindow.Height - lngHeight
        sswSlideShowWindow.Top = mpreBanner.SlideShowWindow.Top + lngHeight
    Else
        lngHeight = (mpreBanner.SlideMaster.Shapes.Title.Height * mpreBanner.SlideShowWindow.Height) / mpreBanner.PageSetup.SlideHeight
        sswSlideShowWindow.Top = mpreBanner.SlideShowWindow.Top + lngHeight
        sswSlideShowWindow.Height = mpreBanner.SlideShowWindow.Height - lngHeight
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub gCreate _
( _
)
    Dim sswSlideShowWindow As PowerPoint.SlideShowWindow
    Dim lngIndex As Long
    
    Set mpreBanner = Nothing
    
    For lngIndex = 1 To Application.Presentations.Count Step 1
        If (modBanner.gblnIsBanner(Application.Presentations(lngIndex)) = True) Then
            Exit For
        End If
    Next
    If (lngIndex <= Application.Presentations.Count) Then
        Set mpreBanner = Application.Presentations(lngIndex)
    Else
        Set mpreBanner = Application.Presentations.Add(Office.msoTrue)
        mpreBanner.Tags.Add "WorshipServiceAssistantType", "Banner"
        modBanner.gblnEnabled = False
        With mpreBanner.PageSetup
            .SlideSize = PowerPoint.ppSlideSizeOnScreen
            .SlideOrientation = Office.msoOrientationHorizontal
        End With
        For lngIndex = mpreBanner.Slides.Count To 1 Step -1
            mpreBanner.Slides(lngIndex).Delete
        Next
        With mpreBanner.SlideMaster
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
                .Width = mpreBanner.PageSetup.SlideWidth
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
        mpreBanner.Slides.Add 1, PowerPoint.ppLayoutTitleOnly
        
        '
        ' There should be no need to fill in the title text.
        ' However, it appears that under PowerPoint 2002,
        ' if there is no title text when the slide show is started,
        ' then no title text can be added later.
        '
        mpreBanner.SlideMaster.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = VBA.RGB(0, 0, 0)
        mpreBanner.Slides(1).Shapes.Title.TextFrame.TextRange.Text = "Welcome"
        mpreBanner.SlideMaster.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = VBA.RGB(0, 0, 0)
        
        modSlideShow.gSetup mpreBanner
    End If
    If (modActive.gblnActiveSlideShowExists(mpreBanner) = True) Then
    Else
        Set sswSlideShowWindow = Nothing
        '
        ' If there is just one slide show window,
        ' then will make sure that it is on top.
        ' Otherwise, we do not know which one is on top.
        '
        If (Application.SlideShowWindows.Count = 1) Then
            Set sswSlideShowWindow = Application.SlideShowWindows(1)
        End If
        
        mpreBanner.SlideShowSettings.Run
        modSlideShow.gSizeSet mpreBanner
        If ((sswSlideShowWindow Is Nothing) = False) Then
            sswSlideShowWindow.Activate
        End If
    End If
    
    mpreBanner.Saved = Office.msoTrue
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub gDelete _
( _
)
    If (modBanner.gblnExists = False) Then
        Exit Sub
    End If
    mpreBanner.Saved = Office.msoTrue
    mpreBanner.Close
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub gLoad _
( _
    ByVal strBanner As String, _
    ByVal intRed As Integer, _
    ByVal intGreen As Integer, _
    ByVal intBlue As Integer _
)
    If (modBanner.gblnExists = False) Then
        Exit Sub
    End If
    If (modBanner.gblnEnabled = False) Then
        Exit Sub
    End If
    
    mpreBanner.SlideMaster.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = VBA.RGB(0, 0, 0)
    mpreBanner.Slides(1).Shapes.Title.TextFrame.TextRange.Text = strBanner
    mpreBanner.SlideMaster.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = VBA.RGB(intRed, intGreen, intBlue)
    modBanner.gblnVisible = True
    mpreBanner.Saved = Office.msoTrue
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Function gblnExists _
( _
) As Boolean
    gblnExists = False
    If (mpreBanner Is Nothing) Then
        Exit Function
    End If
    If (modActive.gblnActiveSlideShowExists(mpreBanner) = False) Then
        Exit Function
    End If
    gblnExists = True
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Property Get gblnVisible _
( _
) As Boolean
    gblnVisible = False
    If (modBanner.gblnExists = False) Then
        Exit Property
    End If
    If (modBanner.gblnEnabled = False) Then
        Exit Property
    End If
    If (mpreBanner.SlideShowWindow.View.State <> PowerPoint.ppSlideShowBlackScreen) Then
        gblnVisible = True
    End If
End Property

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Property Let gblnVisible _
( _
    ByVal blnValue As Boolean _
)
    If (modBanner.gblnExists = False) Then
        Exit Property
    End If
    With mpreBanner.SlideShowWindow.View
        If (blnValue = True) Then
            .State = PowerPoint.ppSlideShowRunning
            .State = PowerPoint.ppSlideShowPaused
        Else
            .State = PowerPoint.ppSlideShowBlackScreen
        End If
        .PointerType = PowerPoint.ppSlideShowPointerArrow
    End With
End Property

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Property Get gblnEnabled _
( _
) As Boolean
    gblnEnabled = False
    If (modBanner.gblnExists = False) Then
        Exit Property
    End If
    If (mpreBanner.Tags("WorshipServiceAssistantBannerEnabled") = "True") Then
        gblnEnabled = True
    End If
End Property

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Property Let gblnEnabled _
( _
    ByVal blnValue As Boolean _
)
    Dim sswSlideShowWindow As PowerPoint.SlideShowWindow
    
    If (modBanner.gblnExists = False) Then
        Exit Property
    End If
    
    If (blnValue = True) Then
        mpreBanner.Tags.Add "WorshipServiceAssistantBannerEnabled", "True"
    Else
        modBanner.gblnVisible = False
        mpreBanner.Tags.Add "WorshipServiceAssistantBannerEnabled", "False"
    End If
    mpreBanner.Saved = Office.msoTrue
    
    For Each sswSlideShowWindow In Application.SlideShowWindows
        modBanner.gApply sswSlideShowWindow
    Next
End Property

'===============================================================================
' Private Subroutines and Functions.
'===============================================================================
