Attribute VB_Name = "Banner"
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
'   1.00.1005:
'     (1) Worked around problem under Windows XP with PowerPoint XP
'         that would cause the banner to fail to display.
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
Private Pres As PowerPoint.Presentation


'===============================================================================
' Public Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Function IsBanner(ByRef P As PowerPoint.Presentation) As Boolean
    If (P.Tags("WorshipServiceAssistantType") = "Banner") Then
        IsBanner = True
    Else
        IsBanner = False
    End If
End Function

Public Sub Apply(ByVal SSW As PowerPoint.SlideShowWindow)
    If (Banner.Exists = False) Then
        Exit Sub
    End If
    If (Banner.IsBanner(SSW.Presentation) = True) Then
        Exit Sub
    End If
    
    Dim Height As Integer
    
    SSW.Left = Pres.SlideShowWindow.Left
    SSW.Width = Pres.SlideShowWindow.Width
    If (Banner.Enabled = False) Then
        Height = 0
        SSW.Height = Pres.SlideShowWindow.Height - Height
        SSW.Top = Pres.SlideShowWindow.Top + Height
    Else
        Height = (Pres.SlideMaster.Shapes.Title.Height * Pres.SlideShowWindow.Height) / Pres.PageSetup.SlideHeight
        SSW.Top = Pres.SlideShowWindow.Top + Height
        SSW.Height = Pres.SlideShowWindow.Height - Height
    End If
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Create()
    Dim i As Integer
    Dim S As PowerPoint.Shape
    
    Set Pres = Nothing
    
    For i = 1 To Application.Presentations.Count Step 1
        If (Banner.IsBanner(Application.Presentations(i)) = True) Then
            Exit For
        End If
    Next
    If (i <= Application.Presentations.Count) Then
        Set Pres = Application.Presentations(i)
    Else
        Set Pres = Application.Presentations.Add(msoTrue)
        Pres.Tags.Add "WorshipServiceAssistantType", "Banner"
        Banner.Enabled = False
        With Pres.PageSetup
            .SlideSize = ppSlideSizeOnScreen
            .SlideOrientation = msoOrientationHorizontal
        End With
        For i = Pres.slides.Count To 1 Step -1
            Pres.slides(i).Delete
        Next
        With Pres.SlideMaster
            With .Background
                With .Fill
                    .Solid
                    .BackColor.RGB = RGB(0, 0, 0)
                    .ForeColor.RGB = RGB(0, 0, 0)
                End With
            End With
            For i = .Shapes.Count To 1 Step -1
                .Shapes(i).Delete
            Next
            .Shapes.AddTitle
            With .Shapes.Title
                .Top = 0
                .Left = 0
                .Width = Pres.PageSetup.SlideWidth
                .Height = 0.5 * 72
            End With
            With .Shapes.Title.TextFrame
                .HorizontalAnchor = msoAnchorCenter
                .VerticalAnchor = msoAnchorTop
                .Orientation = msoTextOrientationHorizontal
                .WordWrap = msoFalse
                .MarginTop = 4
                .MarginRight = 4
                .MarginBottom = 4
                .MarginLeft = 4
                With .TextRange
                    With .ParagraphFormat
                        .Alignment = ppAlignCenter
                        .Bullet = msoFalse
                        .WordWrap = msoFalse
                    End With
                    With .Font
                        .NameAscii = "Arial"
                        .Size = 28
                        .Bold = msoFalse
                        .Italic = msoFalse
                        .Underline = msoFalse
                        .Subscript = msoFalse
                        .Superscript = msoFalse
                        .Emboss = msoFalse
                        .Shadow = msoFalse
                        .Color.RGB = RGB(0, 0, 0)
                    End With
                End With
            End With
        End With
        Pres.slides.Add 1, ppLayoutTitleOnly
        
        '
        ' There should be no need to fill in the title text.
        ' However, it appears that under Windows XP with PowerPoint XP,
        ' if there is no title text when the slide show is started,
        ' then no title text can be added later.
        '
        Pres.SlideMaster.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
        Pres.slides(1).Shapes.Title.TextFrame.TextRange.Text = "Welcome"
        Pres.SlideMaster.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
        
        SlideShow_Setup Pres
    End If
    If (ActiveSlideShowExists(Pres) = True) Then
    Else
        Dim SSW As PowerPoint.SlideShowWindow
        Dim SSWState As Integer
        
        Set SSW = Nothing
        '
        ' If there is just one slide show window,
        ' then will make sure that it is on top.
        ' Otherwise, we do not know which one is on top.
        '
        If (Application.SlideShowWindows.Count = 1) Then
            Set SSW = Application.SlideShowWindows(1)
        End If
        
        Pres.SlideShowSettings.Run
        SlideShow_Scale Pres
        If ((SSW Is Nothing) = False) Then
            SSW.Activate
        End If
    End If
    
    Pres.Saved = msoTrue
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Delete()
    If (Banner.Exists = False) Then
        Exit Sub
    End If
    Pres.Saved = msoTrue
    Pres.Close
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Sub Load(ByVal BannerString As String, ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer)
    If (Banner.Exists = False) Then
        Exit Sub
    End If
    If (Banner.Enabled = False) Then
        Exit Sub
    End If
    
    Pres.SlideMaster.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Pres.slides(1).Shapes.Title.TextFrame.TextRange.Text = BannerString
    Pres.SlideMaster.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(Red, Green, Blue)
    Banner.Visible = True
    Pres.Saved = msoTrue
End Sub

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Function Exists() As Boolean
    Exists = False
    If (Pres Is Nothing) Then
        Exit Function
    End If
    If (ActiveSlideShowExists(Pres) = False) Then
        Exit Function
    End If
    Exists = True
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Property Get Visible() As Boolean
    Visible = False
    If (Banner.Exists = False) Then
        Exit Property
    End If
    If (Banner.Enabled = False) Then
        Exit Property
    End If
    If (Pres.SlideShowWindow.View.State <> ppSlideShowBlackScreen) Then
        Visible = True
    End If
End Property

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Property Let Visible(ByVal Value As Boolean)
    If (Banner.Exists = False) Then
        Exit Property
    End If
    With Pres.SlideShowWindow.View
        If (Value = True) Then
            .State = ppSlideShowRunning
            .State = ppSlideShowPaused
        Else
            .State = ppSlideShowBlackScreen
        End If
        .PointerType = ppSlideShowPointerArrow
    End With
End Property

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Property Get Enabled() As Boolean
    Enabled = False
    If (Banner.Exists = False) Then
        Exit Property
    End If
    If (Pres.Tags("WorshipServiceAssistantBannerEnabled") = "True") Then
        Enabled = True
    End If
End Property

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Public Property Let Enabled(ByVal Value As Boolean)
    If (Banner.Exists = False) Then
        Exit Property
    End If
    
    Dim SSW As PowerPoint.SlideShowWindow
    
    If (Value = True) Then
        Pres.Tags.Add "WorshipServiceAssistantBannerEnabled", "True"
    Else
        Banner.Visible = False
        Pres.Tags.Add "WorshipServiceAssistantBannerEnabled", "False"
    End If
    Pres.Saved = msoTrue
    
    For Each SSW In Application.SlideShowWindows
        Banner.Apply SSW
    Next
End Property

'===============================================================================
' Private Subroutines and Functions.
'===============================================================================
