Attribute VB_Name = "modIndex"
'===============================================================================
' Name:
'   WorshipServiceAssistant.modIndex
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
'   1.03.0002:
'     (1) Made changes to the source code so that it follows Microsoft's
'         Visual Basic coding conventions.
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
' Private Constants.
'===============================================================================


'===============================================================================
' Public Variables.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================


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
Public Sub gRun _
( _
    ByRef dwCurrent As PowerPoint.DocumentWindow _
)
    Dim preCurrent As PowerPoint.Presentation
    Dim preIndex As PowerPoint.Presentation
    Dim strCategory As String
    Dim lngCategoryIndex As Long
    Dim strIndexTitle As String
    Dim strSlideTitle As String
    Dim strSlideTitleLast As String
    Dim lngSlideTitleRemaining As Long
    Dim lngBlockTitleRemaining As Long
    Dim lngColMax As Long
    Dim lngRowMax As Long
    Dim lngColWidth As Long
    Dim lngRowHeight As Long
    Dim lngColIndex As Long
    Dim lngRowIndex As Long
    Dim blnColorState As Boolean
    Dim lngIndex As Long
    Dim sldSlide As PowerPoint.Slide
    Dim shpShape As PowerPoint.Shape
    Dim strShapeText As String
    Dim blnShapeLineChop As Boolean
    Dim lngShapeLineCount As Long
    Dim lngShapeBottomLast As Long
    Dim lngTemp As Long
    Dim astrCategoryList() As String
    
    If ((dwCurrent Is Nothing) = False) Then
        If (dwCurrent.Presentation.Slides.Count >= 1) Then
            Set preCurrent = dwCurrent.Presentation
            
            ' Find all the slide categories that are in the category list,
            ' preserving the order of the category list.
            ReDim astrCategoryList(0)
            For lngCategoryIndex = 1 To modCategory.Count Step 1
                strCategory = modCategory.Item(lngCategoryIndex)
                For lngIndex = 1 To preCurrent.Slides.Count Step 1
                    If (preCurrent.Slides(lngIndex).Tags("Category") = strCategory) Then
                        astrCategoryList(UBound(astrCategoryList)) = strCategory
                        ReDim Preserve astrCategoryList(UBound(astrCategoryList) + 1)
                        Exit For
                    End If
                Next
            Next
            
            ' Find all the slide categories that are not in the category list,
            ' preserving the order of occurance in the slides.
            For lngIndex = 1 To preCurrent.Slides.Count Step 1
                strCategory = preCurrent.Slides(lngIndex).Tags("Category")
                For lngCategoryIndex = LBound(astrCategoryList) To UBound(astrCategoryList) - 1 Step 1
                    If (strCategory = astrCategoryList(lngCategoryIndex)) Then
                        Exit For
                    End If
                Next
                If (lngCategoryIndex >= UBound(astrCategoryList)) Then
                    astrCategoryList(UBound(astrCategoryList)) = strCategory
                    ReDim Preserve astrCategoryList(UBound(astrCategoryList) + 1)
                End If
            Next
            
            ReDim Preserve astrCategoryList(UBound(astrCategoryList) - 1)
            
            Set preIndex = mppIndexAdd()
                
            With preIndex.SlideMaster.TextStyles(ppBodyStyle).Levels(1).ParagraphFormat
                lngRowHeight = .SpaceBefore + .SpaceWithin + .SpaceAfter
            End With
            lngColWidth = 72 * 2.5
            lngColMax = preIndex.PageSetup.SlideWidth
            lngColMax = VBA.Int(lngColMax / lngColWidth)
            lngRowMax = preIndex.PageSetup.SlideHeight
            If (preIndex.SlideMaster.Shapes.HasTitle = Office.msoTrue) Then
                lngRowMax = lngRowMax - preIndex.SlideMaster.Shapes.Title.Height
            End If
            lngRowMax = VBA.Int(lngRowMax / lngRowHeight)
            
            strIndexTitle = preCurrent.Name
            If (VBA.LCase(VBA.Right(strIndexTitle, 4)) = ".ppt") Then
                strIndexTitle = VBA.Left(strIndexTitle, VBA.Len(strIndexTitle) - 4)
            End If
            strIndexTitle = _
                strIndexTitle & " Index" & _
                VBA.Chr(11) & _
                "(Generated on " & Date & " at " & Time & ")"
            
            lngColIndex = lngColMax + 1
            For lngCategoryIndex = LBound(astrCategoryList) To UBound(astrCategoryList) Step 1
                strCategory = astrCategoryList(lngCategoryIndex)
                lngRowIndex = lngRowMax + 1
                blnColorState = False
                strSlideTitleLast = ""
                lngSlideTitleRemaining = 0
                lngBlockTitleRemaining = 0
                If (preIndex.SlideMaster.Shapes.HasTitle = Office.msoTrue) Then
                    lngShapeBottomLast = preIndex.SlideMaster.Shapes.Title.Height
                Else
                    lngShapeBottomLast = 0
                End If
                strSlideTitleLast = ""
                For lngIndex = 1 To preCurrent.Slides.Count Step 1
                    If (preCurrent.Slides(lngIndex).Tags("Category") = strCategory) Then
                        '
                        ' Get the slide's title
                        '
                        If (preCurrent.Slides(lngIndex).Shapes.HasTitle = Office.msoTrue) Then
                            strSlideTitle = mstrIndexTitleClean(preCurrent.Slides(lngIndex).Shapes.Title.TextFrame.TextRange.Text)
                        Else
                            strSlideTitle = ""
                        End If
                        
                        If ((strSlideTitle <> "") And (strSlideTitle <> strSlideTitleLast)) Then
                            If (lngSlideTitleRemaining <= 0) Then
                                lngSlideTitleRemaining = mlngIndexGroupDetermine(preCurrent, lngIndex)
                                blnColorState = Not blnColorState
                                lngBlockTitleRemaining = 0
                                If (lngRowIndex + lngSlideTitleRemaining > lngRowMax) Then
                                    lngRowIndex = lngRowMax + 1
                                    lngTemp = lngSlideTitleRemaining
                                    lngTemp = lngTemp - (lngTemp Mod lngRowMax)
                                    lngTemp = lngTemp / lngRowMax
                                    If (lngTemp <= lngColMax) Then
                                        If (lngColIndex + lngTemp > lngColMax) Then
                                            lngColIndex = lngColMax + 1
                                        End If
                                    End If
                                End If
                            End If
                            
                            If (lngRowIndex > lngRowMax) Then
                                lngRowIndex = 1
                                lngColIndex = lngColIndex + 1
                                If (lngColIndex > lngColMax) Then
                                    lngRowIndex = 1
                                    lngColIndex = 1
                                    Set sldSlide = mppIndexSlideAdd(preIndex)
                                    mppIndexTitleAdd sldSlide, strIndexTitle
                                End If
                                Set shpShape = mppIndexCategoryAdd(sldSlide, lngColIndex, lngRowIndex, 2, strCategory)
                                lngShapeBottomLast = shpShape.Top + shpShape.Height
                                shpShape.Fill.ForeColor.RGB = VBA.RGB(255, 255, 255)
                                lngRowIndex = lngRowIndex + 2
                                lngBlockTitleRemaining = 0
                                If (preIndex.SlideMaster.Shapes.HasTitle = Office.msoTrue) Then
                                    lngShapeBottomLast = _
                                        preIndex.SlideMaster.Shapes.Title.Top + _
                                        preIndex.SlideMaster.Shapes.Title.Height
                                Else
                                    lngShapeBottomLast = 0
                                End If
                            End If
                            
                            If (lngBlockTitleRemaining <= 0) Then
                                lngTemp = lngRowMax - VBA.Int(lngShapeBottomLast / lngRowHeight) + 1
                                If (lngSlideTitleRemaining <= lngTemp) Then
                                    Set shpShape = mppIndexShapeAdd(sldSlide, lngColIndex, lngRowIndex, lngSlideTitleRemaining)
                                    lngBlockTitleRemaining = lngSlideTitleRemaining
                                Else
                                    Set shpShape = mppIndexShapeAdd(sldSlide, lngColIndex, lngRowIndex, lngTemp)
                                    lngBlockTitleRemaining = lngTemp
                                End If
                                lngShapeBottomLast = shpShape.Top + shpShape.Height
                                If (blnColorState = False) Then
                                    shpShape.Fill.ForeColor.RGB = VBA.RGB(220, 220, 220)
                                Else
                                    shpShape.Fill.ForeColor.RGB = VBA.RGB(255, 255, 255)
                                End If
                            End If
                            
                            If (shpShape.TextFrame.HasText = Office.msoFalse) Then
                                strShapeText = ""
                                lngShapeLineCount = 1
                            Else
                                strShapeText = shpShape.TextFrame.TextRange.Text & VBA.vbCrLf
                                lngShapeLineCount = shpShape.TextFrame.TextRange.Lines.Count + 1
                            End If
                            strShapeText = strShapeText & str(lngIndex) & "." & VBA.Chr(9) & strSlideTitle
                            shpShape.TextFrame.TextRange.Text = strShapeText
                            With shpShape.TextFrame.TextRange
                                blnShapeLineChop = False
                                While (.Lines.Count > lngShapeLineCount)
                                    If (blnShapeLineChop = True) Then
                                        .Text = VBA.Left(.Text, VBA.Len(.Text) - 4)
                                    End If
                                    While (VBA.Right(.Text, 1) <> " ")
                                        .Text = VBA.Left(.Text, VBA.Len(.Text) - 1)
                                    Wend
                                    .Text = VBA.Left(.Text, VBA.Len(.Text) - 1)
                                    .Text = .Text & " ..."
                                    blnShapeLineChop = True
                                Wend
                            End With
                            
                            lngRowIndex = lngRowIndex + 1
                            lngBlockTitleRemaining = lngBlockTitleRemaining - 1
                            lngSlideTitleRemaining = lngSlideTitleRemaining - 1
                        End If
                        strSlideTitleLast = strSlideTitle
                    End If
                Next
            Next
        End If
    End If
End Sub


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
Private Function mppIndexAdd _
( _
) As PowerPoint.Presentation
    Dim preIndex As PowerPoint.Presentation
    Dim lngShapeIndex As Long
    Dim lngLevelIndex As Long
    Dim lngRulerIndex As Long
    Dim lngTabStopIndex As Long
    
    Set preIndex = Application.Presentations.Add
    
    With preIndex.PageSetup
        .SlideOrientation = Office.msoOrientationVertical
        .SlideSize = PowerPoint.ppSlideSizeLetterPaper
        .SlideWidth = 72 * 7.5
        .SlideHeight = 72 * 10
    End With
    
    For lngShapeIndex = preIndex.SlideMaster.Shapes.Count To 1 Step -1
        preIndex.SlideMaster.Shapes(lngShapeIndex).Delete
    Next
    
    preIndex.SlideMaster.Shapes.AddTitle
    With preIndex.SlideMaster.Shapes.Title
        .Left = 0
        .Top = 0
        .Width = preIndex.PageSetup.SlideWidth
        .Height = 72 * 0.5
    End With
    
    With preIndex.SlideMaster.TextStyles(ppTitleStyle)
        .TextFrame.AutoSize = PowerPoint.ppAutoSizeNone
        .TextFrame.MarginLeft = 72 * 0.1
        .TextFrame.MarginTop = 0
        .TextFrame.MarginRight = 72 * 0.1
        .TextFrame.MarginBottom = 0
        .TextFrame.Orientation = Office.msoTextOrientationHorizontal
        For lngTabStopIndex = 1 To .TextFrame.Ruler.TabStops.Count Step 1
            .TextFrame.Ruler.TabStops(lngTabStopIndex).Clear
        Next
        For lngRulerIndex = 1 To .TextFrame.Ruler.Levels.Count Step 1
            .TextFrame.Ruler.Levels(lngRulerIndex).FirstMargin = 0
            .TextFrame.Ruler.Levels(lngRulerIndex).LeftMargin = 0
        Next
        .TextFrame.VerticalAnchor = Office.msoAnchorTop
        .TextFrame.WordWrap = Office.msoTrue
        For lngLevelIndex = 1 To .Levels.Count Step 1
            .Levels(lngLevelIndex).Font.AutoRotateNumbers = Office.msoFalse
            .Levels(lngLevelIndex).Font.BaselineOffset = 0
            .Levels(lngLevelIndex).Font.Bold = Office.msoFalse
            .Levels(lngLevelIndex).Font.Color.RGB = VBA.RGB(0, 0, 0)
            .Levels(lngLevelIndex).Font.Emboss = Office.msoFalse
            .Levels(lngLevelIndex).Font.Italic = Office.msoFalse
            .Levels(lngLevelIndex).Font.Name = "Arial"
            .Levels(lngLevelIndex).Font.Shadow = Office.msoFalse
            .Levels(lngLevelIndex).Font.Size = 12
            .Levels(lngLevelIndex).Font.Subscript = Office.msoFalse
            .Levels(lngLevelIndex).Font.Superscript = Office.msoFalse
            .Levels(lngLevelIndex).Font.Underline = Office.msoFalse
            .Levels(lngLevelIndex).ParagraphFormat.Alignment = PowerPoint.ppAlignCenter
            .Levels(lngLevelIndex).ParagraphFormat.BaseLineAlignment = PowerPoint.ppBaselineAlignTop
            .Levels(lngLevelIndex).ParagraphFormat.Bullet = Office.msoFalse
            .Levels(lngLevelIndex).ParagraphFormat.FarEastLineBreakControl = Office.msoFalse
            .Levels(lngLevelIndex).ParagraphFormat.LineRuleAfter = Office.msoFalse
            .Levels(lngLevelIndex).ParagraphFormat.LineRuleBefore = Office.msoFalse
            .Levels(lngLevelIndex).ParagraphFormat.LineRuleWithin = Office.msoFalse
            .Levels(lngLevelIndex).ParagraphFormat.SpaceAfter = 0
            .Levels(lngLevelIndex).ParagraphFormat.SpaceBefore = 0
            .Levels(lngLevelIndex).ParagraphFormat.SpaceWithin = 14
            .Levels(lngLevelIndex).ParagraphFormat.TextDirection = PowerPoint.ppDirectionLeftToRight
            .Levels(lngLevelIndex).ParagraphFormat.WordWrap = Office.msoTrue
        Next
    End With
    
    With preIndex.SlideMaster.TextStyles(PowerPoint.ppBodyStyle)
        For lngLevelIndex = 1 To .Levels.Count Step 1
            .Levels(lngLevelIndex).Font.AutoRotateNumbers = Office.msoFalse
            .Levels(lngLevelIndex).Font.BaselineOffset = 0
            .Levels(lngLevelIndex).Font.Bold = Office.msoFalse
            .Levels(lngLevelIndex).Font.Color.RGB = VBA.RGB(0, 0, 0)
            .Levels(lngLevelIndex).Font.Emboss = Office.msoFalse
            .Levels(lngLevelIndex).Font.Italic = Office.msoFalse
            .Levels(lngLevelIndex).Font.Name = "Arial"
            .Levels(lngLevelIndex).Font.Shadow = Office.msoFalse
            .Levels(lngLevelIndex).Font.Size = 12
            .Levels(lngLevelIndex).Font.Subscript = Office.msoFalse
            .Levels(lngLevelIndex).Font.Superscript = Office.msoFalse
            .Levels(lngLevelIndex).Font.Underline = Office.msoTrue
            .Levels(lngLevelIndex).ParagraphFormat.Alignment = PowerPoint.ppAlignLeft
            .Levels(lngLevelIndex).ParagraphFormat.BaseLineAlignment = PowerPoint.ppBaselineAlignTop
            .Levels(lngLevelIndex).ParagraphFormat.Bullet = Office.msoFalse
            .Levels(lngLevelIndex).ParagraphFormat.FarEastLineBreakControl = Office.msoFalse
            .Levels(lngLevelIndex).ParagraphFormat.LineRuleAfter = Office.msoFalse
            .Levels(lngLevelIndex).ParagraphFormat.LineRuleBefore = Office.msoFalse
            .Levels(lngLevelIndex).ParagraphFormat.LineRuleWithin = Office.msoFalse
            .Levels(lngLevelIndex).ParagraphFormat.SpaceAfter = 0
            .Levels(lngLevelIndex).ParagraphFormat.SpaceBefore = 0
            .Levels(lngLevelIndex).ParagraphFormat.SpaceWithin = 14
            .Levels(lngLevelIndex).ParagraphFormat.TextDirection = PowerPoint.ppDirectionLeftToRight
            .Levels(lngLevelIndex).ParagraphFormat.WordWrap = Office.msoTrue
        Next
    End With
    
    With preIndex.SlideMaster.TextStyles(PowerPoint.ppDefaultStyle)
        .TextFrame.AutoSize = PowerPoint.ppAutoSizeNone
        .TextFrame.MarginLeft = 72 * 0.1
        .TextFrame.MarginTop = 0
        .TextFrame.MarginRight = 72 * 0.1
        .TextFrame.MarginBottom = 0
        .TextFrame.Orientation = Office.msoTextOrientationHorizontal
        .TextFrame.VerticalAnchor = Office.msoAnchorTop
        .TextFrame.WordWrap = Office.msoTrue
        For lngLevelIndex = 1 To .Levels.Count Step 1
            .Levels(lngLevelIndex).Font.AutoRotateNumbers = Office.msoFalse
            .Levels(lngLevelIndex).Font.BaselineOffset = 0
            .Levels(lngLevelIndex).Font.Bold = Office.msoFalse
            .Levels(lngLevelIndex).Font.Color.RGB = VBA.RGB(0, 0, 0)
            .Levels(lngLevelIndex).Font.Emboss = Office.msoFalse
            .Levels(lngLevelIndex).Font.Italic = Office.msoFalse
            .Levels(lngLevelIndex).Font.Name = "Arial"
            .Levels(lngLevelIndex).Font.Shadow = Office.msoFalse
            .Levels(lngLevelIndex).Font.Size = 12
            .Levels(lngLevelIndex).Font.Subscript = Office.msoFalse
            .Levels(lngLevelIndex).Font.Superscript = Office.msoFalse
            .Levels(lngLevelIndex).Font.Underline = Office.msoTrue
            .Levels(lngLevelIndex).ParagraphFormat.Alignment = PowerPoint.ppAlignLeft
            .Levels(lngLevelIndex).ParagraphFormat.BaseLineAlignment = PowerPoint.ppBaselineAlignTop
            .Levels(lngLevelIndex).ParagraphFormat.Bullet = Office.msoFalse
            .Levels(lngLevelIndex).ParagraphFormat.FarEastLineBreakControl = Office.msoFalse
            .Levels(lngLevelIndex).ParagraphFormat.LineRuleAfter = Office.msoFalse
            .Levels(lngLevelIndex).ParagraphFormat.LineRuleBefore = Office.msoFalse
            .Levels(lngLevelIndex).ParagraphFormat.LineRuleWithin = Office.msoFalse
            .Levels(lngLevelIndex).ParagraphFormat.SpaceAfter = 0
            .Levels(lngLevelIndex).ParagraphFormat.SpaceBefore = 0
            .Levels(lngLevelIndex).ParagraphFormat.SpaceWithin = 14
            .Levels(lngLevelIndex).ParagraphFormat.TextDirection = PowerPoint.ppDirectionLeftToRight
            .Levels(lngLevelIndex).ParagraphFormat.WordWrap = Office.msoTrue
        Next
    End With
    
    Set mppIndexAdd = preIndex
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mppIndexSlideAdd _
( _
    ByRef preCurrent As PowerPoint.Presentation _
) As PowerPoint.Slide
    Dim sldIndex As PowerPoint.Slide
    Dim lngIndex As Long
    Dim shpTitle As PowerPoint.Shape
    
    Set sldIndex = preCurrent.Slides.Add(preCurrent.Slides.Count + 1, PowerPoint.ppLayoutText)
    
    Set shpTitle = sldIndex.Shapes.Title
    
    For lngIndex = sldIndex.Shapes.Count To 1 Step -1
        If (sldIndex.Shapes(lngIndex) Is shpTitle) Then
        Else
            sldIndex.Shapes(lngIndex).Delete
        End If
    Next
    
    Set mppIndexSlideAdd = sldIndex
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mppIndexTitleAdd _
( _
    ByRef sldSlide As PowerPoint.Slide, _
    ByRef strTitle As String _
) As PowerPoint.Shape
    Dim shpTitle As PowerPoint.Shape
    Dim preCurrent As PowerPoint.Presentation
    
    Set preCurrent = sldSlide.Parent
    
    Set shpTitle = sldSlide.Shapes.Title
    
    With shpTitle
        .TextFrame.TextRange.Text = strTitle
        .Fill.Solid
        .Fill.ForeColor.RGB = VBA.RGB(255, 255, 255)
        .Fill.BackColor.RGB = VBA.RGB(255, 255, 255)
    End With
    
    Set mppIndexTitleAdd = shpTitle
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mppIndexCategoryAdd _
( _
    ByRef sldSlide As PowerPoint.Slide, _
    ByRef lngColStart As Long, _
    ByRef lngRowStart As Long, _
    ByRef lngRowCount As Long, _
    ByRef strCategory As String _
) As PowerPoint.Shape
    Dim ppCategoryShape As PowerPoint.Shape
    
    Set ppCategoryShape = mppIndexShapeAdd(sldSlide, lngColStart, lngRowStart, lngRowCount)
        
    With ppCategoryShape
        .TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.ppAlignCenter
        .TextFrame.TextRange.Font.Size = .TextFrame.TextRange.Font.Size + 2
        .TextFrame.TextRange.Text = strCategory
    End With
    
    Set mppIndexCategoryAdd = ppCategoryShape
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mppIndexShapeAdd _
( _
    ByRef sldSlide As PowerPoint.Slide, _
    ByRef lngColStart As Long, _
    ByRef lngRowStart As Long, _
    ByRef lngRowCount As Long _
) As PowerPoint.Shape
    Dim shpIndex As PowerPoint.Shape
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngTabStopIndex As Long
    Dim lngRulerIndex As Long
    
    lngLeft = (lngColStart - 1) * 72 * 2.5
    lngTop = (lngRowStart - 1) * 14
    If (sldSlide.Shapes.HasTitle = Office.msoTrue) Then
        lngTop = (sldSlide.Shapes.Title.Top + sldSlide.Shapes.Title.Height) + lngTop
    End If
    lngWidth = 72 * 2.5
    lngHeight = lngRowCount * 14
        
    Set shpIndex = sldSlide.Shapes.AddTextbox _
                       (Orientation:=Office.msoTextOrientationHorizontal, _
                        Left:=lngLeft, _
                        Top:=lngTop, _
                        Width:=lngWidth, _
                        Height:=lngHeight)
             
    With shpIndex
        .Fill.Solid
        .Fill.ForeColor.RGB = VBA.RGB(255, 255, 255)
        .Fill.BackColor.RGB = VBA.RGB(255, 255, 255)
        .TextFrame.AutoSize = PowerPoint.ppAutoSizeNone
        .TextFrame.MarginLeft = 72 * 0.1
        .TextFrame.MarginTop = 0
        .TextFrame.MarginRight = 72 * 0.1
        .TextFrame.MarginBottom = 0
        .TextFrame.Orientation = Office.msoTextOrientationHorizontal
        For lngTabStopIndex = 1 To .TextFrame.Ruler.TabStops.Count Step 1
            .TextFrame.Ruler.TabStops(lngTabStopIndex).Clear
        Next
        .TextFrame.Ruler.TabStops.Add PowerPoint.ppTabStopLeft, 72 * 0.5
        .TextFrame.Ruler.TabStops.Add PowerPoint.ppTabStopLeft, 72 * 0.75
        .TextFrame.Ruler.TabStops.Add PowerPoint.ppTabStopLeft, 72 * 1
        For lngRulerIndex = 1 To .TextFrame.Ruler.Levels.Count Step 1
            .TextFrame.Ruler.Levels(lngRulerIndex).FirstMargin = 0
            .TextFrame.Ruler.Levels(lngRulerIndex).LeftMargin = 0
        Next
        .TextFrame.VerticalAnchor = Office.msoAnchorTop
        .TextFrame.WordWrap = Office.msoTrue
    End With
             
    Set mppIndexShapeAdd = shpIndex
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mstrIndexTitleClean _
( _
    ByRef strTitle As String _
) As String
    Dim strPage As String
    Dim strPart As String
    
    strTitle = VBA.Replace(strTitle, VBA.Chr(9), " ")
    strTitle = VBA.Replace(strTitle, VBA.Chr(11), " ")
    While (VBA.Right(strTitle, 1) = " ")
        strTitle = VBA.Left(strTitle, VBA.Len(strTitle) - 1)
    Wend
    mstrIndexTitleClean = strTitle
    
    If (VBA.Right(strTitle, 1) <> ")") Then
        Exit Function
    End If
    strTitle = VBA.Left(strTitle, VBA.Len(strTitle) - 1)
    strPage = ""
    While ((VBA.Len(strTitle) > 0) And (VBA.Right(strTitle, 1) <> "("))
        strPage = VBA.Right(strTitle, 1) & strPage
        strTitle = VBA.Left(strTitle, VBA.Len(strTitle) - 1)
    Wend
    If (VBA.Right(strTitle, 1) <> "(") Then
        Exit Function
    End If
    strTitle = VBA.Left(strTitle, VBA.Len(strTitle) - 1)
    While (VBA.Right(strTitle, 1) = " ")
        strTitle = VBA.Left(strTitle, VBA.Len(strTitle) - 1)
    Wend
    While (VBA.Right(strPage, 1) = " ")
        strPage = VBA.Left(strPage, VBA.Len(strPage) - 1)
    Wend
    strPart = ""
    While ((VBA.Len(strPage) > 0) And (VBA.Right(strPage, 1) <> " "))
        strPart = VBA.Right(strPage, 1) & strPart
        strPage = VBA.Left(strPage, VBA.Len(strPage) - 1)
    Wend
    If (IsNumeric(strPart) = False) Then
        Exit Function
    End If
    While (VBA.Right(strPage, 1) = " ")
        strPage = VBA.Left(strPage, VBA.Len(strPage) - 1)
    Wend
    strPart = ""
    While ((VBA.Len(strPage) > 0) And (VBA.Right(strPage, 1) <> " "))
        strPart = VBA.Right(strPage, 1) & strPart
        strPage = VBA.Left(strPage, VBA.Len(strPage) - 1)
    Wend
    If (VBA.LCase(strPart) <> "of") Then
        Exit Function
    End If
    While (VBA.Right(strPage, 1) = " ")
        strPage = VBA.Left(strPage, VBA.Len(strPage) - 1)
    Wend
    strPart = ""
    While ((VBA.Len(strPage) > 0) And (VBA.Right(strPage, 1) <> " "))
        strPart = VBA.Right(strPage, 1) & strPart
        strPage = VBA.Left(strPage, VBA.Len(strPage) - 1)
    Wend
    If (IsNumeric(strPart) = False) Then
        Exit Function
    End If
    While (VBA.Right(strPage, 1) = " ")
        strPage = VBA.Left(strPage, VBA.Len(strPage) - 1)
    Wend
    If (strPage <> "") Then
        Exit Function
    End If
    mstrIndexTitleClean = strTitle
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mlngIndexGroupDetermine _
( _
    ByRef preCurrent As PowerPoint.Presentation, _
    ByRef lngStart As Long _
) As Long
    Dim strCategory As String
    Dim strTitle As String
    Dim strTitleLast As String
    Dim intGroup As Integer
    Dim intGroupLast As Integer
    Dim lngIndex As Long
    Dim lngCount As Long

    If (preCurrent.Slides(lngStart).Shapes.HasTitle = Office.msoTrue) Then
        strTitle = mstrIndexTitleClean(preCurrent.Slides(lngStart).Shapes.Title.TextFrame.TextRange.Text)
    Else
        strTitle = ""
    End If
    If (strTitle <> "") Then
        strTitleLast = strTitle
        intGroup = VBA.Asc(VBA.LCase(VBA.Left(strTitle, 1))) - VBA.Asc("a") + 1
        intGroupLast = intGroup
        strCategory = preCurrent.Slides(lngStart).Tags("Category")
        lngCount = 1
        For lngIndex = lngStart To preCurrent.Slides.Count Step 1
            If (preCurrent.Slides(lngIndex).Tags("Category") = strCategory) Then
                If (preCurrent.Slides(lngIndex).Shapes.HasTitle = Office.msoTrue) Then
                    strTitle = mstrIndexTitleClean(preCurrent.Slides(lngIndex).Shapes.Title.TextFrame.TextRange.Text)
                Else
                    strTitle = ""
                End If
                If ((strTitle <> "") And (strTitle <> strTitleLast)) Then
                    intGroup = VBA.Asc(VBA.LCase(VBA.Left(strTitle, 1))) - VBA.Asc("a") + 1
                    If ((intGroup < 1) Or (intGroup > 26)) Then
                        intGroup = 0
                    End If
                    If (intGroup = intGroupLast) Then
                        lngCount = lngCount + 1
                    Else
                        Exit For
                    End If
                    strTitleLast = strTitle
                End If
            End If
        Next
    Else
        lngCount = 0
    End If

    mlngIndexGroupDetermine = lngCount
End Function
