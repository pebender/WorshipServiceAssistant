Attribute VB_Name = "Index"
'===============================================================================
' Name:
'   WorshipServiceAssistant.Index
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
Public Sub Index_Run(ByVal W As PowerPoint.DocumentWindow)
    Dim P As PowerPoint.Presentation
    Dim i As PowerPoint.Presentation
    Dim CategoryIndex As Long
    Dim IndexTitle As String
    Dim SlideTitle, SlideTitleLast As String
    Dim SlideTitleRemaining As String
    Dim BlockTitleRemaining As String
    Dim ColMax, RowMax As Long
    Dim ColWidth, RowHeight As Long
    Dim ColIndex, RowIndex As Long
    Dim ColCount, RowCount As Long
    Dim ColorState As Boolean
    Dim Index As Long
    Dim S As PowerPoint.Slide
    Dim Block As PowerPoint.Shape
    Dim BlockText As String
    Dim BlockLineChop As Boolean
    Dim BlockLineCount As Long
    Dim BlockBottomLast As Long

    Dim Temp As Long
    
    Dim Category, CategoryList() As String
    
    If (ActiveWindowSlideExists(W) = True) Then
        Set P = W.Presentation
        
        '
        ' Find all the slide categories that are in the category list,
        ' preserving the order of the category list.
        '
        ReDim CategoryList(0)
        For Each Category In Project_Categories
            For Index = 1 To P.slides.Count Step 1
                If (P.slides(Index).Tags("Category") = Category) Then
                    CategoryList(UBound(CategoryList)) = Category
                    ReDim Preserve CategoryList(UBound(CategoryList) + 1)
                    Exit For
                End If
            Next
        Next
        
        '
        ' Find all the slide categories that are not in the category list,
        ' preserving the order of occurance in the slides.
        '
        For Index = 1 To P.slides.Count Step 1
            Category = P.slides(Index).Tags("Category")
            For CategoryIndex = LBound(CategoryList) To UBound(CategoryList) - 1 Step 1
                If (Category = CategoryList(CategoryIndex)) Then
                    Exit For
                End If
            Next
            If (CategoryIndex >= UBound(CategoryList)) Then
                CategoryList(UBound(CategoryList)) = Category
                ReDim Preserve CategoryList(UBound(CategoryList) + 1)
            End If
        Next
        
        ReDim Preserve CategoryList(UBound(CategoryList) - 1)
        
        Set i = AddIndex()
            
        With i.SlideMaster.TextStyles(ppBodyStyle).Levels(1).ParagraphFormat
            RowHeight = .SpaceBefore + .SpaceWithin + .SpaceAfter
        End With
        ColWidth = 72 * 2.5
        ColMax = i.PageSetup.SlideWidth
        ColMax = Int(ColMax / ColWidth)
        RowMax = i.PageSetup.SlideHeight
        If (i.SlideMaster.Shapes.HasTitle = msoTrue) Then
            RowMax = RowMax - i.SlideMaster.Shapes.Title.Height
        End If
        RowMax = Int(RowMax / RowHeight)
        
        IndexTitle = P.Name
        If (LCase(Right(IndexTitle, 4)) = ".ppt") Then
            IndexTitle = Left(IndexTitle, Len(IndexTitle) - 4)
        End If
        IndexTitle = _
            IndexTitle & " Index" & _
            Chr(11) & _
            "(Generated on " & Date & " at " & Time & ")"
        
        ColIndex = ColMax + 1
        For Each Category In CategoryList
            RowIndex = RowMax + 1
            ColorState = False
            SlideTitleLast = ""
            SlideTitleRemaining = 0
            BlockTitleRemaining = 0
            If (i.SlideMaster.Shapes.HasTitle = msoTrue) Then
                BlockBottomLast = i.SlideMaster.Shapes.Title.Height
            Else
                BlockBottomLast = 0
            End If
            SlideTitleLast = ""
            For Index = 1 To P.slides.Count Step 1
                If (P.slides(Index).Tags("Category") = Category) Then
                    '
                    ' Get the slide's title
                    '
                    If (P.slides(Index).Shapes.HasTitle = msoTrue) Then
                        SlideTitle = CleanIndexTitle(P.slides(Index).Shapes.Title.TextFrame.TextRange.Text)
                    Else
                        SlideTitle = ""
                    End If
                    
                    If ((SlideTitle <> "") And (SlideTitle <> SlideTitleLast)) Then
                        If (SlideTitleRemaining <= 0) Then
                            SlideTitleRemaining = DetermineIndexGroup(P, Index)
                            ColorState = Not ColorState
                            BlockTitleRemaining = 0
                            If (RowIndex + SlideTitleRemaining > RowMax) Then
                                RowIndex = RowMax + 1
                                Temp = SlideTitleRemaining
                                Temp = Temp - (Temp Mod RowMax)
                                Temp = Temp / RowMax
                                If (Temp <= ColMax) Then
                                    If (ColIndex + Temp > ColMax) Then
                                        ColIndex = ColMax + 1
                                    End If
                                End If
                            End If
                        End If
                        
                        If (RowIndex > RowMax) Then
                            RowIndex = 1
                            ColIndex = ColIndex + 1
                            If (ColIndex > ColMax) Then
                                RowIndex = 1
                                ColIndex = 1
                                Set S = AddIndexSlide(i)
                                AddIndexTitle S, IndexTitle
                            End If
                            Set Block = AddIndexCategory(S, ColIndex, RowIndex, 2, Category)
                            BlockBottomLast = Block.Top + Block.Height
                            Block.Fill.ForeColor.RGB = RGB(255, 255, 255)
                            RowIndex = RowIndex + 2
                            BlockTitleRemaining = 0
                            If (i.SlideMaster.Shapes.HasTitle = msoTrue) Then
                                BlockBottomLast = _
                                    i.SlideMaster.Shapes.Title.Top + _
                                    i.SlideMaster.Shapes.Title.Height
                            Else
                                BlockBottomLast = 0
                            End If
                        End If
                        
                        If (BlockTitleRemaining <= 0) Then
                            Temp = RowMax - Int(BlockBottomLast / RowHeight) + 1
                            If (SlideTitleRemaining <= Temp) Then
                                Set Block = AddIndexBlock(S, ColIndex, RowIndex, SlideTitleRemaining)
                                BlockTitleRemaining = SlideTitleRemaining
                            Else
                                Set Block = AddIndexBlock(S, ColIndex, RowIndex, Temp)
                                BlockTitleRemaining = Temp
                            End If
                            BlockBottomLast = Block.Top + Block.Height
                            If (ColorState = False) Then
                                Block.Fill.ForeColor.RGB = RGB(220, 220, 220)
                            Else
                                Block.Fill.ForeColor.RGB = RGB(255, 255, 255)
                            End If
                        End If
                        
                        If (Block.TextFrame.HasText = msoFalse) Then
                            BlockText = ""
                            BlockLineCount = 1
                        Else
                            BlockText = Block.TextFrame.TextRange.Text & Chr(13) & Chr(10)
                            BlockLineCount = Block.TextFrame.TextRange.Lines.Count + 1
                        End If
                        BlockText = BlockText & Str(Index) & "." & Chr(9) & SlideTitle
                        Block.TextFrame.TextRange.Text = BlockText
                        With Block.TextFrame.TextRange
                            BlockLineChop = False
                            While (.Lines.Count > BlockLineCount)
                                If (BlockLineChop = True) Then
                                    .Text = Left(.Text, Len(.Text) - 4)
                                End If
                                While (Right(.Text, 1) <> " ")
                                    .Text = Left(.Text, Len(.Text) - 1)
                                Wend
                                .Text = Left(.Text, Len(.Text) - 1)
                                .Text = .Text & " ..."
                                BlockLineChop = True
                            Wend
                        End With
                        
                        RowIndex = RowIndex + 1
                        BlockTitleRemaining = BlockTitleRemaining - 1
                        SlideTitleRemaining = SlideTitleRemaining - 1
                    End If
                    SlideTitleLast = SlideTitle
                End If
            Next
        Next
    End If
End Sub


'===============================================================================
' Private Subroutines and Functions.
'===============================================================================

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Function AddIndex() As PowerPoint.Presentation
    Dim P As PowerPoint.Presentation
    Dim ShapeIndex As Long
    Dim StyleIndex As Long
    Dim LevelIndex As Long
    Dim RulerIndex As Long
    Dim TabStopIndex As Long
    
    Set P = Application.Presentations.Add
    
    With P.PageSetup
        .SlideOrientation = msoOrientationVertical
        .SlideSize = ppSlideSizeLetterPaper
        .SlideWidth = 72 * 7.5
        .SlideHeight = 72 * 10
    End With
    
    For ShapeIndex = P.SlideMaster.Shapes.Count To 1 Step -1
        P.SlideMaster.Shapes(ShapeIndex).Delete
    Next
    
    P.SlideMaster.Shapes.AddTitle
    With P.SlideMaster.Shapes.Title
        .Left = 0
        .Top = 0
        .Width = P.PageSetup.SlideWidth
        .Height = 72 * 0.5
    End With
    
    With P.SlideMaster.TextStyles(ppTitleStyle)
        .TextFrame.AutoSize = ppAutoSizeNone
        .TextFrame.MarginLeft = 72 * 0.1
        .TextFrame.MarginTop = 0
        .TextFrame.MarginRight = 72 * 0.1
        .TextFrame.MarginBottom = 0
        .TextFrame.Orientation = msoTextOrientationHorizontal
        For TabStopIndex = 1 To .TextFrame.Ruler.TabStops.Count Step 1
            .TextFrame.Ruler.TabStops(TabStopIndex).Clear
        Next
        For RulerIndex = 1 To .TextFrame.Ruler.Levels.Count Step 1
            .TextFrame.Ruler.Levels(RulerIndex).FirstMargin = 0
            .TextFrame.Ruler.Levels(RulerIndex).LeftMargin = 0
        Next
        .TextFrame.VerticalAnchor = msoAnchorTop
        .TextFrame.WordWrap = msoTrue
        For LevelIndex = 1 To .Levels.Count Step 1
            .Levels(LevelIndex).Font.AutoRotateNumbers = msoFalse
            .Levels(LevelIndex).Font.BaselineOffset = 0
            .Levels(LevelIndex).Font.Bold = msoFalse
            .Levels(LevelIndex).Font.Color.RGB = RGB(0, 0, 0)
            .Levels(LevelIndex).Font.Emboss = msoFalse
            .Levels(LevelIndex).Font.Italic = msoFalse
            .Levels(LevelIndex).Font.Name = "Arial"
            .Levels(LevelIndex).Font.Shadow = msoFalse
            .Levels(LevelIndex).Font.Size = 12
            .Levels(LevelIndex).Font.Subscript = msoFalse
            .Levels(LevelIndex).Font.Superscript = msoFalse
            .Levels(LevelIndex).Font.Underline = msoFalse
            .Levels(LevelIndex).ParagraphFormat.Alignment = ppAlignCenter
            .Levels(LevelIndex).ParagraphFormat.BaseLineAlignment = ppBaselineAlignTop
            .Levels(LevelIndex).ParagraphFormat.Bullet = msoFalse
            .Levels(LevelIndex).ParagraphFormat.FarEastLineBreakControl = msoFalse
            .Levels(LevelIndex).ParagraphFormat.LineRuleAfter = msoFalse
            .Levels(LevelIndex).ParagraphFormat.LineRuleBefore = msoFalse
            .Levels(LevelIndex).ParagraphFormat.LineRuleWithin = msoFalse
            .Levels(LevelIndex).ParagraphFormat.SpaceAfter = 0
            .Levels(LevelIndex).ParagraphFormat.SpaceBefore = 0
            .Levels(LevelIndex).ParagraphFormat.SpaceWithin = 14
            .Levels(LevelIndex).ParagraphFormat.TextDirection = ppDirectionLeftToRight
            .Levels(LevelIndex).ParagraphFormat.WordWrap = msoTrue
        Next
    End With
    
    With P.SlideMaster.TextStyles(ppBodyStyle)
        For LevelIndex = 1 To .Levels.Count Step 1
            .Levels(LevelIndex).Font.AutoRotateNumbers = msoFalse
            .Levels(LevelIndex).Font.BaselineOffset = 0
            .Levels(LevelIndex).Font.Bold = msoFalse
            .Levels(LevelIndex).Font.Color.RGB = RGB(0, 0, 0)
            .Levels(LevelIndex).Font.Emboss = msoFalse
            .Levels(LevelIndex).Font.Italic = msoFalse
            .Levels(LevelIndex).Font.Name = "Arial"
            .Levels(LevelIndex).Font.Shadow = msoFalse
            .Levels(LevelIndex).Font.Size = 12
            .Levels(LevelIndex).Font.Subscript = msoFalse
            .Levels(LevelIndex).Font.Superscript = msoFalse
            .Levels(LevelIndex).Font.Underline = msoTrue
            .Levels(LevelIndex).ParagraphFormat.Alignment = ppAlignLeft
            .Levels(LevelIndex).ParagraphFormat.BaseLineAlignment = ppBaselineAlignTop
            .Levels(LevelIndex).ParagraphFormat.Bullet = msoFalse
            .Levels(LevelIndex).ParagraphFormat.FarEastLineBreakControl = msoFalse
            .Levels(LevelIndex).ParagraphFormat.LineRuleAfter = msoFalse
            .Levels(LevelIndex).ParagraphFormat.LineRuleBefore = msoFalse
            .Levels(LevelIndex).ParagraphFormat.LineRuleWithin = msoFalse
            .Levels(LevelIndex).ParagraphFormat.SpaceAfter = 0
            .Levels(LevelIndex).ParagraphFormat.SpaceBefore = 0
            .Levels(LevelIndex).ParagraphFormat.SpaceWithin = 14
            .Levels(LevelIndex).ParagraphFormat.TextDirection = ppDirectionLeftToRight
            .Levels(LevelIndex).ParagraphFormat.WordWrap = msoTrue
        Next
    End With
    
    With P.SlideMaster.TextStyles(ppDefaultStyle)
        .TextFrame.AutoSize = ppAutoSizeNone
        .TextFrame.MarginLeft = 72 * 0.1
        .TextFrame.MarginTop = 0
        .TextFrame.MarginRight = 72 * 0.1
        .TextFrame.MarginBottom = 0
        .TextFrame.Orientation = msoTextOrientationHorizontal
        .TextFrame.VerticalAnchor = msoAnchorTop
        .TextFrame.WordWrap = msoTrue
        For LevelIndex = 1 To .Levels.Count Step 1
            .Levels(LevelIndex).Font.AutoRotateNumbers = msoFalse
            .Levels(LevelIndex).Font.BaselineOffset = 0
            .Levels(LevelIndex).Font.Bold = msoFalse
            .Levels(LevelIndex).Font.Color.RGB = RGB(0, 0, 0)
            .Levels(LevelIndex).Font.Emboss = msoFalse
            .Levels(LevelIndex).Font.Italic = msoFalse
            .Levels(LevelIndex).Font.Name = "Arial"
            .Levels(LevelIndex).Font.Shadow = msoFalse
            .Levels(LevelIndex).Font.Size = 12
            .Levels(LevelIndex).Font.Subscript = msoFalse
            .Levels(LevelIndex).Font.Superscript = msoFalse
            .Levels(LevelIndex).Font.Underline = msoTrue
            .Levels(LevelIndex).ParagraphFormat.Alignment = ppAlignLeft
            .Levels(LevelIndex).ParagraphFormat.BaseLineAlignment = ppBaselineAlignTop
            .Levels(LevelIndex).ParagraphFormat.Bullet = msoFalse
            .Levels(LevelIndex).ParagraphFormat.FarEastLineBreakControl = msoFalse
            .Levels(LevelIndex).ParagraphFormat.LineRuleAfter = msoFalse
            .Levels(LevelIndex).ParagraphFormat.LineRuleBefore = msoFalse
            .Levels(LevelIndex).ParagraphFormat.LineRuleWithin = msoFalse
            .Levels(LevelIndex).ParagraphFormat.SpaceAfter = 0
            .Levels(LevelIndex).ParagraphFormat.SpaceBefore = 0
            .Levels(LevelIndex).ParagraphFormat.SpaceWithin = 14
            .Levels(LevelIndex).ParagraphFormat.TextDirection = ppDirectionLeftToRight
            .Levels(LevelIndex).ParagraphFormat.WordWrap = msoTrue
        Next
    End With
    
    Set AddIndex = P
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Function AddIndexSlide(ByVal P As PowerPoint.Presentation) As PowerPoint.Slide
    Dim S As PowerPoint.Slide
    Dim Index As Long
    Dim TitleShape As PowerPoint.Shape
    
    Set S = P.slides.Add(P.slides.Count + 1, ppLayoutText)
    
    Set TitleShape = S.Shapes.Title
    
    For Index = S.Shapes.Count To 1 Step -1
        If (S.Shapes(Index) Is TitleShape) Then
        Else
            S.Shapes(Index).Delete
        End If
    Next
    
    Set AddIndexSlide = S
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Function AddIndexTitle(ByVal S As PowerPoint.Slide, ByVal Title As String) As PowerPoint.Shape
    Dim T As PowerPoint.Shape
    Dim P As PowerPoint.Presentation
    
    Set P = S.Parent
    
    Set T = S.Shapes.Title
    
    With T
        .TextFrame.TextRange.Text = Title
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Fill.BackColor.RGB = RGB(255, 255, 255)
    End With
    
    Set AddIndexTitle = T
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Function AddIndexCategory _
( _
    ByVal S As PowerPoint.Slide, _
    ByVal ColStart As Long, _
    ByVal RowStart As Long, _
    ByVal RowCount As Long, _
    ByVal Category As String _
) As Shape
    Dim B As PowerPoint.Shape
    Dim P As PowerPoint.Presentation
    Dim CategoryEnabled As Boolean
    
    Set B = AddIndexBlock(S, ColStart, RowStart, RowCount)
        
    With B
        .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
        .TextFrame.TextRange.Font.Size = .TextFrame.TextRange.Font.Size + 2
        .TextFrame.TextRange.Text = Category
    End With
    
    Set AddIndexCategory = B
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Function AddIndexBlock _
( _
    ByVal S As PowerPoint.Slide, _
    ByVal ColStart As Long, _
    ByVal RowStart As Long, _
    ByVal RowCount As Long _
) As Shape
    Dim B As PowerPoint.Shape
    Dim Left, Top, Width, Height As Long
    Dim TabStopIndex As Long
    Dim RulerIndex As Long
    
    Left = (ColStart - 1) * 72 * 2.5
    Top = (RowStart - 1) * 14
    If (S.Shapes.HasTitle = msoTrue) Then
        Top = (S.Shapes.Title.Top + S.Shapes.Title.Height) + Top
    End If
    Width = 72 * 2.5
    Height = RowCount * 14
        
    Set B = S.Shapes.AddTextbox _
            (Orientation:=msoTextOrientationHorizontal, _
             Left:=Left, _
             Top:=Top, _
             Width:=Width, _
             Height:=Height)
             
    With B
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Fill.BackColor.RGB = RGB(255, 255, 255)
        .TextFrame.AutoSize = ppAutoSizeNone
        .TextFrame.MarginLeft = 72 * 0.1
        .TextFrame.MarginTop = 0
        .TextFrame.MarginRight = 72 * 0.1
        .TextFrame.MarginBottom = 0
        .TextFrame.Orientation = msoTextOrientationHorizontal
        For TabStopIndex = 1 To .TextFrame.Ruler.TabStops.Count Step 1
            .TextFrame.Ruler.TabStops(TabStopIndex).Clear
        Next
        .TextFrame.Ruler.TabStops.Add ppTabStopLeft, 72 * 0.5
        .TextFrame.Ruler.TabStops.Add ppTabStopLeft, 72 * 0.75
        .TextFrame.Ruler.TabStops.Add ppTabStopLeft, 72 * 1
        For RulerIndex = 1 To .TextFrame.Ruler.Levels.Count Step 1
            .TextFrame.Ruler.Levels(RulerIndex).FirstMargin = 0
            .TextFrame.Ruler.Levels(RulerIndex).LeftMargin = 0
        Next
        .TextFrame.VerticalAnchor = msoAnchorTop
        .TextFrame.WordWrap = msoTrue
    End With
             
    Set AddIndexBlock = B
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Function CleanIndexTitle(ByVal Title As String) As String
    Dim Page As String
    Dim Part As String
    
    Title = Replace(Title, Chr(9), " ")
    Title = Replace(Title, Chr(11), " ")
    While (Right(Title, 1) = " ")
        Title = Left(Title, Len(Title) - 1)
    Wend
    CleanIndexTitle = Title
    
    If (Right(Title, 1) <> ")") Then
        Exit Function
    End If
    Title = Left(Title, Len(Title) - 1)
    Page = ""
    While ((Len(Title) > 0) And (Right(Title, 1) <> "("))
        Page = Right(Title, 1) & Page
        Title = Left(Title, Len(Title) - 1)
    Wend
    If (Right(Title, 1) <> "(") Then
        Exit Function
    End If
    Title = Left(Title, Len(Title) - 1)
    While (Right(Title, 1) = " ")
        Title = Left(Title, Len(Title) - 1)
    Wend
    While (Right(Page, 1) = " ")
        Page = Left(Page, Len(Page) - 1)
    Wend
    Part = ""
    While ((Len(Page) > 0) And (Right(Page, 1) <> " "))
        Part = Right(Page, 1) & Part
        Page = Left(Page, Len(Page) - 1)
    Wend
    If (IsNumeric(Part) = False) Then
        Exit Function
    End If
    While (Right(Page, 1) = " ")
        Page = Left(Page, Len(Page) - 1)
    Wend
    Part = ""
    While ((Len(Page) > 0) And (Right(Page, 1) <> " "))
        Part = Right(Page, 1) & Part
        Page = Left(Page, Len(Page) - 1)
    Wend
    If (LCase(Part) <> "of") Then
        Exit Function
    End If
    While (Right(Page, 1) = " ")
        Page = Left(Page, Len(Page) - 1)
    Wend
    Part = ""
    While ((Len(Page) > 0) And (Right(Page, 1) <> " "))
        Part = Right(Page, 1) & Part
        Page = Left(Page, Len(Page) - 1)
    Wend
    If (IsNumeric(Part) = False) Then
        Exit Function
    End If
    While (Right(Page, 1) = " ")
        Page = Left(Page, Len(Page) - 1)
    Wend
    If (Page <> "") Then
        Exit Function
    End If
    CleanIndexTitle = Title
End Function

'-------------------------------------------------------------------------------
' Description:
'-------------------------------------------------------------------------------
Private Function DetermineIndexGroup _
( _
    ByVal P As PowerPoint.Presentation, _
    ByVal Start As Long _
) As Long
Dim Category As String
Dim Title, TitleLast As String
Dim Group, GroupLast As String
Dim Index As Long
Dim Count As Long

    If (P.slides(Start).Shapes.HasTitle = msoTrue) Then
        Title = CleanIndexTitle(P.slides(Start).Shapes.Title.TextFrame.TextRange.Text)
    Else
        Title = ""
    End If
    If (Title <> "") Then
        TitleLast = Title
        Group = Asc(LCase(Left(Title, 1))) - Asc("a") + 1
        GroupLast = Group
        Category = P.slides(Start).Tags("Category")
        Count = 1
        For Index = Start To P.slides.Count Step 1
            If (P.slides(Index).Tags("Category") = Category) Then
                If (P.slides(Index).Shapes.HasTitle = msoTrue) Then
                    Title = CleanIndexTitle(P.slides(Index).Shapes.Title.TextFrame.TextRange.Text)
                Else
                    Title = ""
                End If
                If ((Title <> "") And (Title <> TitleLast)) Then
                    Group = Asc(LCase(Left(Title, 1))) - Asc("a") + 1
                    If ((Group < 1) Or (Group > 26)) Then
                        Group = 0
                    End If
                    If (Group = GroupLast) Then
                        Count = Count + 1
                    Else
                        Exit For
                    End If
                    TitleLast = Title
                End If
            End If
        Next
    Else
        Count = 0
    End If

    DetermineIndexGroup = Count
End Function
