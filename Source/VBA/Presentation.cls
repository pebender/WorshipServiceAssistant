VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Presentation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Name:
'   WorshipServiceAssistant.Presentation
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
' Public Type Declarations.
'===============================================================================


'===============================================================================
' Private Type Declarations.
'===============================================================================

Private Type Area
    Top As Long
    Left As Long
    Height As Long
    Width As Long
    RowHeight As Long
    ColumnWidth As Long
    RowCount As Long
    ColumnCount As Long
End Type


'===============================================================================
' Public Variables.
'===============================================================================


'===============================================================================
' Private Variables.
'===============================================================================
Private mblnInvariant As Boolean
Private mudtApplication As WorshipServiceAssistant.Application
Private mudtParent As WorshipServiceAssistant.Application
Private mpreDoppelganger As PowerPoint.Presentation
Private mudtSlides As New WorshipServiceAssistant.Slides
Private mudtWindows As New WorshipServiceAssistant.DocumentWindows
Private mudtSlideShowWindow As New WorshipServiceAssistant.SlideShowWindow


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
    ByRef udtParent As WorshipServiceAssistant.Application, _
    ByRef preDoppelganger As PowerPoint.Presentation _
)
    mblnInvariant = False
    
    Set mudtApplication = udtApplication
    Set mudtParent = udtParent
    Set mpreDoppelganger = preDoppelganger
    
    mudtSlides.Initialize udtApplication, Me, mpreDoppelganger.Slides
    mudtWindows.Initialize udtApplication, Me, mpreDoppelganger.Windows
    mudtSlideShowWindow.Initialize udtApplication, Me
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
    mudtSlideShowWindow.Terminate
    mudtWindows.Terminate
    mudtSlides.Terminate
    
    Set mpreDoppelganger = Nothing
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
            mblnInvariant = blnInvariant
        End If
        mudtSlides.Invariant = blnInvariant
        mudtWindows.Invariant = blnInvariant
        mudtSlideShowWindow.Invariant = blnInvariant
        mblnInvariant = blnInvariant
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
) As PowerPoint.Presentation
    Set Doppelganger = mpreDoppelganger
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
Public Property Get ActiveWindow _
( _
) As WorshipServiceAssistant.DocumentWindow
    Dim lng As Long
    
    Set ActiveWindow = mudtWindows.Active
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get ActiveSlide _
( _
) As WorshipServiceAssistant.Slide
    Dim udtActiveWindow As WorshipServiceAssistant.DocumentWindow
    Dim lng As Long
    
    Set udtActiveWindow = Me.ActiveWindow
    If (udtActiveWindow Is Nothing) Then
        Set ActiveSlide = Nothing
    Else
        Set ActiveSlide = udtActiveWindow.ActiveSlide
    End If
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get ActiveSlideSelection _
( _
) As WorshipServiceAssistant.SlideRange
    Dim udtActiveWindow As WorshipServiceAssistant.DocumentWindow
    Dim lng As Long
    
    Set udtActiveWindow = Me.ActiveWindow
    If (udtActiveWindow Is Nothing) Then
        Set ActiveSlideSelection = New WorshipServiceAssistant.SlideRange
    Else
        Set ActiveSlideSelection = udtActiveWindow.ActiveSlideSelection
    End If
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Windows _
( _
) As WorshipServiceAssistant.DocumentWindows
    Set Windows = mudtWindows
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Slides _
( _
) As WorshipServiceAssistant.Slides
    Set Slides = mudtSlides
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get SlideShowWindow _
( _
) As WorshipServiceAssistant.SlideShowWindow
    Set SlideShowWindow = mudtSlideShowWindow
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Name _
( _
) As String
    Name = mpreDoppelganger.Name
End Property

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Public Property Get Active _
( _
) As Office.MsoTriState
    If (mudtWindows.Active Is Nothing) Then
        Active = Office.msoFalse
    Else
        Active = Office.msoTrue
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
    With mudtWindows
        If (.Count > 0) Then
            .Item(1).Activate
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
Public Sub CreateIndex _
( _
)
    mIndexCreate
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
Private Sub mIndexCreate _
( _
)
    Dim preIndex As PowerPoint.Presentation
    Dim sldPage As PowerPoint.Slide
    Dim shpTitle As PowerPoint.Shape
    Dim shpCategory As PowerPoint.Shape
    Dim shpBlock As PowerPoint.Shape
    Dim udtPageArea As Area
    Dim udtTitleArea As Area
    Dim udtCategoryArea As Area
    Dim udtBlockArea As Area
    Dim strTitleText As String
    Dim strCategoryText As String
    Dim strBlockText As String
    
    Dim cstrCategories As VBA.Collection
    Dim cobjCategories As VBA.Collection
    Dim cobjLetters As VBA.Collection
    Dim cobjPages As VBA.Collection
    Dim cobjColumns As VBA.Collection
    Dim cobjBlocks As VBA.Collection
    Dim cobjSongs As VBA.Collection
    Dim udtSlide As WorshipServiceAssistant.Slide
    
    Dim strCategory As String
    Dim strLetter As String
    
    Dim lngCategory As Long
    Dim lngLetter As Long
    Dim lngPage As Long
    Dim lngColumn As Long
    Dim lngBlock As Long
    Dim lngSong As Long
    Dim lngSlide As Long
    Dim lngRow As Long
    
    Dim strSongCategory As String
    Dim strSongLetter As String
    Dim strSongTitle As String
    Dim strSongTitleLast As String
    Dim strSongPage As String
    Dim strSongIndex As String
    
    Dim blnBlockColorState As Boolean
    
    ' Add PowerPoint presentation that will contain the index.
    Set preIndex = mpreIndexAdd
    
    ' Determine the Area of the different elements on an
    ' index page.
    With udtPageArea
        .RowHeight = _
            preIndex.SlideMaster.TextStyles(PowerPoint.ppDefaultStyle).Levels(1).ParagraphFormat.SpaceBefore + _
            preIndex.SlideMaster.TextStyles(PowerPoint.ppDefaultStyle).Levels(1).ParagraphFormat.SpaceWithin + _
            preIndex.SlideMaster.TextStyles(PowerPoint.ppDefaultStyle).Levels(1).ParagraphFormat.SpaceAfter
        .ColumnWidth = 72 * 2.5
        .Top = 0
        .Left = 0
        .Height = preIndex.PageSetup.SlideHeight
        .Width = preIndex.PageSetup.SlideWidth
        .RowCount = VBA.Fix(VBA.CDbl(.Height) / VBA.CDbl(.RowHeight))
        .ColumnCount = VBA.Fix(VBA.CDbl(.Width) / VBA.CDbl(.ColumnWidth))
    End With
    udtTitleArea = udtPageArea
    With udtTitleArea
        .Height = 3 * udtPageArea.RowHeight
        .Width = 3 * udtPageArea.ColumnWidth
        .RowCount = 3
        .ColumnCount = 3
    End With
    udtCategoryArea = udtPageArea
    With udtCategoryArea
        .Top = udtTitleArea.Top + udtTitleArea.Height
        .Height = 2 * udtPageArea.RowHeight
        .Left = udtPageArea.Left
        .Width = udtPageArea.ColumnWidth
        .RowCount = 2
        .ColumnCount = 1
    End With
    udtBlockArea = udtPageArea
    With udtBlockArea
        .Top = udtCategoryArea.Top + udtCategoryArea.Height
        .Height = udtPageArea.Height - .Top
        .Left = udtPageArea.Left
        .Width = udtPageArea.ColumnWidth
        .RowCount = VBA.Fix(VBA.CDbl(.Height) / VBA.CDbl(.RowHeight))
        .ColumnCount = 1
    End With
    
    ' Determine the category name list, by starting with the
    ' default category list, and adding any categories that are in
    ' the slides and are not in the default category list.
    Set cstrCategories = New VBA.Collection
    With mudtApplication.Options.CategoryList
        For lngCategory = 1 To .Count Step 1
            cstrCategories.Add .Item(lngCategory)
        Next
    End With
    With Me.Slides
        For lngSlide = 1 To .Count Step 1
            strCategory = .Item(lngSlide).Category
            For lngCategory = 1 To cstrCategories.Count Step 1
                If (strCategory = cstrCategories.Item(lngCategory)) Then
                    Exit For
                End If
            Next
            If (lngCategory > cstrCategories.Count) Then
                cstrCategories.Add strCategory
            End If
        Next
    End With
    
    ' Arrange the songs into a category->letter->song hierarchy,
    ' where the category is the category of the song, and letter
    ' is the letter of the alphabet with which the song title starts.
    Set cobjCategories = New VBA.Collection
    For lngCategory = 1 To cstrCategories.Count Step 1
    ' Create the category->letter hierarchy
        strCategory = cstrCategories.Item(lngCategory)
        cobjCategories.Add New VBA.Collection, strCategory
        Set cobjLetters = cobjCategories.Item(strCategory)
        cobjLetters.Add New VBA.Collection, "<none>"
        For lngLetter = VBA.Asc("a") To VBA.Asc("z") Step 1
            Set cobjLetters = cobjCategories.Item(strCategory)
            cobjLetters.Add New VBA.Collection, VBA.Chr(lngLetter)
        Next
    Next
    strSongTitleLast = ""
    For lngSlide = 1 To Me.Slides.Count Step 1
    ' Add the songs to the category->letter hierarchy.
        Set udtSlide = Me.Slides.Item(lngSlide)
        strSongTitle = mstrIndexSongTitleClean(udtSlide.Title)
        If ((strSongTitle <> "") And (strSongTitle <> strSongTitleLast)) Then
            strSongCategory = udtSlide.Category
            strSongLetter = VBA.LCase(VBA.Left(strSongTitle, 1))
            If ((strSongLetter < "a") And (strSongLetter > "z")) Then
                strSongLetter = "<none>"
            End If
            Set cobjLetters = cobjCategories.Item(strSongCategory)
            Set cobjSongs = cobjLetters.Item(strSongLetter)
            cobjSongs.Add udtSlide
        End If
        strSongTitleLast = strSongTitle
    Next
    
    ' Destroy the now unneeded catetory name list.
    Set cstrCategories = Nothing
    
    ' Arrange the songs into a column->block->song hierarchy,
    ' where column is an index page column and block is a block
    ' of songs that have the same category and letter and that
    ' fit in the same column.
    Set cobjColumns = New VBA.Collection
    ' A new column start is forced.
    lngRow = udtBlockArea.RowCount + 1
    For lngCategory = 1 To cobjCategories.Count Step 1
        Set cobjLetters = cobjCategories.Item(lngCategory)
        ' When a new category is started, a new column start is forced.
        lngRow = udtBlockArea.RowCount + 1
        For lngLetter = 1 To cobjLetters.Count Step 1
            Set cobjSongs = cobjLetters.Item(lngLetter)
           If ((lngRow + cobjSongs.Count - 1) > udtBlockArea.RowCount) Then
            ' All the songs for a letter of the alphabet cannot fit in the
            ' remaining rows of the column. Therefore, a new column start
            ' is forced. If old column was empty, then there will be an empty
            ' column which step (2) will remove.
                lngRow = udtBlockArea.RowCount + 1
            End If
            If (Not lngRow > udtBlockArea.RowCount) Then
            ' While the new letter will not start a new column,
            ' a new letter always starts a new block.
                cobjBlocks.Add New VBA.Collection
            End If
            For lngSong = 1 To cobjSongs.Count Step 1
                Set udtSlide = cobjSongs.Item(lngSong)
                If (lngRow > udtBlockArea.RowCount) Then
                ' The column is full, so start a new column and a new block.
                    If (lngSong > 1) Then
                    ' The new column is a continuation of an existing letter.
                        ' Insert a dummy block so that the color for the letter
                        ' will stay the same (because it will toggle twice).
                        cobjBlocks.Add Nothing
                    End If
                    With cobjColumns
                        .Add New VBA.Collection
                        Set cobjBlocks = .Item(.Count)
                        lngRow = 1
                    End With
                    cobjBlocks.Add New VBA.Collection
                End If
                With cobjBlocks
                    .Item(.Count).Add udtSlide
                End With
                lngRow = lngRow + 1
            Next
        Next
    Next
    
    ' Destroy the now unneeded category->letter->song hierarchy.
    Set cobjCategories = Nothing
    
    ' Remove empty blocks and empty columns from column->block->song
    ' hierarchy.
    For lngColumn = cobjColumns.Count To 1 Step -1
        Set cobjBlocks = cobjColumns.Item(lngColumn)
        For lngBlock = cobjBlocks.Count To 1 Step -1
            Set cobjSongs = cobjBlocks.Item(lngBlock)
            If (Not cobjSongs Is Nothing) Then
            ' This is not a dummy block.
                If (cobjSongs.Count = 0) Then
                ' There are no songs in this block,
                ' so remove the block.
                    cobjBlocks.Remove lngBlock
                End If
            End If
        Next
        If (cobjBlocks.Count = 0) Then
        ' There are no blocks in this column,
        ' so remove the column.
            cobjColumns.Remove lngColumn
        End If
    Next
    
    ' Arrange the songs into a page->column->block->song hierarchy.
    Set cobjPages = New VBA.Collection
    For lngColumn = 1 To cobjColumns.Count Step 1
        Set cobjBlocks = cobjColumns.Item(lngColumn)
        If (lngColumn > (cobjPages.Count * udtPageArea.ColumnCount)) Then
        ' All the columns on the page have been filled,
        ' so add another page.
            cobjPages.Add New VBA.Collection
        End If
        With cobjPages
            ' Add the column to the page.
            .Item(.Count).Add cobjBlocks
        End With
    Next
    
    ' Destroy the now unneeded column->block->song hierarchy.
    Set cobjColumns = Nothing
    
    strTitleText = Me.Name
    If (VBA.LCase(VBA.Right(strTitleText, 4)) = ".ppt") Then
        strTitleText = VBA.Left(strTitleText, VBA.Len(strTitleText) - 4)
    End If
    strTitleText = _
        strTitleText & " Index" & _
        VBA.vbNewLine & _
        "(Generated on " & Date & " at " & Time & ")"

    ' Generate index pages using the page->column->block->song hierarchy.
    blnBlockColorState = False
    For lngPage = 1 To cobjPages.Count Step 1
        Set cobjColumns = cobjPages.Item(lngPage)
        Set sldPage = msldIndexPageAdd(preIndex)
        ' Add the title to the page.
        mshpIndexTitleAdd sldPage, udtTitleArea, strTitleText
        For lngColumn = 1 To cobjColumns.Count Step 1
            Set cobjBlocks = cobjColumns.Item(lngColumn)
            lngRow = 1
            ' Add the category header to the column.
            Set cobjSongs = cobjBlocks.Item(1)
            Set udtSlide = cobjSongs.Item(1)
            strCategoryText = udtSlide.Category
            Set shpCategory = mshpIndexCategoryAdd _
                (sldPage, _
                 udtCategoryArea, _
                 lngColumn - 1, _
                 strCategoryText)
            For lngBlock = 1 To cobjBlocks.Count Step 1
                Set cobjSongs = cobjBlocks.Item(lngBlock)
                blnBlockColorState = Not blnBlockColorState
                If (cobjSongs Is Nothing) Then
                ' This is a dummy block, so do not add a block.
                    Exit For
                End If
                ' Add the song block.
                Set shpBlock = mshpIndexBlockAdd _
                    (sldPage, _
                     udtBlockArea, _
                     lngRow - 1, _
                     lngColumn - 1, _
                     cobjSongs.Count)
                If (blnBlockColorState = False) Then
                    With shpBlock.Fill.ForeColor
                        .RGB = .RGB Xor VBA.RGB(32, 32, 32)
                    End With
                End If
                ' Add the song indexes to the song block.
                For lngSong = 1 To cobjSongs.Count Step 1
                    Set udtSlide = cobjSongs.Item(lngSong)
                    strSongTitle = mstrIndexSongTitleClean(udtSlide.Title)
                    strSongPage = VBA.CStr(udtSlide.SlideIndex)
                    strSongIndex = strSongPage & "." & VBA.vbTab & strSongTitle
                    If (shpBlock.TextFrame.HasText = Office.msoFalse) Then
                        strBlockText = ""
                    Else
                        strBlockText = shpBlock.TextFrame.TextRange.Text & VBA.vbNewLine
                    End If
                    strBlockText = strBlockText & strSongIndex
                    With shpBlock.TextFrame.TextRange
                        .Text = strBlockText
                        Do While (.Lines.Count > lngSong)
                            If (VBA.Right(.Text, 4) = " ...") Then
                                .Text = VBA.Left(.Text, VBA.Len(.Text) - 4)
                            End If
                            .Text = VBA.Left(.Text, VBA.InStrRev(.Text, " ") - 1) & " ..."
                        Loop
                    End With
                    lngRow = lngRow + 1
                Next
            Next
        Next
    Next
    
    ' Destroy the now unneeded page->column->block->song hierarchy.
    Set cobjPages = Nothing
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mpreIndexAdd _
( _
) As PowerPoint.Presentation
    Dim preIndex As PowerPoint.Presentation
    Dim alngTextStyles() As PowerPoint.PpTextStyleType
    Dim lngTextStyle As Long
    Dim lng As Long
    
    Set preIndex = mudtApplication.Doppelganger.Presentations.Add
    
    With preIndex
        With .PageSetup
            .SlideOrientation = Office.msoOrientationVertical
            .SlideSize = PowerPoint.ppSlideSizeLetterPaper
            .SlideWidth = 72 * 7.5
            .SlideHeight = 72 * 10
        End With
        With .SlideMaster
            With .Shapes
                Do While (.Count > 0)
                    .Item(.Count).Delete
                Loop
                .AddTitle
            End With
            
            ReDim alngTextStyles(0 To 2)
            alngTextStyles(0) = PowerPoint.ppDefaultStyle
            alngTextStyles(1) = PowerPoint.ppTitleStyle
            alngTextStyles(2) = PowerPoint.ppBodyStyle
            For lngTextStyle = LBound(alngTextStyles) To UBound(alngTextStyles) Step 1
                With .TextStyles(alngTextStyles(lngTextStyle))
                    With .Levels
                        For lng = 1 To .Count Step 1
                            With .Item(lng)
                                With .Font
                                    .AutoRotateNumbers = Office.msoFalse
                                    .BaselineOffset = 0
                                    .Bold = Office.msoFalse
                                    .Color.RGB = VBA.RGB(0, 0, 0)
                                    .Emboss = Office.msoFalse
                                    .Italic = Office.msoFalse
                                    .Name = "Arial"
                                    .Shadow = Office.msoFalse
                                    .Size = 12
                                    .Subscript = Office.msoFalse
                                    .Superscript = Office.msoFalse
                                    .Underline = Office.msoFalse
                                End With
                                With .ParagraphFormat
                                    .Alignment = PowerPoint.ppAlignCenter
                                    .BaseLineAlignment = PowerPoint.ppBaselineAlignTop
                                    .Bullet = Office.msoFalse
                                    .FarEastLineBreakControl = Office.msoFalse
                                    .LineRuleAfter = Office.msoFalse
                                    .LineRuleBefore = Office.msoFalse
                                    .LineRuleWithin = Office.msoFalse
                                    .SpaceAfter = 0
                                    .SpaceBefore = 0
                                    .SpaceWithin = 14
                                    .TextDirection = PowerPoint.ppDirectionLeftToRight
                                    .WordWrap = Office.msoTrue
                                End With
                            End With
                        Next
                    End With
                End With
            Next
        End With
    End With
    
    Set mpreIndexAdd = preIndex
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function msldIndexPageAdd _
( _
    ByRef preCurrent As PowerPoint.Presentation _
) As PowerPoint.Slide
    Dim sldIndex As PowerPoint.Slide
    Dim lngIndex As Long
    Dim shpTitle As PowerPoint.Shape
    
    Set sldIndex = preCurrent.Slides.Add(preCurrent.Slides.Count + 1, PowerPoint.ppLayoutText)
    
    With sldIndex.Shapes
        For lngIndex = .Count To 1 Step -1
            If (Not .Item(lngIndex) Is sldIndex.Shapes.Title) Then
                .Item(lngIndex).Delete
            End If
        Next
    End With
    
    Set msldIndexPageAdd = sldIndex
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mshpIndexTitleAdd _
( _
    ByRef sldPage As PowerPoint.Slide, _
    ByRef udtTitleArea As Area, _
    ByRef strTitle As String _
) As PowerPoint.Shape
    Dim shpTitle As PowerPoint.Shape
    Dim lngTop As Long
    Dim lngLeft As Long
    Dim lngHeight As Long
    Dim lngWidth As Long
    
    Set shpTitle = sldPage.Shapes.Title
    
    With udtTitleArea
        shpTitle.Top = .Top
        shpTitle.Left = .Left
        shpTitle.Height = .Height
        shpTitle.Width = .Width
    End With
    
    shpTitle.TextFrame.TextRange.Text = strTitle
    
    Set mshpIndexTitleAdd = shpTitle
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mshpIndexCategoryAdd _
( _
    ByRef sldPage As PowerPoint.Slide, _
    ByRef udtCategoryArea As Area, _
    ByRef lngColumnOffset As Long, _
    ByRef strCategory As String _
) As PowerPoint.Shape
    Dim shpCategory As PowerPoint.Shape
    
    Set shpCategory = sldPage.Shapes.AddTextbox(Office.msoTextOrientationHorizontal, 0, 0, 0, 0)
    
    With udtCategoryArea
        shpCategory.Top = .Top
        shpCategory.Left = .Left + lngColumnOffset * .ColumnWidth
        shpCategory.Height = .Height
        shpCategory.Width = .Width
    End With
        
    mIndexShapeConfigure shpCategory
    With shpCategory
        With .TextFrame
            With .TextRange
                With .Font
                    .Size = .Size + 2
                End With
            End With
        End With
    End With
    
    shpCategory.TextFrame.TextRange.Text = strCategory
    
    Set mshpIndexCategoryAdd = shpCategory
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mshpIndexBlockAdd _
( _
    ByRef sldPage As PowerPoint.Slide, _
    ByRef udtBlockArea As Area, _
    ByRef lngRowOffset As Long, _
    ByRef lngColumnOffset As Long, _
    ByRef lngRows As Long _
) As PowerPoint.Shape
    Dim shpBlock As PowerPoint.Shape
    
    Set shpBlock = sldPage.Shapes.AddTextbox(Office.msoTextOrientationHorizontal, 0, 0, 0, 0)
    
    With udtBlockArea
        shpBlock.Top = .Top + lngRowOffset * .RowHeight
        shpBlock.Left = .Left + lngColumnOffset * .ColumnWidth
        shpBlock.Height = lngRows * .RowHeight
        shpBlock.Width = .Width
    End With
        
    mIndexShapeConfigure shpBlock
    With shpBlock
        With .TextFrame
            With .Ruler
                With .TabStops
                    .Add PowerPoint.ppTabStopLeft, 72 * 0.5
                    .Add PowerPoint.ppTabStopLeft, 72 * 0.75
                    .Add PowerPoint.ppTabStopLeft, 72 * 1
                End With
            End With
            With .TextRange
                With .ParagraphFormat
                    .Alignment = PowerPoint.ppAlignLeft
                End With
            End With
        End With
    End With
             
    Set mshpIndexBlockAdd = shpBlock
End Function

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Sub mIndexShapeConfigure _
( _
    ByRef shpShape As PowerPoint.Shape _
)
    Dim lng As Long
    
    With shpShape
        With .Fill
            .Solid
            .ForeColor.RGB = VBA.RGB(255, 255, 255)
            .BackColor.RGB = VBA.RGB(255, 255, 255)
        End With
        With .TextFrame
            .AutoSize = PowerPoint.ppAutoSizeNone
            .MarginLeft = 72 * 0.1
            .MarginTop = 0
            .MarginRight = 72 * 0.1
            .MarginBottom = 0
            .Orientation = Office.msoTextOrientationHorizontal
            With .Ruler
                With .Levels
                    For lng = 1 To .Count Step 1
                        .Item(lng).FirstMargin = 0
                        .Item(lng).LeftMargin = 0
                    Next
                End With
                With .TabStops
                    For lng = 1 To .Count Step 1
                        .Item(lng).Clear
                    Next
                End With
            End With
            .VerticalAnchor = Office.msoAnchorTop
            .WordWrap = Office.msoTrue
        End With
    End With
End Sub

'-------------------------------------------------------------------------------
' Purpose:
' Assumptions:
' Effects:
' Inputs:
' Returns:
'-------------------------------------------------------------------------------
Private Function mstrIndexSongTitleClean _
( _
    ByRef strDirty As String _
) As String
    Dim strClean As String
    Dim strPage As String
    Dim strPart As String
    Dim lngM As Long
    Dim lngN As Long
    
    mstrIndexSongTitleClean = strDirty

    strClean = strDirty
    
    Do While (VBA.Right(strClean, 1) = " ")
        strClean = VBA.Left(strClean, VBA.Len(strClean) - 1)
    Loop
    If (Not VBA.Right(strClean, 1) = ")") Then Exit Function
    
    strClean = VBA.Left(strClean, VBA.Len(strClean) - 1)
    
    strPage = ""
    Do While ((VBA.Len(strClean) > 0) And (Not VBA.Right(strClean, 1) = "("))
        strPage = VBA.Right(strClean, 1) & strPage
        strClean = VBA.Left(strClean, VBA.Len(strClean) - 1)
    Loop
    If (Not VBA.Right(strClean, 1) = "(") Then Exit Function
    
    strClean = VBA.Left(strClean, VBA.Len(strClean) - 1)
    
    ' Remove trailing whitespace.
    strClean = VBA.Left(strClean, VBA.Len(strClean) - 1)
    Do While (VBA.Right(strClean, 1) = " ")
        strClean = VBA.Left(strClean, VBA.Len(strClean) - 1)
    Loop
    
    strPart = ""
    Do While (VBA.Right(strPage, 1) = " ")
        strPage = VBA.Left(strPage, VBA.Len(strPage) - 1)
    Loop
    Do While ((VBA.Len(strPage) > 0) And (Not VBA.Right(strPage, 1) = " "))
        strPart = VBA.Right(strPage, 1) & strPart
        strPage = VBA.Left(strPage, VBA.Len(strPage) - 1)
    Loop
    If (Not IsNumeric(strPart)) Then Exit Function
    If (Not VBA.CDbl(VBA.Fix(VBA.CDbl(strPart))) = VBA.CDbl(strPart)) Then Exit Function
    
    lngN = VBA.Fix(strPart)
    If (Not lngN > 0) Then Exit Function
        
    strPart = ""
    Do While (VBA.Right(strPage, 1) = " ")
        strPage = VBA.Left(strPage, VBA.Len(strPage) - 1)
    Loop
    Do While ((VBA.Len(strPage) > 0) And (Not VBA.Right(strPage, 1) = " "))
        strPart = VBA.Right(strPage, 1) & strPart
        strPage = VBA.Left(strPage, VBA.Len(strPage) - 1)
    Loop
    If (Not VBA.LCase(strPart) = "of") Then Exit Function
    
    strPart = ""
    Do While (VBA.Right(strPage, 1) = " ")
        strPage = VBA.Left(strPage, VBA.Len(strPage) - 1)
    Loop
    Do While ((VBA.Len(strPage) > 0) And (Not VBA.Right(strPage, 1) = " "))
        strPart = VBA.Right(strPage, 1) & strPart
        strPage = VBA.Left(strPage, VBA.Len(strPage) - 1)
    Loop
    If (Not IsNumeric(strPart)) Then Exit Function
    If (Not VBA.CDbl(VBA.Fix(VBA.CDbl(strPart))) = VBA.CDbl(strPart)) Then Exit Function
    
    lngM = VBA.Fix(VBA.CDbl(strPart))
    If (Not lngM > 0) Then Exit Function
    
    lngM = VBA.Fix(VBA.CDbl(strPart))
    If (Not lngM <= lngN) Then Exit Function
    
    Do While (VBA.Right(strPage, 1) = " ")
        strPage = VBA.Left(strPage, VBA.Len(strPage) - 1)
    Loop
    If (Not strPage = "") Then Exit Function
    
    mstrIndexSongTitleClean = strClean
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
