VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub AddImageCaptionTables()
Dim iShp As InlineShape, Rng As Range, Tbl As Table
Dim i As Long, PicWdth As Single, VPos As Single
Dim HPos As Single, VRel As Long, HRel As Long
With ActiveDocument
  For i = 1 To .InlineShapes.Count
    If .InlineShapes(i).Range.Information(wdWithInTable) = False Then
      PicWdth = .InlineShapes(i).Width
      Set Rng = .InlineShapes(i).Range
      With Rng
        If .Characters.Last.Next.Text = vbCr Then .Characters.Last.Next.Text = vbNullString
        .InlineShapes(1).Range.Cut
      End With
      Set Tbl = .Tables.Add(Range:=Rng, Numrows:=2, NumColumns:=1)
      With Tbl
        With .Cell(1, 1).Range.ParagraphFormat
          .SpaceBefore = 0
          .SpaceAfter = 0
          .LeftIndent = 0
          .RightIndent = 0
          .FirstLineIndent = 0
          .KeepWithNext = True
        End With
        .Cell(1, 1).Range.Paste
        Set Rng = .Cell(2, 1).Range
        With Rng
          .Style = "Caption"
          .End = .End - 1
          .InsertAfter vbCr
          .InsertCaption Label:="Figure", TitleAutoText:="InsertCaption1", _
            Title:="", Position:=wdCaptionPositionBelow, ExcludeLabel:=0
          .Characters.First.Text = vbNullString
          .Paragraphs.First.Range.Characters.Last.Text = vbNullString
        End With
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .TopPadding = 0
        .BottomPadding = 0
        .LeftPadding = 0
        .RightPadding = 0
        .Spacing = 0
        .Columns.Width = PicWdth
      End With
    End If
  Next
  While .Shapes.Count > 0
    With .Shapes(1)
      VRel = .RelativeVerticalPosition
      HRel = .RelativeHorizontalPosition
      VPos = .Top
      HPos = .Left
      PicWdth = .Width
      Set iShp = .ConvertToInlineShape
    End With
    With iShp
      Set Rng = .Range
      .Range.Cut
    End With
    Set Tbl = .Tables.Add(Range:=Rng, Numrows:=2, NumColumns:=1)
    With Tbl
      With .Rows
        .LeftIndent = 0
        .WrapAroundText = True
        .HorizontalPosition = HPos
        .RelativeHorizontalPosition = HRel
        .VerticalPosition = VPos
        .RelativeVerticalPosition = VRel
        .AllowOverlap = False
      End With
      With .Cell(1, 1).Range.ParagraphFormat
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LeftIndent = 0
        .RightIndent = 0
        .FirstLineIndent = 0
        .KeepWithNext = True
      End With
      .Cell(1, 1).Range.Paste
      Set Rng = .Cell(2, 1).Range
      With Rng
        .Style = "Caption"
        .End = .End - 1
        .InsertAfter vbCr
        .InsertCaption Label:="Figure", TitleAutoText:="InsertCaption1", _
          Title:="", Position:=wdCaptionPositionBelow, ExcludeLabel:=0
        .Characters.First.Text = vbNullString
        .Paragraphs.First.Range.Characters.Last.Text = vbNullString
      End With
      .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
      .Borders(wdBorderRight).LineStyle = wdLineStyleNone
      .Borders(wdBorderTop).LineStyle = wdLineStyleNone
      .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
      .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
      .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
      .TopPadding = 0
      .BottomPadding = 0
      .LeftPadding = 0
      .RightPadding = 0
      .Spacing = 0
      .Columns.Width = PicWdth
    End With
  Wend
End With
End Sub

