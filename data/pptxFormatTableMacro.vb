Sub tidytable()
    With Application.ActiveWindow.View.Slide.Shapes
        For i = 1 To .Count
            If .Item(i).HasTable Then
                Set Table = .Item(i).Table
                ' Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Cell 1"
                ' Table.Cell(1, 1).Shape.TextFrame.TextRange.Font.Color = RGB(0, 0, 0)
                ' Table.Cell(1, 1).Shape.TextFrame.TextRange.Font.Size = 20
                ' Table.Cell(1, 1).Shape.TextFrame.TextRange.Font.Name = "Arial"
                For c = 1 To Table.Columns.Count
                    For r = 1 To Table.Rows.Count
                        With Table.Rows(r).Cells(c)
                            .Borders.Item(ppBorderBottom).ForeColor.RGB = RGB(0, 0, 0)
                            .Borders.Item(ppBorderTop).ForeColor.RGB = RGB(0, 0, 0)
                            .Borders.Item(ppBorderRight).ForeColor.RGB = RGB(0, 0, 0)
                            .Borders.Item(ppBorderLeft).ForeColor.RGB = RGB(0, 0, 0)
                            .Borders.Item(ppBorderBottom).Weight = 1
                            .Borders.Item(ppBorderRight).Weight = 1
                            .Shape.Fill.ForeColor.RGB = RGB(235, 241, 233)
                            .Shape.TextFrame.TextRange.Font.Color = RGB(0, 0, 0)
                            If r = 1 Or c = 1 Then
                                .Shape.Fill.ForeColor.RGB = RGB(112, 173, 71)
                            End If
                            If r = 1 Then
                                .Borders.Item(ppBorderTop).Weight = 1
                                .Borders.Item(ppBorderBottom).Weight = 3
                            End If
                            If c = 1 Then
                                .Borders.Item(ppBorderLeft).Weight = 1
                                .Borders.Item(ppBorderRight).Weight = 3
                            End If
                        End With
                    Next
                Next
            End If
        Next
    End With
End Sub



Sub tidytable_noidx()
    With Application.ActiveWindow.View.Slide.Shapes
        For i = 1 To .Count
            If .Item(i).HasTable Then
                Set Table = .Item(i).Table
                With Table.Rows(1).Cells(1)
                    .Shape.Fill.Transparency = 1
                    .Shape.TextFrame.TextRange.Font.Color = RGB(0, 0, 0)
                    .Borders.Item(ppBorderTop).ForeColor.RGB = RGB(255, 255, 255)
                    .Borders.Item(ppBorderLeft).ForeColor.RGB = RGB(255, 255, 255)
                    .Borders.Item(ppBorderRight).ForeColor.RGB = RGB(255, 255, 255)
                End With
                For c = 1 To Table.Columns.Count
                    For r = 2 To Table.Rows.Count
                        With Table.Rows(r).Cells(c)
                            .Borders.Item(ppBorderBottom).ForeColor.RGB = RGB(0, 0, 0)
                            .Borders.Item(ppBorderTop).ForeColor.RGB = RGB(0, 0, 0)
                            .Borders.Item(ppBorderRight).ForeColor.RGB = RGB(0, 0, 0)
                            .Borders.Item(ppBorderLeft).ForeColor.RGB = RGB(0, 0, 0)
                            .Borders.Item(ppBorderBottom).Weight = 1
                            .Borders.Item(ppBorderRight).Weight = 1
                            .Shape.Fill.ForeColor.RGB = RGB(235, 241, 233)
                            .Shape.TextFrame.TextRange.Font.Color = RGB(0, 0, 0)
                            If r = 2 Then
                                .Shape.Fill.ForeColor.RGB = RGB(112, 173, 71)
                                .Borders.Item(ppBorderTop).Weight = 1
                                .Borders.Item(ppBorderBottom).Weight = 3
                            End If
'                            If c = 1 Then
'                                .Borders.Item(ppBorderLeft).Weight = 1
'                                .Borders.Item(ppBorderRight).Weight = 3
'                            End If
                        End With
                    Next
                Next
            End If
        Next
    End With
End Sub
