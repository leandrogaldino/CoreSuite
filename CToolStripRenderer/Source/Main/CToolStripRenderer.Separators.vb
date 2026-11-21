Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public Partial Class CToolStripRenderer
    Protected Overrides Sub OnRenderSeparator(e As ToolStripSeparatorRenderEventArgs)
        If e Is Nothing OrElse Not _ShowSeparators Then Return
        Dim Palette As ToolStripPalette = EffectivePalette()
        Using Pen As New Pen(Palette.Separator, 1.0F)
            If e.Vertical Then
                Dim x As Integer = e.Item.Width \ 2
                e.Graphics.DrawLine(Pen, x, ToolStripDrawing.Scale(e.Graphics, 6),
                                    x, Math.Max(0, e.Item.Height - ToolStripDrawing.Scale(e.Graphics, 7)))
            Else
                Dim Inset As Integer = If(e.Item.IsOnDropDown, ToolStripDrawing.Scale(e.Graphics, 32), ToolStripDrawing.Scale(e.Graphics, 7))
                Dim Y As Integer = e.Item.Height \ 2
                e.Graphics.DrawLine(Pen, Inset, Y, Math.Max(Inset, e.Item.Width - ToolStripDrawing.Scale(e.Graphics, 8)), Y)
            End If
        End Using
    End Sub

    Protected Overrides Sub OnRenderGrip(e As ToolStripGripRenderEventArgs)
        If e Is Nothing OrElse Not _showGrip Then Return
        Dim Rect As Rectangle = e.GripBounds
        If Rect.Width < 4 OrElse Rect.Height < 4 Then Return
        Dim OldMode As SmoothingMode = e.Graphics.SmoothingMode
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
        Try
            Using Brush As New SolidBrush(EffectivePalette().Grip)
                Dim DotSize As Integer = ToolStripDrawing.Scale(e.Graphics, 2)
                Dim Spacing As Integer = ToolStripDrawing.Scale(e.Graphics, 4)
                If Rect.Height > Rect.Width Then
                    For Row As Integer = 0 To 3
                        For Col As Integer = 0 To 1
                            e.Graphics.FillEllipse(Brush, Rect.Left + Col * Spacing + 2, Rect.Top + Row * Spacing + 4, DotSize, DotSize)
                        Next
                    Next
                Else
                    For Row As Integer = 0 To 1
                        For Col As Integer = 0 To 3
                            e.Graphics.FillEllipse(Brush, Rect.Left + Col * Spacing + 4, Rect.Top + Row * Spacing + 2, DotSize, DotSize)
                        Next
                    Next
                End If
            End Using
        Finally
            e.Graphics.SmoothingMode = OldMode
        End Try
    End Sub
    Protected Overrides Sub OnRenderStatusStripSizingGrip(e As ToolStripRenderEventArgs)
        If e Is Nothing OrElse e.ToolStrip Is Nothing Then Return
        Dim Palette As ToolStripPalette = EffectivePalette()
        Dim Size As Integer = ToolStripDrawing.Scale(e.Graphics, 3)
        Dim Distance As Integer = ToolStripDrawing.Scale(e.Graphics, 4)
        Dim Width As Integer = e.ToolStrip.Width
        Dim Height As Integer = e.ToolStrip.Height
        Dim IsRtl As Boolean = (e.ToolStrip.RightToLeft = RightToLeft.Yes)
        Using Pen As New Pen(Palette.Grip, 1.0F)
            For i As Integer = 0 To 2
                Dim dx As Integer = Size + i * Distance
                If IsRtl Then
                    e.Graphics.DrawLine(Pen, dx, Height - Size, Size, Height - dx)
                Else
                    e.Graphics.DrawLine(Pen, Width - dx, Height - Size, Width - Size, Height - dx)
                End If
            Next
        End Using
    End Sub
End Class
