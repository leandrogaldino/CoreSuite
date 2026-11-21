Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public Partial Class CToolStripRenderer
    Protected Overrides Sub OnRenderToolStripBackground(e As ToolStripRenderEventArgs)
        If e Is Nothing OrElse e.ToolStrip Is Nothing Then Return
        Dim Rectangle As New Rectangle(Point.Empty, e.ToolStrip.Size)
        If Rectangle.Width <= 0 OrElse Rectangle.Height <= 0 Then Return
        Dim Palette As ToolStripPalette = EffectivePalette()
        Dim Fill As Color = If(TypeOf e.ToolStrip Is ToolStripDropDown, Palette.DropDownBackground, Palette.Background)
        If ShouldUseGradients() Then
            Using brush As New LinearGradientBrush(Rectangle, ToolStripDrawing.Blend(Fill, Color.White, 0.1F), Fill, LinearGradientMode.Vertical)
                e.Graphics.FillRectangle(brush, Rectangle)
            End Using
        Else
            Using brush As New SolidBrush(Fill)
                e.Graphics.FillRectangle(brush, Rectangle)
            End Using
        End If
    End Sub

    Protected Overrides Sub OnRenderToolStripBorder(e As ToolStripRenderEventArgs)
        If e Is Nothing OrElse e.ToolStrip Is Nothing Then Return
        Dim Palette As ToolStripPalette = EffectivePalette()
        Dim Bounds As New Rectangle(0, 0, e.ToolStrip.Width - 1, e.ToolStrip.Height - 1)
        If Bounds.Width <= 0 OrElse Bounds.Height <= 0 Then Return
        If TypeOf e.ToolStrip Is ToolStripDropDown Then
            If _showDropDownBorder Then
                Using pen As New Pen(Palette.DropDownBorder)
                    e.Graphics.DrawRectangle(pen, Bounds)
                End Using
            End If
        ElseIf _ShowToolStripBorder Then
            Using pen As New Pen(Palette.Border)
                If TypeOf e.ToolStrip Is StatusStrip Then
                    e.Graphics.DrawLine(pen, Bounds.Left, Bounds.Top, Bounds.Right, Bounds.Top)
                Else
                    e.Graphics.DrawLine(pen, Bounds.Left, Bounds.Bottom, Bounds.Right, Bounds.Bottom)
                End If
            End Using
        End If
    End Sub

    Protected Overrides Sub OnRenderImageMargin(e As ToolStripRenderEventArgs)
        If e Is Nothing Then Return
        If Not _ShowImageMargin Then Return
        Dim Rect As Rectangle = e.AffectedBounds
        If Rect.Width <= 0 OrElse Rect.Height <= 0 Then Return
        Using brush As New SolidBrush(EffectivePalette().ImageMarginBackground)
            e.Graphics.FillRectangle(brush, Rect)
        End Using
    End Sub
End Class
