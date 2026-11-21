Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public Partial Class CToolStripRenderer
    Protected Overrides Sub OnRenderMenuItemBackground(e As ToolStripItemRenderEventArgs)
        If e Is Nothing OrElse e.Item Is Nothing OrElse Not e.Item.Enabled Then Return
        If Not (e.Item.Selected OrElse e.Item.Pressed) Then Return
        Dim Palette As ToolStripPalette = EffectivePalette()
        Dim Rectangle As New Rectangle(2, 1, e.Item.Width - 5, e.Item.Height - 3)
        Dim Fill As Color = If(e.Item.Pressed, Palette.PressedBackground, Palette.HoverBackground)
        ToolStripDrawing.FillShape(e.Graphics, Rectangle, ToolStripDrawing.Scale(e.Graphics, _CornerRadius), Fill, ShouldUseGradients())
    End Sub

    Protected Overrides Sub OnRenderItemCheck(e As ToolStripItemImageRenderEventArgs)
        If e Is Nothing Then Return
        Dim Palette As ToolStripPalette = EffectivePalette()
        Dim Rect As Rectangle = e.ImageRectangle
        If Rect.Width <= 0 OrElse Rect.Height <= 0 Then Return
        Dim Size As Integer = Math.Min(Rect.Width, Rect.Height)
        Dim BoxRect As New Rectangle(Rect.Left + (Rect.Width - Size) \ 2,
                                     Rect.Top + (Rect.Height - Size) \ 2, Size, Size)
        ToolStripDrawing.FillShape(e.Graphics, BoxRect, ToolStripDrawing.Scale(e.Graphics, 3),
                                  Palette.CheckedBackground, False, Palette.Accent)
        Dim OldMode As SmoothingMode = e.Graphics.SmoothingMode
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
        Try
            Dim CenterX As Integer = BoxRect.Left + BoxRect.Width \ 2
            Dim CenterY As Integer = BoxRect.Top + BoxRect.Height \ 2
            Using Pen As New Pen(Palette.Accent, Math.Max(1.5F, CSng(ToolStripDrawing.Scale(e.Graphics, 2))))
                Pen.StartCap = LineCap.Round
                Pen.EndCap = LineCap.Round
                e.Graphics.DrawLines(Pen, New Point() {
                    New Point(CenterX - Size \ 4, CenterY),
                    New Point(CenterX - Size \ 12, CenterY + Size \ 5),
                    New Point(CenterX + Size \ 3, CenterY - Size \ 4)})
            End Using
        Finally
            e.Graphics.SmoothingMode = OldMode
        End Try
    End Sub

    Protected Overrides Sub OnRenderItemText(e As ToolStripItemTextRenderEventArgs)
        If e Is Nothing Then Return
        Dim Palette As ToolStripPalette = EffectivePalette()
        If Not e.Item.Enabled Then
            e.TextColor = Palette.DisabledForeground
        ElseIf IsSystemPaletteActive() AndAlso (e.Item.Selected OrElse e.Item.Pressed) Then
            e.TextColor = Palette.CheckedForeground
        Else
            Dim Button As ToolStripButton = TryCast(e.Item, ToolStripButton)
            If Button IsNot Nothing AndAlso Button.Checked AndAlso _CheckedButtonStyle = CheckedButtonStyle.Accent AndAlso
                Not Button.Selected AndAlso Not Button.Pressed Then
                e.TextColor = Palette.CheckedForeground
            Else
                e.TextColor = Palette.Foreground
            End If
        End If
        MyBase.OnRenderItemText(e)
    End Sub

    Protected Overrides Sub OnRenderArrow(e As ToolStripArrowRenderEventArgs)
        If e Is Nothing Then Return
        Dim Palette As ToolStripPalette = EffectivePalette()
        e.ArrowColor = If(Not e.Item.Enabled, Palette.DisabledForeground, If(IsSystemPaletteActive() AndAlso (e.Item.Selected OrElse e.Item.Pressed), Palette.CheckedForeground, Palette.Foreground))
        MyBase.OnRenderArrow(e)
    End Sub
End Class
