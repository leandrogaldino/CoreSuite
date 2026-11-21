Imports System.Drawing
Imports System.Windows.Forms

Public Partial Class CToolStripRenderer
    Private Sub PaintItemBackground(Graphics As Graphics, Item As ToolStripItem, Checked As Boolean, Optional PressedOverride As Boolean = False)
        If Item Is Nothing OrElse Not Item.Enabled Then Return
        Dim Palette As ToolStripPalette = EffectivePalette()
        Dim Rectangle As New Rectangle(1, 1, Item.Width - 3, Item.Height - 3)
        If Rectangle.Width <= 0 OrElse Rectangle.Height <= 0 Then Return
        Dim Radius As Integer = ToolStripDrawing.Scale(Graphics, _CornerRadius)
        Dim HasHotState As Boolean = PressedOverride OrElse Item.Pressed OrElse Item.Selected
        If Checked Then
            Select Case _checkedButtonStyle
                Case CheckedButtonStyle.None
                Case CheckedButtonStyle.Outline
                    ToolStripDrawing.FillShape(Graphics, Rectangle, Radius, Palette.Background, False, Palette.Accent)
                    If Not HasHotState Then Return
                Case CheckedButtonStyle.Filled
                    ToolStripDrawing.FillShape(Graphics, Rectangle, Radius, Palette.CheckedBackground, ShouldUseGradients())
                    If Not HasHotState Then Return
                Case CheckedButtonStyle.Underline
                    Using pen As New Pen(Palette.Accent, ToolStripDrawing.Scale(Graphics, 3))
                        Graphics.DrawLine(pen, Rectangle.Left + Radius, Rectangle.Bottom - 1, Rectangle.Right - Radius, Rectangle.Bottom - 1)
                    End Using
                    If Not HasHotState Then Return
                Case CheckedButtonStyle.Accent
                    ToolStripDrawing.FillShape(Graphics, Rectangle, Radius, Palette.Accent, ShouldUseGradients())
                    If Not HasHotState Then Return
            End Select
        End If
        If Item.Pressed OrElse PressedOverride Then
            ToolStripDrawing.FillShape(Graphics, Rectangle, Radius, Palette.PressedBackground, ShouldUseGradients())
        ElseIf Item.Selected Then
            ToolStripDrawing.FillShape(Graphics, Rectangle, Radius, Palette.HoverBackground, ShouldUseGradients())
        End If
    End Sub

    Protected Overrides Sub OnRenderButtonBackground(e As ToolStripItemRenderEventArgs)
        If e Is Nothing Then Return
        Dim Button As ToolStripButton = TryCast(e.Item, ToolStripButton)
        PaintItemBackground(e.Graphics, e.Item, Button IsNot Nothing AndAlso Button.Checked)
    End Sub

    Protected Overrides Sub OnRenderDropDownButtonBackground(e As ToolStripItemRenderEventArgs)
        If e Is Nothing Then Return
        PaintItemBackground(e.Graphics, e.Item, False)
        Dim Button As ToolStripDropDownButton = TryCast(e.Item, ToolStripDropDownButton)
        If Button IsNot Nothing AndAlso Button.ShowDropDownArrow Then
            Dim Palette As ToolStripPalette = EffectivePalette()
            Dim Color As Color = If(Button.Enabled, Palette.Foreground, Palette.DisabledForeground)
            Dim ArrowX As Integer = Button.Width - ToolStripDrawing.Scale(e.Graphics, 9)
            ToolStripDrawing.DrawDownArrow(e.Graphics, New Point(ArrowX, Button.Height \ 2), Color, ToolStripDrawing.Scale(e.Graphics, 3))
        End If
    End Sub

    Protected Overrides Sub OnRenderSplitButtonBackground(e As ToolStripItemRenderEventArgs)
        If e Is Nothing Then Return
        Dim Button As ToolStripSplitButton = TryCast(e.Item, ToolStripSplitButton)
        If Button Is Nothing Then Return
        PaintItemBackground(e.Graphics, Button, False)
        Dim Palette As ToolStripPalette = EffectivePalette()
        Dim ArrowRect As Rectangle = Button.DropDownButtonBounds
        If Button.Enabled AndAlso Button.Selected AndAlso ArrowRect.Width > 0 Then
            Using pen As New Pen(Palette.Separator)
                e.Graphics.DrawLine(pen, ArrowRect.Left, 4, ArrowRect.Left, Button.Height - 5)
            End Using
        End If
        Dim Color As Color = If(Button.Enabled, Palette.Foreground, Palette.DisabledForeground)
        ToolStripDrawing.DrawDownArrow(e.Graphics, New Point(ArrowRect.Left + ArrowRect.Width \ 2, ArrowRect.Top + ArrowRect.Height \ 2), Color, ToolStripDrawing.Scale(e.Graphics, 3))
    End Sub

    Protected Overrides Sub OnRenderOverflowButtonBackground(e As ToolStripItemRenderEventArgs)
        If e Is Nothing Then Return
        PaintItemBackground(e.Graphics, e.Item, False)
        Dim Palette As ToolStripPalette = EffectivePalette()
        Dim Color As Color = If(e.Item.Enabled, Palette.Foreground, Palette.DisabledForeground)
        ToolStripDrawing.DrawDownArrow(e.Graphics, New Point(e.Item.Width \ 2, e.Item.Height \ 2), Color, ToolStripDrawing.Scale(e.Graphics, 3))
    End Sub
End Class
