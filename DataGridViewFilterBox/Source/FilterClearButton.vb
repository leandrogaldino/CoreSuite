Imports System.Drawing.Drawing2D
''' <summary>
''' Represents the internal button used to clear text and render either a custom image or a built-in close glyph.
''' </summary>
Friend Class FilterClearButton
    Inherits Button
    ''' <summary>
    ''' Initializes a new instance of the <see cref="FilterClearButton"/> class.
    ''' </summary>
    Public Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint, True)
        FlatStyle = FlatStyle.Flat
        FlatAppearance.BorderSize = 0
        TabStop = False
        Text = String.Empty
        Cursor = Cursors.Hand
        AccessibleName = "Clear filter"
        AccessibleRole = AccessibleRole.PushButton
    End Sub
    ''' <summary>
    ''' Paints the assigned image or the built-in close glyph using the current enabled and color state.
    ''' </summary>
    ''' <param name="e">The paint event data.</param>
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        e.Graphics.Clear(BackColor)
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
        If Image IsNot Nothing Then
            Dim ImageSize As Integer = Math.Max(1, Math.Min(Width - 6, Height - 6))
            Dim ImageBounds As New Rectangle((Width - ImageSize) \ 2, (Height - ImageSize) \ 2, ImageSize, ImageSize)
            If Enabled Then
                e.Graphics.DrawImage(Image, ImageBounds)
            Else
                ControlPaint.DrawImageDisabled(e.Graphics, Image, ImageBounds.X, ImageBounds.Y, BackColor)
            End If
            Return
        End If
        Dim GlyphColor As Color = If(Enabled, ForeColor, SystemColors.GrayText)
        Dim Margin As Single = Math.Max(5.0F, Math.Min(Width, Height) * 0.32F)
        Using GlyphPen As New Pen(GlyphColor, 1.6F)
            GlyphPen.StartCap = LineCap.Round
            GlyphPen.EndCap = LineCap.Round
            e.Graphics.DrawLine(GlyphPen, Margin, Margin, Width - Margin, Height - Margin)
            e.Graphics.DrawLine(GlyphPen, Width - Margin, Margin, Margin, Height - Margin)
        End Using
    End Sub
    ''' <summary>
    ''' Invalidates the button when the pointer enters its bounds.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnMouseEnter(e As EventArgs)
        MyBase.OnMouseEnter(e)
        Invalidate()
    End Sub
    ''' <summary>
    ''' Invalidates the button when the pointer leaves its bounds.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        Invalidate()
    End Sub
End Class
