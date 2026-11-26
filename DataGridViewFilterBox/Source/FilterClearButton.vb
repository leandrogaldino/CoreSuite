Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms
''' <summary>
''' Represents the internal button used to clear text and render either a custom image or a built-in close glyph.
''' </summary>
Friend Class FilterClearButton
    Inherits NoFocusCueButton
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
    Friend Sub RenderContent(G As Graphics, SurfaceSize As Size)
        G.SmoothingMode = SmoothingMode.AntiAlias
        If Image IsNot Nothing Then
            EmbeddedTextButtonLayout.DrawCenteredImage(G, Image, SurfaceSize, CInt(Math.Round(3 * DeviceDpi / 96.0R)), Enabled, BackColor)
            Return
        End If
        Dim GlyphColor As Color = If(Enabled, ForeColor, SystemColors.GrayText)
        Dim Margin As Single = Math.Max(5.0F, Math.Min(SurfaceSize.Width, SurfaceSize.Height) * 0.32F)
        Using GlyphPen As New Pen(GlyphColor, 1.6F)
            GlyphPen.StartCap = LineCap.Round
            GlyphPen.EndCap = LineCap.Round
            G.DrawLine(GlyphPen, Margin, Margin, SurfaceSize.Width - Margin, SurfaceSize.Height - Margin)
            G.DrawLine(GlyphPen, SurfaceSize.Width - Margin, Margin, Margin, SurfaceSize.Height - Margin)
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
