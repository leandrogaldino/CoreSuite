Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D

''' <summary>Small GDI+ primitives used only by the renderer.</summary>
Friend NotInheritable Class ToolStripDrawing
    Private Sub New()
    End Sub

    Friend Shared Function Scale(Graphics As Graphics, Value As Integer) As Integer
        Return Math.Max(1, CInt(Math.Round(CDbl(Value) * Graphics.DpiX / 96.0R)))
    End Function

    Friend Shared Function Blend(First As Color, Second As Color, Weight As Single) As Color
        Dim Factor As Single = Math.Max(0.0F, Math.Min(1.0F, Weight))
        Return Color.FromArgb(CInt(First.A + (Second.A - First.A) * Factor),
                              CInt(First.R + (Second.R - First.R) * Factor),
                              CInt(First.G + (Second.G - First.G) * Factor),
                              CInt(First.B + (Second.B - First.B) * Factor))
    End Function

    Friend Shared Function RoundedPath(Rect As Rectangle, Radius As Integer) As GraphicsPath
        Dim Path As New GraphicsPath()
        If Rect.Width <= 0 OrElse Rect.Height <= 0 Then Return Path
        Dim Rounded As Integer = Math.Max(0, Math.Min(Radius, Math.Min(Rect.Width, Rect.Height) \ 2))
        If Rounded = 0 Then
            Path.AddRectangle(Rect)
            Return Path
        End If
        Dim Diameter As Integer = Rounded * 2
        Path.AddArc(Rect.Left, Rect.Top, Diameter, Diameter, 180, 90)
        Path.AddArc(Rect.Right - Diameter, Rect.Top, Diameter, Diameter, 270, 90)
        Path.AddArc(Rect.Right - Diameter, Rect.Bottom - Diameter, Diameter, Diameter, 0, 90)
        Path.AddArc(Rect.Left, Rect.Bottom - Diameter, Diameter, Diameter, 90, 90)
        Path.CloseFigure()
        Return Path
    End Function

    Friend Shared Sub FillShape(Graphics As Graphics, Rect As Rectangle, Radius As Integer, Color As Color, Gradient As Boolean, Optional Outline As Color? = Nothing)
        If Rect.Width <= 0 OrElse Rect.Height <= 0 Then Return
        Dim Previous As SmoothingMode = Graphics.SmoothingMode
        Graphics.SmoothingMode = SmoothingMode.AntiAlias
        Try
            Using Path As GraphicsPath = RoundedPath(Rect, Radius)
                If Gradient AndAlso Rect.Height > 1 Then
                    Using Brush As New LinearGradientBrush(Rect, Blend(Color, Color.White, 0.1F), Color, LinearGradientMode.Vertical)
                        Graphics.FillPath(Brush, Path)
                    End Using
                Else
                    Using Brush As New SolidBrush(Color)
                        Graphics.FillPath(Brush, Path)
                    End Using
                End If
                If Outline.HasValue Then
                    Using pen As New Pen(Outline.Value, 1.0F)
                        Graphics.DrawPath(pen, Path)
                    End Using
                End If
            End Using
        Finally
            Graphics.SmoothingMode = Previous
        End Try
    End Sub

    Friend Shared Sub DrawDownArrow(Graphics As Graphics, Center As Point, Color As Color, Size As Integer)
        Dim OldMode As SmoothingMode = Graphics.SmoothingMode
        Graphics.SmoothingMode = SmoothingMode.AntiAlias
        Try
            Dim Triangle As Point() = {
                New Point(Center.X - Size, Center.Y - Size \ 2),
                New Point(Center.X + Size, Center.Y - Size \ 2),
                New Point(Center.X, Center.Y + Size \ 2)}
            Using brush As New SolidBrush(Color)
                Graphics.FillPolygon(brush, Triangle)
            End Using
        Finally
            Graphics.SmoothingMode = OldMode
        End Try
    End Sub
    ''' <summary>Draws a centered, anti-aliased down chevron for a split button.</summary>
    Friend Shared Sub DrawDownChevron(Graphics As Graphics, Bounds As Rectangle, Color As Color)
        If Bounds.Width <= 0 OrElse Bounds.Height <= 0 Then Return

        Dim DpiScale As Single = CSng(Graphics.DpiX / 96.0F)
        Dim ArrowWidth As Single = Math.Min(6.0F * DpiScale, Bounds.Width - 2.0F * DpiScale)
        Dim ArrowHeight As Single = Math.Min(3.5F * DpiScale, Bounds.Height - 2.0F * DpiScale)
        If ArrowWidth <= 0 OrElse ArrowHeight <= 0 Then Return

        Dim CenterX As Single = Bounds.Left + Bounds.Width / 2.0F
        Dim CenterY As Single = Bounds.Top + Bounds.Height / 2.0F
        Dim PreviousSmoothing As SmoothingMode = Graphics.SmoothingMode
        Dim PreviousPixelOffset As PixelOffsetMode = Graphics.PixelOffsetMode
        Graphics.SmoothingMode = SmoothingMode.AntiAlias
        Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality
        Try
            Using ChevronPen As New Pen(Color, 1.4F * DpiScale)
                ChevronPen.StartCap = LineCap.Round
                ChevronPen.EndCap = LineCap.Round
                ChevronPen.LineJoin = LineJoin.Round
                Graphics.DrawLines(ChevronPen, New PointF() {
                    New PointF(CenterX - ArrowWidth / 2.0F, CenterY - ArrowHeight / 2.0F),
                    New PointF(CenterX, CenterY + ArrowHeight / 2.0F),
                    New PointF(CenterX + ArrowWidth / 2.0F, CenterY - ArrowHeight / 2.0F)})
            End Using
        Finally
            Graphics.SmoothingMode = PreviousSmoothing
            Graphics.PixelOffsetMode = PreviousPixelOffset
        End Try
    End Sub

End Class
