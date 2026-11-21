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
End Class
