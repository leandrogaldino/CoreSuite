Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public Partial Class ColoredProgressBar

    ' --------------------------------------------------
    ' Painting
    ' --------------------------------------------------

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        If ClientSize.Width > 1 AndAlso ClientSize.Height > 1 Then
            Dim G As Graphics = e.Graphics
            Dim OriginalState As GraphicsState = G.Save()

            Try
                G.SmoothingMode = SmoothingMode.AntiAlias
                G.PixelOffsetMode = PixelOffsetMode.HighQuality

                Dim MaximumBorder As Integer =
                    Math.Max(
                        0,
                        Math.Min(ClientSize.Width, ClientSize.Height) - 1)

                Dim EffectiveBorder As Single =
                    CSng(Math.Min(_BorderThickness, MaximumBorder))

                Dim Inset As Single =
                    If(EffectiveBorder > 0.0F,
                       EffectiveBorder / 2.0F,
                       0.5F)

                Dim TrackRectangle As New RectangleF(
                    Inset,
                    Inset,
                    CSng(ClientSize.Width) - Inset * 2.0F,
                    CSng(ClientSize.Height) - Inset * 2.0F)

                If TrackRectangle.Width > 0.0F AndAlso
                   TrackRectangle.Height > 0.0F Then

                    DrawProgressBar(G, TrackRectangle, EffectiveBorder)
                End If

            Finally
                G.Restore(OriginalState)
            End Try
        End If

        MyBase.OnPaint(e)
    End Sub

    Private Sub DrawProgressBar(
        G As Graphics,
        TrackRectangle As RectangleF,
        EffectiveBorder As Single)

        ' Use the animated progress rather than the target value.

        Dim FillWidth As Single =
            CSng(TrackRectangle.Width * _DisplayedProgress)

        Dim FillRectangle As New RectangleF(
            TrackRectangle.X,
            TrackRectangle.Y,
            FillWidth,
            TrackRectangle.Height)

        Using TrackPath As GraphicsPath =
            CreateRoundedPath(TrackRectangle, _CornerRadius)

            Using TrackBrush As New SolidBrush(_TrackColor)
                G.FillPath(TrackBrush, TrackPath)
            End Using

            If FillRectangle.Width > 0.5F Then
                Using FillPath As GraphicsPath =
                    CreateRoundedPath(FillRectangle, _CornerRadius)

                    Using FillBrush As New LinearGradientBrush(
                        FillRectangle,
                        _ProgressStartColor,
                        _ProgressEndColor,
                        _GradientDirection)

                        G.FillPath(FillBrush, FillPath)
                    End Using
                End Using
            End If

            If EffectiveBorder > 0.0F Then
                Using BorderPen As New Pen(_BorderColor, EffectiveBorder)
                    G.DrawPath(BorderPen, TrackPath)
                End Using
            End If
        End Using

        If _ShowPercentage Then
            DrawPercentage(G, FillRectangle)
        End If
    End Sub

    Private Sub DrawPercentage(
        G As Graphics,
        FillRectangle As RectangleF)

        ' Display the animated percentage, not the target percentage.

        Dim PercentageText As String =
            (_DisplayedProgress * 100.0R).ToString("0") & "%"

        Using TextFormat As New StringFormat()
            TextFormat.Alignment = StringAlignment.Center
            TextFormat.LineAlignment = StringAlignment.Center
            TextFormat.FormatFlags = StringFormatFlags.NoWrap
            TextFormat.Trimming = StringTrimming.EllipsisCharacter

            Using TextBrush As New SolidBrush(ForeColor)
                G.DrawString(
                    PercentageText,
                    Font,
                    TextBrush,
                    ClientRectangle,
                    TextFormat)
            End Using

            If FillRectangle.Width > 0.5F Then
                Using FillPath As GraphicsPath =
                    CreateRoundedPath(FillRectangle, _CornerRadius)

                    Dim SavedState As GraphicsState = G.Save()

                    Try
                        G.SetClip(FillPath, CombineMode.Intersect)

                        Using FilledTextBrush As New SolidBrush(_ProgressTextColor)
                            G.DrawString(
                                PercentageText,
                                Font,
                                FilledTextBrush,
                                ClientRectangle,
                                TextFormat)
                        End Using
                    Finally
                        G.Restore(SavedState)
                    End Try
                End Using
            End If
        End Using
    End Sub

    Private Shared Function CreateRoundedPath(
        Rect As RectangleF,
        Radius As Integer) As GraphicsPath

        Dim Path As New GraphicsPath()

        If Rect.Width <= 0.0F OrElse Rect.Height <= 0.0F Then
            Return Path
        End If

        Dim EffectiveRadius As Single =
            Math.Min(
                CSng(Math.Max(0, Radius)),
                Math.Min(Rect.Width, Rect.Height) / 2.0F)

        If EffectiveRadius <= 0.0F Then
            Path.AddRectangle(Rect)
            Return Path
        End If

        Dim Diameter As Single = EffectiveRadius * 2.0F

        Path.AddArc(
            Rect.Left,
            Rect.Top,
            Diameter,
            Diameter,
            180.0F,
            90.0F)

        Path.AddArc(
            Rect.Right - Diameter,
            Rect.Top,
            Diameter,
            Diameter,
            270.0F,
            90.0F)

        Path.AddArc(
            Rect.Right - Diameter,
            Rect.Bottom - Diameter,
            Diameter,
            Diameter,
            0.0F,
            90.0F)

        Path.AddArc(
            Rect.Left,
            Rect.Bottom - Diameter,
            Diameter,
            Diameter,
            90.0F,
            90.0F)

        Path.CloseFigure()

        Return Path
    End Function

End Class
