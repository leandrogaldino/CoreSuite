Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms
''' <summary>
''' Represents the internal button that indicates selection, clears the lookup, or displays and cancels active searching.
''' </summary>
Friend Class LookupActionButton
    Inherits Control
    Private Const WM_MOUSEACTIVATE As Integer = &H21
    Private Const MA_NOACTIVATE As Integer = 3
    Private ReadOnly _AnimationTimer As Timer
    Private _IsSearching As Boolean
    Private _IsSelected As Boolean
    Private _AnimationAngle As Integer
    Private _Image As Image
    ''' <summary>
    ''' Initializes a new instance of the <see cref="LookupActionButton"/> class.
    ''' </summary>
    Public Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.Selectable, False)
        TabStop = False
        Text = String.Empty
        Cursor = Cursors.Hand
        AccessibleName = "Clear lookup text"
        AccessibleRole = AccessibleRole.PushButton
        _AnimationTimer = New Timer With {.Interval = 70}
        AddHandler _AnimationTimer.Tick, AddressOf AnimationTimerTick
    End Sub
    ''' <summary>
    ''' Gets or sets the optional image rendered by the action surface.
    ''' </summary>
    Friend Property Image As Image
        Get
            Return _Image
        End Get
        Set(value As Image)
            If ReferenceEquals(_Image, value) Then Return
            _Image = value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Prevents the embedded action surface from taking input focus when it is clicked.
    ''' </summary>
    ''' <param name="m">The Windows message to process.</param>
    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = WM_MOUSEACTIVATE Then
            m.Result = New IntPtr(MA_NOACTIVATE)
            Return
        End If
        MyBase.WndProc(m)
    End Sub
    ''' <summary>
    ''' Gets or sets a value indicating whether the button renders a progress indicator.
    ''' </summary>
    Friend Property IsSearching As Boolean
        Get
            Return _IsSearching
        End Get
        Set(value As Boolean)
            If _IsSearching = value Then Return
            _IsSearching = value
            UpdateAccessibleName()
            If value Then
                _AnimationTimer.Start()
            Else
                _AnimationTimer.Stop()
                _AnimationAngle = 0
            End If
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the button represents a selected lookup item.
    ''' </summary>
    Friend Property IsSelected As Boolean
        Get
            Return _IsSelected
        End Get
        Set(value As Boolean)
            If _IsSelected = value Then Return
            _IsSelected = value
            UpdateAccessibleName()
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Paints the progress indicator, assigned image, built-in selected-state check mark, or clear glyph.
    ''' </summary>
    Friend Sub RenderContent(G As Graphics, SurfaceSize As Size)
        G.SmoothingMode = SmoothingMode.AntiAlias
        If IsSearching Then
            PaintProgress(G, SurfaceSize)
            Return
        End If
        If Image IsNot Nothing Then
            EmbeddedTextButtonLayout.DrawCenteredImage(G, Image, SurfaceSize, CInt(Math.Round(3 * DeviceDpi / 96.0R)), Enabled, BackColor)
            Return
        End If
        Dim GlyphColor As Color = If(Enabled, ForeColor, SystemColors.GrayText)
        If IsSelected Then
            PaintSelectedGlyph(G, GlyphColor, SurfaceSize)
            Return
        End If
        Dim Margin As Single = Math.Max(5.0F, Math.Min(SurfaceSize.Width, SurfaceSize.Height) * 0.32F)
        Using GlyphPen As New Pen(GlyphColor, 1.6F)
            GlyphPen.StartCap = LineCap.Round
            GlyphPen.EndCap = LineCap.Round
            G.DrawLine(GlyphPen, Margin, Margin, SurfaceSize.Width - Margin, SurfaceSize.Height - Margin)
            G.DrawLine(GlyphPen, SurfaceSize.Width - Margin, Margin, Margin, SurfaceSize.Height - Margin)
        End Using
    End Sub
    ''' <summary>
    ''' Releases the animation timer used by the progress indicator.
    ''' </summary>
    ''' <param name="Disposing"><see langword="True"/> to release managed resources; otherwise, <see langword="False"/>.</param>
    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing Then
            _AnimationTimer.Stop()
            RemoveHandler _AnimationTimer.Tick, AddressOf AnimationTimerTick
            _AnimationTimer.Dispose()
        End If
        MyBase.Dispose(Disposing)
    End Sub
    Private Sub PaintProgress(Graphics As Graphics, SurfaceSize As Size)
        Dim IndicatorSize As Integer = Math.Max(6, Math.Min(SurfaceSize.Width, SurfaceSize.Height) - 10)
        Dim IndicatorBounds As New Rectangle((SurfaceSize.Width - IndicatorSize) \ 2, (SurfaceSize.Height - IndicatorSize) \ 2, IndicatorSize, IndicatorSize)
        Dim IndicatorColor As Color = If(Enabled, ForeColor, SystemColors.GrayText)
        Using IndicatorPen As New Pen(IndicatorColor, 1.8F)
            IndicatorPen.StartCap = LineCap.Round
            IndicatorPen.EndCap = LineCap.Round
            Graphics.DrawArc(IndicatorPen, IndicatorBounds, _AnimationAngle, 250)
        End Using
    End Sub
    Private Sub PaintSelectedGlyph(Graphics As Graphics, GlyphColor As Color, SurfaceSize As Size)
        Dim Left As Single = SurfaceSize.Width * 0.27F
        Dim MiddleX As Single = SurfaceSize.Width * 0.44F
        Dim MiddleY As Single = SurfaceSize.Height * 0.65F
        Dim Right As Single = SurfaceSize.Width * 0.75F
        Using GlyphPen As New Pen(GlyphColor, 2.0F)
            GlyphPen.StartCap = LineCap.Round
            GlyphPen.EndCap = LineCap.Round
            GlyphPen.LineJoin = LineJoin.Round
            Graphics.DrawLines(GlyphPen, {New PointF(Left, SurfaceSize.Height * 0.52F), New PointF(MiddleX, MiddleY), New PointF(Right, SurfaceSize.Height * 0.34F)})
        End Using
    End Sub
    Private Sub UpdateAccessibleName()
        If IsSearching Then
            AccessibleName = "Cancel lookup search"
        ElseIf IsSelected Then
            AccessibleName = "Clear selected lookup item"
        Else
            AccessibleName = "Clear lookup text"
        End If
    End Sub
    Private Sub AnimationTimerTick(Sender As Object, E As EventArgs)
        _AnimationAngle = (_AnimationAngle + 24) Mod 360
        Invalidate()
    End Sub
End Class
