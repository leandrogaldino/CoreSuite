Imports System.ComponentModel
Imports System.Drawing.Drawing2D
''' <summary>
''' Represents the run-time surface used internally to block and paint over the configured target.
''' </summary>
<DesignerCategory("Code")>
<ToolboxItem(False)>
Friend NotInheritable Class BusyOverlayView
    Inherits Control
    Private Const OuterMargin As Integer = 12
    Private ReadOnly _Owner As BusyOverlay
    Private ReadOnly _AnimationTimer As System.Windows.Forms.Timer
    Private ReadOnly _CancelButton As Button
    Private _Snapshot As Bitmap
    Private _AnimationFrame As Integer
    ''' <summary>
    ''' Occurs when the internal cancellation button is selected.
    ''' </summary>
    Friend Event CancellationClick As EventHandler
    ''' <summary>
    ''' Initializes a new view owned by the specified component.
    ''' </summary>
    ''' <param name="Owner">The component that supplies content and appearance settings.</param>
    Public Sub New(Owner As BusyOverlay)
        ArgumentNullException.ThrowIfNull(Owner)
        _Owner = Owner
        _AnimationTimer = New System.Windows.Forms.Timer()
        AddHandler _AnimationTimer.Tick, AddressOf AnimationTimer_Tick
        _CancelButton = New Button With {.FlatStyle = FlatStyle.System, .TabStop = False, .UseVisualStyleBackColor = True, .Visible = False}
        AddHandler _CancelButton.Click, AddressOf CancelButton_Click
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.SupportsTransparentBackColor Or ControlStyles.Selectable, True)
        BackColor = Color.Transparent
        TabStop = False
        AccessibleRole = AccessibleRole.Pane
        Controls.Add(_CancelButton)
        Visible = False
    End Sub
    ''' <summary>
    ''' Applies the owner's current settings and refreshes layout and animation.
    ''' </summary>
    Friend Sub ApplySettings()
        _AnimationTimer.Interval = _Owner.AnimationInterval
        _CancelButton.Text = _Owner.CancelButtonText
        _CancelButton.Font = _Owner.DetailFont
        _CancelButton.Size = _Owner.CancelButtonSize
        _CancelButton.Visible = _Owner.CanCancel
        _CancelButton.Enabled = _Owner.CanCancel
        _CancelButton.Cursor = Cursors.Default
        Cursor = If(_Owner.UseWaitCursor, Cursors.WaitCursor, Cursors.Default)
        AccessibleName = If(String.IsNullOrWhiteSpace(_Owner.MessageText), "Busy", _Owner.MessageText)
        UpdateAnimationState()
        PerformLayout()
        Invalidate()
    End Sub
    ''' <summary>
    ''' Replaces the visual snapshot drawn beneath the overlay tint.
    ''' </summary>
    ''' <param name="Snapshot">The new owned snapshot.</param>
    Friend Sub SetSnapshot(Snapshot As Bitmap)
        ClearSnapshot()
        _Snapshot = Snapshot
        Invalidate()
    End Sub
    ''' <summary>
    ''' Releases the current target snapshot.
    ''' </summary>
    Friend Sub ClearSnapshot()
        If _Snapshot Is Nothing Then Return
        _Snapshot.Dispose()
        _Snapshot = Nothing
        Invalidate()
    End Sub
    ''' <summary>
    ''' Releases the timer, button subscriptions, and captured target image.
    ''' </summary>
    ''' <param name="Disposing"><see langword="True"/> to release managed resources; otherwise, <see langword="False"/>.</param>
    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing Then
            _AnimationTimer.Stop()
            RemoveHandler _AnimationTimer.Tick, AddressOf AnimationTimer_Tick
            _AnimationTimer.Dispose()
            RemoveHandler _CancelButton.Click, AddressOf CancelButton_Click
            ClearSnapshot()
        End If
        MyBase.Dispose(Disposing)
    End Sub
    ''' <summary>
    ''' Updates animation when the surface becomes visible or hidden.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnVisibleChanged(e As EventArgs)
        MyBase.OnVisibleChanged(e)
        UpdateAnimationState()
    End Sub
    ''' <summary>
    ''' Positions the cancellation button in the calculated content layout.
    ''' </summary>
    ''' <param name="Levent">The layout event data.</param>
    Protected Overrides Sub OnLayout(Levent As LayoutEventArgs)
        MyBase.OnLayout(Levent)
        Dim layout As OverlayLayout = CalculateLayout()
        If layout.HasCancelButton Then _CancelButton.Bounds = layout.CancelButtonBounds
    End Sub
    ''' <summary>
    ''' Draws the target snapshot, overlay tint, centered content, and current indicator.
    ''' </summary>
    ''' <param name="e">The paint event data.</param>
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
        e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality
        If _Snapshot IsNot Nothing Then e.Graphics.DrawImage(_Snapshot, ClientRectangle)
        Using overlayBrush As New SolidBrush(Color.FromArgb(_Owner.OverlayOpacity, _Owner.OverlayColor))
            e.Graphics.FillRectangle(overlayBrush, ClientRectangle)
        End Using
        Dim layout As OverlayLayout = CalculateLayout()
        If _Owner.ShowContentPanel AndAlso layout.ContentPanelBounds.Width > 0 AndAlso layout.ContentPanelBounds.Height > 0 Then DrawContentPanel(e.Graphics, layout.ContentPanelBounds)
        DrawIndicator(e.Graphics, layout)
        Dim textFlags As TextFormatFlags = TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter Or TextFormatFlags.WordBreak Or TextFormatFlags.NoPadding
        If layout.HasMessage Then TextRenderer.DrawText(e.Graphics, _Owner.MessageText, _Owner.MessageFont, layout.MessageBounds, _Owner.MessageForeColor, textFlags)
        If layout.HasDetail Then TextRenderer.DrawText(e.Graphics, _Owner.DetailText, _Owner.DetailFont, layout.DetailBounds, _Owner.DetailForeColor, textFlags)
        If layout.HasPercentage Then TextRenderer.DrawText(e.Graphics, $"{_Owner.ProgressPercentage:0}%", _Owner.DetailFont, layout.PercentageBounds, _Owner.DetailForeColor, textFlags)
    End Sub
    Private Sub DrawContentPanel(Graphics As Graphics, Bounds As Rectangle)
        Using Path As GraphicsPath = CreateRoundedRectangle(Bounds, _Owner.ContentCornerRadius)
            Using BackgroundBrush As New SolidBrush(Color.FromArgb(_Owner.ContentOpacity, _Owner.ContentBackColor))
                Graphics.FillPath(BackgroundBrush, Path)
            End Using
            If _Owner.ContentBorderThickness > 0 Then
                Using BorderPen As New Pen(_Owner.ContentBorderColor, _Owner.ContentBorderThickness)
                    BorderPen.Alignment = PenAlignment.Inset
                    Graphics.DrawPath(BorderPen, Path)
                End Using
            End If
        End Using
    End Sub
    Private Sub DrawIndicator(Graphics As Graphics, Layout As OverlayLayout)
        Select Case _Owner.IndicatorStyle
            Case BusyOverlayIndicatorStyle.Spinner
                If Not Layout.HasIndicator Then Return
                Dim Inset As Single = _Owner.IndicatorThickness / 2.0F
                Dim SpinnerBounds As New RectangleF(Layout.IndicatorBounds.X + Inset, Layout.IndicatorBounds.Y + Inset, Layout.IndicatorBounds.Width - _Owner.IndicatorThickness, Layout.IndicatorBounds.Height - _Owner.IndicatorThickness)
                Using SpinnerPen As New Pen(_Owner.IndicatorColor, _Owner.IndicatorThickness)
                    SpinnerPen.StartCap = LineCap.Round
                    SpinnerPen.EndCap = LineCap.Round
                    Graphics.DrawArc(SpinnerPen, SpinnerBounds, _AnimationFrame * 12.0F, 265.0F)
                End Using
            Case BusyOverlayIndicatorStyle.MarqueeBar
                If Not Layout.HasIndicator Then Return
                DrawBarTrack(Graphics, Layout.IndicatorBounds)
                Dim SegmentWidth As Integer = Math.Max(Layout.IndicatorBounds.Height, Layout.IndicatorBounds.Width \ 3)
                Dim TravelWidth As Integer = Layout.IndicatorBounds.Width + SegmentWidth
                Dim SegmentX As Integer = Layout.IndicatorBounds.X - SegmentWidth + CInt(TravelWidth * (_AnimationFrame / 100.0R))
                Dim SegmentBounds As Rectangle = Rectangle.Intersect(Layout.IndicatorBounds, New Rectangle(SegmentX, Layout.IndicatorBounds.Y, SegmentWidth, Layout.IndicatorBounds.Height))
                If SegmentBounds.Width > 0 Then FillRoundedBar(Graphics, SegmentBounds, _Owner.IndicatorColor)
            Case BusyOverlayIndicatorStyle.ProgressBar
                If Not Layout.HasIndicator Then Return
                DrawBarTrack(Graphics, Layout.IndicatorBounds)
                Dim ProgressWidth As Integer = CInt(Math.Round(Layout.IndicatorBounds.Width * _Owner.ProgressPercentage / 100.0R))
                If ProgressWidth > 0 Then FillRoundedBar(Graphics, New Rectangle(Layout.IndicatorBounds.X, Layout.IndicatorBounds.Y, ProgressWidth, Layout.IndicatorBounds.Height), _Owner.IndicatorColor)
        End Select
    End Sub
    Private Sub DrawBarTrack(Graphics As Graphics, Bounds As Rectangle)
        FillRoundedBar(Graphics, Bounds, _Owner.IndicatorTrackColor)
    End Sub
    Private Shared Sub FillRoundedBar(Graphics As Graphics, Bounds As Rectangle, Color As Color)
        If Bounds.Width <= 0 OrElse Bounds.Height <= 0 Then Return
        Using Path As GraphicsPath = CreateRoundedRectangle(Bounds, Bounds.Height \ 2)
            Using brush As New SolidBrush(Color)
                Graphics.FillPath(brush, Path)
            End Using
        End Using
    End Sub
    Private Function CalculateLayout() As OverlayLayout
        Dim Result As New OverlayLayout()
        If ClientSize.Width <= 0 OrElse ClientSize.Height <= 0 Then Return Result
        Dim EffectivePadding As Integer = If(_Owner.ShowContentPanel, _Owner.ContentPadding, 0)
        Dim AvailablePanelWidth As Integer = Math.Max(1, ClientSize.Width - OuterMargin * 2)
        Dim PanelMaximumWidth As Integer = Math.Min(_Owner.ContentMaximumWidth, AvailablePanelWidth)
        Dim InnerMaximumWidth As Integer = Math.Max(1, PanelMaximumWidth - EffectivePadding * 2)
        Dim TextMeasureFlags As TextFormatFlags = TextFormatFlags.SingleLine Or TextFormatFlags.NoPadding
        Dim MessagePreferredWidth As Integer = If(String.IsNullOrWhiteSpace(_Owner.MessageText), 0, TextRenderer.MeasureText(_Owner.MessageText, _Owner.MessageFont, New Size(100000, 100000), TextMeasureFlags).Width)
        Dim DetailPreferredWidth As Integer = If(String.IsNullOrWhiteSpace(_Owner.DetailText), 0, TextRenderer.MeasureText(_Owner.DetailText, _Owner.DetailFont, New Size(100000, 100000), TextMeasureFlags).Width)
        Dim IndicatorPreferredWidth As Integer = 0
        Select Case _Owner.IndicatorStyle
            Case BusyOverlayIndicatorStyle.Spinner
                IndicatorPreferredWidth = _Owner.IndicatorSize
            Case BusyOverlayIndicatorStyle.MarqueeBar, BusyOverlayIndicatorStyle.ProgressBar
                IndicatorPreferredWidth = _Owner.ProgressBarWidth
        End Select
        Dim CancelPreferredWidth As Integer = If(_Owner.CanCancel, _Owner.CancelButtonSize.Width, 0)
        Dim DesiredInnerWidth As Integer = Math.Max(Math.Max(MessagePreferredWidth, DetailPreferredWidth), Math.Max(IndicatorPreferredWidth, CancelPreferredWidth))
        DesiredInnerWidth = Math.Max(Math.Min(160, InnerMaximumWidth), DesiredInnerWidth)
        Dim InnerWidth As Integer = Math.Min(InnerMaximumWidth, DesiredInnerWidth)
        Dim WrappedTextFlags As TextFormatFlags = TextFormatFlags.HorizontalCenter Or TextFormatFlags.WordBreak Or TextFormatFlags.NoPadding
        Dim MessageSize As Size = If(String.IsNullOrWhiteSpace(_Owner.MessageText), Size.Empty, TextRenderer.MeasureText(_Owner.MessageText, _Owner.MessageFont, New Size(InnerWidth, 100000), WrappedTextFlags))
        Dim DetailSize As Size = If(String.IsNullOrWhiteSpace(_Owner.DetailText), Size.Empty, TextRenderer.MeasureText(_Owner.DetailText, _Owner.DetailFont, New Size(InnerWidth, 100000), WrappedTextFlags))
        Dim PercentageSize As Size = If(_Owner.IndicatorStyle = BusyOverlayIndicatorStyle.ProgressBar AndAlso _Owner.ShowProgressPercentage, TextRenderer.MeasureText("100%", _Owner.DetailFont, New Size(InnerWidth, 100000), TextMeasureFlags), Size.Empty)
        Dim IndicatorHeight As Integer = 0
        Select Case _Owner.IndicatorStyle
            Case BusyOverlayIndicatorStyle.Spinner
                IndicatorHeight = _Owner.IndicatorSize
            Case BusyOverlayIndicatorStyle.MarqueeBar, BusyOverlayIndicatorStyle.ProgressBar
                IndicatorHeight = _Owner.ProgressBarHeight
        End Select
        Dim ItemHeights As New List(Of Integer)
        If IndicatorHeight > 0 Then ItemHeights.Add(IndicatorHeight)
        If MessageSize.Height > 0 Then ItemHeights.Add(MessageSize.Height)
        If DetailSize.Height > 0 Then ItemHeights.Add(DetailSize.Height)
        If PercentageSize.Height > 0 Then ItemHeights.Add(PercentageSize.Height)
        If _Owner.CanCancel Then ItemHeights.Add(_Owner.CancelButtonSize.Height)
        Dim ContentHeight As Integer = ItemHeights.Sum()
        If ItemHeights.Count > 1 Then ContentHeight += (ItemHeights.Count - 1) * _Owner.ContentSpacing
        Dim PanelWidth As Integer = InnerWidth + EffectivePadding * 2
        Dim PanelHeight As Integer = ContentHeight + EffectivePadding * 2
        Dim PanelX As Integer = Math.Max(0, (ClientSize.Width - PanelWidth) \ 2)
        Dim PanelY As Integer = Math.Max(0, (ClientSize.Height - PanelHeight) \ 2)
        Result.ContentPanelBounds = New Rectangle(PanelX, PanelY, Math.Min(PanelWidth, ClientSize.Width), Math.Min(PanelHeight, ClientSize.Height))
        Dim ContentX As Integer = PanelX + EffectivePadding
        Dim CurrentY As Integer = PanelY + EffectivePadding
        Dim ItemsRemaining As Integer = ItemHeights.Count
        If IndicatorHeight > 0 Then
            Dim IndicatorWidth As Integer = If(_Owner.IndicatorStyle = BusyOverlayIndicatorStyle.Spinner, _Owner.IndicatorSize, Math.Min(_Owner.ProgressBarWidth, InnerWidth))
            Result.IndicatorBounds = New Rectangle(ContentX + (InnerWidth - IndicatorWidth) \ 2, CurrentY, IndicatorWidth, IndicatorHeight)
            Result.HasIndicator = True
            CurrentY += IndicatorHeight
            ItemsRemaining -= 1
            If ItemsRemaining > 0 Then CurrentY += _Owner.ContentSpacing
        End If
        If MessageSize.Height > 0 Then
            Result.MessageBounds = New Rectangle(ContentX, CurrentY, InnerWidth, MessageSize.Height)
            Result.HasMessage = True
            CurrentY += MessageSize.Height
            ItemsRemaining -= 1
            If ItemsRemaining > 0 Then CurrentY += _Owner.ContentSpacing
        End If
        If DetailSize.Height > 0 Then
            Result.DetailBounds = New Rectangle(ContentX, CurrentY, InnerWidth, DetailSize.Height)
            Result.HasDetail = True
            CurrentY += DetailSize.Height
            ItemsRemaining -= 1
            If ItemsRemaining > 0 Then CurrentY += _Owner.ContentSpacing
        End If
        If PercentageSize.Height > 0 Then
            Result.PercentageBounds = New Rectangle(ContentX, CurrentY, InnerWidth, PercentageSize.Height)
            Result.HasPercentage = True
            CurrentY += PercentageSize.Height
            ItemsRemaining -= 1
            If ItemsRemaining > 0 Then CurrentY += _Owner.ContentSpacing
        End If
        If _Owner.CanCancel Then
            Result.CancelButtonBounds = New Rectangle(ContentX + (InnerWidth - _Owner.CancelButtonSize.Width) \ 2, CurrentY, _Owner.CancelButtonSize.Width, _Owner.CancelButtonSize.Height)
            Result.HasCancelButton = True
        End If
        Return Result
    End Function
    Private Shared Function CreateRoundedRectangle(Bounds As Rectangle, Radius As Integer) As GraphicsPath
        Dim Path As New GraphicsPath()
        If Bounds.Width <= 0 OrElse Bounds.Height <= 0 Then Return Path
        Dim EffectiveRadius As Integer = Math.Max(0, Math.Min(Radius, Math.Min(Bounds.Width, Bounds.Height) \ 2))
        If EffectiveRadius = 0 Then
            Path.AddRectangle(Bounds)
            Return Path
        End If
        Dim Diameter As Integer = EffectiveRadius * 2
        Path.AddArc(Bounds.Left, Bounds.Top, Diameter, Diameter, 180, 90)
        Path.AddArc(Bounds.Right - Diameter, Bounds.Top, Diameter, Diameter, 270, 90)
        Path.AddArc(Bounds.Right - Diameter, Bounds.Bottom - Diameter, Diameter, Diameter, 0, 90)
        Path.AddArc(Bounds.Left, Bounds.Bottom - Diameter, Diameter, Diameter, 90, 90)
        Path.CloseFigure()
        Return Path
    End Function
    Private Sub UpdateAnimationState()
        If _AnimationTimer Is Nothing OrElse _Owner Is Nothing OrElse IsDisposed OrElse Disposing Then Return
        _AnimationTimer.Enabled = Visible AndAlso (_Owner.IndicatorStyle = BusyOverlayIndicatorStyle.Spinner OrElse _Owner.IndicatorStyle = BusyOverlayIndicatorStyle.MarqueeBar)
    End Sub
    Private Sub AnimationTimer_Tick(sender As Object, e As EventArgs)
        If _Owner.IndicatorStyle = BusyOverlayIndicatorStyle.Spinner Then
            _AnimationFrame = (_AnimationFrame + 1) Mod 30
        Else
            _AnimationFrame = (_AnimationFrame + 4) Mod 101
        End If
        Invalidate()
    End Sub
    Private Sub CancelButton_Click(sender As Object, e As EventArgs)
        RaiseEvent CancellationClick(Me, EventArgs.Empty)
    End Sub
    Private Structure OverlayLayout
        Public ContentPanelBounds As Rectangle
        Public IndicatorBounds As Rectangle
        Public MessageBounds As Rectangle
        Public DetailBounds As Rectangle
        Public PercentageBounds As Rectangle
        Public CancelButtonBounds As Rectangle
        Public HasIndicator As Boolean
        Public HasMessage As Boolean
        Public HasDetail As Boolean
        Public HasPercentage As Boolean
        Public HasCancelButton As Boolean
    End Structure
End Class
