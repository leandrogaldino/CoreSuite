Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Windows.Forms
Public Partial Class TileView
    Private Const PreviewIntervalMilliseconds As Integer = 30
    Private Const ScrollSettlingMilliseconds As Integer = 160
    Private Const MaximumPreviewCacheEntries As Integer = 64
    Private Const MaximumPreviewCacheBytes As Long = 24L * 1024L * 1024L
    Private ReadOnly _PreviewFrameTimer As New Timer() With {.Interval = PreviewIntervalMilliseconds}
    Private ReadOnly _PreviewIdleTimer As New Timer() With {.Interval = ScrollSettlingMilliseconds}
    Private ReadOnly _PreviewCache As New Dictionary(Of Integer, Bitmap)()
    Private ReadOnly _PreviewCacheOrder As New Queue(Of Integer)()
    Private _PreviewCacheBytes As Long
    Private _PreviewTemplate As TileViewItem
    Private _PreviewActive As Boolean
    Private _PreviewDirty As Boolean

    Private Sub InitializePreviewTimers()
        AddHandler _PreviewFrameTimer.Tick, AddressOf PreviewFrameTimer_Tick
        AddHandler _PreviewIdleTimer.Tick, AddressOf PreviewIdleTimer_Tick
    End Sub

    ''' <summary>Hide native windows only once per gesture; do not move or recycle them during scroll tracking.</summary>
    Private Sub BeginScrollPreview()
        If IsDisposed OrElse Disposing Then Return
        If Not _PreviewActive Then
            _PreviewActive = True
            If ContainsFocus Then Focus()
            SuspendLayout()
            Try
                For Each Tile As TileViewItem In _Realized.Values
                    Tile.Visible = False
                Next
            Finally
                ResumeLayout(False)
            End Try
            ClearPreviewCache()
            Invalidate()
            RaiseEvent ViewportChanged(Me, EventArgs.Empty)
        End If
        _PreviewDirty = True
        If Not _PreviewFrameTimer.Enabled Then _PreviewFrameTimer.Start()
        If Not _ThumbTracking Then
            _PreviewIdleTimer.Stop()
            _PreviewIdleTimer.Start()
        End If
    End Sub

    Private Sub PreviewFrameTimer_Tick(sender As Object, e As EventArgs)
        If Not _PreviewActive Then
            _PreviewFrameTimer.Stop()
            Return
        End If
        If _PreviewDirty Then
            _PreviewDirty = False
            Invalidate(New Rectangle(0, 0, Math.Max(1, ClientSize.Width - If(_VerticalBar.Visible, _VerticalBar.Width, 0)), ClientSize.Height))
        End If
    End Sub

    Private Sub PreviewIdleTimer_Tick(sender As Object, e As EventArgs)
        If _ThumbTracking Then Return
        _PreviewIdleTimer.Stop()
        FinishScrollPreview()
    End Sub

    ''' <summary>Finish a gesture and restore the live interactive tile controls.</summary>
    Private Sub FinishScrollPreview()
        If Not _PreviewActive Then Return
        CancelScrollPreview()
        If Not IsDisposed AndAlso Not Disposing Then
            UpdateViewport()
            Invalidate()
        End If
    End Sub

    ''' <summary>Cancel preview before changing data/layout or disposing; callers will update the viewport.</summary>
    Private Sub CancelScrollPreview()
        _PreviewActive = False
        _PreviewDirty = False
        _ThumbTracking = False
        _PreviewFrameTimer.Stop()
        _PreviewIdleTimer.Stop()
        ClearPreviewCache()
    End Sub

    Private Sub DisposeScrollPreview()
        CancelScrollPreview()
        RemoveHandler _PreviewFrameTimer.Tick, AddressOf PreviewFrameTimer_Tick
        RemoveHandler _PreviewIdleTimer.Tick, AddressOf PreviewIdleTimer_Tick
        _PreviewFrameTimer.Dispose()
        _PreviewIdleTimer.Dispose()
        DisposePreviewTemplate()
    End Sub

    Private Sub ClearPreviewCache()
        For Each Image As Bitmap In _PreviewCache.Values
            Image.Dispose()
        Next
        _PreviewCache.Clear()
        _PreviewCacheOrder.Clear()
        _PreviewCacheBytes = 0
    End Sub

    Private Sub DisposePreviewTemplate()
        ClearPreviewCache()
        If _PreviewTemplate Is Nothing Then Return
        _PreviewTemplate.UnbindInternal()
        _PreviewTemplate.Dispose()
        _PreviewTemplate = Nothing
    End Sub

    ''' <summary>
    ''' Render all virtual items onto a single surface during scrolling. A custom
    ''' painter is preferred; the designer-based UserControl snapshot is the fallback.
    ''' </summary>
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        If Not _PreviewActive OrElse _Factory Is Nothing OrElse _VisibleSourceIndices.Count = 0 Then Return
        Dim DrawingWidth As Integer = ClientSize.Width - If(_VerticalBar.Visible, _VerticalBar.Width, 0)
        If DrawingWidth <= 0 OrElse ClientSize.Height <= 0 OrElse _rowCount = 0 Then Return

        Dim Saved As Drawing2D.GraphicsState = e.Graphics.Save()
        Try
            e.Graphics.SetClip(New Rectangle(0, 0, DrawingWidth, ClientSize.Height))
            DrawPreview(e.Graphics, DrawingWidth)
        Finally
            e.Graphics.Restore(Saved)
        End Try
    End Sub

    Private Sub DrawPreview(Graphics As Graphics, DrawingWidth As Integer)
        Dim Stride As Long = CLng(_ItemSize.Height) + _ItemSpacing
        Dim FirstRow As Long = Math.Max(0L, (CLng(_ScrollOffset) - _ContentPadding.Top) \ Stride)
        Dim LastRow As Long = Math.Min(CLng(_rowCount - 1),
                                     (CLng(_ScrollOffset) + ClientSize.Height - _ContentPadding.Top) \ Stride)
        If LastRow < FirstRow Then Return

        Dim FirstIndex As Long = FirstRow * _Columns
        Dim FinalIndex As Long = Math.Min(CLng(_VisibleSourceIndices.Count - 1), (LastRow + 1L) * _Columns - 1L)
        For Index As Long = FirstIndex To FinalIndex
            Dim DisplayIndex As Integer = CInt(Index)
            Dim TileBounds As Rectangle = GetTileBounds(DisplayIndex)
            If TileBounds.Width <= 0 OrElse TileBounds.Height <= 0 OrElse TileBounds.Right <= 0 OrElse
               TileBounds.Left >= DrawingWidth OrElse TileBounds.Bottom <= 0 OrElse TileBounds.Top >= ClientSize.Height Then Continue For

            Dim SourceIndex As Integer = _VisibleSourceIndices(DisplayIndex)
            Dim Selected As Boolean = SourceIndex = _SelectedSourceIndex
            If _PreviewPainter IsNot Nothing Then
                Dim State As Drawing2D.GraphicsState = Graphics.Save()
                Try
                    Graphics.SetClip(TileBounds, Drawing2D.CombineMode.Intersect)
                    _PreviewPainter(Graphics, TileBounds, _Data(SourceIndex), Selected)
                Finally
                    Graphics.Restore(State)
                End Try
            Else
                Dim Snapshot As Bitmap = GetPreviewSnapshot(DisplayIndex, SourceIndex, TileBounds.Size, Selected)
                Graphics.DrawImageUnscaled(Snapshot, TileBounds.Location)
                If Not _PreviewCache.ContainsKey(DisplayIndex) Then Snapshot.Dispose()
            End If
        Next
    End Sub

    Private Function GetTileBounds(DisplayIndex As Integer) As Rectangle
        Dim RowIndex As Long = DisplayIndex \ _Columns
        Dim ColumnIndex As Integer = DisplayIndex Mod _Columns
        Dim X As Long = CLng(_ContentPadding.Left) + CLng(ColumnIndex) * (_TileWidth + _ItemSpacing)
        Dim Y As Long = CLng(_ContentPadding.Top) + RowIndex * (CLng(_ItemSize.Height) + _ItemSpacing) - _ScrollOffset
        Dim Width As Integer = _TileWidth
        If _StretchItems AndAlso ColumnIndex = _Columns - 1 Then Width += _LastColumnExtraWidth
        Return New Rectangle(CInt(X), CInt(Y), Width, _ItemSize.Height)
    End Function

    Private Function GetPreviewSnapshot(DisplayIndex As Integer, SourceIndex As Integer, Size As Size, IsSelected As Boolean) As Bitmap
        Dim Wxisting As Bitmap = Nothing
        If _PreviewCache.TryGetValue(DisplayIndex, Wxisting) Then Return Wxisting

        If _PreviewTemplate Is Nothing Then
            _PreviewTemplate = _Factory()
            If _PreviewTemplate Is Nothing OrElse _PreviewTemplate.IsDisposed OrElse _PreviewTemplate.Parent IsNot Nothing Then
                Throw New InvalidOperationException("The tile factory must return a new, unparented TileViewItem for snapshots.")
            End If
            _PreviewTemplate.SetReadOnlyInternal(_ReadOnly)
        End If

        Dim Snapshot As New Bitmap(Math.Max(1, Size.Width), Math.Max(1, Size.Height), PixelFormat.Format32bppPArgb)
        Try
            _PreviewTemplate.Size = Size
            _PreviewTemplate.BindInternal(_Data(SourceIndex), SourceIndex, DisplayIndex)
            Try
                _Binder(_PreviewTemplate, _Data(SourceIndex))
                _PreviewTemplate.SetSelectedInternal(IsSelected)
                _PreviewTemplate.DrawToBitmap(Snapshot, New Rectangle(Point.Empty, Size))
            Finally
                _PreviewTemplate.UnbindInternal()
            End Try

            Dim Bytes As Long = CLng(Snapshot.Width) * Snapshot.Height * 4L
            If Bytes > MaximumPreviewCacheBytes Then Return Snapshot
            While _PreviewCacheOrder.Count >= MaximumPreviewCacheEntries OrElse
                  _PreviewCacheBytes + Bytes > MaximumPreviewCacheBytes
                Dim Oldest As Integer = _PreviewCacheOrder.Dequeue()
                Dim OldImage As Bitmap = _PreviewCache(Oldest)
                _PreviewCacheBytes -= CLng(OldImage.Width) * OldImage.Height * 4L
                _PreviewCache.Remove(Oldest)
                OldImage.Dispose()
            End While
            _PreviewCache.Add(DisplayIndex, Snapshot)
            _PreviewCacheOrder.Enqueue(DisplayIndex)
            _PreviewCacheBytes += Bytes
            Return Snapshot
        Catch
            Snapshot.Dispose()
            Throw
        End Try
    End Function

    Private Function HitTestDisplayIndex(Location As Point) As Integer
        Dim SurfaceWidth As Integer = ClientSize.Width - If(_VerticalBar.Visible, _VerticalBar.Width, 0)
        If Location.X < 0 OrElse Location.X >= SurfaceWidth OrElse Location.Y < 0 OrElse _rowCount = 0 Then Return -1
        Dim Y As Long = CLng(Location.Y) + _ScrollOffset - _ContentPadding.Top
        If Y < 0 Then Return -1
        Dim Row As Long = Y \ (CLng(_ItemSize.Height) + _ItemSpacing)
        If Row < 0 OrElse Row >= _rowCount Then Return -1
        For Column As Integer = 0 To _Columns - 1
            Dim candidate As Long = Row * _Columns + Column
            If candidate >= _visibleSourceIndices.Count Then Exit For
            If GetTileBounds(CInt(candidate)).Contains(Location) Then Return CInt(candidate)
        Next
        Return -1
    End Function
End Class
