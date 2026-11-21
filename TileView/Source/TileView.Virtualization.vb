Imports System.Drawing
Imports System.Windows.Forms

Public Partial Class TileView
    Private ReadOnly _Realized As New Dictionary(Of Integer, TileViewItem)()
    Private ReadOnly _TilePool As New List(Of TileViewItem)()
    Private ReadOnly _FreeTiles As New Stack(Of TileViewItem)()
    Private _Columns As Integer = 1
    Private _TileWidth As Integer = 260
    Private _LastColumnExtraWidth As Integer
    Private _RowCount As Integer
    Private _ContentHeight As Integer
    Private _LayoutDirty As Boolean = True
    Private _Refreshing As Boolean
    Private _RefreshPending As Boolean

    Private Sub InvalidateVirtualLayout()
        CancelScrollPreview()
        _LayoutDirty = True
        UpdateViewport()
    End Sub

    Private Sub CalculateLayout()
        Dim Available As Integer = Math.Max(1, ClientSize.Width - _ContentPadding.Horizontal)
        Dim ColumnCount As Integer = DetermineColumns(Available)
        Dim RowCount As Integer = GetRowCount(ColumnCount)
        Dim TotalHeight As Integer = GetContentHeight(RowCount)
        If TotalHeight > ClientSize.Height Then
            Available = Math.Max(1, Available - _VerticalBar.Width)
            ColumnCount = DetermineColumns(Available)
            RowCount = GetRowCount(ColumnCount)
            TotalHeight = GetContentHeight(RowCount)
        End If
        _Columns = ColumnCount
        _RowCount = RowCount
        _ContentHeight = TotalHeight
        ' For fixed columns, limit each tile to its available column slot. When
        ' stretching, assign any remainder to the last column to fill the row.
        Dim GapWidth As Long = CLng(ColumnCount - 1) * _ItemSpacing
        Dim WidthForTiles As Long = CLng(Available) - GapWidth
        Dim SlotWidth As Integer = CInt(Math.Max(1L, WidthForTiles \ ColumnCount))
        _TileWidth = If(_StretchItems, SlotWidth, Math.Min(_ItemSize.Width, SlotWidth))
        _LastColumnExtraWidth = If(_StretchItems AndAlso WidthForTiles >= ColumnCount, CInt(WidthForTiles Mod ColumnCount), 0)
        _LayoutDirty = False
        ConfigureVirtualScrollBar(TotalHeight)
    End Sub

    Private Function DetermineColumns(AvailableWidth As Integer) As Integer
        If _RequestedColumns > 0 Then Return _RequestedColumns
        Dim Stride As Long = CLng(_ItemSize.Width) + _ItemSpacing
        Return CInt(Math.Max(1L, (CLng(AvailableWidth) + _ItemSpacing) \ Stride))
    End Function

    Private Function GetRowCount(Columns As Integer) As Integer
        If _VisibleSourceIndices.Count = 0 Then Return 0
        Return CInt((CLng(_VisibleSourceIndices.Count) + Columns - 1L) \ Columns)
    End Function

    Private Function GetContentHeight(Rows As Integer) As Integer
        If Rows = 0 Then Return 0
        Dim Height As Long = CLng(_ContentPadding.Top) + CLng(Rows) * (CLng(_ItemSize.Height) + _ItemSpacing) - _ItemSpacing + _ContentPadding.Bottom
        If Height > Integer.MaxValue - 128L Then
            Throw New InvalidOperationException("The virtual content exceeds the WinForms scrollbar coordinate range.")
        End If
        Return CInt(Height)
    End Function

    ''' <summary>
    ''' Rebuilds only the viewport and overscan rows. The number of visual controls is
    ''' independent of the number of data records.
    ''' </summary>
    Private Sub UpdateViewport()
        If _Realized Is Nothing OrElse _VisibleSourceIndices Is Nothing OrElse
           IsDisposed OrElse Disposing OrElse _UpdateDepth > 0 OrElse _PreviewActive Then Return
        If _Refreshing Then
            _RefreshPending = True
            Return
        End If
        _Refreshing = True
        Try
            Do
                _RefreshPending = False
                If _LayoutDirty Then CalculateLayout()
                RealizeViewport()
            Loop While _RefreshPending
        Finally
            _Refreshing = False
        End Try
    End Sub

    Private Sub RealizeViewport()
        If _Factory Is Nothing OrElse _VisibleSourceIndices.Count = 0 OrElse ClientSize.Width <= 0 OrElse ClientSize.Height <= 0 Then
            ReleaseAllTiles()
            RaiseEvent ViewportChanged(Me, EventArgs.Empty)
            Return
        End If

        Dim Stride As Long = CLng(_ItemSize.Height) + _ItemSpacing
        Dim ScrollTop As Integer = GetScrollTop()
        Dim FirstRow As Integer = CInt(Math.Max(0L, (CLng(ScrollTop) - _ContentPadding.Top) \ Stride - _OverscanRows))
        Dim LastRow As Integer = CInt(Math.Min(CLng(_RowCount - 1), (CLng(ScrollTop) + ClientSize.Height - _ContentPadding.Top) \ Stride + _OverscanRows))
        If LastRow < FirstRow Then
            ReleaseAllTiles()
            RaiseEvent ViewportChanged(Me, EventArgs.Empty)
            Return
        End If
        Dim First As Integer = FirstRow * _Columns
        Dim Last As Integer = CInt(Math.Min(CLng(_VisibleSourceIndices.Count - 1), (CLng(LastRow) + 1L) * _Columns - 1L))
        SuspendLayout()
        Try
            For Each Pair In _Realized.ToArray()
                If Pair.Key < First OrElse Pair.Key > Last Then ReleaseTile(Pair.Key, Pair.Value)
            Next
            For i As Integer = First To Last
                Dim Tile As TileViewItem = Nothing
                If Not _Realized.TryGetValue(i, Tile) Then
                    Tile = AcquireTile(i)
                End If
                Dim RowNumber As Integer = i \ _Columns
                Dim ColumnNumber As Integer = i Mod _Columns
                Dim X As Integer = _ContentPadding.Left + ColumnNumber * (_TileWidth + _ItemSpacing)
                Dim Y As Integer = CInt(CLng(_ContentPadding.Top) + CLng(RowNumber) * (CLng(_ItemSize.Height) + _ItemSpacing) - ScrollTop)
                Dim TileWidth As Integer = _TileWidth
                If _StretchItems AndAlso ColumnNumber = _Columns - 1 Then
                    TileWidth += _LastColumnExtraWidth
                End If
                Dim Target As New Rectangle(X, Y, TileWidth, _ItemSize.Height)
                If Tile.Bounds <> Target Then Tile.Bounds = Target
                If Not Tile.Visible Then Tile.Visible = True
                Tile.SetSelectedInternal(_VisibleSourceIndices(i) = _SelectedSourceIndex)
            Next
        Finally
            ResumeLayout(False)
        End Try
        Invalidate()
        RaiseEvent ViewportChanged(Me, EventArgs.Empty)
    End Sub

    Private Function AcquireTile(DisplayIndex As Integer) As TileViewItem
        Dim Tile As TileViewItem
        If _FreeTiles.Count > 0 Then
            Tile = _FreeTiles.Pop()
        Else
            Tile = _Factory()
            If Tile Is Nothing Then Throw New InvalidOperationException("Tile factory returned Nothing.")
            If Tile.IsDisposed OrElse Tile.Parent IsNot Nothing OrElse _TilePool.Contains(Tile) Then
                Throw New InvalidOperationException("Tile factory must return a new, unparented TileViewItem.")
            End If
            Tile.Visible = False
            Tile.SetOwner(Me)
            Tile.SetReadOnlyInternal(_ReadOnly)
            _TilePool.Add(Tile)
            Controls.Add(Tile)
            _VerticalBar.BringToFront()
            RaiseEvent TileCreated(Me, New TileViewItemEventArgs(Tile))
        End If

        Dim SourceIndex As Integer = _VisibleSourceIndices(DisplayIndex)
        Try
            Tile.BindInternal(_Data(SourceIndex), SourceIndex, DisplayIndex)
            _Binder(Tile, _Data(SourceIndex))
            _Realized.Add(DisplayIndex, Tile)
        Catch
            Tile.UnbindInternal()
            Tile.Visible = False
            _FreeTiles.Push(Tile)
            Throw
        End Try
        Return Tile
    End Function

    Private Sub ReleaseTile(DisplayIndex As Integer, Tile As TileViewItem)
        If Tile.ContainsFocus Then Focus()
        _Realized.Remove(DisplayIndex)
        Tile.Visible = False
        Tile.UnbindInternal()
        _FreeTiles.Push(Tile)
    End Sub

    Private Sub ReleaseAllTiles()
        For Each Pair In _Realized.ToArray()
            ReleaseTile(Pair.Key, Pair.Value)
        Next
    End Sub

    Private Sub DisposeTilePool()
        DisposePreviewTemplate()
        _Realized.Clear()
        _FreeTiles.Clear()
        For Each Tile In _TilePool
            Tile.UnbindInternal()
            Tile.SetOwner(Nothing)
            Controls.Remove(Tile)
            Tile.Dispose()
        Next
        _TilePool.Clear()
    End Sub

    Private Function GetScrollTop() As Integer
        Return _ScrollOffset
    End Function

    ''' <summary>Scrolls a filtered record into view, realizing its tile as necessary.</summary>
    Public Sub EnsureVisible(DisplayIndex As Integer)
        FinishScrollPreview()
        If DisplayIndex < 0 OrElse DisplayIndex >= _VisibleSourceIndices.Count Then
            Throw New ArgumentOutOfRangeException(NameOf(DisplayIndex))
        End If
        If _LayoutDirty Then UpdateViewport()
        Dim RowNumber As Integer = DisplayIndex \ _Columns
        Dim Top As Long = CLng(_ContentPadding.Top) + CLng(RowNumber) * (CLng(_ItemSize.Height) + _ItemSpacing)
        Dim ScrollTop As Integer = GetScrollTop()
        Dim Target As Long = ScrollTop
        If Top < ScrollTop Then
            Target = Top
        ElseIf Top + _ItemSize.Height > CLng(ScrollTop) + ClientSize.Height Then
            Target = Top + _ItemSize.Height - ClientSize.Height
        End If
        Target = Math.Max(0L, Math.Min(Target, CLng(Math.Max(0, _ContentHeight - ClientSize.Height))))
        If Target <> ScrollTop Then ScrollToOffset(CInt(Target))
        UpdateViewport()
    End Sub

    Protected Overrides Sub OnMouseWheel(e As MouseEventArgs)
        If _VerticalBar IsNot Nothing AndAlso _VerticalBar.Visible Then
            ScrollByWheelDelta(e.Delta)
            Dim Handled As HandledMouseEventArgs = TryCast(e, HandledMouseEventArgs)
            If Handled IsNot Nothing Then Handled.Handled = True
        End If
        MyBase.OnMouseWheel(e)
    End Sub

    Protected Overrides Sub OnResize(e As EventArgs)
        If _PreviewFrameTimer IsNot Nothing Then CancelScrollPreview()
        MyBase.OnResize(e)
        If _Realized IsNot Nothing AndAlso _VisibleSourceIndices IsNot Nothing Then InvalidateVirtualLayout()
    End Sub

    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        If _Realized IsNot Nothing AndAlso _VisibleSourceIndices IsNot Nothing Then InvalidateVirtualLayout()
    End Sub

    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        Dim ClickIndex As Integer = -1
        If _PreviewActive AndAlso e.Button = MouseButtons.Left Then
            ClickIndex = HitTestDisplayIndex(e.Location)
            FinishScrollPreview()
        End If
        If e.Button = MouseButtons.Left Then
            If ClickIndex >= 0 Then
                SelectIndex(ClickIndex)
            Else
                If _ClearSelectionOnBlankClick Then ClearSelection()
                If CanFocus Then Focus()
            End If
        End If
        MyBase.OnMouseDown(e)
    End Sub
End Class
