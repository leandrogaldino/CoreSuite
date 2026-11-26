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
    Private _RowOffsets As Long()
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
        CalculateColumnsAndRows(Available)
        If _ContentHeight > ClientSize.Height Then
            Available = Math.Max(1, Available - _VerticalBar.Width)
            CalculateColumnsAndRows(Available)
        End If
        _LayoutDirty = False
        ConfigureVirtualScrollBar(_ContentHeight)
    End Sub

    Private Sub CalculateColumnsAndRows(Available As Integer)
        _Columns = DetermineColumns(Available)
        _RowCount = GetRowCount(_Columns)
        Dim GapWidth As Long = CLng(_Columns - 1) * _ItemSpacing
        Dim WidthForTiles As Long = CLng(Available) - GapWidth
        Dim SlotWidth As Integer = CInt(Math.Max(1L, WidthForTiles \ _Columns))
        _TileWidth = If(_StretchItems, SlotWidth, Math.Min(_ItemSize.Width, SlotWidth))
        _LastColumnExtraWidth = If(_StretchItems AndAlso WidthForTiles >= _Columns, CInt(WidthForTiles Mod _Columns), 0)
        _RowOffsets = Nothing
        If _MeasureHeight IsNot Nothing AndAlso _RowCount > 0 Then
            Dim Offsets(_RowCount) As Long
            For Row As Integer = 0 To _RowCount - 1
                Dim RowHeight As Integer = _ItemSize.Height
                Dim First As Integer = Row * _Columns
                Dim Last As Integer = CInt(Math.Min(CLng(_VisibleSourceIndices.Count - 1), CLng(First) + _Columns - 1L))
                For Index As Integer = First To Last
                    Dim Width As Integer = GetColumnWidth(Index Mod _Columns)
                    RowHeight = Math.Max(RowHeight, _MeasureHeight(_Data(_VisibleSourceIndices(Index)), Width, _ItemSize.Height))
                Next
                Offsets(Row + 1) = Offsets(Row) + CLng(RowHeight) + _ItemSpacing
            Next
            _RowOffsets = Offsets
        End If
        _ContentHeight = GetContentHeight(_RowCount)
    End Sub

    Private Function GetColumnWidth(Column As Integer) As Integer
        Return _TileWidth + If(_StretchItems AndAlso Column = _Columns - 1, _LastColumnExtraWidth, 0)
    End Function

    Private Function GetRowOffset(Row As Integer) As Long
        If _RowOffsets IsNot Nothing Then Return _RowOffsets(Row)
        Return CLng(Row) * (CLng(_ItemSize.Height) + _ItemSpacing)
    End Function

    Private Function GetRowHeight(Row As Integer) As Integer
        If _RowOffsets Is Nothing Then Return _ItemSize.Height
        Return CInt(_RowOffsets(Row + 1) - _RowOffsets(Row) - _ItemSpacing)
    End Function

    ' Binary search in the prefix offsets; no controls are created to measure rows.
    Private Function GetRowAtOffset(Offset As Long) As Integer
        If _RowCount = 0 Then Return -1
        If Offset <= 0 Then Return 0
        If _RowOffsets Is Nothing Then
            Return CInt(Math.Min(CLng(_RowCount - 1), Offset \ (CLng(_ItemSize.Height) + _ItemSpacing)))
        End If
        Dim Low As Integer = 0
        Dim High As Integer = _RowCount - 1
        While Low < High
            Dim Middle As Integer = Low + (High - Low + 1) \ 2
            If _RowOffsets(Middle) <= Offset Then
                Low = Middle
            Else
                High = Middle - 1
            End If
        End While
        Return Low
    End Function

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
        Dim Height As Long = CLng(_ContentPadding.Top) + GetRowOffset(Rows) - _ItemSpacing + _ContentPadding.Bottom
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

        Dim ScrollTop As Integer = GetScrollTop()
        Dim FirstRow As Integer = Math.Max(0, GetRowAtOffset(CLng(ScrollTop) - _ContentPadding.Top) - _OverscanRows)
        Dim LastRow As Integer = CInt(Math.Min(CLng(_RowCount - 1), CLng(GetRowAtOffset(CLng(ScrollTop) + ClientSize.Height - _ContentPadding.Top)) + _OverscanRows))
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
                Dim Target As Rectangle = GetTileBounds(i)
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
        Dim Top As Long = CLng(_ContentPadding.Top) + GetRowOffset(RowNumber)
        Dim RowHeight As Integer = GetRowHeight(RowNumber)
        Dim ScrollTop As Integer = GetScrollTop()
        Dim Target As Long = ScrollTop
        If Top < ScrollTop Then
            Target = Top
        ElseIf Top + RowHeight > CLng(ScrollTop) + ClientSize.Height Then
            Target = If(RowHeight > ClientSize.Height, Top, Top + RowHeight - ClientSize.Height)
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
