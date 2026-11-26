Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Virtualized tile list for Windows Forms. Data records are kept separate from recycled TileViewItem controls.
''' Create a designer-editable UserControl inheriting TileViewItem, then call SetItems.
''' </summary>
<DefaultEvent("ItemActivated"), ToolboxItem(True)>
Public Partial Class TileView
    Inherits Panel

    Private ReadOnly _Data As New List(Of Object)()
    Private ReadOnly _VisibleSourceIndices As New List(Of Integer)()
    Private _Factory As Func(Of TileViewItem)
    Private _Binder As Action(Of TileViewItem, Object)
    Private _PreviewPainter As Action(Of Graphics, Rectangle, Object, Boolean)
    Private _Filter As Func(Of Object, Boolean)
    Private _MeasureHeight As Func(Of Object, Integer, Integer, Integer)
    Private _SelectedSourceIndex As Integer = -1

    Private _ItemSize As New Size(260, 112)
    Private _RequestedColumns As Integer
    Private _StretchItems As Boolean
    Private _ItemSpacing As Integer = 10
    Private _ContentPadding As New Padding(10)
    Private _OverscanRows As Integer = 1
    Private _ActivateOnDoubleClick As Boolean = True
    Private _ClearSelectionOnBlankClick As Boolean = True
    Private _AutoSelectFirst As Boolean
    Private _WrapNavigation As Boolean
    Private _ReadOnly As Boolean
    Private _SmoothScrolling As Boolean = True
    Private _UpdateDepth As Integer

    Public Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.Selectable, True)
        AutoScroll = False
        InitializeVirtualScrollBar()
        InitializePreviewTimers()
        BackColor = SystemColors.Window
        BorderStyle = BorderStyle.FixedSingle
        AccessibleRole = AccessibleRole.List
        TabStop = True
    End Sub

    ''' <summary>Current vertical scroll position in pixels (logical, not a child-window offset).</summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property ScrollOffset As Integer
        Get
            Return _ScrollOffset
        End Get
    End Property

    ''' <summary>Largest valid vertical scroll offset for the current data and viewport.</summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property MaximumScrollOffset As Integer
        Get
            If _layoutDirty Then UpdateViewport()
            Return Math.Max(0, _ContentHeight - ClientSize.Height)
        End Get
    End Property

    ''' <summary>Whether the TileView's own vertical scrollbar is currently shown.</summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property VerticalScrollBarVisible As Boolean
        Get
            Return _VerticalBar.Visible
        End Get
    End Property

    ''' <summary>The actual width of the internal virtual scrollbar, in pixels.</summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property VerticalScrollBarWidth As Integer
        Get
            Return _VerticalBar.Width
        End Get
    End Property

    <Category("Property Changed"), Description("Raised when the logical selection changes, including off-screen records.")>
    Public Event SelectionChanged As EventHandler(Of TileViewSelectionChangedEventArgs)

    <Category("Property Changed")>
    Public Event SelectedItemChanged As EventHandler

    <Category("Action")>
    Public Event ItemClick As EventHandler(Of TileViewItemEventArgs)

    <Category("Action")>
    Public Event ItemDoubleClick As EventHandler(Of TileViewItemEventArgs)

    <Category("Action")>
    Public Event ItemActivated As EventHandler(Of TileViewItemEventArgs)

    ''' <summary>Raised once when a new visual tile is instantiated (not once per record).</summary>
    <Category("Action")>
    Public Event TileCreated As EventHandler(Of TileViewItemEventArgs)

    ''' <summary>Raised after the realized viewport changes (useful for diagnostics).</summary>
    <Category("Property Changed")>
    Public Event ViewportChanged As EventHandler

    ''' <summary>The original number of records, including filtered-out records.</summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property ItemCount As Integer
        Get
            Return _Data.Count
        End Get
    End Property

    ''' <summary>The number of records left after applying the filter.</summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property VisibleItemCount As Integer
        Get
            Return _VisibleSourceIndices.Count
        End Get
    End Property

    ''' <summary>Number of currently bound, on-screen or overscan tile controls.</summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property RealizedItemCount As Integer
        Get
            Return _Realized.Count
        End Get
    End Property

    ''' <summary>Total tiles allocated, including reusable cached tiles; should remain independent of data count.</summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property CreatedTileCount As Integer
        Get
            Return _TilePool.Count + If(_PreviewTemplate Is Nothing, 0, 1)
        End Get
    End Property

    ''' <summary>A snapshot of realized tile controls only, not the entire data set.</summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property Items As IReadOnlyList(Of TileViewItem)
        Get
            Dim Result As New List(Of TileViewItem)()
            For Each Pair In _Realized.OrderBy(Function(p) p.Key)
                Result.Add(Pair.Value)
            Next
            Return Result.AsReadOnly()
        End Get
    End Property

    ''' <summary>Position among filtered/visible records, or -1 if none is selected.</summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property SelectedIndex As Integer
        Get
            Return GetDisplayIndex(_SelectedSourceIndex)
        End Get
        Set(value As Integer)
            If value = -1 Then
                ClearSelection()
            Else
                SelectIndex(value)
            End If
        End Set
    End Property

    ''' <summary>Original data index (independent of filter and virtualization), or -1.</summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property SelectedSourceIndex As Integer
        Get
            Return _SelectedSourceIndex
        End Get
    End Property

    ''' <summary>The selected model object, even if its tile is not realized.</summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property SelectedDataItem As Object
        Get
            Return GetSourceItemOrNothing(_SelectedSourceIndex)
        End Get
    End Property

    ''' <summary>The selected tile ONLY IF realized; use SelectedDataItem for off-screen selections.</summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property SelectedItem As TileViewItem
        Get
            Dim Tile As TileViewItem = Nothing
            _Realized.TryGetValue(SelectedIndex, Tile)
            Return Tile
        End Get
    End Property

    <Category("Layout"), Description("Dimensions of each logical tile, before narrow-viewport width clamping."), DefaultValue(GetType(Size), "260, 112")>
    Public Property ItemSize As Size
        Get
            Return _ItemSize
        End Get
        Set(value As Size)
            If value.Width <= 0 OrElse value.Height <= 0 Then Throw New ArgumentOutOfRangeException(NameOf(value))
            If _ItemSize = value Then Return
            _ItemSize = value
            InvalidateVirtualLayout()
        End Set
    End Property

    ''' <summary>
    ''' Sets the number of tiles per row. Zero automatically calculates columns from ItemSize.Width.
    ''' Fixed columns are retained when resizing, and tile widths are clamped to the viewport.
    ''' </summary>
    <Category("Layout"), Description("Number of tiles per row. 0 = automatic; 1 = one tile per row."), DefaultValue(0)>
    Public Property Columns As Integer
        Get
            Return _RequestedColumns
        End Get
        Set(value As Integer)
            If value < 0 Then Throw New ArgumentOutOfRangeException(NameOf(value), "Columns must be zero or greater.")
            If _RequestedColumns = value Then Return
            _RequestedColumns = value
            InvalidateVirtualLayout()
        End Set
    End Property

    ''' <summary>
    ''' Stretches tiles across the available horizontal space, respecting Columns
    ''' (or the automatic column calculation when Columns is zero).
    ''' </summary>
    <Category("Layout"), Description("Stretches tiles to use the available width of each row."), DefaultValue(False)>
    Public Property StretchItems As Boolean
        Get
            Return _StretchItems
        End Get
        Set(value As Boolean)
            If _StretchItems = value Then Return
            _StretchItems = value
            InvalidateVirtualLayout()
        End Set
    End Property

    <Category("Layout"), Description("Distance between adjacent tiles in pixels."), DefaultValue(10)>
    Public Property ItemSpacing As Integer
        Get
            Return _ItemSpacing
        End Get
        Set(value As Integer)
            If value < 0 Then Throw New ArgumentOutOfRangeException(NameOf(value))
            If _ItemSpacing = value Then Return
            _ItemSpacing = value
            InvalidateVirtualLayout()
        End Set
    End Property

    <Category("Layout"), Description("Padding around the complete tile collection."), DefaultValue(GetType(Padding), "10, 10, 10, 10")>
    Public Property ContentPadding As Padding
        Get
            Return _ContentPadding
        End Get
        Set(value As Padding)
            If value.Left < 0 OrElse value.Top < 0 OrElse value.Right < 0 OrElse value.Bottom < 0 Then
                Throw New ArgumentOutOfRangeException(NameOf(value))
            End If
            _ContentPadding = value
            InvalidateVirtualLayout()
        End Set
    End Property

    <Category("Behavior"), Description("Extra logical rows rendered outside the viewport."), DefaultValue(1)>
    Public Property OverscanRows As Integer
        Get
            Return _OverscanRows
        End Get
        Set(value As Integer)
            If value < 0 OrElse value > 10 Then Throw New ArgumentOutOfRangeException(NameOf(value), "Use a value from 0 to 10.")
            _OverscanRows = value
            UpdateViewport()
        End Set
    End Property

    <Category("Behavior"), DefaultValue(True)>
    Public Property ActivateOnDoubleClick As Boolean
        Get
            Return _ActivateOnDoubleClick
        End Get
        Set(value As Boolean)
            _ActivateOnDoubleClick = value
        End Set
    End Property

    <Category("Behavior"), DefaultValue(True)>
    Public Property ClearSelectionOnBlankClick As Boolean
        Get
            Return _ClearSelectionOnBlankClick
        End Get
        Set(value As Boolean)
            _ClearSelectionOnBlankClick = value
        End Set
    End Property

    <Category("Behavior"), DefaultValue(False)>
    Public Property AutoSelectFirst As Boolean
        Get
            Return _AutoSelectFirst
        End Get
        Set(value As Boolean)
            _AutoSelectFirst = value
            If value AndAlso _SelectedSourceIndex = -1 AndAlso _VisibleSourceIndices.Count > 0 Then SelectIndex(0)
        End Set
    End Property

    <Category("Behavior"), DefaultValue(False)>
    Public Property WrapNavigation As Boolean
        Get
            Return _WrapNavigation
        End Get
        Set(value As Boolean)
            _WrapNavigation = value
        End Set
    End Property

    ''' <summary>During user scrolling, draw tiles on one buffered surface instead of moving live child controls.</summary>
    <Category("Behavior"), Description("Draws a single buffered preview while scrolling and restores interactive tiles afterwards."), DefaultValue(True)>
    Public Property SmoothScrolling As Boolean
        Get
            Return _SmoothScrolling
        End Get
        Set(value As Boolean)
            If _SmoothScrolling = value Then Return
            If Not value Then FinishScrollPreview()
            _SmoothScrolling = value
        End Set
    End Property

    ''' <summary>True only during a user's active scrollbar or mouse-wheel movement.</summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property IsPreviewRendering As Boolean
        Get
            Return _PreviewActive
        End Get
    End Property

    ''' <summary>Propagated to pooled visual tiles. Template decides which editing controls to disable.</summary>
    <Category("Behavior"), DefaultValue(False)>
    Public Property [ReadOnly] As Boolean
        Get
            Return _ReadOnly
        End Get
        Set(value As Boolean)
            If _ReadOnly = value Then Return
            _ReadOnly = value
            For Each tile In _TilePool
                tile.SetReadOnlyInternal(value)
            Next
            If _PreviewTemplate IsNot Nothing Then _PreviewTemplate.SetReadOnlyInternal(value)
            ClearPreviewCache()
        End Set
    End Property

    ''' <summary>
    ''' Captures the source in memory once. Only tiles near the viewport are created.
    ''' The binding delegate must overwrite ALL visual data that can vary between records.
    ''' </summary>
    Public Sub SetItems(Of T)(Source As IEnumerable(Of T), TileFactory As Func(Of TileViewItem), BindItem As Action(Of TileViewItem, T), Optional PaintPreview As Action(Of Graphics, Rectangle, T, Boolean) = Nothing)
        SetItems(Of T)(Source, TileFactory, BindItem, PaintPreview, Nothing)
    End Sub

    ''' <summary>Optionally measures each record from its data, available width and minimum item height.
    ''' Items in the same row share the largest measured height. The callback must be side-effect free.</summary>
    Public Sub SetItems(Of T)(Source As IEnumerable(Of T), TileFactory As Func(Of TileViewItem), BindItem As Action(Of TileViewItem, T), PaintPreview As Action(Of Graphics, Rectangle, T, Boolean), MeasureHeight As Func(Of T, Integer, Integer, Integer))
        If Source Is Nothing Then Throw New ArgumentNullException(NameOf(Source))
        If TileFactory Is Nothing Then Throw New ArgumentNullException(NameOf(TileFactory))
        If BindItem Is Nothing Then Throw New ArgumentNullException(NameOf(BindItem))

        Dim Snapshot As New List(Of Object)()
        For Each value As T In Source
            Snapshot.Add(value)
        Next
        Dim Binding As Action(Of TileViewItem, Object) = Sub(tile As TileViewItem, dataItem As Object)
                                                             BindItem(tile, CType(dataItem, T))
                                                         End Sub
        Dim Painter As Action(Of Graphics, Rectangle, Object, Boolean) = Nothing
        If PaintPreview IsNot Nothing Then
            Painter = Sub(g As Graphics, Bounds As Rectangle, Model As Object, IsSelected As Boolean)
                          PaintPreview(g, Bounds, CType(Model, T), IsSelected)
                      End Sub
        End If
        Dim HeightSelector As Func(Of Object, Integer, Integer, Integer) = Nothing
        If MeasureHeight IsNot Nothing Then
            HeightSelector = Function(Model As Object, Width As Integer, MinimumHeight As Integer) MeasureHeight(CType(Model, T), Width, MinimumHeight)
        End If
        ReplaceData(Snapshot, TileFactory, Binding, Painter, HeightSelector)
    End Sub

    ''' <summary>Applies a model filter, without creating a visual control for each record.</summary>
    Public Sub ApplyFilter(Of T)(Predicate As Predicate(Of T))
        If Predicate Is Nothing Then Throw New ArgumentNullException(NameOf(Predicate))
        Dim PreviousFilter = _Filter
        _Filter = Function(value As Object) Predicate(CType(value, T))
        Try
            RebuildFilteredIndices()
        Catch
            _Filter = PreviousFilter
            Throw
        End Try
    End Sub

    Public Sub ClearFilter()
        If _Filter Is Nothing Then Return
        _Filter = Nothing
        RebuildFilteredIndices()
    End Sub

    ''' <summary>Rebinds realized tiles after mutations to the model objects.</summary>
    Public Sub RefreshVisibleItems()
        If _Binder Is Nothing Then Return
        FinishScrollPreview()
        ClearPreviewCache()
        For Each Pair In _Realized.ToArray()
            Dim Tile = Pair.Value
            If Pair.Key >= 0 AndAlso Pair.Key < _VisibleSourceIndices.Count Then
                _Binder(Tile, _Data(_VisibleSourceIndices(Pair.Key)))
                Tile.Invalidate()
            End If
        Next
        If _MeasureHeight IsNot Nothing Then InvalidateVirtualLayout()
    End Sub

    ''' <summary>Discard all data and dispose every visual control owned by this view.</summary>
    Public Sub ClearItems()
        Dim OldIndex = _SelectedSourceIndex
        Dim OldData = GetSourceItemOrNothing(OldIndex)
        Dim OldTile = SelectedItem
        _SelectedSourceIndex = -1
        CancelScrollPreview()
        DisposeTilePool()
        ResetVirtualScroll()
        _Data.Clear()
        _VisibleSourceIndices.Clear()
        _Filter = Nothing
        _Factory = Nothing
        _Binder = Nothing
        _PreviewPainter = Nothing
        _MeasureHeight = Nothing
        _RowOffsets = Nothing
        _layoutDirty = True
        If OldIndex <> -1 Then RaiseSelectionChanged(OldIndex, -1, OldData, Nothing, OldTile, Nothing)
        UpdateViewport()
    End Sub

    ''' <summary>Suspends viewport updates until the matching EndUpdate call.</summary>
    Public Sub BeginUpdate()
        _UpdateDepth += 1
        If _UpdateDepth = 1 Then SuspendLayout()
    End Sub

    Public Sub EndUpdate()
        If _UpdateDepth = 0 Then Throw New InvalidOperationException("EndUpdate requires a preceding BeginUpdate.")
        _UpdateDepth -= 1
        If _UpdateDepth = 0 Then
            ResumeLayout(False)
            UpdateViewport()
        End If
    End Sub

    Private Sub ReplaceData(Snapshot As List(Of Object), Factory As Func(Of TileViewItem), Binder As Action(Of TileViewItem, Object), Painter As Action(Of Graphics, Rectangle, Object, Boolean), HeightSelector As Func(Of Object, Integer, Integer, Integer))
        BeginUpdate()
        Try
            ClearItems()
            _Factory = Factory
            _Binder = Binder
            _PreviewPainter = Painter
            _MeasureHeight = HeightSelector
            _Data.AddRange(Snapshot)
            RebuildFilteredIndices()
        Finally
            EndUpdate()
        End Try
    End Sub

    Private Sub RebuildFilteredIndices()
        CancelScrollPreview()
        Dim NextIndices As New List(Of Integer)()
        For i As Integer = 0 To _Data.Count - 1
            If _Filter Is Nothing OrElse _Filter(_Data(i)) Then NextIndices.Add(i)
        Next
        Dim PreviousSource As Integer = _SelectedSourceIndex
        Dim PreviousData As Object = GetSourceItemOrNothing(PreviousSource)
        Dim PreviousTile As TileViewItem = SelectedItem
        ReleaseAllTiles()
        ResetVirtualScroll()
        _VisibleSourceIndices.Clear()
        _VisibleSourceIndices.AddRange(NextIndices)
        If PreviousSource >= 0 AndAlso GetDisplayIndex(PreviousSource) = -1 Then
            _SelectedSourceIndex = -1
            RaiseSelectionChanged(PreviousSource, -1, PreviousData, Nothing, PreviousTile, Nothing)
        End If
        InvalidateVirtualLayout()
        If _AutoSelectFirst AndAlso _SelectedSourceIndex = -1 AndAlso _VisibleSourceIndices.Count > 0 Then SelectIndex(0)
    End Sub

    Private Function GetDisplayIndex(SourceIndex As Integer) As Integer
        If SourceIndex < 0 Then Return -1
        Dim Result As Integer = _VisibleSourceIndices.BinarySearch(SourceIndex)
        Return If(Result < 0, -1, Result)
    End Function

    Private Function GetSourceItemOrNothing(Index As Integer) As Object
        If Index < 0 OrElse Index >= _Data.Count Then Return Nothing
        Return _Data(Index)
    End Function

    Private Sub RaiseSelectionChanged(PreviousSource As Integer, NextSource As Integer, PreviousData As Object, NextData As Object, PreviousTile As TileViewItem, NextTile As TileViewItem)
        RaiseEvent SelectedItemChanged(Me, EventArgs.Empty)
        RaiseEvent SelectionChanged(Me, New TileViewSelectionChangedEventArgs(PreviousSource, NextSource, PreviousData, NextData, PreviousTile, NextTile))
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            DisposeScrollPreview()
            DisposeTilePool()
            _Data.Clear()
            _VisibleSourceIndices.Clear()
            _Factory = Nothing
            _Binder = Nothing
            _PreviewPainter = Nothing
            _MeasureHeight = Nothing
            _RowOffsets = Nothing
            _Filter = Nothing
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class
