Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Linq
Imports System.Windows.Forms

''' <summary>
''' Provides a Windows Forms tile list and grid with centralized single selection, keyboard navigation, activation, filtering, scrolling, and designer-editable <see cref="TileViewItem"/> controls.
''' </summary>
<DefaultEvent("ItemActivated"), ToolboxItem(True), ToolboxItemFilter("CoreSuite")>
Public Class TileView
    Inherits FlowLayoutPanel
    ''' <summary>
    ''' Stores the currently selected item.
    ''' </summary>
    Private _SelectedItem As TileViewItem
    ''' <summary>
    ''' Stores whether double-clicking an item activates it.
    ''' </summary>
    Private _ActivateOnDoubleClick As Boolean = True
    ''' <summary>
    ''' Stores whether clicking an unused area clears the current selection.
    ''' </summary>
    Private _ClearSelectionOnBlankClick As Boolean = True
    ''' <summary>
    ''' Stores whether the first item added to an empty view is selected automatically.
    ''' </summary>
    Private _AutoSelectFirst As Boolean
    ''' <summary>
    ''' Stores whether keyboard navigation wraps from the final item to the first item and vice versa.
    ''' </summary>
    Private _WrapNavigation As Boolean
    ''' <summary>
    ''' Stores the current layout suspension depth used by <see cref="BeginUpdate"/> and <see cref="EndUpdate"/>.
    ''' </summary>
    Private _UpdateCount As Integer
    ''' <summary>
    ''' Initializes a new instance of the <see cref="TileView"/> class.
    ''' </summary>
    Public Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
        AutoScroll = True
        BackColor = SystemColors.Window
        BorderStyle = BorderStyle.FixedSingle
        TabStop = True
        AccessibleRole = AccessibleRole.List
    End Sub
    ''' <summary>
    ''' Occurs when the selected item changes.
    ''' </summary>
    <Category("Property Changed"), Description("Occurs when the selected item changes.")>
    Public Event SelectedItemChanged As EventHandler
    ''' <summary>
    ''' Occurs when the selection changes and provides both the previous and current items.
    ''' </summary>
    <Category("Property Changed"), Description("Occurs when the selection changes and provides both the previous and current items.")>
    Public Event SelectionChanged As EventHandler(Of TileViewSelectionChangedEventArgs)
    ''' <summary>
    ''' Occurs when an item receives a click anywhere inside its visual hierarchy.
    ''' </summary>
    <Category("Action"), Description("Occurs when an item receives a click anywhere inside its visual hierarchy.")>
    Public Event ItemClick As EventHandler(Of TileViewItemEventArgs)
    ''' <summary>
    ''' Occurs when an item receives a double-click anywhere inside its visual hierarchy.
    ''' </summary>
    <Category("Action"), Description("Occurs when an item receives a double-click anywhere inside its visual hierarchy.")>
    Public Event ItemDoubleClick As EventHandler(Of TileViewItemEventArgs)
    ''' <summary>
    ''' Occurs when an item is activated by double-clicking it, pressing Enter, or calling an activation method.
    ''' </summary>
    <Category("Action"), Description("Occurs when an item is activated by double-clicking it, pressing Enter, or calling an activation method.")>
    Public Event ItemActivated As EventHandler(Of TileViewItemEventArgs)
    ''' <summary>
    ''' Occurs after a <see cref="TileViewItem"/> is added to the view.
    ''' </summary>
    <Category("Action"), Description("Occurs after a TileViewItem is added to the view.")>
    Public Event ItemAdded As EventHandler(Of TileViewItemEventArgs)
    ''' <summary>
    ''' Occurs after a <see cref="TileViewItem"/> is removed from the view.
    ''' </summary>
    <Category("Action"), Description("Occurs after a TileViewItem is removed from the view.")>
    Public Event ItemRemoved As EventHandler(Of TileViewItemEventArgs)
    ''' <summary>
    ''' Gets or sets the currently selected item.
    ''' </summary>
    ''' <exception cref="ArgumentException">The assigned item is not contained by this view.</exception>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property SelectedItem As TileViewItem
        Get
            Return _SelectedItem
        End Get
        Set(value As TileViewItem)
            If value Is Nothing Then
                ClearSelection()
                Return
            End If
            SelectItem(value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the zero-based index of the currently selected tile among the tile items in the view.
    ''' </summary>
    ''' <exception cref="ArgumentOutOfRangeException">The assigned index is less than -1 or greater than or equal to <see cref="ItemCount"/>.</exception>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property SelectedIndex As Integer
        Get
            If _SelectedItem Is Nothing Then Return -1
            Return Items.ToList().IndexOf(_SelectedItem)
        End Get
        Set(value As Integer)
            If value = -1 Then
                ClearSelection()
                Return
            End If
            Dim currentItems = Items
            If value < 0 OrElse value >= currentItems.Count Then Throw New ArgumentOutOfRangeException(NameOf(value), "SelectedIndex must refer to an existing tile item or be -1.")
            SelectItem(currentItems(value))
        End Set
    End Property
    ''' <summary>
    ''' Gets a read-only snapshot of the tile items currently contained by the view.
    ''' </summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property Items As IReadOnlyList(Of TileViewItem)
        Get
            Return Controls.OfType(Of TileViewItem)().ToList().AsReadOnly()
        End Get
    End Property
    ''' <summary>
    ''' Gets the number of <see cref="TileViewItem"/> controls currently contained by the view.
    ''' </summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property ItemCount As Integer
        Get
            Return Controls.OfType(Of TileViewItem)().Count()
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether double-clicking an item activates it.
    ''' </summary>
    <Category("Behavior"), Description("Specifies whether double-clicking an item activates it."), DefaultValue(True)>
    Public Property ActivateOnDoubleClick As Boolean
        Get
            Return _ActivateOnDoubleClick
        End Get
        Set(value As Boolean)
            _ActivateOnDoubleClick = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether clicking an unused area clears the current selection.
    ''' </summary>
    <Category("Behavior"), Description("Specifies whether clicking an unused area clears the current selection."), DefaultValue(True)>
    Public Property ClearSelectionOnBlankClick As Boolean
        Get
            Return _ClearSelectionOnBlankClick
        End Get
        Set(value As Boolean)
            _ClearSelectionOnBlankClick = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the first item added to an empty view is selected automatically.
    ''' </summary>
    <Category("Behavior"), Description("Specifies whether the first item added to an empty view is selected automatically."), DefaultValue(False)>
    Public Property AutoSelectFirst As Boolean
        Get
            Return _AutoSelectFirst
        End Get
        Set(value As Boolean)
            _AutoSelectFirst = value
            If value AndAlso _SelectedItem Is Nothing Then SelectFirst()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether keyboard navigation wraps at the beginning and end of the item collection.
    ''' </summary>
    <Category("Behavior"), Description("Specifies whether keyboard navigation wraps at the beginning and end of the item collection."), DefaultValue(False)>
    Public Property WrapNavigation As Boolean
        Get
            Return _WrapNavigation
        End Get
        Set(value As Boolean)
            _WrapNavigation = value
        End Set
    End Property
    ''' <summary>
    ''' Adds a tile item to the view.
    ''' </summary>
    ''' <param name="item">The item to add.</param>
    Public Sub Add(item As TileViewItem)
        ArgumentNullException.ThrowIfNull(item)
        Controls.Add(item)
    End Sub
    ''' <summary>
    ''' Adds multiple tile items to the view while minimizing intermediate layout work.
    ''' </summary>
    ''' <param name="items">The items to add.</param>
    Public Sub AddRange(items As IEnumerable(Of TileViewItem))
        ArgumentNullException.ThrowIfNull(items)
        BeginUpdate()
        Try
            For Each item In items
                Add(item)
            Next
        Finally
            EndUpdate()
        End Try
    End Sub
    ''' <summary>
    ''' Replaces all controls in the view with tile items created from a data source.
    ''' </summary>
    ''' <typeparam name="T">The type of source item.</typeparam>
    ''' <param name="source">The data source used to create the tile items.</param>
    ''' <param name="factory">A function that creates a <see cref="TileViewItem"/> for each source item.</param>
    Public Sub SetItems(Of T)(source As IEnumerable(Of T), factory As Func(Of T, TileViewItem))
        ArgumentNullException.ThrowIfNull(source)
        ArgumentNullException.ThrowIfNull(factory)
        BeginUpdate()
        Try
            ClearItems()
            For Each value In source
                Dim item = factory(value)
                If item Is Nothing Then Throw New InvalidOperationException("The TileView item factory returned Nothing.")
                Add(item)
            Next
        Finally
            EndUpdate()
        End Try
    End Sub
    ''' <summary>
    ''' Removes a tile item from the view.
    ''' </summary>
    ''' <param name="item">The item to remove.</param>
    ''' <returns><see langword="True"/> when the item was contained and removed; otherwise, <see langword="False"/>.</returns>
    Public Function Remove(item As TileViewItem) As Boolean
        If item Is Nothing OrElse item.Parent IsNot Me Then Return False
        Controls.Remove(item)
        Return True
    End Function
    ''' <summary>
    ''' Removes all controls from the view and clears the current selection.
    ''' </summary>
    Public Sub ClearItems()
        ClearSelection()
        Controls.Clear()
    End Sub
    ''' <summary>
    ''' Selects the specified tile item.
    ''' </summary>
    ''' <param name="Item">The item to select.</param>
    ''' <exception cref="ArgumentNullException">The item is <see langword="Nothing"/>.</exception>
    ''' <exception cref="ArgumentException">The item is not contained by this view.</exception>
    Public Sub SelectItem(Item As TileViewItem)
        ArgumentNullException.ThrowIfNull(Item)
        If Item.Parent IsNot Me Then Throw New ArgumentException("The specified item is not contained by this TileView.", NameOf(Item))
        If ReferenceEquals(_SelectedItem, Item) Then
            If Item.Visible Then EnsureVisible(Item)
            Return
        End If
        Dim PreviousItem = _SelectedItem
        If PreviousItem IsNot Nothing Then PreviousItem.SetSelectedInternal(False)
        _SelectedItem = Item
        _SelectedItem.SetSelectedInternal(True)
        If Item.Visible Then EnsureVisible(Item)
        RaiseEvent SelectedItemChanged(Me, EventArgs.Empty)
        RaiseEvent SelectionChanged(Me, New TileViewSelectionChangedEventArgs(PreviousItem, _SelectedItem))
    End Sub
    ''' <summary>
    ''' Clears the current selection.
    ''' </summary>
    Public Sub ClearSelection()
        If _SelectedItem Is Nothing Then Return
        Dim PreviousItem = _SelectedItem
        _SelectedItem = Nothing
        PreviousItem.SetSelectedInternal(False)
        RaiseEvent SelectedItemChanged(Me, EventArgs.Empty)
        RaiseEvent SelectionChanged(Me, New TileViewSelectionChangedEventArgs(PreviousItem, Nothing))
    End Sub
    ''' <summary>
    ''' Selects and focuses the first visible and enabled tile item.
    ''' </summary>
    ''' <returns><see langword="True"/> when an item was selected; otherwise, <see langword="False"/>.</returns>
    Public Function SelectFirst() As Boolean
        Dim Item = GetNavigableItems().FirstOrDefault()
        Return SelectAndFocus(Item)
    End Function
    ''' <summary>
    ''' Selects and focuses the last visible and enabled tile item.
    ''' </summary>
    ''' <returns><see langword="True"/> when an item was selected; otherwise, <see langword="False"/>.</returns>
    Public Function SelectLast() As Boolean
        Dim Item = GetNavigableItems().LastOrDefault()
        Return SelectAndFocus(Item)
    End Function
    ''' <summary>
    ''' Selects and focuses the next visible and enabled tile item according to control order.
    ''' </summary>
    ''' <returns><see langword="True"/> when the selection moved; otherwise, <see langword="False"/>.</returns>
    Public Function SelectNext() As Boolean
        Return MoveSequential(1)
    End Function
    ''' <summary>
    ''' Selects and focuses the previous visible and enabled tile item according to control order.
    ''' </summary>
    ''' <returns><see langword="True"/> when the selection moved; otherwise, <see langword="False"/>.</returns>
    Public Function SelectPrevious() As Boolean
        Return MoveSequential(-1)
    End Function
    ''' <summary>
    ''' Scrolls the specified item into the visible client area.
    ''' </summary>
    ''' <param name="Item">The item to reveal.</param>
    Public Sub EnsureVisible(Item As TileViewItem)
        ArgumentNullException.ThrowIfNull(Item)
        If Item.Parent IsNot Me Then Throw New ArgumentException("The specified item is not contained by this TileView.", NameOf(Item))
        If Item.Visible Then ScrollControlIntoView(Item)
    End Sub
    ''' <summary>
    ''' Activates the currently selected item.
    ''' </summary>
    ''' <returns><see langword="True"/> when an item was activated; otherwise, <see langword="False"/>.</returns>
    Public Function ActivateSelected() As Boolean
        If _SelectedItem Is Nothing Then Return False
        Return ActivateItem(_SelectedItem)
    End Function
    ''' <summary>
    ''' Selects and activates the specified item.
    ''' </summary>
    ''' <param name="Item">The item to activate.</param>
    ''' <returns><see langword="True"/> when the item was activated; otherwise, <see langword="False"/>.</returns>
    Public Function ActivateItem(Item As TileViewItem) As Boolean
        If Item Is Nothing OrElse Item.Parent IsNot Me OrElse Not Item.Enabled OrElse Not Item.Visible Then Return False
        SelectItem(Item)
        Item.RaiseActivated()
        RaiseEvent ItemActivated(Me, New TileViewItemEventArgs(Item))
        Return True
    End Function
    ''' <summary>
    ''' Applies a visibility filter to the current tile items.
    ''' </summary>
    ''' <param name="Predicate">A function that returns <see langword="True"/> for items that should remain visible.</param>
    Public Sub Filter(Predicate As Predicate(Of TileViewItem))
        ArgumentNullException.ThrowIfNull(Predicate)
        BeginUpdate()
        Try
            For Each Item In Items
                Item.Visible = Predicate(Item)
            Next
            If _SelectedItem IsNot Nothing AndAlso Not _SelectedItem.Visible Then ClearSelection()
        Finally
            EndUpdate()
        End Try
    End Sub
    ''' <summary>
    ''' Makes every tile item visible, removing any filter previously applied through <see cref="Filter"/>.
    ''' </summary>
    Public Sub ClearFilter()
        BeginUpdate()
        Try
            For Each Item In Items
                Item.Visible = True
            Next
        Finally
            EndUpdate()
        End Try
    End Sub
    ''' <summary>
    ''' Suspends layout updates until a matching call to <see cref="EndUpdate"/> is made.
    ''' </summary>
    Public Sub BeginUpdate()
        _UpdateCount += 1
        If _UpdateCount = 1 Then SuspendLayout()
    End Sub
    ''' <summary>
    ''' Resumes layout updates suspended by <see cref="BeginUpdate"/>.
    ''' </summary>
    ''' <exception cref="InvalidOperationException">No matching <see cref="BeginUpdate"/> call is active.</exception>
    Public Sub EndUpdate()
        If _UpdateCount <= 0 Then Throw New InvalidOperationException("EndUpdate cannot be called without a matching BeginUpdate call.")
        _UpdateCount -= 1
        If _UpdateCount = 0 Then
            ResumeLayout(True)
            If _AutoSelectFirst AndAlso _SelectedItem Is Nothing Then
                Dim item = GetNavigableItems().FirstOrDefault()
                If item IsNot Nothing Then SelectItem(item)
            End If
            Invalidate(True)
            Update()
        End If
    End Sub
    ''' <summary>
    ''' Handles keyboard commands forwarded by a hosted tile item.
    ''' </summary>
    ''' <param name="Source">The item that currently owns keyboard focus.</param>
    ''' <param name="KeyData">The key data to process.</param>
    ''' <returns><see langword="True"/> when the command was handled by the view; otherwise, <see langword="False"/>.</returns>
    Friend Function ProcessItemCommand(Source As TileViewItem, KeyData As Keys) As Boolean
        If Source Is Nothing OrElse Source.Parent IsNot Me Then Return False
        If (KeyData And Keys.Alt) = Keys.Alt OrElse (KeyData And Keys.Control) = Keys.Control Then Return False
        Select Case KeyData And Keys.KeyCode
            Case Keys.Up, Keys.Down, Keys.Left, Keys.Right
                Return MoveDirectional(Source, KeyData And Keys.KeyCode)
            Case Keys.Home
                Return SelectFirst()
            Case Keys.End
                Return SelectLast()
            Case Keys.Enter
                SelectItem(Source)
                Return ActivateItem(Source)
            Case Keys.Space
                SelectItem(Source)
                Return True
        End Select
        Return False
    End Function
    ''' <summary>
    ''' Returns the visible and enabled items in control order.
    ''' </summary>
    Private Function GetNavigableItems() As List(Of TileViewItem)
        Return Controls.OfType(Of TileViewItem)().Where(Function(item) item.Visible AndAlso item.Enabled).ToList()
    End Function
    ''' <summary>
    ''' Selects and focuses an item when the supplied reference is valid.
    ''' </summary>
    Private Function SelectAndFocus(Item As TileViewItem) As Boolean
        If Item Is Nothing Then Return False
        SelectItem(Item)
        If Item.CanFocus Then Item.Focus()
        Return True
    End Function
    ''' <summary>
    ''' Moves selection sequentially through the current navigable item list.
    ''' </summary>
    Private Function MoveSequential(offset As Integer) As Boolean
        Dim Items = GetNavigableItems()
        If Items.Count = 0 Then Return False
        If _SelectedItem Is Nothing OrElse Not Items.Contains(_SelectedItem) Then Return SelectAndFocus(If(offset >= 0, Items(0), Items(Items.Count - 1)))
        Dim Index = Items.IndexOf(_SelectedItem) + offset
        If Index < 0 OrElse Index >= Items.Count Then
            If Not _WrapNavigation Then Return False
            Index = If(Index < 0, Items.Count - 1, 0)
        End If
        Return SelectAndFocus(Items(Index))
    End Function
    ''' <summary>
    ''' Moves selection to the closest tile located in the requested visual direction.
    ''' </summary>
    Private Function MoveDirectional(Source As TileViewItem, Direction As Keys) As Boolean
        Dim Items = GetNavigableItems()
        If Items.Count = 0 Then Return False
        Dim SourceCenter = GetItemCenter(Source)
        Dim BestItem As TileViewItem = Nothing
        Dim BestScore As Long = Long.MaxValue
        For Each Candidate In Items
            If ReferenceEquals(Candidate, Source) Then Continue For
            Dim CandidateCenter = GetItemCenter(Candidate)
            Dim DeltaX = CandidateCenter.X - SourceCenter.X
            Dim DeltaY = CandidateCenter.Y - SourceCenter.Y
            Dim PrimaryDistance As Integer
            Dim SecondaryDistance As Integer
            Select Case Direction
                Case Keys.Up
                    If DeltaY >= 0 Then Continue For
                    PrimaryDistance = -DeltaY
                    SecondaryDistance = Math.Abs(DeltaX)
                Case Keys.Down
                    If DeltaY <= 0 Then Continue For
                    PrimaryDistance = DeltaY
                    SecondaryDistance = Math.Abs(DeltaX)
                Case Keys.Left
                    If DeltaX >= 0 Then Continue For
                    PrimaryDistance = -DeltaX
                    SecondaryDistance = Math.Abs(DeltaY)
                Case Keys.Right
                    If DeltaX <= 0 Then Continue For
                    PrimaryDistance = DeltaX
                    SecondaryDistance = Math.Abs(DeltaY)
                Case Else
                    Return False
            End Select
            Dim Score = CLng(PrimaryDistance) * 10000L + SecondaryDistance
            If Score >= BestScore Then Continue For
            BestScore = Score
            BestItem = Candidate
        Next
        If BestItem IsNot Nothing Then Return SelectAndFocus(BestItem)
        If Not _WrapNavigation Then Return False
        Dim VisualOrder = Items.OrderBy(Function(item) item.Top).ThenBy(Function(item) item.Left).ToList()
        If Direction = Keys.Up OrElse Direction = Keys.Left Then Return SelectAndFocus(VisualOrder.Last())
        Return SelectAndFocus(VisualOrder.First())
    End Function
    ''' <summary>
    ''' Gets the center point of an item in TileView client coordinates.
    ''' </summary>
    Private Shared Function GetItemCenter(Item As TileViewItem) As Point
        Return New Point(Item.Left + Item.Width \ 2, Item.Top + Item.Height \ 2)
    End Function
    ''' <summary>
    ''' Raises the item-click event after a hosted tile receives a click.
    ''' </summary>
    Private Sub HostedItem_Click(sender As Object, e As EventArgs)
        Dim Item = TryCast(sender, TileViewItem)
        If Item Is Nothing Then Return
        RaiseEvent ItemClick(Me, New TileViewItemEventArgs(Item))
    End Sub
    ''' <summary>
    ''' Raises the item-double-click event and optionally activates the tile.
    ''' </summary>
    Private Sub HostedItem_DoubleClick(sender As Object, e As EventArgs)
        Dim Item = TryCast(sender, TileViewItem)
        If Item Is Nothing Then Return
        RaiseEvent ItemDoubleClick(Me, New TileViewItemEventArgs(Item))
        If _ActivateOnDoubleClick Then ActivateItem(Item)
    End Sub
    ''' <summary>
    ''' Attaches selection, click, and navigation behavior when a tile item is added.
    ''' </summary>
    Protected Overrides Sub OnControlAdded(e As ControlEventArgs)
        MyBase.OnControlAdded(e)
        Dim item = TryCast(e.Control, TileViewItem)
        If item Is Nothing Then Return
        If item.OwnerView IsNot Nothing AndAlso item.OwnerView IsNot Me Then Throw New InvalidOperationException("A TileViewItem cannot belong to more than one TileView.")
        item.SetOwner(Me)
        AddHandler item.Click, AddressOf HostedItem_Click
        AddHandler item.DoubleClick, AddressOf HostedItem_DoubleClick
        RaiseEvent ItemAdded(Me, New TileViewItemEventArgs(item))
        If _UpdateCount = 0 AndAlso _AutoSelectFirst AndAlso _SelectedItem Is Nothing AndAlso item.Visible AndAlso item.Enabled Then SelectItem(item)
    End Sub
    ''' <summary>
    ''' Detaches tile behavior and clears selection when the selected item is removed.
    ''' </summary>
    Protected Overrides Sub OnControlRemoved(e As ControlEventArgs)
        Dim Item = TryCast(e.Control, TileViewItem)
        If Item IsNot Nothing Then
            RemoveHandler Item.Click, AddressOf HostedItem_Click
            RemoveHandler Item.DoubleClick, AddressOf HostedItem_DoubleClick
            If ReferenceEquals(_SelectedItem, Item) Then ClearSelection()
            Item.SetOwner(Nothing)
        End If
        MyBase.OnControlRemoved(e)
        If Item IsNot Nothing Then RaiseEvent ItemRemoved(Me, New TileViewItemEventArgs(Item))
    End Sub
    ''' <summary>
    ''' Clears the selection when the user clicks unused view space and the corresponding behavior is enabled.
    ''' </summary>
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        If _ClearSelectionOnBlankClick Then ClearSelection()
        If CanFocus Then Focus()
        MyBase.OnMouseDown(e)
    End Sub
    ''' <summary>
    ''' Handles list-navigation and activation keys when the TileView itself owns keyboard focus.
    ''' </summary>
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        If (keyData And Keys.Alt) <> Keys.Alt AndAlso (keyData And Keys.Control) <> Keys.Control Then
            Select Case keyData And Keys.KeyCode
                Case Keys.Up, Keys.Down, Keys.Left, Keys.Right
                    If _SelectedItem Is Nothing Then Return SelectFirst()
                    If MoveDirectional(_SelectedItem, keyData And Keys.KeyCode) Then Return True
                Case Keys.Home
                    Return SelectFirst()
                Case Keys.End
                    Return SelectLast()
                Case Keys.Enter
                    If ActivateSelected() Then Return True
                Case Keys.Space
                    If _SelectedItem Is Nothing Then Return SelectFirst()
                    Return True
            End Select
        End If
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function
End Class
