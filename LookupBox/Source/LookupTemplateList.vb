Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Virtualized, fixed-row-height host for designer-created UserControl templates.
''' The scrollbar operates in pixels and only the visible row controls are created.
''' </summary>
<DesignerCategory("Code")>
Friend NotInheritable Class LookupTemplateList
    Inherits UserControl

    Private ReadOnly _Owner As LookupBox
    Private ReadOnly _Viewport As Panel
    Private ReadOnly _VerticalScroll As VScrollBar
    Private ReadOnly _Rows As New List(Of LookupTemplateRow)()
    Private _Items As IReadOnlyList(Of Object) = Array.Empty(Of Object)()
    Private _TemplateType As Type
    Private _SelectedIndex As Integer = -1
    Private _RowHeight As Integer = 72
    Private _RowSpacing As Integer = 2
    Private _UpdatingLayout As Boolean
    Private _NeedsScroll As Boolean

    Friend Event ItemActivated As EventHandler(Of LookupItemActivatedEventArgs)

    Friend Sub New(Owner As LookupBox)
        If Owner Is Nothing Then Throw New ArgumentNullException(NameOf(Owner))
        _Owner = Owner
        DoubleBuffered = True
        TabStop = False
        _Viewport = New Panel With {
            .Dock = DockStyle.Fill,
            .Margin = System.Windows.Forms.Padding.Empty,
            .Padding = System.Windows.Forms.Padding.Empty,
            .TabStop = False
        }
        _VerticalScroll = New VScrollBar With {
            .Dock = DockStyle.Right,
            .Visible = False,
            .TabStop = False
        }
        Controls.Add(_Viewport)
        Controls.Add(_VerticalScroll)
        AddHandler _VerticalScroll.ValueChanged, AddressOf ScrollValueChanged
        AddHandler _VerticalScroll.MouseWheel, AddressOf ChildMouseWheel
        AddHandler _Viewport.MouseWheel, AddressOf ChildMouseWheel
        AddHandler _Viewport.Resize, AddressOf ViewportResized
    End Sub

    Friend Sub ApplyOwnerSettings()
        _RowHeight = _Owner.TemplateItemHeight
        _RowSpacing = _Owner.TemplateItemSpacing
        BackColor = _Owner.DropDownBackColor
        _Viewport.BackColor = _Owner.DropDownBackColor
        Font = _Owner.Font
        RebuildLayout()
    End Sub

    Friend Sub ShowItems(Items As IReadOnlyList(Of Object))
        _Items = If(Items, Array.Empty(Of Object)())
        If Not Object.Equals(_TemplateType, _Owner.ItemTemplateType) Then
            DisposeRows()
            _TemplateType = _Owner.ItemTemplateType
        End If
        _SelectedIndex = If(_Items.Count = 0, -1, 0)
        For Each Host As LookupTemplateRow In _Rows
            Host.Tag = Nothing   'A new search may reuse the same index with a different object.
        Next
        RebuildLayout(True)
    End Sub

    Friend ReadOnly Property SelectedItem As Object
        Get
            If _SelectedIndex < 0 OrElse _SelectedIndex >= _Items.Count Then Return Nothing
            Return _Items(_SelectedIndex)
        End Get
    End Property

    Friend Sub MoveSelection(Offset As Integer)
        If _Items.Count = 0 Then Return
        SetSelectedIndex(Math.Max(0, Math.Min(_Items.Count - 1, _SelectedIndex + Offset)))
    End Sub

    Friend Sub JumpSelection(ToLast As Boolean)
        If _Items.Count = 0 Then Return
        SetSelectedIndex(If(ToLast, _Items.Count - 1, 0))
    End Sub

    Friend Sub PageSelection(Down As Boolean)
        Dim StepCount As Integer = Math.Max(1, _Viewport.ClientSize.Height \ Math.Max(1, RowStride))
        MoveSelection(If(Down, StepCount, -StepCount))
    End Sub

    Friend Sub ScrollByWheel(Delta As Integer)
        If Not _NeedsScroll OrElse Delta = 0 Then Return
        Dim Notches As Integer = Math.Max(1, CInt(Math.Abs(CLng(Delta)) \ 120L))
        Dim Direction As Integer = If(Delta > 0, -1, 1)
        ScrollTo(_VerticalScroll.Value + Direction * Notches * RowStride * 3)
    End Sub

    Private ReadOnly Property RowStride As Integer
        Get
            Return _RowHeight + _RowSpacing
        End Get
    End Property

    Private Function MaximumScrollValue() As Integer
        Return Math.Max(0, _VerticalScroll.Maximum - _VerticalScroll.LargeChange + 1)
    End Function

    Private Sub ScrollTo(Value As Integer)
        If Not _NeedsScroll Then Return
        Dim Clamped As Integer = Math.Max(_VerticalScroll.Minimum, Math.Min(MaximumScrollValue(), Value))
        If _VerticalScroll.Value <> Clamped Then _VerticalScroll.Value = Clamped
    End Sub

    Private Sub SetSelectedIndex(Index As Integer)
        If Index < 0 OrElse Index >= _Items.Count Then Return
        _SelectedIndex = Index
        EnsureSelectionVisible()
        RenderRows()
    End Sub

    Private Sub EnsureSelectionVisible()
        If Not _NeedsScroll OrElse _SelectedIndex < 0 Then Return
        Dim Top As Long = CLng(_SelectedIndex) * RowStride
        Dim Bottom As Long = Top + _RowHeight
        Dim ScrollTop As Integer = _VerticalScroll.Value
        If Top < ScrollTop Then
            ScrollTo(CInt(Top))
        ElseIf Bottom > CLng(ScrollTop) + _Viewport.ClientSize.Height Then
            ScrollTo(CInt(Bottom - _Viewport.ClientSize.Height))
        End If
    End Sub

    Private Sub RebuildLayout(Optional ResetScroll As Boolean = False)
        If _UpdatingLayout OrElse IsDisposed Then Return
        _UpdatingLayout = True
        Try
            Dim TotalHeight As Long = CLng(_Items.Count) * RowStride
            Dim NeedsScroll As Boolean = TotalHeight > _Viewport.ClientSize.Height
            _NeedsScroll = NeedsScroll
            _VerticalScroll.Visible = NeedsScroll
            Dim ViewportHeight As Integer = Math.Max(1, _Viewport.ClientSize.Height)
            _VerticalScroll.Minimum = 0
            _VerticalScroll.LargeChange = ViewportHeight
            _VerticalScroll.SmallChange = Math.Max(1, RowStride)
            _VerticalScroll.Maximum = CInt(Math.Min(Integer.MaxValue - 1L, Math.Max(0L, TotalHeight - 1L)))
            If ResetScroll OrElse Not NeedsScroll Then
                _VerticalScroll.Value = 0
            ElseIf _VerticalScroll.Value > MaximumScrollValue() Then
                _VerticalScroll.Value = MaximumScrollValue()
            End If
            Dim NeededRows As Integer = Math.Min(_Items.Count, CInt(Math.Ceiling(CDbl(ViewportHeight) / RowStride)) + 2)
            While _Rows.Count > NeededRows
                Dim Last As LookupTemplateRow = _Rows(_Rows.Count - 1)
                _Rows.RemoveAt(_Rows.Count - 1)
                _Viewport.Controls.Remove(Last)
                UnregisterClicks(Last)
                Last.Dispose()
            End While
            While _Rows.Count < NeededRows
                If _TemplateType Is Nothing Then Throw New InvalidOperationException("Configure ItemTemplateType before using Template view mode.")
                Dim View As UserControl = DirectCast(Activator.CreateInstance(_TemplateType), UserControl)
                View.Dock = DockStyle.Fill
                View.TabStop = False
                Dim Host As New LookupTemplateRow With {
                    .Height = _RowHeight,
                    .Padding = New Padding(4, 1, 1, 1),
                    .Margin = System.Windows.Forms.Padding.Empty,
                    .TabStop = False
                }
                Host.Controls.Add(View)
                _Rows.Add(Host)
                _Viewport.Controls.Add(Host)
                RegisterClicks(Host)
            End While
        Finally
            _UpdatingLayout = False
        End Try
        RenderRows()
    End Sub

    Private Sub RenderRows()
        If _UpdatingLayout OrElse IsDisposed Then Return
        Dim ScrollOffset As Integer = If(_NeedsScroll, _VerticalScroll.Value, 0)
        Dim FirstIndex As Integer = ScrollOffset \ RowStride
        Dim WidthAvailable As Integer = Math.Max(1, _Viewport.ClientSize.Width)
        For Slot As Integer = 0 To _Rows.Count - 1
            Dim Host As LookupTemplateRow = _Rows(Slot)
            Dim Index As Integer = FirstIndex + Slot
            Dim ShowHost As Boolean = Index < _Items.Count
            Host.Visible = ShowHost
            If Not ShowHost Then Continue For
            Host.Bounds = New Rectangle(0, Index * RowStride - ScrollOffset, WidthAvailable, _RowHeight)
            Dim WasIndex As Integer = If(TypeOf Host.Tag Is Integer, CInt(Host.Tag), -1)
            Host.Tag = Index
            Dim IsSelected As Boolean = Index = _SelectedIndex
            Host.Highlighted = IsSelected
            Host.BackColor = If(IsSelected, _Owner.SelectionBackColor, _Owner.DropDownBackColor)
            Host.ForeColor = If(IsSelected, _Owner.SelectionForeColor, _Owner.DropDownForeColor)
            Dim View As UserControl = DirectCast(Host.Controls(0), UserControl)
            If WasIndex <> Index Then
                _Owner.BindTemplateControl(View, _Items(Index), Index)
            End If
            Dim Aware As ILookupItemSelectionAware = TryCast(View, ILookupItemSelectionAware)
            If Aware IsNot Nothing Then Aware.SetHighlighted(IsSelected)
        Next
    End Sub

    Private Sub ScrollValueChanged(Sender As Object, E As EventArgs)
        If Not _UpdatingLayout Then RenderRows()
    End Sub

    Private Sub ViewportResized(Sender As Object, E As EventArgs)
        RebuildLayout()
    End Sub

    Private Sub ChildMouseWheel(Sender As Object, E As MouseEventArgs)
        ScrollByWheel(E.Delta)
        Dim Handled As HandledMouseEventArgs = TryCast(E, HandledMouseEventArgs)
        If Handled IsNot Nothing Then Handled.Handled = True
    End Sub

    Protected Overrides Sub OnMouseWheel(E As MouseEventArgs)
        ScrollByWheel(E.Delta)
        Dim Handled As HandledMouseEventArgs = TryCast(E, HandledMouseEventArgs)
        If Handled IsNot Nothing Then Handled.Handled = True
    End Sub

    Private Sub RegisterClicks(Root As Control)
        AddHandler Root.Click, AddressOf TemplateChildClick
        AddHandler Root.MouseWheel, AddressOf ChildMouseWheel
        AddHandler Root.ControlAdded, AddressOf TemplateControlAdded
        AddHandler Root.ControlRemoved, AddressOf TemplateControlRemoved
        For Each Child As Control In Root.Controls
            Child.TabStop = False
            RegisterClicks(Child)
        Next
    End Sub

    Private Sub UnregisterClicks(Root As Control)
        RemoveHandler Root.Click, AddressOf TemplateChildClick
        RemoveHandler Root.MouseWheel, AddressOf ChildMouseWheel
        RemoveHandler Root.ControlAdded, AddressOf TemplateControlAdded
        RemoveHandler Root.ControlRemoved, AddressOf TemplateControlRemoved
        For Each Child As Control In Root.Controls
            UnregisterClicks(Child)
        Next
    End Sub

    Private Sub TemplateControlAdded(Sender As Object, E As ControlEventArgs)
        E.Control.TabStop = False
        RegisterClicks(E.Control)
    End Sub

    Private Sub TemplateControlRemoved(Sender As Object, E As ControlEventArgs)
        UnregisterClicks(E.Control)
    End Sub

    Private Sub TemplateChildClick(Sender As Object, E As EventArgs)
        Dim Current As Control = TryCast(Sender, Control)
        While Current IsNot Nothing AndAlso Not TypeOf Current Is LookupTemplateRow
            Current = Current.Parent
        End While
        If Current Is Nothing OrElse Not TypeOf Current.Tag Is Integer Then Return
        Dim Index As Integer = CInt(Current.Tag)
        If Index < 0 OrElse Index >= _Items.Count Then Return
        _SelectedIndex = Index
        RenderRows()
        RaiseEvent ItemActivated(Me, New LookupItemActivatedEventArgs(_Items(Index)))
    End Sub

    Private Sub DisposeRows()
        For Each Host As LookupTemplateRow In _Rows
            _Viewport.Controls.Remove(Host)
            UnregisterClicks(Host)
            Host.Dispose()
        Next
        _Rows.Clear()
    End Sub

    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing Then
            RemoveHandler _VerticalScroll.ValueChanged, AddressOf ScrollValueChanged
            RemoveHandler _VerticalScroll.MouseWheel, AddressOf ChildMouseWheel
            RemoveHandler _Viewport.MouseWheel, AddressOf ChildMouseWheel
            RemoveHandler _Viewport.Resize, AddressOf ViewportResized
            DisposeRows()
        End If
        MyBase.Dispose(Disposing)
    End Sub
End Class

''' <summary>Recyclable row shell that keeps the active selection visible around its template.</summary>
<DesignerCategory("Code")>
Friend NotInheritable Class LookupTemplateRow
    Inherits Panel

    Friend Property Highlighted As Boolean

    Friend Sub New()
        DoubleBuffered = True
        TabStop = False
    End Sub
End Class
