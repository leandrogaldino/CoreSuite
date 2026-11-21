Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Designer-editable, reusable visual template for TileView. Override OnUnbound to release
''' record-specific resources; always fully overwrite fields in the SetItems binding callback.
''' </summary>
<DefaultEvent("Activated"), ToolboxItem(True)>
Public Class TileViewItem
    Inherits UserControl

    Private ReadOnly _HookedControls As New HashSet(Of Control)()
    Private _Owner As TileView
    Private _IsSelected As Boolean
    Private _ReadOnly As Boolean
    Private _BorderColor As Color = SystemColors.ControlDark
    Private _SelectedBorderColor As Color = SystemColors.Highlight
    Private _BorderWidth As Integer = 1
    Private _SelectedBorderWidth As Integer = 2
    Private _ShowFocusRectangle As Boolean = True
    Private _SelectOnFocus As Boolean = True
    Private _FocusOnClick As Boolean = True
    Private _BoundData As Object
    Private _DataIndex As Integer = -1
    Private _DisplayIndex As Integer = -1

    Public Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.UserPaint Or ControlStyles.ResizeRedraw Or ControlStyles.Selectable, True)
        MyBase.BorderStyle = BorderStyle.None
        Padding = New Padding(2)
        BackColor = SystemColors.Window
        TabStop = True
        AccessibleRole = AccessibleRole.ListItem
    End Sub

    <Category("Action")>
    Public Event Activated As EventHandler

    <Category("Property Changed")>
    Public Event ReadOnlyChanged As EventHandler

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property OwnerView As TileView
        Get
            Return _Owner
        End Get
    End Property

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property DataItem As Object
        Get
            Return _BoundData
        End Get
    End Property

    ''' <summary>Original position in the source collection, or -1 when unbound.</summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property DataIndex As Integer
        Get
            Return _DataIndex
        End Get
    End Property

    ''' <summary>Position after filtering, or -1 when unbound.</summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property DisplayIndex As Integer
        Get
            Return _DisplayIndex
        End Get
    End Property

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property IsSelected As Boolean
        Get
            Return _IsSelected
        End Get
    End Property

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property [ReadOnly] As Boolean
        Get
            Return _ReadOnly
        End Get
    End Property

    <Browsable(False), EditorBrowsable(EditorBrowsableState.Never), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Shadows Property BorderStyle As BorderStyle
        Get
            Return MyBase.BorderStyle
        End Get
        Set(value As BorderStyle)
            MyBase.BorderStyle = BorderStyle.None
        End Set
    End Property

    <Category("Appearance"), DefaultValue(GetType(Color), "ControlDark")>
    Public Property BorderColor As Color
        Get
            Return _BorderColor
        End Get
        Set(value As Color)
            If _BorderColor = value Then Return
            _BorderColor = value
            Invalidate()
        End Set
    End Property

    <Category("Appearance"), DefaultValue(GetType(Color), "Highlight")>
    Public Property SelectedBorderColor As Color
        Get
            Return _SelectedBorderColor
        End Get
        Set(value As Color)
            If _SelectedBorderColor = value Then Return
            _SelectedBorderColor = value
            Invalidate()
        End Set
    End Property

    <Category("Appearance"), DefaultValue(1)>
    Public Property BorderWidth As Integer
        Get
            Return _BorderWidth
        End Get
        Set(value As Integer)
            If value < 0 Then Throw New ArgumentOutOfRangeException(NameOf(value))
            _BorderWidth = value
            Invalidate()
        End Set
    End Property

    <Category("Appearance"), DefaultValue(2)>
    Public Property SelectedBorderWidth As Integer
        Get
            Return _SelectedBorderWidth
        End Get
        Set(value As Integer)
            If value < 0 Then Throw New ArgumentOutOfRangeException(NameOf(value))
            _SelectedBorderWidth = value
            Invalidate()
        End Set
    End Property

    <Category("Appearance"), DefaultValue(True)>
    Public Property ShowFocusRectangle As Boolean
        Get
            Return _ShowFocusRectangle
        End Get
        Set(value As Boolean)
            _ShowFocusRectangle = value
            Invalidate()
        End Set
    End Property

    <Category("Behavior"), DefaultValue(True)>
    Public Property SelectOnFocus As Boolean
        Get
            Return _SelectOnFocus
        End Get
        Set(value As Boolean)
            _SelectOnFocus = value
        End Set
    End Property

    ''' <summary>Only clicking the empty tile surface moves focus. Clicking an input control does not.</summary>
    <Category("Behavior"), DefaultValue(True)>
    Public Property FocusOnClick As Boolean
        Get
            Return _FocusOnClick
        End Get
        Set(value As Boolean)
            _FocusOnClick = value
        End Set
    End Property

    Public Sub SelectItem()
        If _Owner Is Nothing Then
            SetSelectedInternal(True)
        ElseIf _DisplayIndex >= 0 Then
            _Owner.SelectItem(Me)
        End If
    End Sub

    Public Function ActivateItem() As Boolean
        If _Owner Is Nothing Then
            RaiseActivated()
            Return True
        End If
        Return _Owner.ActivateItem(Me)
    End Function

    Friend Sub SetOwner(owner As TileView)
        _Owner = owner
    End Sub

    Friend Sub BindInternal(data As Object, dataIndex As Integer, displayIndex As Integer)
        _BoundData = data
        _DataIndex = dataIndex
        _DisplayIndex = displayIndex
    End Sub

    Friend Sub UnbindInternal()
        If _DataIndex = -1 Then Return
        _BoundData = Nothing
        _DataIndex = -1
        _DisplayIndex = -1
        SetSelectedInternal(False)
        OnUnbound(EventArgs.Empty)
    End Sub

    ''' <summary>Override to clear row-specific subscriptions/images before a tile is reused.</summary>
    Protected Overridable Sub OnUnbound(e As EventArgs)
    End Sub

    Friend Sub SetReadOnlyInternal(value As Boolean)
        If _ReadOnly = value Then Return
        _ReadOnly = value
        OnReadOnlyChanged(EventArgs.Empty)
    End Sub

    Protected Overridable Sub OnReadOnlyChanged(e As EventArgs)
        RaiseEvent ReadOnlyChanged(Me, e)
    End Sub

    Friend Sub SetSelectedInternal(value As Boolean)
        If _IsSelected = value Then Return
        _IsSelected = value
        Invalidate()
    End Sub

    Friend Sub RaiseActivated()
        RaiseEvent Activated(Me, EventArgs.Empty)
    End Sub

    Private Sub HookTree(Control As Control)
        If _HookedControls Is Nothing Then Return
        If _HookedControls.Add(Control) Then
            AddHandler Control.Click, AddressOf Descendant_Click
            AddHandler Control.DoubleClick, AddressOf Descendant_DoubleClick
            AddHandler Control.ControlAdded, AddressOf Descendant_ControlAdded
            AddHandler Control.ControlRemoved, AddressOf Descendant_ControlRemoved
        End If
        For Each child As Control In Control.Controls
            HookTree(child)
        Next
    End Sub

    Private Sub UnhookTree(Control As Control)
        If _HookedControls Is Nothing Then Return
        If _HookedControls.Remove(Control) Then
            RemoveHandler Control.Click, AddressOf Descendant_Click
            RemoveHandler Control.DoubleClick, AddressOf Descendant_DoubleClick
            RemoveHandler Control.ControlAdded, AddressOf Descendant_ControlAdded
            RemoveHandler Control.ControlRemoved, AddressOf Descendant_ControlRemoved
        End If
        For Each child As Control In Control.Controls
            UnhookTree(child)
        Next
    End Sub

    Private Sub Descendant_ControlAdded(sender As Object, e As ControlEventArgs)
        HookTree(e.Control)
    End Sub

    Private Sub Descendant_ControlRemoved(sender As Object, e As ControlEventArgs)
        UnhookTree(e.Control)
    End Sub

    Private Sub Descendant_Click(sender As Object, e As EventArgs)
        If _Owner IsNot Nothing Then
            _Owner.NotifyItemClick(Me, False)
        Else
            SetSelectedInternal(True)
        End If
        MyBase.OnClick(e)
    End Sub

    Private Sub Descendant_DoubleClick(sender As Object, e As EventArgs)
        If _Owner IsNot Nothing Then _Owner.NotifyItemDoubleClick(Me, False)
        MyBase.OnDoubleClick(e)
    End Sub

    Protected Overrides Sub OnControlAdded(e As ControlEventArgs)
        MyBase.OnControlAdded(e)
        HookTree(e.Control)
    End Sub

    Protected Overrides Sub OnControlRemoved(e As ControlEventArgs)
        UnhookTree(e.Control)
        MyBase.OnControlRemoved(e)
    End Sub

    Protected Overrides Sub OnClick(e As EventArgs)
        If _Owner IsNot Nothing Then
            _Owner.NotifyItemClick(Me, True)
        Else
            SetSelectedInternal(True)
        End If
        MyBase.OnClick(e)
    End Sub

    Protected Overrides Sub OnDoubleClick(e As EventArgs)
        If _Owner IsNot Nothing Then _Owner.NotifyItemDoubleClick(Me, True)
        MyBase.OnDoubleClick(e)
    End Sub

    Protected Overrides Sub OnEnter(e As EventArgs)
        MyBase.OnEnter(e)
        If _SelectOnFocus Then SelectItem()
        Invalidate()
    End Sub

    Protected Overrides Sub OnLeave(e As EventArgs)
        MyBase.OnLeave(e)
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseWheel(e As MouseEventArgs)
        If _Owner IsNot Nothing Then
            _Owner.ScrollFromItem(e.Delta)
            Dim handled As HandledMouseEventArgs = TryCast(e, HandledMouseEventArgs)
            If handled IsNot Nothing Then handled.Handled = True
        End If
        MyBase.OnMouseWheel(e)
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        Dim LineWidth As Integer = If(_IsSelected, _SelectedBorderWidth, _BorderWidth)
        Dim LineColor As Color = If(_IsSelected, _SelectedBorderColor, _BorderColor)
        If LineWidth > 0 AndAlso ClientSize.Width > LineWidth AndAlso ClientSize.Height > LineWidth Then
            Using Pen As New Pen(LineColor, CSng(LineWidth))
                Dim Inset As Single = LineWidth / 2.0F
                e.Graphics.DrawRectangle(Pen, Inset, Inset, ClientSize.Width - LineWidth, ClientSize.Height - LineWidth)
            End Using
        End If
        If _ShowFocusRectangle AndAlso Focused AndAlso ShowFocusCues AndAlso
           ClientSize.Width > 6 AndAlso ClientSize.Height > 6 Then
            ControlPaint.DrawFocusRectangle(e.Graphics, Rectangle.Inflate(ClientRectangle, -3, -3), ForeColor, BackColor)
        End If
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef Msg As Message, KeyData As Keys) As Boolean
        If _Owner IsNot Nothing AndAlso Focused AndAlso _Owner.ProcessItemCommand(Me, KeyData) Then Return True
        Return MyBase.ProcessCmdKey(Msg, KeyData)
    End Function

    Protected Overrides ReadOnly Property DefaultSize As Size
        Get
            Return New Size(260, 112)
        End Get
    End Property

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            For Each child As Control In Controls
                UnhookTree(child)
            Next
            _Owner = Nothing
            _BoundData = Nothing
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class
