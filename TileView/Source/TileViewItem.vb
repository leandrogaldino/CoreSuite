Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Represents a designer-editable tile that can be hosted by a <see cref="TileView"/>.
''' </summary>
''' <remarks>
''' Inherit from this class instead of <see cref="UserControl"/> to create application-specific tiles while preserving the standard Windows Forms UserControl designer experience.
''' </remarks>
<DefaultEvent("Activated"), ToolboxItem(True), ToolboxItemFilter("CoreSuite")>
Public Class TileViewItem
    Inherits UserControl
    ''' <summary>
    ''' Defines the default border width used by an unselected item.
    ''' </summary>
    Private Const DefaultBorderWidth As Integer = 1
    ''' <summary>
    ''' Defines the default border width used by a selected item.
    ''' </summary>
    Private Const DefaultSelectedBorderWidth As Integer = 2
    ''' <summary>
    ''' Stores the current selection state.
    ''' </summary>
    Private _IsSelected As Boolean
    ''' <summary>
    ''' Stores the border color used when the item is not selected.
    ''' </summary>
    Private _BorderColor As Color = SystemColors.ControlDark
    ''' <summary>
    ''' Stores the border color used when the item is selected.
    ''' </summary>
    Private _SelectedBorderColor As Color = SystemColors.Highlight
    ''' <summary>
    ''' Stores the border width used when the item is not selected.
    ''' </summary>
    Private _BorderWidth As Integer = DefaultBorderWidth
    ''' <summary>
    ''' Stores the border width used when the item is selected.
    ''' </summary>
    Private _SelectedBorderWidth As Integer = DefaultSelectedBorderWidth
    ''' <summary>
    ''' Stores whether the item draws the standard Windows Forms focus rectangle while focused.
    ''' </summary>
    Private _ShowFocusRectangle As Boolean = True
    ''' <summary>
    ''' Stores whether receiving focus automatically selects the item.
    ''' </summary>
    Private _SelectOnFocus As Boolean = True
    ''' <summary>
    ''' Stores whether a mouse click moves keyboard focus to the item.
    ''' </summary>
    Private _FocusOnClick As Boolean = True
    ''' <summary>
    ''' Stores the <see cref="TileView"/> that currently owns the item.
    ''' </summary>
    Private _OwnerView As TileView
    ''' <summary>
    ''' Initializes a new instance of the <see cref="TileViewItem"/> class.
    ''' </summary>
    Public Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.Selectable, True)
        BorderStyle = BorderStyle.None
        Padding = New Padding(DefaultSelectedBorderWidth)
        TabStop = True
        AccessibleRole = AccessibleRole.ListItem
    End Sub
    ''' <summary>
    ''' Occurs when the item is activated by the owning <see cref="TileView"/>.
    ''' </summary>
    <Category("Action"), Description("Occurs when the item is activated by the owning TileView.")>
    Public Event Activated As EventHandler
    ''' <summary>
    ''' Gets a value indicating whether the item is currently selected.
    ''' </summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property IsSelected As Boolean
        Get
            Return _IsSelected
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets the border color used when the item is not selected.
    ''' </summary>
    <Category("Appearance"), Description("Specifies the border color used when the item is not selected."), DefaultValue(GetType(Color), "ControlDark")>
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
    ''' <summary>
    ''' Gets or sets the border color used when the item is selected.
    ''' </summary>
    <Category("Appearance"), Description("Specifies the border color used when the item is selected."), DefaultValue(GetType(Color), "Highlight")>
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
    ''' <summary>
    ''' Gets or sets the border width, in pixels, used when the item is not selected.
    ''' </summary>
    <Category("Appearance"), Description("Specifies the border width used when the item is not selected."), DefaultValue(DefaultBorderWidth)>
    Public Property BorderWidth As Integer
        Get
            Return _BorderWidth
        End Get
        Set(value As Integer)
            If value < 0 Then Throw New ArgumentOutOfRangeException(NameOf(value), "BorderWidth cannot be negative.")
            If _BorderWidth = value Then Return
            _BorderWidth = value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the border width, in pixels, used when the item is selected.
    ''' </summary>
    <Category("Appearance"), Description("Specifies the border width used when the item is selected."), DefaultValue(DefaultSelectedBorderWidth)>
    Public Property SelectedBorderWidth As Integer
        Get
            Return _SelectedBorderWidth
        End Get
        Set(value As Integer)
            If value < 0 Then Throw New ArgumentOutOfRangeException(NameOf(value), "SelectedBorderWidth cannot be negative.")
            If _SelectedBorderWidth = value Then Return
            _SelectedBorderWidth = value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the standard focus rectangle is drawn when the item has keyboard focus.
    ''' </summary>
    <Category("Appearance"), Description("Specifies whether the standard focus rectangle is drawn when the item has keyboard focus."), DefaultValue(True)>
    Public Property ShowFocusRectangle As Boolean
        Get
            Return _ShowFocusRectangle
        End Get
        Set(value As Boolean)
            If _ShowFocusRectangle = value Then Return
            _ShowFocusRectangle = value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the item becomes selected when it receives keyboard focus.
    ''' </summary>
    <Category("Behavior"), Description("Specifies whether the item becomes selected when it receives keyboard focus."), DefaultValue(True)>
    Public Property SelectOnFocus As Boolean
        Get
            Return _SelectOnFocus
        End Get
        Set(value As Boolean)
            _SelectOnFocus = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether clicking the item or one of its child controls moves keyboard focus to the item.
    ''' </summary>
    <Category("Behavior"), Description("Specifies whether clicking the item or one of its child controls moves keyboard focus to the item."), DefaultValue(True)>
    Public Property FocusOnClick As Boolean
        Get
            Return _FocusOnClick
        End Get
        Set(value As Boolean)
            _FocusOnClick = value
        End Set
    End Property
    ''' <summary>
    ''' Gets the <see cref="TileView"/> that currently owns the item.
    ''' </summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property OwnerView As TileView
        Get
            Return _OwnerView
        End Get
    End Property
    ''' <summary>
    ''' Requests selection of this item through its owning <see cref="TileView"/>.
    ''' </summary>
    Public Sub SelectItem()
        RequestSelection()
    End Sub
    ''' <summary>
    ''' Requests activation of this item through its owning <see cref="TileView"/>.
    ''' </summary>
    ''' <returns><see langword="True"/> when the item was activated; otherwise, <see langword="False"/>.</returns>
    Public Function ActivateItem() As Boolean
        If _OwnerView Is Nothing Then
            RaiseActivated()
            Return True
        End If
        Return _OwnerView.ActivateItem(Me)
    End Function
    ''' <summary>
    ''' Sets the owning <see cref="TileView"/> used for centralized selection and navigation.
    ''' </summary>
    ''' <param name="owner">The new owner, or <see langword="Nothing"/> when the item is detached.</param>
    Friend Sub SetOwner(owner As TileView)
        _OwnerView = owner
    End Sub
    ''' <summary>
    ''' Updates the selection state without notifying the owning <see cref="TileView"/>.
    ''' </summary>
    ''' <param name="value">The new selection state.</param>
    Friend Sub SetSelectedInternal(value As Boolean)
        If _IsSelected = value Then Return
        _IsSelected = value
        Refresh()
    End Sub
    ''' <summary>
    ''' Raises the <see cref="Activated"/> event.
    ''' </summary>
    Friend Sub RaiseActivated()
        RaiseEvent Activated(Me, EventArgs.Empty)
    End Sub
    ''' <summary>
    ''' Handles selection requests generated by mouse or keyboard interaction.
    ''' </summary>
    Private Sub RequestSelection()
        If _OwnerView Is Nothing Then
            SetSelectedInternal(True)
            Return
        End If
        _OwnerView.SelectItem(Me)
    End Sub
    ''' <summary>
    ''' Registers click forwarding for a child control and its descendants.
    ''' </summary>
    ''' <param name="control">The child control to register.</param>
    Private Sub RegisterControlHandlers(control As Control)
        RemoveHandler control.Click, AddressOf ChildControl_Click
        RemoveHandler control.DoubleClick, AddressOf ChildControl_DoubleClick
        RemoveHandler control.ControlAdded, AddressOf ChildControl_ControlAdded
        RemoveHandler control.ControlRemoved, AddressOf ChildControl_ControlRemoved
        AddHandler control.Click, AddressOf ChildControl_Click
        AddHandler control.DoubleClick, AddressOf ChildControl_DoubleClick
        AddHandler control.ControlAdded, AddressOf ChildControl_ControlAdded
        AddHandler control.ControlRemoved, AddressOf ChildControl_ControlRemoved
        For Each Child As Control In control.Controls
            RegisterControlHandlers(Child)
        Next
    End Sub
    ''' <summary>
    ''' Removes click forwarding from a child control and its descendants.
    ''' </summary>
    ''' <param name="control">The child control to unregister.</param>
    Private Sub UnregisterControlHandlers(control As Control)
        RemoveHandler control.Click, AddressOf ChildControl_Click
        RemoveHandler control.DoubleClick, AddressOf ChildControl_DoubleClick
        RemoveHandler control.ControlAdded, AddressOf ChildControl_ControlAdded
        RemoveHandler control.ControlRemoved, AddressOf ChildControl_ControlRemoved
        For Each Child As Control In control.Controls
            UnregisterControlHandlers(Child)
        Next
    End Sub
    ''' <summary>
    ''' Forwards a child control click to the tile itself.
    ''' </summary>
    Private Sub ChildControl_Click(sender As Object, e As EventArgs)
        OnClick(e)
    End Sub
    ''' <summary>
    ''' Forwards a child control double-click to the tile itself.
    ''' </summary>
    Private Sub ChildControl_DoubleClick(sender As Object, e As EventArgs)
        OnDoubleClick(e)
    End Sub
    ''' <summary>
    ''' Registers a child control that was added to an already registered descendant.
    ''' </summary>
    Private Sub ChildControl_ControlAdded(sender As Object, e As ControlEventArgs)
        RegisterControlHandlers(e.Control)
    End Sub
    ''' <summary>
    ''' Unregisters a child control that was removed from a registered descendant.
    ''' </summary>
    Private Sub ChildControl_ControlRemoved(sender As Object, e As ControlEventArgs)
        UnregisterControlHandlers(e.Control)
    End Sub
    ''' <summary>
    ''' Registers newly added direct child controls for whole-tile mouse interaction.
    ''' </summary>
    Protected Overrides Sub OnControlAdded(e As ControlEventArgs)
        MyBase.OnControlAdded(e)
        RegisterControlHandlers(e.Control)
    End Sub
    ''' <summary>
    ''' Unregisters removed direct child controls from whole-tile mouse interaction.
    ''' </summary>
    Protected Overrides Sub OnControlRemoved(e As ControlEventArgs)
        UnregisterControlHandlers(e.Control)
        MyBase.OnControlRemoved(e)
    End Sub
    ''' <summary>
    ''' Selects and focuses the item before raising the standard click event.
    ''' </summary>
    Protected Overrides Sub OnClick(e As EventArgs)
        RequestSelection()
        If _FocusOnClick AndAlso CanFocus Then Focus()
        MyBase.OnClick(e)
    End Sub
    ''' <summary>
    ''' Selects and focuses the item before raising the standard double-click event.
    ''' </summary>
    Protected Overrides Sub OnDoubleClick(e As EventArgs)
        RequestSelection()
        If _FocusOnClick AndAlso CanFocus Then Focus()
        MyBase.OnDoubleClick(e)
    End Sub
    ''' <summary>
    ''' Selects the item when it receives keyboard focus when <see cref="SelectOnFocus"/> is enabled.
    ''' </summary>
    Protected Overrides Sub OnEnter(e As EventArgs)
        If _SelectOnFocus Then RequestSelection()
        Invalidate()
        MyBase.OnEnter(e)
    End Sub
    ''' <summary>
    ''' Invalidates the item when keyboard focus leaves it so the focus cue is removed.
    ''' </summary>
    Protected Overrides Sub OnLeave(e As EventArgs)
        Invalidate()
        MyBase.OnLeave(e)
    End Sub
    ''' <summary>
    ''' Draws the standard and selected item borders and the optional Windows Forms focus rectangle.
    ''' </summary>
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        Dim Width = If(_IsSelected, _SelectedBorderWidth, _BorderWidth)
        Dim Color = If(_IsSelected, _SelectedBorderColor, _BorderColor)
        If Width > 0 AndAlso ClientSize.Width > Width AndAlso ClientSize.Height > Width Then
            Using Pen As New Pen(Color, Width)
                Dim Offset As Single = Width / 2.0F
                e.Graphics.DrawRectangle(Pen, Offset, Offset, ClientSize.Width - Width, ClientSize.Height - Width)
            End Using
        End If
        If _ShowFocusRectangle AndAlso Focused AndAlso ShowFocusCues AndAlso ClientSize.Width > 6 AndAlso ClientSize.Height > 6 Then
            Dim FocusBounds = Rectangle.Inflate(ClientRectangle, -3, -3)
            ControlPaint.DrawFocusRectangle(e.Graphics, FocusBounds, ForeColor, BackColor)
        End If
    End Sub
    ''' <summary>
    ''' Delegates list-navigation and activation keys to the owning <see cref="TileView"/>.
    ''' </summary>
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        If _OwnerView IsNot Nothing AndAlso _OwnerView.ProcessItemCommand(Me, keyData) Then Return True
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function
    ''' <summary>
    ''' Gets the default size used for newly created tile items.
    ''' </summary>
    Protected Overrides ReadOnly Property DefaultSize As Size
        Get
            Return New Size(240, 96)
        End Get
    End Property
End Class
