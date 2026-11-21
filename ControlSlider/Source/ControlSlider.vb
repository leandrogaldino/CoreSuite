Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Adds mouse-driven dragging to a Windows Forms control, with optional handle,
''' movement constraints, grid snapping and animated snap-to-grid.
''' </summary>
<DesignerCategory("Component")>
<DefaultProperty("Child")>
<ToolboxItem(True)>
Public Partial Class ControlSlider
    Inherits Component

    Private _Parent As Control
    Private _Child As Control
    Private _DragHandle As Control
    Private _Enabled As Boolean = True
    Private _MovementAxis As SliderMovementAxis = SliderMovementAxis.Both
    Private _ConstrainToParent As Boolean = True
    Private _BoundsPadding As Padding = Padding.Empty
    Private _SnapToGrid As Boolean
    Private _GridSize As Integer = 8
    Private _SnapAnimation As Boolean
    Private _AnimationDuration As Integer = 180
    Private _BringToFrontOnDrag As Boolean = True
    Private _DragThreshold As Integer = 4
    Private _IncludeHandleChildren As Boolean = True

    ''' <summary>Occurs when the pointer has moved far enough to begin dragging.</summary>
    Public Event DragStarted As EventHandler(Of ControlSliderDragEventArgs)

    ''' <summary>Occurs after the dragged control changes location.</summary>
    Public Event Dragging As EventHandler(Of ControlSliderDragEventArgs)

    ''' <summary>Occurs when dragging finishes, including any snap animation.</summary>
    Public Event DragCompleted As EventHandler(Of ControlSliderDragEventArgs)

    Public Sub New()
    End Sub

    ''' <summary>Creates the component and registers it with a designer container.</summary>
    Public Sub New(container As IContainer)
        Me.New()
        If container Is Nothing Then Throw New ArgumentNullException(NameOf(container))
        container.Add(Me)
    End Sub

    <Category("ControlSlider - Targets")>
    <Description("The ancestor container whose client area defines the movement bounds.")>
    Public Property Parent As Control
        Get
            Return _Parent
        End Get
        Set(value As Control)
            If ReferenceEquals(_Parent, value) Then Return
            CancelCurrentOperation()
            If _Parent IsNot Nothing Then RemoveHandler _Parent.Disposed, AddressOf ParentDisposed
            _Parent = value
            If _Parent IsNot Nothing Then AddHandler _Parent.Disposed, AddressOf ParentDisposed
            RefreshInputHandlers()
        End Set
    End Property

    <Category("ControlSlider - Targets")>
    <Description("The control to move. It must be a descendant of Parent.")>
    Public Property Child As Control
        Get
            Return _Child
        End Get
        Set(value As Control)
            If ReferenceEquals(_Child, value) Then Return
            CancelCurrentOperation()
            DetachInputHandlers()
            If _Child IsNot Nothing Then
                RemoveHandler _Child.ParentChanged, AddressOf ChildParentChanged
                RemoveHandler _Child.Disposed, AddressOf ChildDisposed
            End If
            _Child = value
            If _Child IsNot Nothing Then
                AddHandler _Child.ParentChanged, AddressOf ChildParentChanged
                AddHandler _Child.Disposed, AddressOf ChildDisposed
            End If
            RefreshInputHandlers()
        End Set
    End Property

    <Category("ControlSlider - Targets")>
    <Description("Optional handle within Child. Only this area starts dragging. Leave empty to drag using Child itself.")>
    Public Property DragHandle As Control
        Get
            Return _DragHandle
        End Get
        Set(value As Control)
            If ReferenceEquals(_DragHandle, value) Then Return
            _DragHandle = value
            RefreshInputHandlers()
        End Set
    End Property

    <Category("ControlSlider - Behavior")>
    <Description("Enables mouse-driven dragging.")>
    <DefaultValue(True)>
    Public Property Enabled As Boolean
        Get
            Return _Enabled
        End Get
        Set(value As Boolean)
            If _Enabled = value Then Return
            _Enabled = value
            If Not value Then CancelCurrentOperation()
        End Set
    End Property

    <Category("ControlSlider - Behavior")>
    <Description("Controls whether movement is horizontal, vertical or unrestricted.")>
    <DefaultValue(GetType(SliderMovementAxis), "Both")>
    Public Property MovementAxis As SliderMovementAxis
        Get
            Return _MovementAxis
        End Get
        Set(value As SliderMovementAxis)
            If Not [Enum].IsDefined(GetType(SliderMovementAxis), value) Then
                Throw New InvalidEnumArgumentException(NameOf(value), CInt(value), GetType(SliderMovementAxis))
            End If
            _MovementAxis = value
        End Set
    End Property

    <Category("ControlSlider - Behavior")>
    <Description("Restricts the child to the parent's client area, respecting BoundsPadding.")>
    <DefaultValue(True)>
    Public Property ConstrainToParent As Boolean
        Get
            Return _ConstrainToParent
        End Get
        Set(value As Boolean)
            _ConstrainToParent = value
        End Set
    End Property

    <Category("ControlSlider - Behavior")>
    <Description("Brings the child to the front when dragging begins.")>
    <DefaultValue(True)>
    Public Property BringToFrontOnDrag As Boolean
        Get
            Return _BringToFrontOnDrag
        End Get
        Set(value As Boolean)
            _BringToFrontOnDrag = value
        End Set
    End Property

    <Category("ControlSlider - Behavior")>
    <Description("Minimum mouse travel in pixels before dragging begins.")>
    <DefaultValue(4)>
    Public Property DragThreshold As Integer
        Get
            Return _DragThreshold
        End Get
        Set(value As Integer)
            If value < 0 Then Throw New ArgumentOutOfRangeException(NameOf(value), "DragThreshold must be zero or greater.")
            _DragThreshold = value
        End Set
    End Property

    <Category("ControlSlider - Behavior")>
    <Description("Allows passive child controls such as labels inside DragHandle to initiate dragging.")>
    <DefaultValue(True)>
    Public Property IncludeHandleChildren As Boolean
        Get
            Return _IncludeHandleChildren
        End Get
        Set(value As Boolean)
            If _IncludeHandleChildren = value Then Return
            _IncludeHandleChildren = value
            RefreshInputHandlers()
        End Set
    End Property

    <Category("ControlSlider - Bounds")>
    <Description("Space reserved inside Parent's client area when clamping or snapping.")>
    Public Property BoundsPadding As Padding
        Get
            Return _BoundsPadding
        End Get
        Set(value As Padding)
            If value.Left < 0 OrElse value.Top < 0 OrElse value.Right < 0 OrElse value.Bottom < 0 Then
                Throw New ArgumentOutOfRangeException(NameOf(value), "BoundsPadding cannot contain negative values.")
            End If
            _BoundsPadding = value
        End Set
    End Property

    Public Function ShouldSerializeBoundsPadding() As Boolean
        Return _BoundsPadding <> Padding.Empty
    End Function

    Public Sub ResetBoundsPadding()
        _BoundsPadding = Padding.Empty
    End Sub

    <Category("ControlSlider - Snap")>
    <Description("Aligns the child to the nearest grid point when dragging ends.")>
    <DefaultValue(False)>
    Public Property SnapToGrid As Boolean
        Get
            Return _SnapToGrid
        End Get
        Set(value As Boolean)
            _SnapToGrid = value
        End Set
    End Property

    <Category("ControlSlider - Snap")>
    <Description("Grid spacing in pixels.")>
    <DefaultValue(8)>
    Public Property GridSize As Integer
        Get
            Return _GridSize
        End Get
        Set(value As Integer)
            If value < 1 Then Throw New ArgumentOutOfRangeException(NameOf(value), "GridSize must be greater than zero.")
            _GridSize = value
        End Set
    End Property

    <Category("ControlSlider - Snap")>
    <Description("Animates movement to the snapped position after mouse release.")>
    <DefaultValue(False)>
    Public Property SnapAnimation As Boolean
        Get
            Return _SnapAnimation
        End Get
        Set(value As Boolean)
            _SnapAnimation = value
        End Set
    End Property

    <Category("ControlSlider - Snap")>
    <Description("Duration of the snap animation, in milliseconds. Zero disables animation.")>
    <DefaultValue(180)>
    Public Property AnimationDuration As Integer
        Get
            Return _AnimationDuration
        End Get
        Set(value As Integer)
            If value < 0 Then Throw New ArgumentOutOfRangeException(NameOf(value), "AnimationDuration cannot be negative.")
            _AnimationDuration = value
        End Set
    End Property

    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property IsDragging As Boolean
        Get
            Return _IsDragging
        End Get
    End Property

    Private Sub ParentDisposed(sender As Object, e As EventArgs)
        Parent = Nothing
    End Sub

    Private Sub ChildDisposed(sender As Object, e As EventArgs)
        Child = Nothing
    End Sub

    Private Sub ChildParentChanged(sender As Object, e As EventArgs)
        RefreshInputHandlers()
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            CancelCurrentOperation()
            DetachInputHandlers()
            If _Parent IsNot Nothing Then RemoveHandler _Parent.Disposed, AddressOf ParentDisposed
            If _Child IsNot Nothing Then
                RemoveHandler _Child.ParentChanged, AddressOf ChildParentChanged
                RemoveHandler _Child.Disposed, AddressOf ChildDisposed
            End If
            DisposeAnimationResources()
            _Parent = Nothing
            _Child = Nothing
            _DragHandle = Nothing
        End If
        MyBase.Dispose(disposing)
    End Sub

    Private Sub OnDragStarted()
        RaiseEvent DragStarted(Me, CreateDragEventArgs())
    End Sub

    Private Sub OnDragging()
        RaiseEvent Dragging(Me, CreateDragEventArgs())
    End Sub

    Private Sub OnDragCompleted(startLocation As Point)
        If _Child IsNot Nothing AndAlso Not _Child.IsDisposed Then
            RaiseEvent DragCompleted(Me, New ControlSliderDragEventArgs(startLocation, _Child.Location))
        End If
    End Sub

    Private Function CreateDragEventArgs() As ControlSliderDragEventArgs
        Return New ControlSliderDragEventArgs(_StartChildLocation, _Child.Location)
    End Function
End Class
