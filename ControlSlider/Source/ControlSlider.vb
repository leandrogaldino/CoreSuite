Imports System.ComponentModel
''' <summary>
''' Provides drag-and-drop sliding capability for a child control within the boundary limits of a parent container control.
''' </summary>
Public Class ControlSlider
    Inherits Component
    Private _Offset As Point = Point.Empty
    Private _Parent As Control
    Private _Child As Control
    ''' <summary>
    ''' Gets or sets the parent container control that defines the movement boundaries.
    ''' </summary>
    <Category("ControlSlider")>
    <Description("Defines the parent container control that constrains the movement bounds of the child control.")>
    Public Property Parent As Control
        Get
            Return _Parent
        End Get
        Set(ByVal Value As Control)
            _Parent = Value
            AttachEvents()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the target child control that can be interactively dragged.
    ''' </summary>
    <Category("ControlSlider")>
    <Description("Defines the child control that can be interactively dragged within the parent container.")>
    Public Property Child As Control
        Get
            Return _Child
        End Get
        Set(ByVal Value As Control)
            DetachEvents()
            _Child = Value
            AttachEvents()
        End Set
    End Property
    ''' <summary>
    ''' Attaches the mouse interaction event handlers to the child control if both parent and child controls are set.
    ''' </summary>
    Private Sub AttachEvents()
        If _Parent IsNot Nothing AndAlso _Child IsNot Nothing Then
            DetachEvents()
            AddHandler _Child.MouseDown, AddressOf Ctrl_MouseDown
            AddHandler _Child.MouseUp, AddressOf Ctrl_MouseUp
            AddHandler _Child.MouseMove, AddressOf Ctrl_MouseMove
        End If
    End Sub
    ''' <summary>
    ''' Detaches the mouse interaction event handlers from the current child control to prevent memory leaks or duplicate handlers.
    ''' </summary>
    Private Sub DetachEvents()
        If _Child IsNot Nothing Then
            RemoveHandler _Child.MouseDown, AddressOf Ctrl_MouseDown
            RemoveHandler _Child.MouseUp, AddressOf Ctrl_MouseUp
            RemoveHandler _Child.MouseMove, AddressOf Ctrl_MouseMove
        End If
    End Sub
    ''' <summary>
    ''' Handles the MouseDown event to record the initial mouse cursor offset when starting a drag operation.
    ''' </summary>
    ''' <param name="Sender">The event source.</param>
    ''' <param name="E">The mouse event arguments containing mouse coordinates and button state.</param>
    Private Sub Ctrl_MouseDown(ByVal Sender As Object, ByVal E As MouseEventArgs)
        If E.Button = MouseButtons.Left Then
            _Offset = New Point(E.X, E.Y)
        End If
    End Sub
    ''' <summary>
    ''' Handles the MouseUp event to reset the drag offset when the mouse button is released.
    ''' </summary>
    ''' <param name="Sender">The event source.</param>
    ''' <param name="E">The mouse event arguments containing mouse coordinates and button state.</param>
    Private Sub Ctrl_MouseUp(ByVal Sender As Object, ByVal E As MouseEventArgs)
        If E.Button = MouseButtons.Left Then
            _Offset = Point.Empty
        End If
    End Sub
    ''' <summary>
    ''' Handles the MouseMove event to update the position of the child control while clamping its coordinates within the parent container's bounds.
    ''' </summary>
    ''' <param name="Sender">The event source.</param>
    ''' <param name="E">The mouse event arguments containing current mouse coordinates.</param>
    Private Sub Ctrl_MouseMove(ByVal Sender As Object, ByVal E As MouseEventArgs)
        If E.Button = MouseButtons.Left AndAlso _Offset <> Point.Empty AndAlso _Parent IsNot Nothing AndAlso _Child IsNot Nothing Then
            Dim TargetX As Integer = _Child.Left + (E.X - _Offset.X)
            Dim TargetY As Integer = _Child.Top + (E.Y - _Offset.Y)
            Dim MaxX As Integer = Math.Max(0, _Parent.ClientSize.Width - _Child.Width)
            Dim MaxY As Integer = Math.Max(0, _Parent.ClientSize.Height - _Child.Height)
            Dim ClampedX As Integer = Math.Max(0, Math.Min(TargetX, MaxX))
            Dim ClampedY As Integer = Math.Max(0, Math.Min(TargetY, MaxY))
            _Child.Location = New Point(ClampedX, ClampedY)
        End If
    End Sub
End Class