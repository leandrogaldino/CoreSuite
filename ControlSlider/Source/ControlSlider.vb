Imports System.ComponentModel
''' <summary>
''' Provides drag-and-drop movement for a child control within the bounds of a parent control.
''' </summary>
Public Class ControlSlider
    Inherits Component
    Private _Offset As Point = Point.Empty
    Private _Parent As Control
    Private _Child As Control
    ''' <summary>
    ''' Gets or sets the parent control that defines the movement boundaries.
    ''' </summary>
    Public Property Parent As Control
        Get
            Return _Parent
        End Get
        Set(value As Control)
            _Parent = value
            If _Parent IsNot Nothing And _Child IsNot Nothing Then
                AddHandler Child.MouseDown, AddressOf Ctrl_MouseDown
                AddHandler Child.MouseUp, AddressOf Ctrl_MouseUp
                AddHandler Child.MouseMove, AddressOf Ctrl_MouseMove
            End If
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the control that can be dragged within the parent control.
    ''' </summary>
    Public Property Child As Control
        Get
            Return _Child
        End Get
        Set(value As Control)
            _Child = value
            If _Parent IsNot Nothing And _Child IsNot Nothing Then
                AddHandler Child.MouseDown, AddressOf Ctrl_MouseDown
                AddHandler Child.MouseUp, AddressOf Ctrl_MouseUp
                AddHandler Child.MouseMove, AddressOf Ctrl_MouseMove
            End If
        End Set
    End Property
    ''' <summary>
    ''' Records the initial mouse position when dragging begins.
    ''' </summary>
    ''' <param name="sender">The event source.</param>
    ''' <param name="e">Mouse event data.</param>
    Private Sub Ctrl_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            _Offset = New Point(e.X, e.Y)
        End If
    End Sub
    ''' <summary>
    ''' Ends the drag operation and ensures the child control remains fully visible within the parent.
    ''' </summary>
    ''' <param name="sender">The event source.</param>
    ''' <param name="e">Mouse event data.</param>
    Private Sub Ctrl_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)
        _Offset = Point.Empty
        If Not ControlHelper.IsControlFullyVisible(_Parent, _Child) Then
            Dim myBounds As Rectangle = Parent.ClientRectangle
            If Child.Left < myBounds.Left Then
                Child.Location = New Point(0, Child.Location.Y)
            End If
            If Child.Right > myBounds.Right Then
                Child.Location = New Point(myBounds.Right - Child.Width, Child.Location.Y)
            End If
            If Child.Top < myBounds.Top Then
                Child.Location = New Point(Child.Location.X, 0)
            End If
            If Child.Bottom > myBounds.Bottom Then
                Child.Location = New Point(Child.Location.X, myBounds.Bottom - Child.Height)
            End If
        End If
    End Sub
    ''' <summary>
    ''' Updates the child control position while it is being dragged.
    ''' </summary>
    ''' <param name="sender">The event source.</param>
    ''' <param name="e">Mouse event data.</param>
    Private Sub Ctrl_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)
        If _Offset <> Point.Empty Then
            Dim newlocation As Point = Child.Location
            newlocation.X += e.X - _Offset.X
            newlocation.Y += e.Y - _Offset.Y
            Child.Location = newlocation
        End If
    End Sub
End Class



