Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports CoreSuite.Helpers
Public Class ControlSlider
    Inherits Component
    Private _Offset As Point = Point.Empty
    Private _Parent As Control
    Private _Child As Control
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
    Private Sub Ctrl_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            _Offset = New Point(e.X, e.Y)
        End If
    End Sub
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
    Private Sub Ctrl_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)
        If _Offset <> Point.Empty Then
            Dim newlocation As Point = Child.Location
            newlocation.X += e.X - _Offset.X
            newlocation.Y += e.Y - _Offset.Y
            Child.Location = newlocation
        End If
    End Sub
End Class



