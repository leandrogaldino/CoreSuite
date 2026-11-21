Imports System.Drawing

''' <summary>Positions of the dragged child before and during/after dragging.</summary>
Public NotInheritable Class ControlSliderDragEventArgs
    Inherits EventArgs

    Public Sub New(startLocation As Point, currentLocation As Point)
        Me.StartLocation = startLocation
        Me.CurrentLocation = currentLocation
    End Sub

    ''' <summary>Child.Location when the mouse button was pressed.</summary>
    Public ReadOnly Property StartLocation As Point

    ''' <summary>Child.Location at the moment the event is raised.</summary>
    Public ReadOnly Property CurrentLocation As Point
End Class
