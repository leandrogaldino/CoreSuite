Imports System.ComponentModel

''' <summary>
''' Provides data for the <see cref="FluidResizer.ResizeEnd"/> event.
''' </summary>
Public Class ResizeEndEventArgs
    Inherits EventArgs
    ''' <summary>
    ''' Gets the initial <see cref="Size"/> of the control before the resizing operation started.
    ''' </summary>
    ''' <value>
    ''' A <see cref="Size"/> structure representing the starting dimensions.
    ''' </value>
    <Category("ResizeEndEventArgs")>
    <Description("Gets the initial size of the control before the fluid resizing operation started.")>
    Public ReadOnly Property StartSize As Size
    ''' <summary>
    ''' Gets the final target <see cref="Size"/> of the control after the resizing operation completed.
    ''' </summary>
    ''' <value>
    ''' A <see cref="Size"/> structure representing the final dimensions.
    ''' </value>
    <Category("ResizeEndEventArgs")>
    <Description("Gets the final target size reached when the fluid resizing operation completed.")>
    Public ReadOnly Property EndSize As Size
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ResizeEndEventArgs"/> class with the specified initial and final sizes.
    ''' </summary>
    ''' <param name="StartSize">The initial <see cref="Size"/> of the control before resizing.</param>
    ''' <param name="EndSize">The target <see cref="Size"/> reached after resizing.</param>
    Public Sub New(StartSize As Size, EndSize As Size)
        Me.StartSize = StartSize
        Me.EndSize = EndSize
    End Sub
End Class
