''' <summary>
''' Internal representation of a single animation frame.
''' </summary>
Public Class AnimationFrame
    ''' <summary>
    ''' Gets or sets the image associated with this frame.
    ''' </summary>
    Public Property Image As Image

    ''' <summary>
    ''' Gets or sets the delay (in seconds) before advancing to the next frame.
    ''' </summary>
    Public Property Delay As Double
    ''' <summary>
    ''' Initializes a new instance of the <see cref="AnimationFrameInfo"/> class.
    ''' </summary>
    ''' <param name="Image">The image to be displayed.</param>
    ''' <param name="Delay">
    ''' The frame delay in seconds. Default value is 0.03 (≈33 FPS).
    ''' </param>
    Public Sub New(Image As Image, Optional Delay As Double = 0.03)
        Me.Image = Image
        Me.Delay = Delay
    End Sub
End Class
