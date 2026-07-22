''' <summary>
''' Defines how animation frames are scaled inside the <see cref="AnimatedBox"/> control.
''' </summary>
Public Enum AnimationScaleMode
    ''' <summary>
    ''' Draws the image using its original size, centered within the control.
    ''' </summary>
    Normal
    ''' <summary>
    ''' Stretches the image to completely fill the control,
    ''' ignoring the original aspect ratio.
    ''' </summary>
    Fill
    ''' <summary>
    ''' Scales the image proportionally to fit inside the control
    ''' while keeping its aspect ratio and centering it.
    ''' </summary>
    Centrer
End Enum