''' <summary>
''' Specifies the indicator drawn by a <see cref="BusyOverlay"/> while it is visible.
''' </summary>
Public Enum BusyOverlayIndicatorStyle
    ''' <summary>
    ''' Draws a rotating circular indicator for work whose completion percentage is unknown.
    ''' </summary>
    Spinner
    ''' <summary>
    ''' Draws an animated horizontal segment for work whose completion percentage is unknown.
    ''' </summary>
    MarqueeBar
    ''' <summary>
    ''' Draws a horizontal bar based on the configured progress range and value.
    ''' </summary>
    ProgressBar
    ''' <summary>
    ''' Does not draw an indicator.
    ''' </summary>
    None
End Enum
