''' <summary>
''' Specifies the position of the caption relative to the toggle switch.
''' </summary>
Public Enum ToggleButtonTextPosition
    ''' <summary>
    ''' Places the caption to the left of the switch.
    ''' </summary>
    Left
    ''' <summary>
    ''' Places the caption to the right of the switch.
    ''' </summary>
    Right
End Enum
''' <summary>
''' Specifies the easing function used by the toggle animation.
''' </summary>
Public Enum ToggleButtonAnimationEasing
    ''' <summary>
    ''' Uses constant-speed interpolation.
    ''' </summary>
    Linear
    ''' <summary>
    ''' Starts quickly and slows down near the destination.
    ''' </summary>
    EaseOutCubic
    ''' <summary>
    ''' Accelerates at the beginning and decelerates near the destination.
    ''' </summary>
    EaseInOutCubic
End Enum
''' <summary>
''' Determines whether the switch uses its logical size or grows with the control.
''' </summary>
Public Enum ToggleButtonTrackSizingMode
    ''' <summary>Stretches when there is no caption; uses TrackSize when a caption is present.</summary>
    Automatic
    ''' <summary>Uses TrackSize regardless of the control size.</summary>
    Fixed
    ''' <summary>Resizes the switch as its containing control is resized.</summary>
    Stretch
End Enum

''' <summary>Predefined accent palettes that can be applied via ApplyColorPreset.</summary>
Public Enum ToggleButtonColorPreset
    Blue
    Emerald
    Violet
    Coral
    Slate
End Enum
