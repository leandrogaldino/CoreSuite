''' <summary>
''' Specifies the possible states of the drop-down container's visibility.
''' </summary>
Public Enum ControlContainerDropDownState
    ''' <summary>
    ''' The drop-down is closed and not visible.
    ''' </summary>
    Closed
    ''' <summary>
    ''' The drop-down is in the process of closing.
    ''' </summary>
    Closing
    ''' <summary>
    ''' The drop-down is in the process of opening.
    ''' </summary>
    Dropping
    ''' <summary>
    ''' The drop-down is fully open and visible.
    ''' </summary>
    Dropped
End Enum