''' <summary>
''' Specifies how a <see cref="SplitButton"/> responds when activated.
''' </summary>
Public Enum SplitButtonMode
    ''' <summary>
    ''' The main area raises the Click event, while the drop-down area opens the associated menu.
    ''' </summary>
    Split
    ''' <summary>
    ''' The entire button opens the associated menu.
    ''' </summary>
    DropDown
End Enum