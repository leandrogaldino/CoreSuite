''' <summary>
''' Defines the type of message displayed by the CMessageBox.
''' </summary>
''' <remarks>
''' The selected message type determines the visual appearance, icon and intended
''' purpose of the message shown to the user.
''' </remarks>
Public Enum CMessageBoxType

    ''' <summary>
    ''' Represents an informational message intended to provide details or guidance
    ''' to the user.
    ''' </summary>
    Information

    ''' <summary>
    ''' Represents a success message indicating that an operation was completed successfully.
    ''' </summary>
    Success

    ''' <summary>
    ''' Represents an error message indicating that an operation failed or an unexpected
    ''' problem occurred.
    ''' </summary>
    [Error]

    ''' <summary>
    ''' Represents a warning message used to notify the user about a situation that
    ''' requires attention or caution.
    ''' </summary>
    Warning

    ''' <summary>
    ''' Represents a question message used when user confirmation or a decision is required.
    ''' </summary>
    Question

End Enum