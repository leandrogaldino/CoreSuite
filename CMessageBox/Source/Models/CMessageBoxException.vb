''' <summary>
''' Represents the information collected from an exception displayed by the CMessageBox.
''' </summary>
''' <remarks>
''' This class stores the exception details and additional context information used
''' for displaying, logging or sending error reports.
''' </remarks>
Public Class CMessageBoxException

    ''' <summary>
    ''' Gets or sets the title associated with the exception message.
    ''' </summary>
    Public Property Title As String

    ''' <summary>
    ''' Gets or sets the user-friendly message describing the error.
    ''' </summary>
    Public Property Message As String

    ''' <summary>
    ''' Gets or sets the main exception message provided by the exception instance.
    ''' </summary>
    Public Property ExceptionMessage As String

    ''' <summary>
    ''' Gets or sets the message from the inner exception, if available.
    ''' </summary>
    Public Property ExceptionInnerMessage As String

    ''' <summary>
    ''' Gets or sets the stack trace containing the execution path where the exception occurred.
    ''' </summary>
    Public Property StackTrace As String

    ''' <summary>
    ''' Gets or sets additional information associated with the exception.
    ''' </summary>
    ''' <remarks>
    ''' This collection can be used to provide extra contextual data to assist with
    ''' troubleshooting and error analysis.
    ''' </remarks>
    Public Property AdditionalInformations As Dictionary(Of String, Object)

    ''' <summary>
    ''' Gets or sets the date and time when the exception was captured.
    ''' </summary>
    Public Property ExceptionDate As Date

End Class