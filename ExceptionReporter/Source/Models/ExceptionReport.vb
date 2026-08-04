''' <summary>
''' Represents a structured exception report that can be displayed, saved or sent.
''' </summary>
Public Class ExceptionReport
    ''' <summary>
    ''' Gets or sets the title associated with the error.
    ''' </summary>
    Public Property Title As String
    ''' <summary>
    ''' Gets or sets the user-friendly message associated with the error.
    ''' </summary>
    Public Property Message As String
    ''' <summary>
    ''' Gets or sets the main exception message.
    ''' </summary>
    Public Property ExceptionMessage As String
    ''' <summary>
    ''' Gets or sets the immediate inner-exception message, when available.
    ''' </summary>
    Public Property ExceptionInnerMessage As String
    ''' <summary>
    ''' Gets or sets the stack trace of the main exception.
    ''' </summary>
    Public Property StackTrace As String
    ''' <summary>
    ''' Gets or sets the steps or additional description supplied by the user.
    ''' </summary>
    Public Property UserSteps As String
    ''' <summary>
    ''' Gets or sets additional contextual information associated with the report.
    ''' </summary>
    Public Property AdditionalInformations As New Dictionary(Of String, Object)()
    ''' <summary>
    ''' Gets or sets the complete exception and inner-exception details.
    ''' </summary>
    Public Property ExceptionDetails As ExceptionReportException
    ''' <summary>
    ''' Gets or sets the local date and time at which the exception was captured.
    ''' </summary>
    Public Property ExceptionDate As Date
End Class
