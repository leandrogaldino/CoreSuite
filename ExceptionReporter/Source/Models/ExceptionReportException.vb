''' <summary>
''' Represents one exception in the exception and inner-exception chain.
''' </summary>
Public Class ExceptionReportException
    ''' <summary>
    ''' Gets or sets the fully qualified exception type name.
    ''' </summary>
    Public Property Type As String
    ''' <summary>
    ''' Gets or sets the exception message.
    ''' </summary>
    Public Property Message As String
    ''' <summary>
    ''' Gets or sets the name of the application or object that caused the exception.
    ''' </summary>
    Public Property Source As String
    ''' <summary>
    ''' Gets or sets the exception stack trace.
    ''' </summary>
    Public Property StackTrace As String
    ''' <summary>
    ''' Gets or sets the exception HRESULT.
    ''' </summary>
    Public Property HResult As Integer
    ''' <summary>
    ''' Gets or sets the help link associated with the exception.
    ''' </summary>
    Public Property HelpLink As String
    ''' <summary>
    ''' Gets or sets the custom data attached to the exception.
    ''' </summary>
    Public Property Data As New Dictionary(Of String, String)()
    ''' <summary>
    ''' Gets or sets the exception that caused this exception, when available.
    ''' </summary>
    Public Property InnerException As ExceptionReportException
End Class
