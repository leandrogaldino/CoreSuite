''' <summary>
''' The exception that is thrown when an attempt is made to resolve a service
''' that has not been registered.
''' </summary>
Public Class ServiceNotRegisteredException
    Inherits InvalidOperationException
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ServiceNotRegisteredException"/> class
    ''' with a specified error message.
    ''' </summary>
    ''' <param name="Message">The message that describes the error.</param>
    Public Sub New(Message As String)
        MyBase.New(Message)
    End Sub
End Class
