Imports System.IO
Imports System.Net.Http
Imports System.Threading

''' <summary>Provides cancellable HTTP content reads.</summary>
Public NotInheritable Class HttpHelper
    Private Sub New()
    End Sub
    ''' <summary>Reads HTTP content as text, disposing the content if cancelled.</summary>
    Public Shared Async Function ReadAsStringAsync(Content As HttpContent, Token As CancellationToken) As Task(Of String)
        If Content Is Nothing Then Throw New ArgumentNullException(NameOf(Content))
        Token.ThrowIfCancellationRequested()
        Using Registration = Token.Register(Sub() Content.Dispose())
            Try
                Dim Result = Await Content.ReadAsStringAsync().ConfigureAwait(False)
                Token.ThrowIfCancellationRequested()
                Return Result
            Catch When Token.IsCancellationRequested
                Throw New OperationCanceledException(Token)
            End Try
        End Using
    End Function
    ''' <summary>Obtains an HTTP content stream with cancellation while awaiting it.</summary>
    Public Shared Async Function ReadAsStreamAsync(Content As HttpContent, Token As CancellationToken) As Task(Of Stream)
        If Content Is Nothing Then Throw New ArgumentNullException(NameOf(Content))
        Token.ThrowIfCancellationRequested()
        Using Registration = Token.Register(Sub() Content.Dispose())
            Try
                Dim Result = Await Content.ReadAsStreamAsync().ConfigureAwait(False)
                Token.ThrowIfCancellationRequested()
                Return Result
            Catch When Token.IsCancellationRequested
                Throw New OperationCanceledException(Token)
            End Try
        End Using
    End Function
End Class
