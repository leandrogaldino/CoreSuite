Imports System.Data
Imports System.Threading
Imports MySql.Data.MySqlClient
Friend NotInheritable Class MySqlConnectionScope
    Implements IDisposable
    Private ReadOnly _OwnsConnection As Boolean
    Private _OpenedConnection As Boolean
    Friend ReadOnly Property Connection As MySqlConnection
    Private Sub New(connection As MySqlConnection, ownsConnection As Boolean)
        Me.Connection = connection
        _OwnsConnection = ownsConnection
    End Sub
    Friend Shared Function Create(client As MySqlClient, Connection As MySqlConnection, Transaction As MySqlTransaction) As MySqlConnectionScope
        ArgumentNullException.ThrowIfNull(client)
        If Transaction IsNot Nothing Then
            Dim TransactionConnection As MySqlConnection = Transaction.Connection
            If TransactionConnection Is Nothing Then Throw New InvalidOperationException("The transaction is no longer associated with a connection.")
            If Connection Is Nothing Then
                Connection = TransactionConnection
            ElseIf Not ReferenceEquals(Connection, TransactionConnection) Then
                Throw New ArgumentException("The supplied connection does not belong to the supplied transaction.", NameOf(Connection))
            End If
        End If
        If Connection Is Nothing Then Return New MySqlConnectionScope(client.CreateDatabaseConnection(), True)
        Return New MySqlConnectionScope(Connection, False)
    End Function
    Friend Sub Open()
        If Connection.State = ConnectionState.Open Then Return
        Connection.Open()
        _OpenedConnection = True
    End Sub
    Friend Async Function OpenAsync(CancellationToken As CancellationToken) As Task
        If Connection.State = ConnectionState.Open Then Return
        Await Connection.OpenAsync(CancellationToken).ConfigureAwait(False)
        _OpenedConnection = True
    End Function
    Public Sub Dispose() Implements IDisposable.Dispose
        If _OwnsConnection Then
            Connection.Dispose()
        ElseIf _OpenedConnection AndAlso Connection.State <> ConnectionState.Closed Then
            Connection.Close()
        End If
    End Sub
End Class
