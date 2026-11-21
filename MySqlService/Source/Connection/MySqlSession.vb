Imports System.Data
Imports System.Threading
Imports MySql.Data.MySqlClient

''' <summary>An explicitly owned connection. Dispose it after use; commands borrow it without taking ownership.</summary>
Public NotInheritable Class MySqlSession
    Implements IDisposable

    Private ReadOnly _Connection As MySqlConnection
    Private _Disposed As Boolean

    Friend Sub New(connection As MySqlConnection)
        _Connection = connection
    End Sub

    Friend ReadOnly Property ProviderConnection As MySqlConnection
        Get
            If _Disposed Then Throw New ObjectDisposedException(NameOf(MySqlSession))
            Return _Connection
        End Get
    End Property

    Public ReadOnly Property State As ConnectionState
        Get
            Return ProviderConnection.State
        End Get
    End Property

    Public Sub Open()
        Try
            ProviderConnection.Open()
        Catch ex As MySqlException
            Throw New MySqlServiceException(ex)
        End Try
    End Sub

    Public Async Function OpenAsync(Optional CancellationToken As CancellationToken = Nothing) As Task
        Try
            Await ProviderConnection.OpenAsync(CancellationToken).ConfigureAwait(False)
        Catch ex As MySqlException
            Throw New MySqlServiceException(ex)
        End Try
    End Function

    Public Sub Close()
        Try
            ProviderConnection.Close()
        Catch ex As MySqlException
            Throw New MySqlServiceException(ex)
        End Try
    End Sub

    Public Function BeginTransaction(Optional IsolationLevel As IsolationLevel = IsolationLevel.ReadCommitted) As MySqlTransactionContext
        Try
            Return New MySqlTransactionContext(ProviderConnection.BeginTransaction(IsolationLevel))
        Catch ex As MySqlException
            Throw New MySqlServiceException(ex)
        End Try
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        Try
            If _Disposed Then Return
            _Disposed = True
            _Connection.Dispose()
        Catch ex As MySqlException
            Throw New MySqlServiceException(ex)
        End Try
    End Sub
End Class
