Imports System.Data
Imports System.Threading
Imports MySql.Data.MySqlClient

''' <summary>A local transaction owned by the service or by an explicitly opened session.</summary>
''' <remarks>ExecuteInTransaction owns commit, rollback and disposal. In its callback, only pass this context to command options.</remarks>
Public NotInheritable Class MySqlTransactionContext
    Implements IDisposable

    Private ReadOnly _Transaction As MySqlTransaction
    Private _Disposed As Boolean

    Friend Sub New(transaction As MySqlTransaction)
        _Transaction = transaction
    End Sub

    Friend ReadOnly Property ProviderTransaction As MySqlTransaction
        Get
            If _Disposed Then Throw New ObjectDisposedException(NameOf(MySqlTransactionContext))
            If _Transaction.Connection Is Nothing Then Throw New InvalidOperationException("The transaction has already completed or been disposed.")
            Return _Transaction
        End Get
    End Property

    Public ReadOnly Property IsolationLevel As IsolationLevel
        Get
            Return ProviderTransaction.IsolationLevel
        End Get
    End Property

    Public Sub Commit()
        Try
            ProviderTransaction.Commit()
        Catch ex As MySqlException
            Throw New MySqlServiceException(ex)
        End Try
    End Sub

    Public Async Function CommitAsync(Optional CancellationToken As CancellationToken = Nothing) As Task
        Try
            Await ProviderTransaction.CommitAsync(CancellationToken).ConfigureAwait(False)
        Catch ex As MySqlException
            Throw New MySqlServiceException(ex)
        End Try
    End Function

    Public Sub Rollback()
        Try
            ProviderTransaction.Rollback()
        Catch ex As MySqlException
            Throw New MySqlServiceException(ex)
        End Try
    End Sub

    Public Async Function RollbackAsync(Optional CancellationToken As CancellationToken = Nothing) As Task
        Try
            Await ProviderTransaction.RollbackAsync(CancellationToken).ConfigureAwait(False)
        Catch ex As MySqlException
            Throw New MySqlServiceException(ex)
        End Try
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        Try
            If _Disposed Then Return
            _Disposed = True
            _Transaction.Dispose()
        Catch ex As MySqlException
            Throw New MySqlServiceException(ex)
        End Try
    End Sub
End Class
