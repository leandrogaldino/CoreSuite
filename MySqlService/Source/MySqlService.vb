Imports System.Threading
Imports MySql.Data.MySqlClient
''' <summary>
''' Provides a single immutable entry point for MySQL connections, commands, CRUD operations, stored procedures, transactions, backup, restore, and database creation.
''' </summary>
Public NotInheritable Class MySqlService
    ''' <summary>
    ''' Initializes a new instance of the <see cref="MySqlService"/> class from individual connection values.
    ''' </summary>
    ''' <param name="Server">The MySQL server address.</param>
    ''' <param name="Database">The default database name.</param>
    ''' <param name="User">The user name.</param>
    ''' <param name="Password">The password.</param>
    ''' <param name="Configure">An optional callback used to configure additional <see cref="MySqlConnectionOptions"/> properties.</param>
    Public Sub New(Server As String, Database As String, User As String, Password As String, Optional Configure As Action(Of MySqlConnectionOptions) = Nothing)
        Me.New(BuildConnectionString(Server, Database, User, Password, Configure))
    End Sub
    ''' <summary>
    ''' Initializes a service by taking a snapshot of the supplied settings.
    ''' </summary>
    ''' <param name="Options">Connection settings to snapshot.</param>
    Public Sub New(Options As MySqlConnectionOptions)
        Me.New(GetOptionsConnectionString(Options))
    End Sub
    Private Shared Function GetOptionsConnectionString(Options As MySqlConnectionOptions) As String
        If Options Is Nothing Then Throw New ArgumentNullException(NameOf(Options))
        Return Options.BuildConnectionString()
    End Function
    ''' <summary>Initializes a service from a complete connection string.</summary>
    Public Sub New(ConnectionString As String)
        Try
            Client = New MySqlClient(ConnectionString)
            Request = New MySqlRequest(Client)
            Maintenance = New MySqlMaintenance(Client)
        Catch ex As MySqlException
            Throw New MySqlServiceException(ex)
        End Try
    End Sub
    ''' <summary>
    ''' Gets the connection factory and immutable server information.
    ''' </summary>
    Public ReadOnly Property Client As MySqlClient
    ''' <summary>
    ''' Gets the command, query, CRUD, and stored procedure service.
    ''' </summary>
    Public ReadOnly Property Request As MySqlRequest
    ''' <summary>
    ''' Gets the database creation, backup, and restore service.
    ''' </summary>
    Public ReadOnly Property Maintenance As MySqlMaintenance
    ''' <summary>
    ''' Executes an operation inside a new local transaction and commits it when the operation succeeds.
    ''' </summary>
    ''' <typeparam name="TResult">The operation result type.</typeparam>
    ''' <param name="Operation">The operation that receives the active transaction.</param>
    ''' <param name="IsolationLevel">The transaction isolation level.</param>
    ''' <returns>The value returned by <paramref name="Operation"/>.</returns>
    Public Function ExecuteInTransaction(Of TResult)(Operation As Func(Of MySqlTransactionContext, TResult), Optional IsolationLevel As IsolationLevel = IsolationLevel.ReadCommitted) As TResult
        Try
            If Operation Is Nothing Then Throw New ArgumentNullException("Operation")
            Using Connection As MySqlConnection = Client.CreateDatabaseProviderConnection()
                Connection.Open()
                Using Transaction As MySqlTransaction = Connection.BeginTransaction(IsolationLevel)
                    Try
                        Dim Result As TResult = Operation(New MySqlTransactionContext(Transaction))
                        Transaction.Commit()
                        Return Result
                    Catch
                        TryRollback(Transaction)
                        Throw
                    End Try
                End Using
            End Using
        Catch ex As MySqlException
            Throw New MySqlServiceException(ex)
        End Try
    End Function
    ''' <summary>
    ''' Executes an action inside a new local transaction and commits it when the action succeeds.
    ''' </summary>
    ''' <param name="Operation">The action that receives the active transaction.</param>
    ''' <param name="IsolationLevel">The transaction isolation level.</param>
    Public Sub ExecuteInTransaction(Operation As Action(Of MySqlTransactionContext), Optional IsolationLevel As IsolationLevel = IsolationLevel.ReadCommitted)
        Try
            If Operation Is Nothing Then Throw New ArgumentNullException("Operation")
            Using Connection As MySqlConnection = Client.CreateDatabaseProviderConnection()
                Connection.Open()
                Using Transaction As MySqlTransaction = Connection.BeginTransaction(IsolationLevel)
                    Try
                        Operation(New MySqlTransactionContext(Transaction))
                        Transaction.Commit()
                    Catch
                        TryRollback(Transaction)
                        Throw
                    End Try
                End Using
            End Using
        Catch ex As MySqlException
            Throw New MySqlServiceException(ex)
        End Try
    End Sub
    ''' <summary>
    ''' Asynchronously executes an operation inside a new local transaction and commits it when the operation succeeds.
    ''' </summary>
    ''' <typeparam name="TResult">The operation result type.</typeparam>
    ''' <param name="Operation">The asynchronous operation that receives the active transaction and cancellation token.</param>
    ''' <param name="IsolationLevel">The transaction isolation level.</param>
    ''' <param name="CancellationToken">The token used to cancel connection opening, the operation, or commit.</param>
    ''' <returns>The value returned by <paramref name="Operation"/>.</returns>
    Public Async Function ExecuteInTransactionAsync(Of TResult)(Operation As Func(Of MySqlTransactionContext, CancellationToken, Task(Of TResult)), Optional IsolationLevel As IsolationLevel = IsolationLevel.ReadCommitted, Optional CancellationToken As CancellationToken = Nothing) As Task(Of TResult)
        Try
            If Operation Is Nothing Then Throw New ArgumentNullException("Operation")
            Using Connection As MySqlConnection = Client.CreateDatabaseProviderConnection()
                Await Connection.OpenAsync(CancellationToken).ConfigureAwait(False)
                Using Transaction As MySqlTransaction = Connection.BeginTransaction(IsolationLevel)
                    Try
                        Dim OperationTask As Task(Of TResult) = Operation(New MySqlTransactionContext(Transaction), CancellationToken)
                        If OperationTask Is Nothing Then Throw New InvalidOperationException("The transaction operation returned Nothing instead of a Task.")
                        Dim Result As TResult = Await OperationTask.ConfigureAwait(False)
                        Await Transaction.CommitAsync(CancellationToken).ConfigureAwait(False)
                        Return Result
                    Catch
                        TryRollback(Transaction)
                        Throw
                    End Try
                End Using
            End Using
        Catch ex As MySqlException
            Throw New MySqlServiceException(ex)
        End Try
    End Function
    ''' <summary>
    ''' Asynchronously executes an action inside a new local transaction and commits it when the action succeeds.
    ''' </summary>
    ''' <param name="Operation">The asynchronous action that receives the active transaction and cancellation token.</param>
    ''' <param name="IsolationLevel">The transaction isolation level.</param>
    ''' <param name="CancellationToken">The token used to cancel connection opening, the action, or commit.</param>
    ''' <returns>A task representing the complete transaction.</returns>
    Public Async Function ExecuteInTransactionAsync(Operation As Func(Of MySqlTransactionContext, CancellationToken, Task), Optional IsolationLevel As IsolationLevel = IsolationLevel.ReadCommitted, Optional CancellationToken As CancellationToken = Nothing) As Task
        Try
            If Operation Is Nothing Then Throw New ArgumentNullException("Operation")
            Using Connection As MySqlConnection = Client.CreateDatabaseProviderConnection()
                Await Connection.OpenAsync(CancellationToken).ConfigureAwait(False)
                Using Transaction As MySqlTransaction = Connection.BeginTransaction(IsolationLevel)
                    Try
                        Dim OperationTask As Task = Operation(New MySqlTransactionContext(Transaction), CancellationToken)
                        If OperationTask Is Nothing Then Throw New InvalidOperationException("The transaction operation returned Nothing instead of a Task.")
                        Await OperationTask.ConfigureAwait(False)
                        Await Transaction.CommitAsync(CancellationToken).ConfigureAwait(False)
                    Catch
                        TryRollback(Transaction)
                        Throw
                    End Try
                End Using
            End Using
        Catch ex As MySqlException
            Throw New MySqlServiceException(ex)
        End Try
    End Function
    Private Shared Function BuildConnectionString(Server As String, Database As String, User As String, Password As String, Configure As Action(Of MySqlConnectionOptions)) As String
        Dim Builder As New MySqlConnectionOptions With {.Server = MySqlSql.RequireValue(Server, NameOf(Server)), .Database = MySqlSql.RequireValue(Database, NameOf(Database)), .UserID = MySqlSql.RequireValue(User, NameOf(User)), .Password = If(Password, String.Empty), .Pooling = True}
        Configure?.Invoke(Builder)
        If String.IsNullOrWhiteSpace(Builder.Server) Then Throw New ArgumentException("The configured connection string must define a server.", NameOf(Server))
        If String.IsNullOrWhiteSpace(Builder.Database) Then Throw New ArgumentException("The configured connection string must define a database.", NameOf(Database))
        Return Builder.BuildConnectionString()
    End Function
    Private Shared Sub TryRollback(Transaction As MySqlTransaction)
        Try
            Transaction.Rollback()
        Catch ex As MySqlException
        Catch ex As InvalidOperationException
        End Try
    End Sub
End Class
