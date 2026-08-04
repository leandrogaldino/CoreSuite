Imports System.Data
Imports System.Threading
Imports System.Threading.Tasks
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
    ''' <param name="Configure">An optional callback used to configure additional <see cref="MySqlConnectionStringBuilder"/> properties.</param>
    Public Sub New(Server As String, Database As String, User As String, Password As String, Optional Configure As Action(Of MySqlConnectionStringBuilder) = Nothing)
        Me.New(BuildConnectionString(Server, Database, User, Password, Configure))
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="MySqlService"/> class from a complete connection string.
    ''' </summary>
    ''' <param name="ConnectionString">A MySQL connection string containing a server and database.</param>
    Public Sub New(ConnectionString As String)
        Client = New MySqlClient(ConnectionString)
        Request = New MySqlRequest(Client)
        Maintenance = New MySqlMaintenance(Client)
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
    Public Function ExecuteInTransaction(Of TResult)(Operation As Func(Of MySqlTransaction, TResult), Optional IsolationLevel As IsolationLevel = IsolationLevel.ReadCommitted) As TResult
        ArgumentNullException.ThrowIfNull(Operation)
        Using Connection As MySqlConnection = Client.CreateDatabaseConnection()
            Connection.Open()
            Using Transaction As MySqlTransaction = Connection.BeginTransaction(IsolationLevel)
                Try
                    Dim Result As TResult = Operation(Transaction)
                    Transaction.Commit()
                    Return Result
                Catch
                    TryRollback(Transaction)
                    Throw
                End Try
            End Using
        End Using
    End Function
    ''' <summary>
    ''' Executes an action inside a new local transaction and commits it when the action succeeds.
    ''' </summary>
    ''' <param name="Operation">The action that receives the active transaction.</param>
    ''' <param name="IsolationLevel">The transaction isolation level.</param>
    Public Sub ExecuteInTransaction(Operation As Action(Of MySqlTransaction), Optional IsolationLevel As IsolationLevel = IsolationLevel.ReadCommitted)
        ArgumentNullException.ThrowIfNull(Operation)
        Using Connection As MySqlConnection = Client.CreateDatabaseConnection()
            Connection.Open()
            Using Transaction As MySqlTransaction = Connection.BeginTransaction(IsolationLevel)
                Try
                    Operation(Transaction)
                    Transaction.Commit()
                Catch
                    TryRollback(Transaction)
                    Throw
                End Try
            End Using
        End Using
    End Sub
    ''' <summary>
    ''' Asynchronously executes an operation inside a new local transaction and commits it when the operation succeeds.
    ''' </summary>
    ''' <typeparam name="TResult">The operation result type.</typeparam>
    ''' <param name="Operation">The asynchronous operation that receives the active transaction and cancellation token.</param>
    ''' <param name="IsolationLevel">The transaction isolation level.</param>
    ''' <param name="CancellationToken">The token used to cancel connection opening, the operation, or commit.</param>
    ''' <returns>The value returned by <paramref name="Operation"/>.</returns>
    Public Async Function ExecuteInTransactionAsync(Of TResult)(Operation As Func(Of MySqlTransaction, CancellationToken, Task(Of TResult)), Optional IsolationLevel As IsolationLevel = IsolationLevel.ReadCommitted, Optional CancellationToken As CancellationToken = Nothing) As Task(Of TResult)
        ArgumentNullException.ThrowIfNull(Operation)
        Using Connection As MySqlConnection = Client.CreateDatabaseConnection()
            Await Connection.OpenAsync(CancellationToken).ConfigureAwait(False)
            Using Transaction As MySqlTransaction = Connection.BeginTransaction(IsolationLevel)
                Try
                    Dim OperationTask As Task(Of TResult) = Operation(Transaction, CancellationToken)
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
    End Function
    ''' <summary>
    ''' Asynchronously executes an action inside a new local transaction and commits it when the action succeeds.
    ''' </summary>
    ''' <param name="Operation">The asynchronous action that receives the active transaction and cancellation token.</param>
    ''' <param name="IsolationLevel">The transaction isolation level.</param>
    ''' <param name="CancellationToken">The token used to cancel connection opening, the action, or commit.</param>
    ''' <returns>A task representing the complete transaction.</returns>
    Public Async Function ExecuteInTransactionAsync(Operation As Func(Of MySqlTransaction, CancellationToken, Task), Optional IsolationLevel As IsolationLevel = IsolationLevel.ReadCommitted, Optional CancellationToken As CancellationToken = Nothing) As Task
        ArgumentNullException.ThrowIfNull(Operation)
        Using Connection As MySqlConnection = Client.CreateDatabaseConnection()
            Await Connection.OpenAsync(CancellationToken).ConfigureAwait(False)
            Using Transaction As MySqlTransaction = Connection.BeginTransaction(IsolationLevel)
                Try
                    Dim OperationTask As Task = Operation(Transaction, CancellationToken)
                    If OperationTask Is Nothing Then Throw New InvalidOperationException("The transaction operation returned Nothing instead of a Task.")
                    Await OperationTask.ConfigureAwait(False)
                    Await Transaction.CommitAsync(CancellationToken).ConfigureAwait(False)
                Catch
                    TryRollback(Transaction)
                    Throw
                End Try
            End Using
        End Using
    End Function
    Private Shared Function BuildConnectionString(Server As String, Database As String, User As String, Password As String, Configure As Action(Of MySqlConnectionStringBuilder)) As String
        Dim Builder As New MySqlConnectionStringBuilder With {.Server = MySqlSql.RequireValue(Server, NameOf(Server)), .Database = MySqlSql.RequireValue(Database, NameOf(Database)), .UserID = MySqlSql.RequireValue(User, NameOf(User)), .Password = If(Password, String.Empty), .Pooling = True}
        Configure?.Invoke(Builder)
        If String.IsNullOrWhiteSpace(Builder.Server) Then Throw New ArgumentException("The configured connection string must define a server.", NameOf(Server))
        If String.IsNullOrWhiteSpace(Builder.Database) Then Throw New ArgumentException("The configured connection string must define a database.", NameOf(Database))
        Return Builder.ConnectionString
    End Function
    Private Shared Sub TryRollback(Transaction As MySqlTransaction)
        Try
            Transaction.Rollback()
        Catch ex As MySqlException
        Catch ex As InvalidOperationException
        End Try
    End Sub
End Class
