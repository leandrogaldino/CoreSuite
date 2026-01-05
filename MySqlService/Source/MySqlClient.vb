Imports System.Data.Common
Imports System.Transactions
Imports MySql.Data.MySqlClient

''' <summary>
''' Class responsible for managing connection information and creating connection objects
''' for MySQL servers and databases using MySqlConnector.
''' </summary>
Public Class MySqlClient
    Friend Property Server As String
    Friend Property Database As String
    Friend Property User As String
    Friend Property Password As String
    Friend Sub New(Server As String, Database As String, User As String, Password As String)
        Me.Server = Server
        Me.Database = Database
        Me.User = User
        Me.Password = Password
    End Sub

    Friend Function GetServerConnectionString() As String
        Return String.Format("Server={0};Uid={1};Pwd={2};Pooling=True", Server, User, Password)
    End Function
    Friend Function GetDatabaseConnectionString() As String
        Return String.Format("Server={0};Database={1};Uid={2};Pwd={3};Pooling=True", Server, Database, User, Password)
    End Function
    ''' <summary>
    ''' Creates a new connection to the MySQL server without associating it with a specific database.
    ''' </summary>
    ''' <remarks>
    ''' <para>
    ''' The returned connection is created using only server information and credentials,
    ''' and is intended for global administrative operations, such as:
    ''' </para>
    ''' <list type="bullet">
    ''' <item><description>Creating or dropping databases</description></item>
    ''' <item><description>Listing schemas</description></item>
    ''' <item><description>Operations that do not depend on a specific database</description></item>
    ''' </list>
    '''
    ''' <para>
    ''' This method only instantiates the connection; the responsibility to open, close,
    ''' and dispose (<c>Dispose</c>) the <see cref="DbConnection"/> lies entirely with the consumer.
    ''' </para>
    '''
    ''' <para>
    ''' This connection can be used within a <see cref="TransactionScope"/>,
    ''' provided that the scope is created before opening the connection.
    ''' </para>
    ''' </remarks>
    ''' <returns>
    ''' A new instance of <see cref="DbConnection"/> configured to connect
    ''' to the MySQL server.
    ''' </returns>
    Public Function CreateServerConnection() As DbConnection
        Dim Connection As New MySqlConnection(GetServerConnectionString())
        Return Connection
    End Function
    ''' <summary>
    ''' Creates a new connection to the MySQL database configured as the default.
    ''' </summary>
    ''' <remarks>
    ''' <para>
    ''' The returned connection is created using the full connection string,
    ''' including the database defined in the client configuration.
    ''' </para>
    '''
    ''' <para>
    ''' This method is the official entry point for creating connections that will be used
    ''' by <see cref="MySqlRequest"/> and <see cref="MySqlMaintenance"/>, especially
    ''' in scenarios involving transactions.
    ''' </para>
    '''
    ''' <para>
    ''' This method only instantiates the connection; the responsibility to open, close,
    ''' and dispose (<c>Dispose</c>) the <see cref="DbConnection"/> lies entirely with the consumer.
    ''' </para>
    '''
    ''' <para>
    ''' When using asynchronous methods within a <see cref="TransactionScope"/>,
    ''' the scope must be created with <see cref="TransactionScopeAsyncFlowOption.Enabled"/>.
    ''' </para>
    ''' </remarks>
    ''' <returns>
    ''' A new instance of <see cref="DbConnection"/> configured to access
    ''' the default MySQL database.
    ''' </returns>
    Public Function CreateDatabaseConnection() As DbConnection
        Dim Connection As New MySqlConnection(GetDatabaseConnectionString())
        Return Connection
    End Function
End Class