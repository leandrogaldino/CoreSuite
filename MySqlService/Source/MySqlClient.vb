Imports MySql.Data.MySqlClient
''' <summary>
''' Stores immutable MySQL connection settings and creates provider-specific connections.
''' </summary>
Public NotInheritable Class MySqlClient
    Private ReadOnly _DatabaseConnectionString As String
    Private ReadOnly _ServerConnectionString As String
    ''' <summary>
    ''' Initializes a new instance of the <see cref="MySqlClient"/> class from a complete connection string.
    ''' </summary>
    ''' <param name="ConnectionString">A MySQL connection string containing a server and database.</param>
    Public Sub New(ConnectionString As String)
        Dim NormalizedConnectionString As String = RequireValue(ConnectionString, NameOf(ConnectionString))
        Dim DatabaseBuilder As New MySqlConnectionStringBuilder(NormalizedConnectionString)
        If String.IsNullOrWhiteSpace(DatabaseBuilder.Server) Then Throw New ArgumentException("The connection string must define a server.", NameOf(ConnectionString))
        If String.IsNullOrWhiteSpace(DatabaseBuilder.Database) Then Throw New ArgumentException("The connection string must define a database.", NameOf(ConnectionString))
        Server = DatabaseBuilder.Server
        Database = DatabaseBuilder.Database
        _DatabaseConnectionString = DatabaseBuilder.ConnectionString
        Dim serverBuilder As New MySqlConnectionStringBuilder(DatabaseBuilder.ConnectionString) With {.Database = String.Empty}
        _ServerConnectionString = serverBuilder.ConnectionString
    End Sub
    ''' <summary>
    ''' Gets the configured server address.
    ''' </summary>
    Public ReadOnly Property Server As String
    ''' <summary>
    ''' Gets the configured default database name.
    ''' </summary>
    Public ReadOnly Property Database As String
    ''' <summary>
    ''' Creates a closed connection targeting the configured database. The caller owns the returned connection.
    ''' </summary>
    ''' <returns>A new <see cref="MySqlConnection"/> instance.</returns>
    Public Function CreateDatabaseConnection() As MySqlConnection
        Return New MySqlConnection(_DatabaseConnectionString)
    End Function
    ''' <summary>
    ''' Creates a closed server-level connection without selecting a default database. The caller owns the returned connection.
    ''' </summary>
    ''' <returns>A new <see cref="MySqlConnection"/> instance.</returns>
    Public Function CreateServerConnection() As MySqlConnection
        Return New MySqlConnection(_ServerConnectionString)
    End Function
End Class
