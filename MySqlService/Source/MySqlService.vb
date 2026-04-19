Imports System.Data.Common
Imports System.Transactions

''' <summary>
''' Main class of the library, responsible for centralizing and exposing
''' all functionalities related to MySQL access and maintenance.
''' </summary>
''' <remarks>
''' <para>
''' This library <b>does not explicitly implement <c>MySqlTransaction</c></b>.
''' Transaction control is handled using <see cref="TransactionScope"/>.
''' </para>
'''
''' <para>
''' The methods of <see cref="MySqlRequest"/> and <see cref="MySqlMaintenance"/>
''' manage connections in an intelligent way:
''' </para>
''' <list type="bullet">
''' <item>
''' <description>
''' When no connection is provided, the method automatically creates,
''' uses, and disposes an internal connection.
''' </description>
''' </item>
''' <item>
''' <description>
''' When a connection is provided, the method only uses it and is not
''' responsible for its lifecycle.
''' </description>
''' </item>
''' </list>
'''
''' <para>
''' To use transactions, the consumer of the library must:
''' </para>
''' <list type="number">
''' <item>
''' <description>
''' Explicitly create the connection using the
''' <see cref="MySqlClient.CreateDatabaseConnection"/> method.
''' </description>
''' </item>
''' <item>
''' <description>
''' Start a <see cref="TransactionScope"/>.
''' For asynchronous methods, use
''' <see cref="TransactionScopeAsyncFlowOption.Enabled"/>.
''' </description>
''' </item>
''' <item>
''' <description>
''' Pass the same <see cref="DbConnection"/> instance to all
''' <see cref="MySqlRequest"/> and <see cref="MySqlMaintenance"/> methods
''' that are part of the transaction.
''' </description>
''' </item>
''' <item>
''' <description>
''' Call <c>Complete()</c> on the scope and, at the end,
''' manually dispose the connection.
''' </description>
''' </item>
''' </list>
'''
''' <para>
''' This model ensures flexibility, allowing simple usage with automatic
''' connection management or explicit transactional control when needed.
''' </para>
''' </remarks>
Public Class MySqlService
    Private _Client As MySqlClient
    Private _Request As MySqlRequest
    Private _Maintenance As MySqlMaintenance

    ''' <summary>
    ''' Instance of <see cref="MySqlClient"/>, responsible for creating
    ''' connections to the MySQL server through the
    ''' <see cref="MySqlClient.CreateDatabaseConnection"/> method.
    ''' </summary>
    Public ReadOnly Property Client As MySqlClient
        Get
            Return _Client
        End Get
    End Property

    ''' <summary>
    ''' Instance of <see cref="MySqlRequest"/>, responsible for executing
    ''' queries, CRUD operations, and stored procedures.
    ''' </summary>
    Public ReadOnly Property Request As MySqlRequest
        Get
            Return _Request
        End Get
    End Property

    ''' <summary>
    ''' Instance of <see cref="MySqlMaintenance"/>, responsible for
    ''' administrative operations such as backup, restore,
    ''' schema creation, and database maintenance.
    ''' </summary>
    Public ReadOnly Property Maintenance As MySqlMaintenance
        Get
            Return _Maintenance
        End Get
    End Property

    ''' <summary>
    ''' Creates a new instance of <see cref="MySqlService"/> without initialization.
    ''' Calling <c>Initialize</c> is required before using any service.
    ''' </summary>
    Public Sub New()
    End Sub

    ''' <summary>
    ''' Initializes a new instance of <see cref="MySqlService"/> with
    ''' connection credentials and configuration.
    ''' </summary>
    ''' <param name="Server">MySQL server address or hostname.</param>
    ''' <param name="Database">Default database name.</param>
    ''' <param name="User">Database access username.</param>
    ''' <param name="Password">Database user password.</param>
    Public Sub New(Server As String, Database As String, User As String, Password As String)
        _Client = New MySqlClient(Server, Database, User, Password)
        _Request = New MySqlRequest(Client)
        _Maintenance = New MySqlMaintenance(Client)
    End Sub

    ''' <summary>
    ''' Initializes the <see cref="MySqlService"/> instance with
    ''' connection credentials and configuration.
    ''' </summary>
    ''' <param name="Server">MySQL server address or hostname.</param>
    ''' <param name="Database">Default database name.</param>
    Public Sub Initialize(Server As String, Database As String, User As String, Password As String)
        _Client = New MySqlClient(Server, Database, User, Password)
        _Request = New MySqlRequest(Client)
        _Maintenance = New MySqlMaintenance(Client)
    End Sub
End Class