Imports System.Collections.ObjectModel
Imports System.Data
Imports System.Text
Imports System.Threading
Imports MySql.Data.MySqlClient
''' <summary>
''' Executes parameterized queries, commands, scalar operations, CRUD operations, and stored procedures against the configured MySQL database.
''' </summary>
Public NotInheritable Class MySqlRequest
    Private ReadOnly _Client As MySqlClient
    Friend Sub New(Client As MySqlClient)
        ArgumentNullException.ThrowIfNull(Client)
        _Client = Client
    End Sub
    ''' <summary>
    ''' Executes SQL expected to return one or more result sets.
    ''' </summary>
    ''' <param name="Sql">The trusted SQL command text.</param>
    ''' <param name="queryArgs">Optional parameter values referenced by <paramref name="Sql"/>.</param>
    ''' <param name="options">Optional connection, transaction, and timeout settings.</param>
    ''' <returns>A response containing every returned result set.</returns>
    Public Function ExecuteQuery(Sql As String, Optional QueryArgs As IDictionary(Of String, Object) = Nothing, Optional Options As MySqlCommandOptions = Nothing) As MySqlResponse
        Dim ActualOptions As MySqlCommandOptions = ResolveOptions(Options)
        ValidateSql(Sql, NameOf(Sql))
        Using ConnectionScope As MySqlConnectionScope = MySqlConnectionScope.Create(_Client, ActualOptions.Connection, ActualOptions.Transaction)
            ConnectionScope.Open()
            Using Command As MySqlCommand = CreateCommand(Sql, CommandType.Text, ConnectionScope.Connection, ActualOptions)
                AddParameters(Command, QueryArgs)
                Using Reader As MySqlDataReader = DirectCast(Command.ExecuteReader(), MySqlDataReader)
                    Dim ResultSets As IReadOnlyList(Of MySqlResultSet) = ReadResultSets(Reader)
                    Dim AffectedRows As Long = NormalizeAffectedRows(Reader.RecordsAffected)
                    Return New MySqlResponse(ResultSets, AffectedRows)
                End Using
            End Using
        End Using
    End Function
    ''' <summary>
    ''' Asynchronously executes SQL expected to return one or more result sets.
    ''' </summary>
    ''' <param name="Sql">The trusted SQL command text.</param>
    ''' <param name="QueryArgs">Optional parameter values referenced by <paramref name="Sql"/>.</param>
    ''' <param name="Options">Optional connection, transaction, and timeout settings.</param>
    ''' <param name="CancellationToken">The token used to cancel connection opening, execution, or result reading.</param>
    ''' <returns>A response containing every returned result set.</returns>
    Public Async Function ExecuteQueryAsync(Sql As String, Optional QueryArgs As IDictionary(Of String, Object) = Nothing, Optional Options As MySqlCommandOptions = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task(Of MySqlResponse)
        Dim ActualOptions As MySqlCommandOptions = ResolveOptions(Options)
        ValidateSql(Sql, NameOf(Sql))
        Using ConnectionScope As MySqlConnectionScope = MySqlConnectionScope.Create(_Client, ActualOptions.Connection, ActualOptions.Transaction)
            Await ConnectionScope.OpenAsync(CancellationToken).ConfigureAwait(False)
            Using Command As MySqlCommand = CreateCommand(Sql, CommandType.Text, ConnectionScope.Connection, ActualOptions)
                AddParameters(Command, QueryArgs)
                Using Reader As MySqlDataReader = DirectCast(Await Command.ExecuteReaderAsync(CancellationToken).ConfigureAwait(False), MySqlDataReader)
                    Dim ResultSets As IReadOnlyList(Of MySqlResultSet) = Await ReadResultSetsAsync(Reader, CancellationToken).ConfigureAwait(False)
                    Dim AffectedRows As Long = NormalizeAffectedRows(Reader.RecordsAffected)
                    Return New MySqlResponse(ResultSets, AffectedRows)
                End Using
            End Using
        End Using
    End Function
    ''' <summary>
    ''' Executes SQL that does not return rows and returns the affected-row count.
    ''' </summary>
    ''' <param name="Sql">The trusted SQL command text.</param>
    ''' <param name="QueryArgs">Optional parameter values referenced by <paramref name="Sql"/>.</param>
    ''' <param name="Options">Optional connection, transaction, and timeout settings.</param>
    ''' <returns>A response containing the affected-row count.</returns>
    Public Function ExecuteNonQuery(Sql As String, Optional QueryArgs As IDictionary(Of String, Object) = Nothing, Optional Options As MySqlCommandOptions = Nothing) As MySqlResponse
        Dim ActualOptions As MySqlCommandOptions = ResolveOptions(Options)
        ValidateSql(Sql, NameOf(Sql))
        Using ConnectionScope As MySqlConnectionScope = MySqlConnectionScope.Create(_Client, ActualOptions.Connection, ActualOptions.Transaction)
            ConnectionScope.Open()
            Using Command As MySqlCommand = CreateCommand(Sql, CommandType.Text, ConnectionScope.Connection, ActualOptions)
                AddParameters(Command, QueryArgs)
                Return New MySqlResponse(AffectedRows:=Command.ExecuteNonQuery())
            End Using
        End Using
    End Function
    ''' <summary>
    ''' Asynchronously executes SQL that does not return rows and returns the affected-row count.
    ''' </summary>
    ''' <param name="Sql">The trusted SQL command text.</param>
    ''' <param name="QueryArgs">Optional parameter values referenced by <paramref name="Sql"/>.</param>
    ''' <param name="Options">Optional connection, transaction, and timeout settings.</param>
    ''' <param name="CancellationToken">The token used to cancel connection opening or execution.</param>
    ''' <returns>A response containing the affected-row count.</returns>
    Public Async Function ExecuteNonQueryAsync(Sql As String, Optional QueryArgs As IDictionary(Of String, Object) = Nothing, Optional Options As MySqlCommandOptions = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task(Of MySqlResponse)
        Dim ActualOptions As MySqlCommandOptions = ResolveOptions(Options)
        ValidateSql(Sql, NameOf(Sql))
        Using ConnectionScope As MySqlConnectionScope = MySqlConnectionScope.Create(_Client, ActualOptions.Connection, ActualOptions.Transaction)
            Await ConnectionScope.OpenAsync(CancellationToken).ConfigureAwait(False)
            Using Command As MySqlCommand = CreateCommand(Sql, CommandType.Text, ConnectionScope.Connection, ActualOptions)
                AddParameters(Command, QueryArgs)
                Dim AffectedRows As Integer = Await Command.ExecuteNonQueryAsync(CancellationToken).ConfigureAwait(False)
                Return New MySqlResponse(AffectedRows:=AffectedRows)
            End Using
        End Using
    End Function
    ''' <summary>
    ''' Executes SQL and returns the first column of the first row, or <see langword="Nothing"/> when no value is returned.
    ''' </summary>
    ''' <param name="Sql">The trusted SQL command text.</param>
    ''' <param name="QueryArgs">Optional parameter values referenced by <paramref name="Sql"/>.</param>
    ''' <param name="Options">Optional connection, transaction, and timeout settings.</param>
    ''' <returns>The first value returned by the command, or <see langword="Nothing"/>.</returns>
    Public Function ExecuteScalar(Sql As String, Optional QueryArgs As IDictionary(Of String, Object) = Nothing, Optional Options As MySqlCommandOptions = Nothing) As Object
        Dim ActualOptions As MySqlCommandOptions = ResolveOptions(Options)
        ValidateSql(Sql, NameOf(Sql))
        Using ConnectionScope As MySqlConnectionScope = MySqlConnectionScope.Create(_Client, ActualOptions.Connection, ActualOptions.Transaction)
            ConnectionScope.Open()
            Using Command As MySqlCommand = CreateCommand(Sql, CommandType.Text, ConnectionScope.Connection, ActualOptions)
                AddParameters(Command, QueryArgs)
                Return ConvertDatabaseValue(Command.ExecuteScalar())
            End Using
        End Using
    End Function
    ''' <summary>
    ''' Asynchronously executes SQL and returns the first column of the first row, or <see langword="Nothing"/> when no value is returned.
    ''' </summary>
    ''' <param name="Sql">The trusted SQL command text.</param>
    ''' <param name="QueryArgs">Optional parameter values referenced by <paramref name="sql"/>.</param>
    ''' <param name="Options">Optional connection, transaction, and timeout settings.</param>
    ''' <param name="CancellationToken">The token used to cancel connection opening or execution.</param>
    ''' <returns>The first value returned by the command, or <see langword="Nothing"/>.</returns>
    Public Async Function ExecuteScalarAsync(Sql As String, Optional QueryArgs As IDictionary(Of String, Object) = Nothing, Optional Options As MySqlCommandOptions = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task(Of Object)
        Dim ActualOptions As MySqlCommandOptions = ResolveOptions(Options)
        ValidateSql(Sql, NameOf(Sql))
        Using ConnectionScope As MySqlConnectionScope = MySqlConnectionScope.Create(_Client, ActualOptions.Connection, ActualOptions.Transaction)
            Await ConnectionScope.OpenAsync(CancellationToken).ConfigureAwait(False)
            Using Command As MySqlCommand = CreateCommand(Sql, CommandType.Text, ConnectionScope.Connection, ActualOptions)
                AddParameters(Command, QueryArgs)
                Return ConvertDatabaseValue(Await Command.ExecuteScalarAsync(CancellationToken).ConfigureAwait(False))
            End Using
        End Using
    End Function
    ''' <summary>
    ''' Executes a stored procedure, reads every result set, and captures output and return-value parameters.
    ''' </summary>
    ''' <param name="ProcedureName">The procedure name, optionally qualified by a database.</param>
    ''' <param name="Parameters">Optional input, output, input/output, and return-value definitions.</param>
    ''' <param name="Options">Optional connection, transaction, and timeout settings.</param>
    ''' <returns>A response containing result sets, affected rows, and output values.</returns>
    Public Function ExecuteProcedure(ProcedureName As String, Optional Parameters As IEnumerable(Of MySqlProcedureParameter) = Nothing, Optional Options As MySqlCommandOptions = Nothing) As MySqlResponse
        Dim ActualOptions As MySqlCommandOptions = ResolveOptions(Options)
        Dim NormalizedProcedureName As String = ValidateRoutineName(ProcedureName, NameOf(ProcedureName))
        Using ConnectionScope As MySqlConnectionScope = MySqlConnectionScope.Create(_Client, ActualOptions.Connection, ActualOptions.Transaction)
            ConnectionScope.Open()
            Using Command As MySqlCommand = CreateCommand(NormalizedProcedureName, CommandType.StoredProcedure, ConnectionScope.Connection, ActualOptions)
                AddProcedureParameters(Command, Parameters)
                Dim ResultSets As IReadOnlyList(Of MySqlResultSet)
                Dim AffectedRows As Long
                Using Reader As MySqlDataReader = DirectCast(Command.ExecuteReader(), MySqlDataReader)
                    ResultSets = ReadResultSets(Reader)
                    AffectedRows = NormalizeAffectedRows(Reader.RecordsAffected)
                End Using
                Return CreateProcedureResponse(Command, ResultSets, AffectedRows)
            End Using
        End Using
    End Function
    ''' <summary>
    ''' Asynchronously executes a stored procedure, reads every result set, and captures output and return-value parameters.
    ''' </summary>
    ''' <param name="ProcedureName">The procedure name, optionally qualified by a database.</param>
    ''' <param name="Parameters">Optional input, output, input/output, and return-value definitions.</param>
    ''' <param name="Options">Optional connection, transaction, and timeout settings.</param>
    ''' <param name="CancellationToken">The token used to cancel connection opening, execution, or result reading.</param>
    ''' <returns>A response containing result sets, affected rows, and output values.</returns>
    Public Async Function ExecuteProcedureAsync(ProcedureName As String, Optional Parameters As IEnumerable(Of MySqlProcedureParameter) = Nothing, Optional Options As MySqlCommandOptions = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task(Of MySqlResponse)
        Dim ActualOptions As MySqlCommandOptions = ResolveOptions(Options)
        Dim NormalizedProcedureName As String = ValidateRoutineName(ProcedureName, NameOf(ProcedureName))
        Using ConnectionScope As MySqlConnectionScope = MySqlConnectionScope.Create(_Client, ActualOptions.Connection, ActualOptions.Transaction)
            Await ConnectionScope.OpenAsync(CancellationToken).ConfigureAwait(False)
            Using Command As MySqlCommand = CreateCommand(NormalizedProcedureName, CommandType.StoredProcedure, ConnectionScope.Connection, ActualOptions)
                AddProcedureParameters(Command, Parameters)
                Dim ResultSets As IReadOnlyList(Of MySqlResultSet)
                Dim AffectedRows As Long
                Using Reader As MySqlDataReader = DirectCast(Await Command.ExecuteReaderAsync(CancellationToken).ConfigureAwait(False), MySqlDataReader)
                    ResultSets = Await ReadResultSetsAsync(Reader, CancellationToken).ConfigureAwait(False)
                    AffectedRows = NormalizeAffectedRows(Reader.RecordsAffected)
                End Using
                Return CreateProcedureResponse(Command, ResultSets, AffectedRows)
            End Using
        End Using
    End Function
    ''' <summary>
    ''' Selects rows from a safely quoted table using structured projection, ordering, and paging options.
    ''' </summary>
    ''' <param name="Table">The table name, optionally qualified by a schema or database.</param>
    ''' <param name="Options">Optional projection, filter, ordering, paging, connection, transaction, and timeout settings.</param>
    ''' <returns>A response containing the selected rows.</returns>
    Public Function ExecuteSelect(Table As String, Optional Options As MySqlSelectOptions = Nothing) As MySqlResponse
        Dim ActualOptions As MySqlSelectOptions = If(Options, New MySqlSelectOptions())
        Dim Sql As String = BuildSelectSql(Table, ActualOptions)
        Return ExecuteQuery(Sql, ActualOptions.QueryArgs, ActualOptions)
    End Function
    ''' <summary>
    ''' Asynchronously selects rows from a safely quoted table using structured projection, ordering, and paging options.
    ''' </summary>
    ''' <param name="Table">The table name, optionally qualified by a schema or database.</param>
    ''' <param name="Options">Optional projection, filter, ordering, paging, connection, transaction, and timeout settings.</param>
    ''' <param name="CancellationToken">The token used to cancel connection opening, execution, or result reading.</param>
    ''' <returns>A response containing the selected rows.</returns>
    Public Function ExecuteSelectAsync(Table As String, Optional Options As MySqlSelectOptions = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task(Of MySqlResponse)
        Dim ActualOptions As MySqlSelectOptions = If(Options, New MySqlSelectOptions())
        Dim Sql As String = BuildSelectSql(Table, ActualOptions)
        Return ExecuteQueryAsync(Sql, ActualOptions.QueryArgs, ActualOptions, CancellationToken)
    End Function
    ''' <summary>
    ''' Inserts one row into a safely quoted table.
    ''' </summary>
    ''' <param name="Table">The table name, optionally qualified by a schema or database.</param>
    ''' <param name="Values">The column names and values to insert.</param>
    ''' <param name="Options">Optional connection, transaction, and timeout settings.</param>
    ''' <returns>A response containing affected rows and the generated identifier when available.</returns>
    Public Function ExecuteInsert(Table As String, Values As IDictionary(Of String, Object), Optional Options As MySqlCommandOptions = Nothing) As MySqlResponse
        ValidateValues(Values)
        Dim ActualOptions As MySqlCommandOptions = ResolveOptions(Options)
        Dim CommandData As CommandBuildResult = BuildInsertSql(Table, Values)
        Using ConnectionScope As MySqlConnectionScope = MySqlConnectionScope.Create(_Client, ActualOptions.Connection, ActualOptions.Transaction)
            ConnectionScope.Open()
            Using Command As MySqlCommand = CreateCommand(CommandData.Sql, CommandType.Text, ConnectionScope.Connection, ActualOptions)
                AddParameters(Command, CommandData.Parameters)
                Dim AffectedRows As Integer = Command.ExecuteNonQuery()
                Return New MySqlResponse(AffectedRows:=AffectedRows, LastInsertedId:=NormalizeLastInsertedId(Command.LastInsertedId))
            End Using
        End Using
    End Function
    ''' <summary>
    ''' Asynchronously inserts one row into a safely quoted table.
    ''' </summary>
    ''' <param name="Table">The table name, optionally qualified by a schema or database.</param>
    ''' <param name="Values">The column names and values to insert.</param>
    ''' <param name="Options">Optional connection, transaction, and timeout settings.</param>
    ''' <param name="CancellationToken">The token used to cancel connection opening or execution.</param>
    ''' <returns>A response containing affected rows and the generated identifier when available.</returns>
    Public Async Function ExecuteInsertAsync(Table As String, Values As IDictionary(Of String, Object), Optional Options As MySqlCommandOptions = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task(Of MySqlResponse)
        ValidateValues(Values)
        Dim ActualOptions As MySqlCommandOptions = ResolveOptions(Options)
        Dim CommandData As CommandBuildResult = BuildInsertSql(Table, Values)
        Using ConnectionScope As MySqlConnectionScope = MySqlConnectionScope.Create(_Client, ActualOptions.Connection, ActualOptions.Transaction)
            Await ConnectionScope.OpenAsync(CancellationToken).ConfigureAwait(False)
            Using Command As MySqlCommand = CreateCommand(CommandData.Sql, CommandType.Text, ConnectionScope.Connection, ActualOptions)
                AddParameters(Command, CommandData.Parameters)
                Dim AffectedRows As Integer = Await Command.ExecuteNonQueryAsync(CancellationToken).ConfigureAwait(False)
                Return New MySqlResponse(AffectedRows:=AffectedRows, LastInsertedId:=NormalizeLastInsertedId(Command.LastInsertedId))
            End Using
        End Using
    End Function
    ''' <summary>
    ''' Updates rows in a safely quoted table. A WHERE clause is required unless <see cref="MySqlMutationOptions.AllowAllRows"/> is explicitly enabled.
    ''' </summary>
    ''' <param name="Table">The table name, optionally qualified by a schema or database.</param>
    ''' <param name="Values">The column names and values to update.</param>
    ''' <param name="Options">Optional filter, safety, connection, transaction, and timeout settings.</param>
    ''' <returns>A response containing the affected-row count.</returns>
    Public Function ExecuteUpdate(Table As String, Values As IDictionary(Of String, Object), Optional Options As MySqlMutationOptions = Nothing) As MySqlResponse
        ValidateValues(Values)
        Dim ActualOptions As MySqlMutationOptions = If(Options, New MySqlMutationOptions())
        ValidateMutationOptions(ActualOptions)
        Dim CommandData As CommandBuildResult = BuildUpdateSql(Table, Values, ActualOptions)
        Return ExecuteNonQuery(CommandData.Sql, CommandData.Parameters, ActualOptions)
    End Function
    ''' <summary>
    ''' Asynchronously updates rows in a safely quoted table. A WHERE clause is required unless <see cref="MySqlMutationOptions.AllowAllRows"/> is explicitly enabled.
    ''' </summary>
    ''' <param name="Table">The table name, optionally qualified by a schema or database.</param>
    ''' <param name="Values">The column names and values to update.</param>
    ''' <param name="Options">Optional filter, safety, connection, transaction, and timeout settings.</param>
    ''' <param name="CancellationToken">The token used to cancel connection opening or execution.</param>
    ''' <returns>A response containing the affected-row count.</returns>
    Public Function ExecuteUpdateAsync(Table As String, Values As IDictionary(Of String, Object), Optional Options As MySqlMutationOptions = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task(Of MySqlResponse)
        ValidateValues(Values)
        Dim ActualOptions As MySqlMutationOptions = If(Options, New MySqlMutationOptions())
        ValidateMutationOptions(ActualOptions)
        Dim CommandData As CommandBuildResult = BuildUpdateSql(Table, Values, ActualOptions)
        Return ExecuteNonQueryAsync(CommandData.Sql, CommandData.Parameters, ActualOptions, CancellationToken)
    End Function
    ''' <summary>
    ''' Deletes rows from a safely quoted table. A WHERE clause is required unless <see cref="MySqlMutationOptions.AllowAllRows"/> is explicitly enabled.
    ''' </summary>
    ''' <param name="Table">The table name, optionally qualified by a schema or database.</param>
    ''' <param name="Options">Optional filter, safety, connection, transaction, and timeout settings.</param>
    ''' <returns>A response containing the affected-row count.</returns>
    Public Function ExecuteDelete(Table As String, Optional Options As MySqlMutationOptions = Nothing) As MySqlResponse
        Dim ActualOptions As MySqlMutationOptions = If(Options, New MySqlMutationOptions())
        ValidateMutationOptions(ActualOptions)
        Dim Sql As String = BuildDeleteSql(Table, ActualOptions)
        Return ExecuteNonQuery(Sql, ActualOptions.QueryArgs, ActualOptions)
    End Function
    ''' <summary>
    ''' Asynchronously deletes rows from a safely quoted table. A WHERE clause is required unless <see cref="MySqlMutationOptions.AllowAllRows"/> is explicitly enabled.
    ''' </summary>
    ''' <param name="Table">The table name, optionally qualified by a schema or database.</param>
    ''' <param name="Options">Optional filter, safety, connection, transaction, and timeout settings.</param>
    ''' <param name="CancellationToken">The token used to cancel connection opening or execution.</param>
    ''' <returns>A response containing the affected-row count.</returns>
    Public Function ExecuteDeleteAsync(Table As String, Optional Options As MySqlMutationOptions = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task(Of MySqlResponse)
        Dim ActualOptions As MySqlMutationOptions = If(Options, New MySqlMutationOptions())
        ValidateMutationOptions(ActualOptions)
        Dim Sql As String = BuildDeleteSql(Table, ActualOptions)
        Return ExecuteNonQueryAsync(Sql, ActualOptions.QueryArgs, ActualOptions, CancellationToken)
    End Function
    Private Shared Function ResolveOptions(Options As MySqlCommandOptions) As MySqlCommandOptions
        Dim ActualOptions As MySqlCommandOptions = If(Options, New MySqlCommandOptions())
        If ActualOptions.CommandTimeout.HasValue AndAlso ActualOptions.CommandTimeout.Value < 0 Then Throw New ArgumentOutOfRangeException(NameOf(Options), ActualOptions.CommandTimeout.Value, "CommandTimeout cannot be negative.")
        Return ActualOptions
    End Function
    Private Shared Sub ValidateSql(Sql As String, ParameterName As String)
        RequireValue(Sql, ParameterName)
    End Sub
    Private Shared Sub ValidateValues(Values As IDictionary(Of String, Object))
        ArgumentNullException.ThrowIfNull(Values)
        If Values.Count = 0 Then Throw New ArgumentException("At least one column value is required.", NameOf(Values))
    End Sub
    Private Shared Sub ValidateMutationOptions(Options As MySqlMutationOptions)
        ResolveOptions(Options)
        If String.IsNullOrWhiteSpace(Options.Where) AndAlso Not Options.AllowAllRows Then Throw New InvalidOperationException("A WHERE clause is required. Set AllowAllRows to True only for an intentional full-table operation.")
    End Sub
    Private Shared Function CreateCommand(Sql As String, CommandType As CommandType, Connection As MySqlConnection, Options As MySqlCommandOptions) As MySqlCommand
        Dim Command As New MySqlCommand With {.CommandText = Sql, .CommandType = CommandType, .Connection = Connection}
        If Options.Transaction IsNot Nothing Then Command.Transaction = Options.Transaction
        If Options.CommandTimeout.HasValue Then Command.CommandTimeout = Options.CommandTimeout.Value
        Return Command
    End Function
    Private Shared Sub AddParameters(Command As MySqlCommand, QueryArgs As IDictionary(Of String, Object))
        If QueryArgs Is Nothing Then Return
        For Each Argument As KeyValuePair(Of String, Object) In QueryArgs
            Dim ParameterName As String = NormalizeParameterName(Argument.Key)
            If Command.Parameters.Contains(ParameterName) Then Throw New ArgumentException($"A parameter named '{ParameterName}' was supplied more than once.", NameOf(QueryArgs))
            Command.Parameters.Add(New MySqlParameter(ParameterName, If(Argument.Value, DBNull.Value)))
        Next Argument
    End Sub
    Private Shared Sub AddProcedureParameters(Command As MySqlCommand, Parameters As IEnumerable(Of MySqlProcedureParameter))
        If Parameters Is Nothing Then Return
        For Each Definition As MySqlProcedureParameter In Parameters
            ArgumentNullException.ThrowIfNull(Definition)
            If Command.Parameters.Contains(Definition.Name) Then Throw New ArgumentException($"A parameter named '{Definition.Name}' was supplied more than once.", NameOf(Parameters))
            If Not [Enum].IsDefined(GetType(ParameterDirection), Definition.Direction) Then Throw New ArgumentOutOfRangeException(NameOf(Parameters), Definition.Direction, "The parameter direction is not valid.")
            If Definition.Direction <> ParameterDirection.Input AndAlso Not Definition.MySqlDbType.HasValue Then Throw New ArgumentException($"The output parameter '{Definition.Name}' must define MySqlDbType.", NameOf(Parameters))
            If Definition.Size.HasValue AndAlso Definition.Size.Value <= 0 Then Throw New ArgumentOutOfRangeException(NameOf(Parameters), Definition.Size.Value, "Size must be greater than zero.")
            If Definition.Scale.HasValue AndAlso Definition.Precision.HasValue AndAlso Definition.Scale.Value > Definition.Precision.Value Then Throw New ArgumentException($"The scale of parameter '{Definition.Name}' cannot exceed its precision.", NameOf(Parameters))
            Dim parameter As New MySqlParameter With {.ParameterName = Definition.Name, .Direction = Definition.Direction, .Value = If(Definition.Value, DBNull.Value)}
            If Definition.MySqlDbType.HasValue Then parameter.MySqlDbType = Definition.MySqlDbType.Value
            If Definition.Size.HasValue Then parameter.Size = Definition.Size.Value
            If Definition.Precision.HasValue Then parameter.Precision = Definition.Precision.Value
            If Definition.Scale.HasValue Then parameter.Scale = Definition.Scale.Value
            Command.Parameters.Add(parameter)
        Next Definition
    End Sub
    Private Shared Function ReadResultSets(Reader As MySqlDataReader) As IReadOnlyList(Of MySqlResultSet)
        Dim ResultSets As New List(Of MySqlResultSet)()
        Do
            If Reader.FieldCount > 0 Then ResultSets.Add(ReadResultSet(Reader))
        Loop While Reader.NextResult()
        Return ResultSets.AsReadOnly()
    End Function
    Private Shared Async Function ReadResultSetsAsync(Reader As MySqlDataReader, cancellationToken As CancellationToken) As Task(Of IReadOnlyList(Of MySqlResultSet))
        Dim ResultSets As New List(Of MySqlResultSet)()
        Do
            If Reader.FieldCount > 0 Then ResultSets.Add(Await ReadResultSetAsync(Reader, cancellationToken).ConfigureAwait(False))
        Loop While Await Reader.NextResultAsync(cancellationToken).ConfigureAwait(False)
        Return ResultSets.AsReadOnly()
    End Function
    Private Shared Function ReadResultSet(Reader As MySqlDataReader) As MySqlResultSet
        Dim Columns As IReadOnlyList(Of String) = GetUniqueColumnNames(Reader)
        Dim Rows As New List(Of IReadOnlyDictionary(Of String, Object))()
        While Reader.Read()
            Rows.Add(ReadRow(Reader, Columns))
        End While
        Return New MySqlResultSet(Columns, Rows)
    End Function
    Private Shared Async Function ReadResultSetAsync(Reader As MySqlDataReader, CancellationToken As CancellationToken) As Task(Of MySqlResultSet)
        Dim Columns As IReadOnlyList(Of String) = GetUniqueColumnNames(Reader)
        Dim Rows As New List(Of IReadOnlyDictionary(Of String, Object))()
        While Await Reader.ReadAsync(CancellationToken).ConfigureAwait(False)
            Rows.Add(ReadRow(Reader, Columns))
        End While
        Return New MySqlResultSet(Columns, Rows)
    End Function
    Private Shared Function GetUniqueColumnNames(Reader As MySqlDataReader) As IReadOnlyList(Of String)
        Dim Columns As New List(Of String)(Reader.FieldCount)
        Dim UsedNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Ordinal As Integer = 0 To Reader.FieldCount - 1
            Dim BaseName As String = Reader.GetName(Ordinal)
            If String.IsNullOrWhiteSpace(BaseName) Then BaseName = $"Column{Ordinal + 1}"
            Dim UniqueName As String = BaseName
            Dim Suffix As Integer = 2
            While Not UsedNames.Add(UniqueName)
                UniqueName = $"{BaseName}_{Suffix}"
                Suffix += 1
            End While
            Columns.Add(UniqueName)
        Next Ordinal
        Return Columns.AsReadOnly()
    End Function
    Private Shared Function ReadRow(Reader As MySqlDataReader, Columns As IReadOnlyList(Of String)) As IReadOnlyDictionary(Of String, Object)
        Dim Row As New Dictionary(Of String, Object)(Columns.Count, StringComparer.OrdinalIgnoreCase)
        For Ordinal As Integer = 0 To Columns.Count - 1
            Row.Add(Columns(Ordinal), ConvertDatabaseValue(Reader.GetValue(Ordinal)))
        Next Ordinal
        Return New ReadOnlyDictionary(Of String, Object)(Row)
    End Function
    Private Shared Function ConvertDatabaseValue(Value As Object) As Object
        If Value Is Nothing OrElse Convert.IsDBNull(Value) Then Return Nothing
        Return Value
    End Function
    Private Shared Function NormalizeAffectedRows(AffectedRows As Integer) As Long
        If AffectedRows < 0 Then Return 0
        Return AffectedRows
    End Function
    Private Shared Function NormalizeLastInsertedId(LastInsertedId As Long) As Long?
        If LastInsertedId = 0 Then Return Nothing
        Return LastInsertedId
    End Function
    Private Shared Function CreateProcedureResponse(Command As MySqlCommand, ResultSets As IReadOnlyList(Of MySqlResultSet), AffectedRows As Long) As MySqlResponse
        Dim OutputParameters As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
        Dim ReturnValue As Object = Nothing
        For Index As Integer = 0 To Command.Parameters.Count - 1
            Dim Parameter As MySqlParameter = Command.Parameters(Index)
            Select Case Parameter.Direction
                Case ParameterDirection.Output, ParameterDirection.InputOutput
                    OutputParameters(Parameter.ParameterName) = ConvertDatabaseValue(Parameter.Value)
                Case ParameterDirection.ReturnValue
                    ReturnValue = ConvertDatabaseValue(Parameter.Value)
            End Select
        Next Index
        Return New MySqlResponse(ResultSets, AffectedRows, OutputParameters:=OutputParameters, returnValue:=ReturnValue)
    End Function
    Private Shared Function BuildSelectSql(Table As String, Options As MySqlSelectOptions) As String
        ResolveOptions(Options)
        If Options.Limit.HasValue AndAlso Options.Limit.Value <= 0 Then Throw New ArgumentOutOfRangeException(NameOf(Options), Options.Limit.Value, "Limit must be greater than zero.")
        If Options.Offset.HasValue AndAlso Options.Offset.Value < 0 Then Throw New ArgumentOutOfRangeException(NameOf(Options), Options.Offset.Value, "Offset cannot be negative.")
        If Options.Offset.GetValueOrDefault() > 0 AndAlso Not Options.Limit.HasValue Then Throw New InvalidOperationException("Offset requires Limit to be specified.")
        Dim Projections As New List(Of String)()
        For Each Column As String In Options.Columns
            Projections.Add(QuoteIdentifier(Column, NameOf(Options.Columns), True))
        Next Column
        For Each Expression As String In Options.TrustedExpressions
            Projections.Add(RequireValue(Expression, NameOf(Options.TrustedExpressions)))
        Next Expression
        If Projections.Count = 0 Then Projections.Add("*")
        Dim Builder As New StringBuilder("SELECT ")
        If Options.Distinct Then Builder.Append("DISTINCT ")
        Builder.Append(String.Join(", ", Projections)).Append(" FROM ").Append(QuoteIdentifier(Table, NameOf(Table)))
        If Not String.IsNullOrWhiteSpace(Options.Where) Then Builder.Append(" WHERE ").Append(Options.Where.Trim())
        If Options.OrderBy.Count > 0 Then
            Dim Ordering As New List(Of String)(Options.OrderBy.Count)
            For Each Item As MySqlOrderBy In Options.OrderBy
                ArgumentNullException.ThrowIfNull(Item)
                Dim direction As String = If(Item.Direction = MySqlSortDirection.Descending, "DESC", "ASC")
                Ordering.Add($"{QuoteIdentifier(Item.Column, NameOf(Options.OrderBy))} {direction}")
            Next Item
            Builder.Append(" ORDER BY ").Append(String.Join(", ", Ordering))
        End If
        If Options.Limit.HasValue Then Builder.Append(" LIMIT ").Append(Options.Limit.Value)
        If Options.Offset.GetValueOrDefault() > 0 Then Builder.Append(" OFFSET ").Append(Options.Offset.Value)
        Return Builder.ToString()
    End Function
    Private Shared Function BuildInsertSql(Table As String, Values As IDictionary(Of String, Object)) As CommandBuildResult
        Dim Columns As New List(Of String)(Values.Count)
        Dim ParameterNames As New List(Of String)(Values.Count)
        Dim Parameters As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
        Dim Index As Integer
        For Each Value As KeyValuePair(Of String, Object) In Values
            Columns.Add(QuoteIdentifier(Value.Key, NameOf(Values)))
            Dim ParameterName As String = $"@__coresuite_value{Index}"
            ParameterNames.Add(ParameterName)
            Parameters.Add(ParameterName, Value.Value)
            Index += 1
        Next Value
        Dim Sql As String = $"INSERT INTO {QuoteIdentifier(Table, NameOf(Table))} ({String.Join(", ", Columns)}) VALUES ({String.Join(", ", ParameterNames)})"
        Return New CommandBuildResult(Sql, Parameters)
    End Function
    Private Shared Function BuildUpdateSql(Table As String, Values As IDictionary(Of String, Object), Options As MySqlMutationOptions) As CommandBuildResult
        Dim Parameters As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
        If Options.QueryArgs IsNot Nothing Then
            For Each Argument As KeyValuePair(Of String, Object) In Options.QueryArgs
                Parameters.Add(NormalizeParameterName(Argument.Key), Argument.Value)
            Next Argument
        End If
        Dim Assignments As New List(Of String)(Values.Count)
        Dim Index As Integer
        For Each Value As KeyValuePair(Of String, Object) In Values
            Dim ParameterName As String = CreateGeneratedParameterName(Index, Parameters)
            Assignments.Add($"{QuoteIdentifier(Value.Key, NameOf(Values))} = {ParameterName}")
            Parameters.Add(ParameterName, Value.Value)
            Index += 1
        Next Value
        Dim Builder As New StringBuilder("UPDATE ")
        Builder.Append(QuoteIdentifier(Table, NameOf(Table)))
        Builder.Append(" SET ")
        Builder.Append(String.Join(", ", Assignments))
        If Not String.IsNullOrWhiteSpace(Options.Where) Then Builder.Append(" WHERE ").Append(Options.Where.Trim())
        Return New CommandBuildResult(Builder.ToString(), Parameters)
    End Function
    Private Shared Function BuildDeleteSql(Table As String, Options As MySqlMutationOptions) As String
        Dim Builder As New StringBuilder("DELETE FROM ")
        Builder.Append(QuoteIdentifier(Table, NameOf(Table)))
        If Not String.IsNullOrWhiteSpace(Options.Where) Then Builder.Append(" WHERE ").Append(Options.Where.Trim())
        Return Builder.ToString()
    End Function
    Private Shared Function CreateGeneratedParameterName(Index As Integer, ExistingParameters As Dictionary(Of String, Object)) As String
        Dim ParameterName As String = $"@__coresuite_value{Index}"
        While ExistingParameters.ContainsKey(ParameterName)
            Index += 1
            ParameterName = $"@__coresuite_value{Index}"
        End While
        Return ParameterName
    End Function
    Private NotInheritable Class CommandBuildResult
        Friend Sub New(sql As String, parameters As IDictionary(Of String, Object))
            Me.Sql = sql
            Me.Parameters = parameters
        End Sub
        Friend ReadOnly Property Sql As String
        Friend ReadOnly Property Parameters As IDictionary(Of String, Object)
    End Class
End Class
