Imports System.Data.Common

''' <summary>
''' Class responsible for executing SQL commands, CRUD operations, and stored procedures
''' on a MySQL database using a <see cref="MySqlClient"/>.
''' </summary>
Public Class MySqlRequest
    Private ReadOnly _Client As MySqlClient

    Friend Sub New(Client As MySqlClient)
        _Client = Client
    End Sub
    ''' <summary>
    ''' Executes a stored procedure on the database.
    ''' </summary>
    ''' <param name="ProcedureName">Name of the stored procedure.</param>
    ''' <param name="Params">Procedure parameters, if any.</param>
    ''' <param name="Connection">Optional connection. If not provided, the class will create one.</param>
    ''' <returns>
    ''' A <see cref="MySqlResponse"/> containing the results.
    ''' </returns>
    Public Function ExecuteProcedure(ProcedureName As String, Optional Params As Dictionary(Of String, Object) = Nothing, Optional Connection As DbConnection = Nothing) As MySqlResponse
        Dim Data As List(Of Dictionary(Of String, Object)) = Nothing
        Dim OwnsConnection As Boolean = (Connection Is Nothing)
        If OwnsConnection Then
            Connection = _Client.CreateDatabaseConnection()
            Connection.Open()
        ElseIf Connection.State <> ConnectionState.Open Then
            Connection.Open()
        End If
        Try
            Using Command As DbCommand = Connection.CreateCommand()
                Command.CommandText = ProcedureName
                Command.CommandType = CommandType.StoredProcedure
                If Params IsNot Nothing Then
                    For Each Arg In Params
                        Dim Param = Command.CreateParameter()
                        Param.ParameterName = Arg.Key
                        Param.Value = If(Arg.Value, DBNull.Value)
                        Command.Parameters.Add(Param)
                    Next Arg
                End If
                Data = New List(Of Dictionary(Of String, Object))()
                Using Reader As DbDataReader = Command.ExecuteReader()
                    While Reader.Read()
                        Dim Item As New Dictionary(Of String, Object)
                        For i = 0 To Reader.FieldCount - 1
                            Item.Add(Reader.GetName(i), Reader.GetValue(i))
                        Next i
                        Data.Add(Item)
                    End While
                End Using
            End Using
            Return New MySqlResponse(Data, 0, Nothing)
        Catch
            Throw
        Finally
            If OwnsConnection Then
                Connection.Dispose()
            End If
        End Try
    End Function
    ''' <summary>
    ''' Executes a stored procedure on the database.
    ''' </summary>
    ''' <param name="ProcedureName">Name of the stored procedure.</param>
    ''' <param name="Params">Procedure parameters, if any.</param>
    ''' <param name="Connection">Optional connection. If not provided, the class will create one.</param>
    ''' <returns>
    ''' A task returning a <see cref="MySqlResponse"/> containing the results.
    ''' </returns>
    Public Async Function ExecuteProcedureAsync(ProcedureName As String, Optional Params As Dictionary(Of String, Object) = Nothing, Optional Connection As DbConnection = Nothing) As Task(Of MySqlResponse)
        Dim Data As List(Of Dictionary(Of String, Object)) = Nothing
        Dim OwnsConnection As Boolean = (Connection Is Nothing)
        If OwnsConnection Then
            Connection = _Client.CreateDatabaseConnection()
            Await Connection.OpenAsync()
        ElseIf Connection.State <> ConnectionState.Open Then
            Await Connection.OpenAsync()
        End If
        Try
            Using Command As DbCommand = Connection.CreateCommand()
                Command.CommandText = ProcedureName
                Command.CommandType = CommandType.StoredProcedure
                If Params IsNot Nothing Then
                    For Each Arg In Params
                        Dim Param = Command.CreateParameter()
                        Param.ParameterName = Arg.Key
                        Param.Value = If(Arg.Value, DBNull.Value)
                        Command.Parameters.Add(Param)
                    Next Arg
                End If
                Data = New List(Of Dictionary(Of String, Object))()
                Using Reader As DbDataReader = Await Command.ExecuteReaderAsync()
                    While Await Reader.ReadAsync()
                        Dim Item As New Dictionary(Of String, Object)
                        For i = 0 To Reader.FieldCount - 1
                            Item.Add(Reader.GetName(i), Reader.GetValue(i))
                        Next i
                        Data.Add(Item)
                    End While
                End Using
            End Using
            Return New MySqlResponse(Data, 0, Nothing)
        Catch
            Throw
        Finally
            If OwnsConnection Then
                Connection.Dispose()
            End If
        End Try
    End Function

    ''' <summary>
    ''' Executes a generic SQL query (SELECT, INSERT, UPDATE, DELETE).
    ''' </summary>
    ''' <param name="Query">SQL query to be executed.</param>
    ''' <param name="QueryArgs">Query parameters, if any.</param>
    ''' <param name="Connection">Optional connection. If not provided, the class will create one.</param>
    ''' <returns>
    ''' A task returning a <see cref="MySqlResponse"/> containing query results
    ''' or the number of affected rows.
    ''' </returns>
    Public Function ExecuteRawQuery(Query As String, Optional QueryArgs As Dictionary(Of String, Object) = Nothing, Optional Connection As DbConnection = Nothing) As MySqlResponse
        Dim IsSelect As Boolean = Query.TrimStart().StartsWith("SELECT", StringComparison.OrdinalIgnoreCase)
        Dim AffectedRows As Integer = 0
        Dim Data As List(Of Dictionary(Of String, Object)) = Nothing
        Dim OwnsConnection As Boolean = (Connection Is Nothing)
        If OwnsConnection Then
            Connection = _Client.CreateDatabaseConnection()
            Connection.Open()
        ElseIf Connection.State <> ConnectionState.Open Then
            Connection.Open()
        End If
        Try
            Using Command As DbCommand = Connection.CreateCommand()
                Command.CommandText = Query
                If QueryArgs IsNot Nothing Then
                    For Each Arg In QueryArgs
                        Dim Param = Command.CreateParameter()
                        Param.ParameterName = Arg.Key
                        Param.Value = If(Arg.Value, DBNull.Value)
                        Command.Parameters.Add(Param)
                    Next Arg
                End If
                If IsSelect Then
                    Data = New List(Of Dictionary(Of String, Object))()
                    Using Reader As DbDataReader = Command.ExecuteReader()
                        While Reader.Read()
                            Dim Item As New Dictionary(Of String, Object)
                            For i = 0 To Reader.FieldCount - 1
                                Item.Add(Reader.GetName(i), Reader.GetValue(i))
                            Next i
                            Data.Add(Item)
                        End While
                    End Using
                Else
                    AffectedRows = Command.ExecuteNonQuery()
                End If
            End Using
            Return New MySqlResponse(Data, AffectedRows, Nothing)
        Catch
            Throw
        Finally
            If OwnsConnection Then
                Connection.Dispose()
            End If
        End Try
    End Function

    ''' <summary>
    ''' Executes a generic SQL query (SELECT, INSERT, UPDATE, DELETE).
    ''' </summary>
    ''' <param name="Query">SQL query to be executed.</param>
    ''' <param name="QueryArgs">Query parameters, if any.</param>
    ''' <param name="Connection">Optional connection. If not provided, the class will create one.</param>
    ''' <returns>
    ''' A task returning a <see cref="MySqlResponse"/> containing query results
    ''' or the number of affected rows.
    ''' </returns>
    Public Async Function ExecuteRawQueryAsync(Query As String, Optional QueryArgs As Dictionary(Of String, Object) = Nothing, Optional Connection As DbConnection = Nothing) As Task(Of MySqlResponse)
        Dim IsSelect As Boolean = Query.TrimStart().StartsWith("SELECT", StringComparison.OrdinalIgnoreCase)
        Dim AffectedRows As Integer = 0
        Dim Data As List(Of Dictionary(Of String, Object)) = Nothing
        Dim OwnsConnection As Boolean = (Connection Is Nothing)
        If OwnsConnection Then
            Connection = _Client.CreateDatabaseConnection()
            Await Connection.OpenAsync()
        ElseIf Connection.State <> ConnectionState.Open Then
            Await Connection.OpenAsync()
        End If
        Try
            Using Command As DbCommand = Connection.CreateCommand()
                Command.CommandText = Query
                If QueryArgs IsNot Nothing Then
                    For Each Arg In QueryArgs
                        Dim Param = Command.CreateParameter()
                        Param.ParameterName = Arg.Key
                        Param.Value = If(Arg.Value, DBNull.Value)
                        Command.Parameters.Add(Param)
                    Next Arg
                End If
                If IsSelect Then
                    Data = New List(Of Dictionary(Of String, Object))()
                    Using Reader As DbDataReader = Await Command.ExecuteReaderAsync()
                        While Await Reader.ReadAsync()
                            Dim Item As New Dictionary(Of String, Object)
                            For i = 0 To Reader.FieldCount - 1
                                Item.Add(Reader.GetName(i), Reader.GetValue(i))
                            Next i
                            Data.Add(Item)
                        End While
                    End Using
                Else
                    AffectedRows = Await Command.ExecuteNonQueryAsync()
                End If
            End Using
            Return New MySqlResponse(Data, AffectedRows, Nothing)
        Catch
            Throw
        Finally
            If OwnsConnection Then
                Connection.Dispose()
            End If
        End Try
    End Function
    ''' <summary>
    ''' Executes a DELETE operation on a table.
    ''' </summary>
    ''' <param name="Table">Table name.</param>
    ''' <param name="Where">Optional WHERE clause.</param>
    ''' <param name="QueryArgs">Query parameters.</param>
    ''' <param name="Connection">Optional connection.</param>
    ''' <returns>
    ''' A task returning a <see cref="MySqlResponse"/> containing the number of affected rows.
    ''' </returns>
    Public Function ExecuteDelete(Table As String, Optional Where As String = Nothing, Optional QueryArgs As Dictionary(Of String, Object) = Nothing, Optional Connection As DbConnection = Nothing) As MySqlResponse
        Dim AffectedRows As Integer
        Dim OwnsConnection As Boolean = (Connection Is Nothing)
        If OwnsConnection Then
            Connection = _Client.CreateDatabaseConnection()
            Connection.Open()
        ElseIf Connection.State <> ConnectionState.Open Then
            Connection.Open()
        End If
        Try
            Dim Query As String = $"DELETE FROM {Table}"
            If Not String.IsNullOrWhiteSpace(Where) Then
                Query &= $" WHERE {Where}"
            End If
            Using Command As DbCommand = Connection.CreateCommand()
                Command.CommandText = Query
                If QueryArgs IsNot Nothing Then
                    For Each Arg In QueryArgs
                        Dim Param = Command.CreateParameter()
                        Param.ParameterName = Arg.Key
                        Param.Value = If(Arg.Value, DBNull.Value)
                        Command.Parameters.Add(Param)
                    Next Arg
                End If
                AffectedRows = Command.ExecuteNonQuery()
            End Using
            Return New MySqlResponse(Nothing, AffectedRows, Nothing)
        Catch
            Throw
        Finally
            If OwnsConnection Then
                Connection.Dispose()
            End If
        End Try
    End Function
    ''' <summary>
    ''' Executes a DELETE operation on a table.
    ''' </summary>
    ''' <param name="Table">Table name.</param>
    ''' <param name="Where">Optional WHERE clause.</param>
    ''' <param name="QueryArgs">Query parameters.</param>
    ''' <param name="Connection">Optional connection.</param>
    ''' <returns>
    ''' A task returning a <see cref="MySqlResponse"/> containing the number of affected rows.
    ''' </returns>
    Public Async Function ExecuteDeleteAsync(Table As String, Optional Where As String = Nothing, Optional QueryArgs As Dictionary(Of String, Object) = Nothing, Optional Connection As DbConnection = Nothing) As Task(Of MySqlResponse)
        Dim AffectedRows As Integer
        Dim OwnsConnection As Boolean = (Connection Is Nothing)
        If OwnsConnection Then
            Connection = _Client.CreateDatabaseConnection()
            Await Connection.OpenAsync()
        ElseIf Connection.State <> ConnectionState.Open Then
            Await Connection.OpenAsync()
        End If
        Try
            Dim Query As String = $"DELETE FROM {Table}"
            If Not String.IsNullOrWhiteSpace(Where) Then
                Query &= $" WHERE {Where}"
            End If
            Using Command As DbCommand = Connection.CreateCommand()
                Command.CommandText = Query
                If QueryArgs IsNot Nothing Then
                    For Each Arg In QueryArgs
                        Dim Param = Command.CreateParameter()
                        Param.ParameterName = Arg.Key
                        Param.Value = If(Arg.Value, DBNull.Value)
                        Command.Parameters.Add(Param)
                    Next Arg
                End If
                AffectedRows = Await Command.ExecuteNonQueryAsync()
            End Using
            Return New MySqlResponse(Nothing, AffectedRows, Nothing)
        Catch
            Throw
        Finally
            If OwnsConnection Then
                Connection.Dispose()
            End If
        End Try
    End Function

    ''' <summary>
    ''' Executes an UPDATE operation on a table.
    ''' </summary>
    ''' <param name="Table">Table name.</param>
    ''' <param name="Values">Dictionary containing columns and values to be updated.</param>
    ''' <param name="Where">Optional WHERE clause.</param>
    ''' <param name="QueryArgs">Query parameters.</param>
    ''' <param name="Connection">Optional connection.</param>
    ''' <returns>
    ''' A task returning a <see cref="MySqlResponse"/> containing the number of affected rows.
    ''' </returns>
    Public Function ExecuteUpdate(Table As String, Values As Dictionary(Of String, Object), Optional Where As String = Nothing, Optional QueryArgs As Dictionary(Of String, Object) = Nothing, Optional Connection As DbConnection = Nothing) As MySqlResponse
        If Values Is Nothing OrElse Values.Count = 0 Then
            Throw New ArgumentException("Values cannot be empty.", NameOf(Values))
        End If
        Dim AffectedRows As Integer
        Dim OwnsConnection As Boolean = (Connection Is Nothing)
        If OwnsConnection Then
            Connection = _Client.CreateDatabaseConnection()
            Connection.Open()
        ElseIf Connection.State <> ConnectionState.Open Then
            Connection.Open()
        End If
        Try
            Using Command As DbCommand = Connection.CreateCommand()
                Dim SetClauses As New List(Of String)
                For Each Kvp In Values
                    Dim ParamName As String = $"@upd_{Kvp.Key}"
                    SetClauses.Add($"{Kvp.Key} = {ParamName}")
                    Dim Param = Command.CreateParameter()
                    Param.ParameterName = ParamName
                    Param.Value = If(Kvp.Value, DBNull.Value)
                    Command.Parameters.Add(Param)
                Next Kvp
                Dim Query As String = $"UPDATE {Table} SET {String.Join(", ", SetClauses)}"
                If Not String.IsNullOrWhiteSpace(Where) Then
                    Query &= $" WHERE {Where}"
                End If
                Command.CommandText = Query
                If QueryArgs IsNot Nothing Then
                    For Each Arg In QueryArgs
                        Dim Param = Command.CreateParameter()
                        Param.ParameterName = Arg.Key
                        Param.Value = If(Arg.Value, DBNull.Value)
                        Command.Parameters.Add(Param)
                    Next Arg
                End If
                AffectedRows = Command.ExecuteNonQuery()
            End Using
            Return New MySqlResponse(Nothing, AffectedRows)
        Catch
            Throw
        Finally
            If OwnsConnection Then
                Connection.Dispose()
            End If
        End Try
    End Function

    ''' <summary>
    ''' Executes an UPDATE operation on a table.
    ''' </summary>
    ''' <param name="Table">Table name.</param>
    ''' <param name="Values">Dictionary containing columns and values to be updated.</param>
    ''' <param name="Where">Optional WHERE clause.</param>
    ''' <param name="QueryArgs">Query parameters.</param>
    ''' <param name="Connection">Optional connection.</param>
    ''' <returns>
    ''' A task returning a <see cref="MySqlResponse"/> containing the number of affected rows.
    ''' </returns>
    Public Async Function ExecuteUpdateAsync(Table As String, Values As Dictionary(Of String, Object), Optional Where As String = Nothing, Optional QueryArgs As Dictionary(Of String, Object) = Nothing, Optional Connection As DbConnection = Nothing) As Task(Of MySqlResponse)
        If Values Is Nothing OrElse Values.Count = 0 Then
            Throw New ArgumentException("Values cannot be empty.", NameOf(Values))
        End If
        Dim AffectedRows As Integer
        Dim OwnsConnection As Boolean = (Connection Is Nothing)
        If OwnsConnection Then
            Connection = _Client.CreateDatabaseConnection()
            Await Connection.OpenAsync()
        ElseIf Connection.State <> ConnectionState.Open Then
            Await Connection.OpenAsync()
        End If
        Try
            Using Command As DbCommand = Connection.CreateCommand()
                Dim SetClauses As New List(Of String)
                For Each Kvp In Values
                    Dim ParamName As String = $"@upd_{Kvp.Key}"
                    SetClauses.Add($"{Kvp.Key} = {ParamName}")
                    Dim Param = Command.CreateParameter()
                    Param.ParameterName = ParamName
                    Param.Value = If(Kvp.Value, DBNull.Value)
                    Command.Parameters.Add(Param)
                Next Kvp
                Dim Query As String = $"UPDATE {Table} SET {String.Join(", ", SetClauses)}"
                If Not String.IsNullOrWhiteSpace(Where) Then
                    Query &= $" WHERE {Where}"
                End If
                Command.CommandText = Query
                If QueryArgs IsNot Nothing Then
                    For Each Arg In QueryArgs
                        Dim Param = Command.CreateParameter()
                        Param.ParameterName = Arg.Key
                        Param.Value = If(Arg.Value, DBNull.Value)
                        Command.Parameters.Add(Param)
                    Next Arg
                End If
                AffectedRows = Await Command.ExecuteNonQueryAsync()
            End Using
            Return New MySqlResponse(Nothing, AffectedRows)
        Catch
            Throw
        Finally
            If OwnsConnection Then
                Connection.Dispose()
            End If
        End Try
    End Function

    ''' <summary>
    ''' Executes an INSERT operation on a table.
    ''' </summary>
    ''' <param name="Table">Table name.</param>
    ''' <param name="Values">Dictionary containing columns and values to be inserted.</param>
    ''' <param name="QueryArgs">Optional query parameters.</param>
    ''' <param name="Connection">Optional connection.</param>
    ''' <returns>
    ''' A task returning a <see cref="MySqlResponse"/> containing the number of affected rows
    ''' and the last inserted ID.
    ''' </returns>
    Public Function ExecuteInsert(Table As String, Values As Dictionary(Of String, Object), Optional QueryArgs As Dictionary(Of String, Object) = Nothing, Optional Connection As DbConnection = Nothing) As MySqlResponse
        If Values Is Nothing OrElse Values.Count = 0 Then
            Throw New ArgumentException("Values cannot be empty.", NameOf(Values))
        End If
        Dim AffectedRows As Integer
        Dim LastInsertedRow As Integer
        Dim OwnsConnection As Boolean = (Connection Is Nothing)
        If OwnsConnection Then
            Connection = _Client.CreateDatabaseConnection()
            Connection.Open()
        ElseIf Connection.State <> ConnectionState.Open Then
            Connection.Open()
        End If
        Try
            Using Command As DbCommand = Connection.CreateCommand()
                Dim Columns As String = String.Join(", ", Values.Keys)
                Dim ParamNames As New List(Of String)
                For Each Kvp In Values
                    Dim ParamName As String = $"@ins_{Kvp.Key}"
                    ParamNames.Add(ParamName)
                    Dim Param As DbParameter = Command.CreateParameter()
                    Param.ParameterName = ParamName
                    Param.Value = If(Kvp.Value, DBNull.Value)
                    Command.Parameters.Add(Param)
                Next Kvp
                Dim Query As String = $"INSERT INTO {Table} ({Columns}) VALUES ({String.Join(", ", ParamNames)})"
                Command.CommandText = Query
                If QueryArgs IsNot Nothing Then
                    For Each Arg In QueryArgs
                        Dim param = Command.CreateParameter()
                        param.ParameterName = Arg.Key
                        param.Value = If(Arg.Value, DBNull.Value)
                        Command.Parameters.Add(param)
                    Next Arg
                End If
                AffectedRows = Command.ExecuteNonQuery()
                Command.CommandText = "SELECT LAST_INSERT_ID();"
                LastInsertedRow = Convert.ToInt32(Command.ExecuteScalar())
            End Using
            Return New MySqlResponse(Nothing, AffectedRows, LastInsertedRow)
        Catch
            Throw
        Finally
            If OwnsConnection Then
                Connection.Dispose()
            End If
        End Try
    End Function

    ''' <summary>
    ''' Executes an INSERT operation on a table.
    ''' </summary>
    ''' <param name="Table">Table name.</param>
    ''' <param name="Values">Dictionary containing columns and values to be inserted.</param>
    ''' <param name="QueryArgs">Optional query parameters.</param>
    ''' <param name="Connection">Optional connection.</param>
    ''' <returns>
    ''' A task returning a <see cref="MySqlResponse"/> containing the number of affected rows
    ''' and the last inserted ID.
    ''' </returns>
    Public Async Function ExecuteInsertAsync(Table As String, Values As Dictionary(Of String, Object), Optional QueryArgs As Dictionary(Of String, Object) = Nothing, Optional Connection As DbConnection = Nothing) As Task(Of MySqlResponse)
        If Values Is Nothing OrElse Values.Count = 0 Then
            Throw New ArgumentException("Values cannot be empty.", NameOf(Values))
        End If
        Dim AffectedRows As Integer
        Dim LastInsertedRow As Integer
        Dim OwnsConnection As Boolean = (Connection Is Nothing)
        If OwnsConnection Then
            Connection = _Client.CreateDatabaseConnection()
            Await Connection.OpenAsync()
        ElseIf Connection.State <> ConnectionState.Open Then
            Await Connection.OpenAsync()
        End If
        Try
            Using Command As DbCommand = Connection.CreateCommand()
                Dim Columns As String = String.Join(", ", Values.Keys)
                Dim ParamNames As New List(Of String)
                For Each Kvp In Values
                    Dim ParamName As String = $"@ins_{Kvp.Key}"
                    ParamNames.Add(ParamName)
                    Dim Param As DbParameter = Command.CreateParameter()
                    Param.ParameterName = ParamName
                    Param.Value = If(Kvp.Value, DBNull.Value)
                    Command.Parameters.Add(Param)
                Next Kvp
                Dim Query As String = $"INSERT INTO {Table} ({Columns}) VALUES ({String.Join(", ", ParamNames)})"
                Command.CommandText = Query
                If QueryArgs IsNot Nothing Then
                    For Each Arg In QueryArgs
                        Dim Param As DbParameter = Command.CreateParameter()
                        Param.ParameterName = Arg.Key
                        Param.Value = If(Arg.Value, DBNull.Value)
                        Command.Parameters.Add(Param)
                    Next Arg
                End If
                AffectedRows = Await Command.ExecuteNonQueryAsync()
                Command.CommandText = "SELECT LAST_INSERT_ID();"
                LastInsertedRow = Convert.ToInt32(Await Command.ExecuteScalarAsync())
            End Using
            Return New MySqlResponse(Nothing, AffectedRows, LastInsertedRow)
        Catch
            Throw
        Finally
            If OwnsConnection Then
                Connection.Dispose()
            End If
        End Try
    End Function

    ''' <summary>
    ''' Executes a SELECT operation on a table with advanced options.
    ''' </summary>
    ''' <param name="Table">Table name.</param>
    ''' <param name="Options">
    ''' Selection options such as columns, filters, ordering, limits, and connection.
    ''' </param>
    ''' <returns>
    ''' A task returning a <see cref="MySqlResponse"/> containing the selected data.
    ''' </returns>
    Public Function ExecuteSelect(Table As String, Optional Options As MySqlSelectOptions = Nothing) As MySqlResponse
        Dim OwnsConnection As Boolean = (Options.Connection Is Nothing)
        Dim Data As New List(Of Dictionary(Of String, Object))
        If OwnsConnection Then
            Options.Connection = _Client.CreateDatabaseConnection()
            Options.Connection.Open()
        ElseIf Options.Connection.State <> ConnectionState.Open Then
            Options.Connection.Open()
        End If
        Try
            Dim Query As String = "SELECT "
            If Options.Columns IsNot Nothing AndAlso Options.Columns.Count > 0 Then
                Query &= String.Join(", ", Options.Columns)
            Else
                Query &= "*"
            End If
            Query &= $" FROM {Table} "
            If Not String.IsNullOrEmpty(Options.Where) Then
                Query &= $"WHERE {Options.Where} "
            End If
            If Not String.IsNullOrEmpty(Options.OrderBy) Then
                Query &= $"ORDER BY {Options.OrderBy} "
            End If
            If Options.Limit IsNot Nothing Then
                Query &= $"LIMIT {Options.Limit.Value} "
            End If
            Using Command As DbCommand = Options.Connection.CreateCommand()
                Command.CommandText = Query
                If Options.QueryArgs IsNot Nothing AndAlso Options.QueryArgs.Count > 0 Then
                    For Each Arg In Options.QueryArgs
                        Dim Param As DbParameter = Command.CreateParameter()
                        Param.ParameterName = Arg.Key
                        Param.Value = If(Arg.Value, DBNull.Value)
                        Command.Parameters.Add(Param)
                    Next Arg
                End If
                Using Reader As DbDataReader = Command.ExecuteReader()
                    While Reader.Read()
                        Dim Item As New Dictionary(Of String, Object)
                        For i = 0 To Reader.FieldCount - 1
                            Item.Add(Reader.GetName(i), Reader.GetValue(i))
                        Next i
                        Data.Add(Item)
                    End While
                End Using
            End Using
            Return New MySqlResponse(Data)
        Catch
            Throw
        Finally
            If OwnsConnection Then
                Options.Connection.Dispose()
            End If
        End Try
    End Function

    ''' <summary>
    ''' Executes a SELECT operation on a table with advanced options.
    ''' </summary>
    ''' <param name="Table">Table name.</param>
    ''' <param name="Options">
    ''' Selection options such as columns, filters, ordering, limits, and connection.
    ''' </param>
    ''' <returns>
    ''' A task returning a <see cref="MySqlResponse"/> containing the selected data.
    ''' </returns>
    Public Async Function ExecuteSelectAsync(Table As String, Optional Options As MySqlSelectOptions = Nothing) As Task(Of MySqlResponse)
        Dim OwnsConnection As Boolean = (Options.Connection Is Nothing)
        Dim Data As New List(Of Dictionary(Of String, Object))
        If OwnsConnection Then
            Options.Connection = _Client.CreateDatabaseConnection()
            Await Options.Connection.OpenAsync()
        ElseIf Options.Connection.State <> ConnectionState.Open Then
            Await Options.Connection.OpenAsync()
        End If
        Try
            Dim Query As String = "SELECT "
            If Options.Columns IsNot Nothing AndAlso Options.Columns.Count > 0 Then
                Query &= String.Join(", ", Options.Columns)
            Else
                Query &= "*"
            End If
            Query &= $" FROM {Table} "
            If Not String.IsNullOrEmpty(Options.Where) Then
                Query &= $"WHERE {Options.Where} "
            End If
            If Not String.IsNullOrEmpty(Options.OrderBy) Then
                Query &= $"ORDER BY {Options.OrderBy} "
            End If
            If Options.Limit IsNot Nothing Then
                Query &= $"LIMIT {Options.Limit.Value} "
            End If
            Using Command As DbCommand = Options.Connection.CreateCommand()
                Command.CommandText = Query
                If Options.QueryArgs IsNot Nothing AndAlso Options.QueryArgs.Count > 0 Then
                    For Each Arg In Options.QueryArgs
                        Dim Param = Command.CreateParameter()
                        Param.ParameterName = Arg.Key
                        Param.Value = If(Arg.Value, DBNull.Value)
                        Command.Parameters.Add(Param)
                    Next Arg
                End If
                Using Reader As DbDataReader = Await Command.ExecuteReaderAsync()
                    While Await Reader.ReadAsync()
                        Dim Item As New Dictionary(Of String, Object)
                        For i = 0 To Reader.FieldCount - 1
                            Item.Add(Reader.GetName(i), Reader.GetValue(i))
                        Next i
                        Data.Add(Item)
                    End While
                End Using
            End Using
            Return New MySqlResponse(Data)
        Catch
            Throw
        Finally
            If OwnsConnection Then
                Options.Connection.Dispose()
            End If
        End Try
    End Function
End Class