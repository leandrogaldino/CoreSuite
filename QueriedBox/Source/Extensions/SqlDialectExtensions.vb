Imports System.Runtime.CompilerServices

''' <summary>
''' Provides extension methods that generate SQL syntax specific to each <see cref="SqlDialect"/>.
''' </summary>
Public Module SqlDialectExtensions

    ''' <summary>
    ''' Returns the parameter prefix used by the specified SQL dialect.
    ''' </summary>
    ''' <param name="Dialect">The SQL dialect.</param>
    ''' <returns>
    ''' The parameter prefix for the specified dialect, such as <c>@</c> or <c>:</c>.
    ''' </returns>
    ''' <exception cref="ArgumentOutOfRangeException">
    ''' Thrown when the specified dialect is not supported.
    ''' </exception>
    <Extension>
    Public Function GetParameterPrefix(Dialect As SqlDialect) As String
        Select Case Dialect
            Case SqlDialect.MySql,
                 SqlDialect.SqlServer,
                 SqlDialect.PostgreSql,
                 SqlDialect.Sqlite,
                 SqlDialect.Firebird
                Return "@"
            Case SqlDialect.Oracle
                Return ":"
            Case Else
                Throw New ArgumentOutOfRangeException(NameOf(Dialect))
        End Select
    End Function

    ''' <summary>
    ''' Returns the SQL expression used to replace <c>NULL</c> values for the specified SQL dialect.
    ''' </summary>
    ''' <param name="Dialect">The SQL dialect.</param>
    ''' <param name="Expression">The expression to evaluate.</param>
    ''' <param name="Replacement">The value to return when the expression is <c>NULL</c>.</param>
    ''' <returns>
    ''' A dialect-specific SQL expression that replaces <c>NULL</c> values.
    ''' </returns>
    ''' <exception cref="ArgumentOutOfRangeException">
    ''' Thrown when the specified dialect is not supported.
    ''' </exception>
    <Extension>
    Public Function GetIfNull(Dialect As SqlDialect, Expression As String, Replacement As String) As String
        Select Case Dialect
            Case SqlDialect.MySql,
                 SqlDialect.Sqlite
                Return $"IFNULL({Expression}, {Replacement})"
            Case SqlDialect.SqlServer
                Return $"ISNULL({Expression}, {Replacement})"
            Case SqlDialect.PostgreSql,
                 SqlDialect.Oracle,
                 SqlDialect.Firebird
                Return $"COALESCE({Expression}, {Replacement})"
            Case Else
                Throw New ArgumentOutOfRangeException(NameOf(Dialect))
        End Select
    End Function

    ''' <summary>
    ''' Returns the SQL clause used to apply row limiting and paging for the specified SQL dialect.
    ''' </summary>
    ''' <param name="Dialect">The SQL dialect.</param>
    ''' <param name="Limit">The maximum number of rows to return, or <see langword="Nothing"/> to omit the limit.</param>
    ''' <param name="Offset">The number of rows to skip, or <see langword="Nothing"/> to start from the first row.</param>
    ''' <returns>
    ''' A dialect-specific <c>LIMIT</c>, <c>OFFSET</c>, or <c>FETCH</c> clause.
    ''' Returns an empty string when neither <paramref name="Limit"/> nor <paramref name="Offset"/> is specified.
    ''' </returns>
    ''' <exception cref="ArgumentOutOfRangeException">
    ''' Thrown when the specified dialect is not supported.
    ''' </exception>
    <Extension>
    Public Function GetLimitOffset(Dialect As SqlDialect, Limit As Integer?, Offset As Integer?) As String
        If Not Limit.HasValue AndAlso Not Offset.HasValue Then
            Return String.Empty
        End If
        Select Case Dialect
            Case SqlDialect.MySql,
                 SqlDialect.PostgreSql
                Return GetStandardLimitOffset(Limit, Offset)
            Case SqlDialect.Sqlite
                Return GetSQLiteLimitOffset(Limit, Offset)
            Case SqlDialect.SqlServer
                Return GetOffsetFetch(Limit, Offset, "FETCH NEXT")
            Case SqlDialect.Oracle
                Return GetOffsetFetch(Limit, Offset, "FETCH NEXT")
            Case SqlDialect.Firebird
                Return GetOffsetFetch(Limit, Offset, "FETCH FIRST")
            Case Else
                Throw New ArgumentOutOfRangeException(NameOf(Dialect))
        End Select
    End Function

    ''' <summary>
    ''' Builds a standard SQL <c>LIMIT</c> and <c>OFFSET</c> clause.
    ''' </summary>
    ''' <param name="Limit">The maximum number of rows to return.</param>
    ''' <param name="Offset">The number of rows to skip.</param>
    ''' <returns>A standard <c>LIMIT</c>/<c>OFFSET</c> clause.</returns>
    Private Function GetStandardLimitOffset(Limit As Integer?, Offset As Integer?) As String
        Dim Parts As New List(Of String)
        If Limit.HasValue Then
            Parts.Add($"LIMIT {Limit.Value}")
        End If
        If Offset.HasValue Then
            Parts.Add($"OFFSET {Offset.Value}")
        End If
        Return String.Join(" ", Parts)
    End Function

    ''' <summary>
    ''' Builds a SQLite-compatible <c>LIMIT</c> and <c>OFFSET</c> clause.
    ''' </summary>
    ''' <param name="Limit">The maximum number of rows to return.</param>
    ''' <param name="Offset">The number of rows to skip.</param>
    ''' <returns>
    ''' A SQLite-compatible paging clause. When only an offset is specified,
    ''' <c>LIMIT -1</c> is emitted to satisfy SQLite syntax requirements.
    ''' </returns>
    Private Function GetSQLiteLimitOffset(Limit As Integer?, Offset As Integer?) As String
        If Limit.HasValue Then
            Return GetStandardLimitOffset(Limit, Offset)
        End If
        If Offset.HasValue Then
            Return $"LIMIT -1 OFFSET {Offset.Value}"
        End If
        Return String.Empty
    End Function

    ''' <summary>
    ''' Builds an <c>OFFSET</c>/<c>FETCH</c> clause for SQL dialects that support this syntax.
    ''' </summary>
    ''' <param name="Limit">The maximum number of rows to return.</param>
    ''' <param name="Offset">The number of rows to skip.</param>
    ''' <param name="FetchSyntax">
    ''' The dialect-specific fetch keyword, such as <c>FETCH NEXT</c> or <c>FETCH FIRST</c>.
    ''' </param>
    ''' <returns>An <c>OFFSET</c>/<c>FETCH</c> paging clause.</returns>
    Private Function GetOffsetFetch(Limit As Integer?, Offset As Integer?, FetchSyntax As String) As String
        Dim Parts As New List(Of String)
        If Offset.HasValue Then
            Parts.Add($"OFFSET {Offset.Value} ROWS")
        ElseIf Limit.HasValue Then
            Parts.Add("OFFSET 0 ROWS")
        End If
        If Limit.HasValue Then
            Parts.Add($"{FetchSyntax} {Limit.Value} ROWS ONLY")
        End If
        Return String.Join(" ", Parts)
    End Function

End Module