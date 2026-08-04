# CoreSuite MySqlService

`CoreSuite.MySqlService` is a structured .NET 8 service for executing MySQL queries, commands, CRUD operations and stored procedures, with explicit connection ownership, local transaction support, cancellation, database creation, backup and restore.

The service uses `MySql.Data` for database access and `MySqlBackup.NET` for managed SQL export and import operations.

## Features

- Immutable service initialization from individual connection values or a complete connection string.
- Connection strings built with `MySqlConnectionStringBuilder` instead of manual string concatenation.
- Synchronous and asynchronous query, non-query and scalar APIs.
- Cancellation support on asynchronous database operations.
- Automatic connection creation and disposal when no connection is supplied.
- Predictable handling of externally supplied open and closed connections.
- Explicit `MySqlTransaction` support and transaction execution helpers.
- Safely quoted table, schema and column identifiers.
- Parameterized insert, update, delete, filter and raw command values.
- Full-table update and delete protection enabled by default.
- Structured `SELECT` projection, ordering, distinct, limit and offset options.
- Multiple result-set support.
- Duplicate result-column names normalized with numeric suffixes.
- Stored procedure input, output, input/output and return-value parameters.
- Last inserted identifier obtained directly from `MySqlCommand.LastInsertedId`.
- Managed database creation with validated character set and collation values.
- Backup and restore progress reporting.
- Background backup and restore operations that do not block the calling thread.
- Atomic backup file replacement through a temporary file.
- Cleanup of incomplete temporary backup files after failure or cancellation.
- XML documentation and NuGet package metadata.

## Requirements

- .NET 8 or a compatible target framework
- MySQL or a compatible server supported by `MySql.Data`
- Database credentials with the permissions required by the requested operations

## Installation

```powershell
dotnet add package CoreSuite.MySqlService
```

## Create the service

Create the service from separate connection values:

```vbnet
Dim mySqlService As New MySqlService(
    "localhost",
    "sample_database",
    "root",
    "password")
```

Additional connection-string settings can be configured through `MySqlConnectionStringBuilder`:

```vbnet
Dim mySqlService As New MySqlService(
    "localhost",
    "sample_database",
    "root",
    "password",
    Sub(builder)
        builder.Port = 3306
        builder.Pooling = True
        builder.ConnectionTimeout = 15
    End Sub)
```

A complete connection string can also be supplied:

```vbnet
Dim connectionString As String =
    "Server=localhost;Database=sample_database;User ID=root;Password=password;Pooling=True;"
Dim mySqlService As New MySqlService(connectionString)
```

The connection string must contain both a server and a database.

## Execute a query

Use `ExecuteQueryAsync` when the SQL is expected to return rows:

```vbnet
Dim queryArgs As New Dictionary(Of String, Object) From {
    {"@minimumId", 10}
}
Dim response As MySqlResponse = Await mySqlService.Request.ExecuteQueryAsync(
    "SELECT id, name FROM customers WHERE id >= @minimumId ORDER BY name",
    queryArgs)
For Each row As IReadOnlyDictionary(Of String, Object) In response.Data
    Dim customerId As Object = row("id")
    Dim customerName As Object = row("name")
Next row
```

`Data` always returns a collection. It is empty when the operation returns no rows.

Database `NULL` values are returned as `Nothing`.

## Execute a non-query command

Use `ExecuteNonQueryAsync` for commands that do not return rows:

```vbnet
Dim queryArgs As New Dictionary(Of String, Object) From {
    {"@isActive", True},
    {"@customerId", 25}
}
Dim response As MySqlResponse = Await mySqlService.Request.ExecuteNonQueryAsync(
    "UPDATE customers SET is_active = @isActive WHERE id = @customerId",
    queryArgs)
Debug.WriteLine($"Affected rows: {response.AffectedRows}")
```

## Execute a scalar command

```vbnet
Dim customerCount As Object = Await mySqlService.Request.ExecuteScalarAsync(
    "SELECT COUNT(*) FROM customers")
```

The result is `Nothing` when the command returns no value or returns database `NULL`.

## Insert a row

Table and column identifiers are quoted by the service, while values are always parameterized:

```vbnet
Dim values As New Dictionary(Of String, Object) From {
    {"name", "CoreSuite"},
    {"email", "contact@example.com"},
    {"is_active", True}
}
Dim response As MySqlResponse = Await mySqlService.Request.ExecuteInsertAsync(
    "customers",
    values)
Debug.WriteLine($"Inserted ID: {response.LastInsertedId}")
```

## Update rows

A `WHERE` expression is required by default:

```vbnet
Dim values As New Dictionary(Of String, Object) From {
    {"name", "Updated customer"}
}
Dim options As New MySqlMutationOptions With {
    .Where = "id = @customerId",
    .QueryArgs = New Dictionary(Of String, Object) From {
        {"@customerId", 25}
    }
}
Dim response As MySqlResponse = Await mySqlService.Request.ExecuteUpdateAsync(
    "customers",
    values,
    options)
```

An intentional full-table update must be explicit:

```vbnet
Dim options As New MySqlMutationOptions With {
    .AllowAllRows = True
}
```

## Delete rows

```vbnet
Dim options As New MySqlMutationOptions With {
    .Where = "id = @customerId",
    .QueryArgs = New Dictionary(Of String, Object) From {
        {"@customerId", 25}
    }
}
Dim response As MySqlResponse = Await mySqlService.Request.ExecuteDeleteAsync(
    "customers",
    options)
```

A delete without `Where` throws `InvalidOperationException` unless `AllowAllRows` is `True`.

## Structured SELECT

`ExecuteSelectAsync` safely quotes the table, selected columns and ordering columns:

```vbnet
Dim options As New MySqlSelectOptions With {
    .Where = "is_active = @isActive",
    .QueryArgs = New Dictionary(Of String, Object) From {
        {"@isActive", True}
    },
    .Distinct = True,
    .Limit = 50,
    .Offset = 0
}
options.Columns.Add("id")
options.Columns.Add("name")
options.OrderBy.Add(New MySqlOrderBy("name"))
Dim response As MySqlResponse = Await mySqlService.Request.ExecuteSelectAsync(
    "customers",
    options)
```

When both `Columns` and `TrustedExpressions` are empty, `SELECT *` is used.

Complex trusted expressions can be added explicitly:

```vbnet
options.TrustedExpressions.Add("COUNT(*) AS total")
```

Only application-controlled SQL should be added to `TrustedExpressions` or `Where`.

## Multiple result sets

Queries and stored procedures may return more than one result set:

```vbnet
Dim response As MySqlResponse = Await mySqlService.Request.ExecuteQueryAsync(
    "SELECT * FROM customers; SELECT * FROM orders;")
For Each resultSet As MySqlResultSet In response.ResultSets
    For Each row As IReadOnlyDictionary(Of String, Object) In resultSet.Rows
        Debug.WriteLine(String.Join(", ", row.Values))
    Next row
Next resultSet
```

When duplicate column names are returned, later occurrences receive suffixes such as `_2`, `_3` and `_4`.

## Stored procedures

```vbnet
Dim parameters As New List(Of MySqlProcedureParameter) From {
    New MySqlProcedureParameter("@customerId", 25),
    New MySqlProcedureParameter("@customerName") With {
        .Direction = ParameterDirection.Output,
        .MySqlDbType = MySqlDbType.VarChar,
        .Size = 200
    }
}
Dim response As MySqlResponse = Await mySqlService.Request.ExecuteProcedureAsync(
    "get_customer",
    parameters)
Dim customerName As Object = response.OutputParameters("@customerName")
```

Every result set is available through `ResultSets`. Output values are populated after the reader is closed.

## Local transactions

Create the connection through `MySqlClient`, begin a local transaction and pass both objects through command options:

```vbnet
Using connection As MySqlConnection = mySqlService.Client.CreateDatabaseConnection()
    Await connection.OpenAsync()
    Using transaction As MySqlTransaction = connection.BeginTransaction()
        Dim commandOptions As New MySqlCommandOptions With {
            .Connection = connection,
            .Transaction = transaction
        }
        Dim firstValues As New Dictionary(Of String, Object) From {
            {"name", "First customer"}
        }
        Dim secondValues As New Dictionary(Of String, Object) From {
            {"name", "Second customer"}
        }
        Await mySqlService.Request.ExecuteInsertAsync("customers", firstValues, commandOptions)
        Await mySqlService.Request.ExecuteInsertAsync("customers", secondValues, commandOptions)
        Await transaction.CommitAsync()
    End Using
End Using
```

The transaction connection is used automatically when `Transaction` is supplied without `Connection`.

If both are supplied, they must refer to the same connection.

For a complete transaction managed by the service, use `ExecuteInTransactionAsync`:

```vbnet
Await mySqlService.ExecuteInTransactionAsync(
    Async Function(transaction, cancellationToken)
        Dim commandOptions As New MySqlCommandOptions With {
            .Transaction = transaction
        }
        Await mySqlService.Request.ExecuteInsertAsync(
            "customers",
            New Dictionary(Of String, Object) From {{"name", "First customer"}},
            commandOptions,
            cancellationToken)
        Await mySqlService.Request.ExecuteInsertAsync(
            "customers",
            New Dictionary(Of String, Object) From {{"name", "Second customer"}},
            commandOptions,
            cancellationToken)
    End Function)
```

The helper commits after successful completion and attempts to roll back before rethrowing any exception. Synchronous `Action` and result-returning overloads are also available.

## Connection ownership

The request service follows these rules:

- When no connection is supplied, it creates, opens and disposes an internal connection.
- When an external connection is already open, it remains open after the operation.
- When an external connection is closed, it is opened for the operation and closed afterward without being disposed.
- A supplied transaction must belong to the supplied connection.
- Connection settings stored by `MySqlService` are immutable after construction.

## Cancellation

Every asynchronous request method accepts `CancellationToken` as its last parameter:

```vbnet
Using cancellationSource As New CancellationTokenSource(TimeSpan.FromSeconds(10))
    Dim response As MySqlResponse = Await mySqlService.Request.ExecuteQueryAsync(
        "SELECT * FROM large_table",
        cancellationToken:=cancellationSource.Token)
End Using
```

The token is passed to connection opening, command execution, reader iteration and result-set transitions.

Backup and restore use the token to interrupt the managed operation by canceling pending work and closing their internally owned connection. Cancellation may be observed at the next provider or backup-library I/O boundary.

## Create the database

The server-level connection does not select the configured database, allowing the database to be created before it exists:

```vbnet
Dim options As New MySqlCreateDatabaseOptions With {
    .IfNotExists = True,
    .CharacterSet = "utf8mb4",
    .Collation = "utf8mb4_unicode_ci"
}
Await mySqlService.Maintenance.ExecuteCreateDatabaseAsync(options)
```

Character set and collation values are validated as single SQL tokens before being included in the command.

## Backup

```vbnet
Dim progress As New Progress(Of Integer)(
    Sub(value) Debug.WriteLine($"Backup: {value}%"))
Dim options As New MySqlBackupOptions With {
    .Progress = progress,
    .Overwrite = True,
    .ExportProcedures = True,
    .ExportFunctions = True,
    .ExportTriggers = True
}
Await mySqlService.Maintenance.ExecuteBackupAsync(
    "C:\Backups\sample_database.sql",
    options)
```

Every backup is written to a temporary file in the destination directory. The destination is replaced only after a successful export, so a failed or canceled backup does not replace an existing valid file.

## Restore

```vbnet
Dim progress As New Progress(Of Integer)(
    Sub(value) Debug.WriteLine($"Restore: {value}%"))
Dim options As New MySqlRestoreOptions With {
    .Progress = progress
}
Await mySqlService.Maintenance.ExecuteRestoreAsync(
    "C:\Backups\sample_database.sql",
    options)
```

The backup file must exist before restore begins.

## Progress events

```vbnet
AddHandler mySqlService.Maintenance.BackupProgressChanged,
    Sub(sender, eventArgs)
        Debug.WriteLine($"Backup: {eventArgs.ProgressPercentage}%")
    End Sub
AddHandler mySqlService.Maintenance.RestoreProgressChanged,
    Sub(sender, eventArgs)
        Debug.WriteLine($"Restore: {eventArgs.ProgressPercentage}%")
    End Sub
```

Asynchronous backup and restore events are raised from a worker thread. Marshal UI updates to the UI thread when required. `Progress(Of Integer)` normally captures the current synchronization context and is preferable for Windows Forms interfaces.

## Main classes

| Class | Purpose |
| --- | --- |
| `MySqlService` | Immutable entry point exposing client, request and maintenance services. |
| `MySqlClient` | Creates database-level and server-level `MySqlConnection` instances. |
| `MySqlRequest` | Executes queries, commands, scalar operations, CRUD operations and stored procedures. |
| `MySqlMaintenance` | Creates the database and performs backup and restore operations. |
| `MySqlCommandOptions` | Defines an optional connection, transaction and command timeout. |
| `MySqlSelectOptions` | Defines projection, filtering, sorting, distinct and paging behavior. |
| `MySqlMutationOptions` | Defines mutation filtering and full-table safety behavior. |
| `MySqlProcedureParameter` | Defines input, output, input/output and return-value procedure parameters. |
| `MySqlResponse` | Contains result sets, rows affected, inserted ID and output parameter values. |
| `MySqlResultSet` | Contains unique column names and read-only rows for one result set. |
| `MySqlBackupOptions` | Configures export content, progress, overwrite behavior and command timeout. |
| `MySqlRestoreOptions` | Configures restore progress and command timeout. |
| `MySqlCreateDatabaseOptions` | Configures database creation behavior, character set and collation. |

## Main request methods

| Method | Description |
| --- | --- |
| `ExecuteQuery` / `ExecuteQueryAsync` | Executes SQL expected to return result sets. |
| `ExecuteNonQuery` / `ExecuteNonQueryAsync` | Executes SQL expected to affect rows without returning result sets. |
| `ExecuteScalar` / `ExecuteScalarAsync` | Returns the first column of the first row. |
| `ExecuteProcedure` / `ExecuteProcedureAsync` | Executes a stored procedure and captures all result sets and output values. |
| `ExecuteSelect` / `ExecuteSelectAsync` | Builds a structured SELECT against a safely quoted table. |
| `ExecuteInsert` / `ExecuteInsertAsync` | Inserts one parameterized row and returns the provider-generated identifier. |
| `ExecuteUpdate` / `ExecuteUpdateAsync` | Updates parameterized values with full-table protection. |
| `ExecuteDelete` / `ExecuteDeleteAsync` | Deletes rows with full-table protection. |

## Security and SQL behavior

- Table, schema and column identifiers supplied to CRUD and structured SELECT methods are escaped with MySQL backticks.
- Identifier backticks are doubled before quoting.
- Insert and update values are assigned generated parameter names that do not depend on column names.
- Dictionary parameter names are validated and normalized to begin with `@`.
- `Where`, `TrustedExpressions` and SQL supplied to query, non-query and scalar methods are trusted SQL surfaces.
- Never concatenate user input into SQL text or trusted fragments.
- Use parameters for every external value.
- Charset and collation names are validated because MySQL does not allow them to be supplied as ordinary command parameters.

## Important behavior

- `MySqlService` has no parameterless constructor and cannot exist in a partially initialized state.
- Reusing an options object does not store or replace its connection.
- `MySqlResponse.Data` and `ResultSets` are never `Nothing`.
- `LastInsertedId` is nullable and is populated by insert methods.
- `Offset` requires `Limit` in structured SELECT operations.
- A command timeout of zero uses the provider convention for no timeout; negative values are rejected.
- Backup and restore use internally owned connections and are not part of a local transaction.
- Every backup uses atomic replacement to protect an existing valid file from partial export results.
- Incomplete temporary backup files are removed after failure or cancellation.
- The permissions required for backup and restore depend on the objects included in the export and the target MySQL server configuration.

## License

MIT License.

## Repository

[CoreSuite on GitHub](https://github.com/leandrogaldino/CoreSuite)
