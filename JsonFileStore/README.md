# JsonFileStore

**A strongly typed JSON file storage service with atomic writes, backups, and recovery for .NET 8, included in CoreSuite.**

> [!NOTE]
> JsonFileStore is one of the independent projects that make up the **CoreSuite** solution. The package contains the generic store, its recovery exception, and the infrastructure required to persist JSON files safely.

## Overview

`JsonFileStore(Of T)` provides a small persistence layer for application settings, local state, caches, and other non-sensitive data. It serializes a value with `System.Text.Json`, writes the new content to a temporary file, and only then replaces the primary file.

The store can maintain a `.bak` copy of the last valid content and automatically recover when the primary file is missing, incomplete, or contains invalid JSON. Synchronous, asynchronous, exception-based, Boolean, and fallback-based APIs are available so each application can choose how storage failures should be handled.

## Key features

- Provides a strongly typed `JsonFileStore(Of T)` API.
- Saves and loads values synchronously or asynchronously.
- Writes through a temporary file in the destination directory.
- Replaces the primary file atomically after serialization succeeds.
- Maintains an automatic `.bak` backup.
- Preserves an older valid backup when the current primary file is invalid.
- Recovers the primary file automatically from a valid backup.
- Exposes `TryLoad` for Boolean-based error handling.
- Exposes `LoadOrDefault` and `LoadOrDefaultAsync` for fallback values.
- Accepts custom `System.Text.Json.JsonSerializerOptions`.
- Supports a custom backup path.
- Serializes operations performed through the same store instance.
- Has no external package dependencies.

## Requirements

- .NET 8 or a compatible target framework
- A reference to `CoreSuite.JsonFileStore`

The service has no runtime dependency on another CoreSuite package.

## Installation

```powershell
dotnet add package CoreSuite.JsonFileStore
```

Or add `JsonFileStore/JsonFileStore.vbproj` as a project reference when working directly with the CoreSuite solution.

## Quick start

Create the type that will be persisted, choose the primary JSON path, and instantiate `JsonFileStore(Of T)`.

```vb
Imports System.IO
Imports CoreSuite.Services
Public Class AppSettings
    Public Property Theme As String
    Public Property PageSize As Integer
End Class
Dim filePath As String = Path.Combine(AppContext.BaseDirectory, "settings.json")
Dim store As New JsonFileStore(Of AppSettings)(filePath)
Dim settings As New AppSettings With {
    .Theme = "Dark",
    .PageSize = 50
}
store.Save(settings)
Dim loadedSettings As AppSettings = store.Load()
```

The first successful save creates the primary file and its backup. With the default path, the example produces `settings.json` and `settings.json.bak`. On later saves, the backup retains the preceding valid primary value.

## Asynchronous operations

Use `SaveAsync` and `LoadAsync` when persistence should not block the calling thread.

```vb
Dim store As New JsonFileStore(Of AppSettings)(filePath)
Await store.SaveAsync(settings)
Dim loadedSettings As AppSettings = Await store.LoadAsync()
```

Both methods accept an optional `CancellationToken`:

```vb
Dim loadedSettings As AppSettings = Await store.LoadAsync(cancellationToken)
Await store.SaveAsync(loadedSettings, cancellationToken)
```

Cancellation is honored while waiting for access to the store and during asynchronous file operations.

## Atomic writes and backups

`Save` and `SaveAsync` never serialize directly over the primary file. The store writes the complete JSON payload to a temporary file beside its destination before committing the result.

For an existing valid primary file, a successful save follows this logical sequence:

1. serialize the new value to a temporary file;
2. retain the existing valid primary content as the backup;
3. atomically replace the primary file with the completed temporary file;
4. remove any temporary artifacts left by the operation.

Keeping temporary and destination files on the same file system allows the final replacement to use the platform's atomic file operation. If the existing primary file is invalid, the store preserves an older valid backup instead of replacing it with corrupted content.

## Automatic recovery

`Load` and `LoadAsync` read the primary file first. If that operation fails and `AutoRecoverFromBackup` is enabled, the store:

1. loads and deserializes the backup;
2. atomically restores that valid backup to the primary path;
3. returns the recovered value.

```vb
Dim store As New JsonFileStore(Of AppSettings)(filePath)
Dim settings As AppSettings = store.Load()
```

If both attempts fail, `JsonFileRecoveryException` exposes the primary and recovery failures separately:

```vb
Try
    Dim settings As AppSettings = store.Load()
Catch ex As JsonFileRecoveryException
    Debug.WriteLine($"Primary: {ex.PrimaryException.Message}")
    Debug.WriteLine($"Recovery: {ex.RecoveryException.Message}")
End Try
```

Disable recovery when the application must receive the original primary-file error directly:

```vb
store.AutoRecoverFromBackup = False
```

## Safe fallback APIs

Use `TryLoad` when a load failure should be represented by a Boolean instead of an exception.

```vb
Dim settings As AppSettings = Nothing
If Not store.TryLoad(settings) Then
    settings = New AppSettings()
End If
```

Use `LoadOrDefault` when the fallback value is already available:

```vb
Dim defaults As New AppSettings With {
    .Theme = "Light",
    .PageSize = 25
}
Dim settings As AppSettings = store.LoadOrDefault(defaults)
```

The asynchronous equivalent is `LoadOrDefaultAsync`:

```vb
Dim settings As AppSettings = Await store.LoadOrDefaultAsync(defaults, cancellationToken)
```

These APIs still use automatic backup recovery when it is enabled. The fallback is returned only when loading and any permitted recovery cannot provide a value.

## Serializer configuration

Pass `JsonSerializerOptions` to the constructor to configure names, formatting, converters, and other `System.Text.Json` behavior.

```vb
Imports System.Text.Json
Imports System.Text.Json.Serialization
Dim options As New JsonSerializerOptions With {
    .WriteIndented = True,
    .PropertyNameCaseInsensitive = True,
    .PropertyNamingPolicy = JsonNamingPolicy.CamelCase
}
options.Converters.Add(New JsonStringEnumConverter())
Dim store As New JsonFileStore(Of AppSettings)(filePath, options)
```

The constructor creates a copy of the supplied options, so later changes to the original `JsonSerializerOptions` instance do not alter the store.

`SerializerOptions` can also be configured through the store before its first save or load:

```vb
store.SerializerOptions.WriteIndented = True
store.SerializerOptions.PropertyNameCaseInsensitive = True
```

`System.Text.Json` makes an options instance read-only after it has been used. Complete the serializer configuration before the first persistence operation.

## Custom backup path

By default, the backup path is the primary path followed by `.bak`. Supply `backupPath` when backups belong in another directory or require a different filename.

```vb
Dim backupPath As String = Path.Combine(AppContext.BaseDirectory, "Backups", "settings.json.bak")
Dim store As New JsonFileStore(Of AppSettings)(filePath, backupPath:=backupPath)
```

The store creates the required destination directories when saving or restoring files.

## Main properties

| Property | Default | Description |
|---|---:|---|
| `FilePath` | Constructor value | Absolute path of the primary JSON file. |
| `BackupPath` | `FilePath & ".bak"` | Absolute path of the backup file. |
| `SerializerOptions` | Standard options | Serializer configuration owned by the store. |
| `AutoRecoverFromBackup` | `True` | Enables backup loading and primary-file restoration. |
| `Exists` | Read-only | Indicates whether the primary file exists. |
| `BackupExists` | Read-only | Indicates whether the backup file exists. |

## Methods

| Method | Description |
|---|---|
| `Save(value)` | Serializes and atomically saves a value. |
| `SaveAsync(value, cancellationToken)` | Asynchronously serializes and saves a value. |
| `Load()` | Loads the primary value or recovers it from the backup. |
| `LoadAsync(cancellationToken)` | Asynchronously loads or recovers a value. |
| `TryLoad(ByRef value)` | Returns `False` instead of propagating expected storage or JSON failures. |
| `LoadOrDefault(defaultValue)` | Returns a fallback value when load and recovery fail. |
| `LoadOrDefaultAsync(defaultValue, cancellationToken)` | Asynchronously loads a value or returns the supplied fallback. |

## Recovery errors

| Member | Description |
|---|---|
| `JsonFileRecoveryException.PrimaryException` | Error produced while reading or deserializing the primary file. |
| `JsonFileRecoveryException.RecoveryException` | Error produced while loading or restoring the backup. |

Catch `JsonFileRecoveryException` when the application needs to report or log both failures. Use `TryLoad` or a fallback-based method when those details are not required by the calling code.

## Concurrency behavior

Operations performed through the same `JsonFileStore(Of T)` instance are serialized. This prevents overlapping saves and loads from that instance from modifying the same files simultaneously.

Separate store instances and separate processes are not coordinated. Applications that share one JSON path across instances or processes must provide an external synchronization strategy.

## Storage notes

- Use the service for settings, local state, caches, and other non-sensitive application data.
- Temporary files are created beside their destination so the final replacement remains on the same file system.
- A newly serialized value is never written directly over the primary file.
- A valid existing primary file becomes the backup before a new value is committed.
- An invalid existing primary file does not replace an older valid backup.
- Automatic recovery restores the backup to the primary path before returning the recovered value.
- `Exists` and `BackupExists` report physical file presence and do not validate JSON content.
- JSON and backup files are not encrypted. Do not store passwords, cryptographic keys, access tokens, or other secrets in plain text.

## License

CoreSuite is licensed under the MIT License.
