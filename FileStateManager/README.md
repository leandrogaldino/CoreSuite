# CoreSuite.FileStateManager

**Track a file's original and pending state, then apply add, replace or remove operations through a transactional file manager.**

> [!NOTE]
> `CoreSuite.FileStateManager` is one of the libraries included in the **CoreSuite** solution. It targets plain .NET 8, works independently of Windows Forms and uses the `TxFileManager` package for file-system transaction enlistment.

## Overview

`CoreSuite.FileStateManager` represents one file-valued application field, such as a customer document, product image, invoice attachment or generated report.

The service remembers two paths:

- `OriginalFile` is the file currently associated with the application record.
- `CurrentFile` is the file that should be associated with that record after pending changes are applied.

Calling `Execute` compares those values and performs the required operation:

- copy a newly selected file into the target directory;
- delete an existing file when the selection is cleared;
- replace the original file when a different file is selected;
- do nothing when the state has not changed.

The actual copy and delete calls are delegated to [`TxFileManager`](https://github.com/chinhdo/txFileManager), allowing them to enlist in an ambient `System.Transactions.TransactionScope`.

## Features

- Tracks original and pending file paths separately.
- Represents add, replace, remove and unchanged states.
- Normalizes empty or whitespace-only selections to `Nothing`.
- Copies selected files into a configured target directory.
- Preserves the selected file name at the destination.
- Deletes an original file when the pending selection is cleared.
- Replaces an original file with a newly selected file.
- Avoids file-system work when the original and current paths are equal.
- Uses `TxFileManager` for ambient transaction participation.
- Supports a shared custom temporary directory for `TxFileManager`.
- Exposes tracked paths through read-only properties.
- Includes English XML documentation for the public API.

## Requirements

- .NET 8 (`net8.0`)
- `TxFileManager` 1.5.0.1, installed automatically by NuGet
- An existing writable target directory

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.FileStateManager
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.FileStateManager
```

## Namespace

Import the service namespace:

```vb
Imports CoreSuite.Services
```

To use ambient transactions, also import:

```vb
Imports System.Transactions
```

## Understand the state model

| `OriginalFile` | `CurrentFile` | `Execute` behavior |
| --- | --- | --- |
| `Nothing` | `Nothing` | Does nothing. |
| `Nothing` | A file path | Copies the current file to `TargetDirectory`. |
| A file path | `Nothing` | Deletes the original file. |
| A file path | The same path | Does nothing. |
| A file path | A different path | Deletes the original and copies the current file to `TargetDirectory`. |

The destination path is always:

```vb
Path.Combine(manager.TargetDirectory, Path.GetFileName(manager.CurrentFile))
```

The source file is not moved; it is copied.

## Track an existing file

When loading an application record that already has an associated file, set it as both current and original:

```vb
Dim manager As New FileStateManager("D:\ApplicationFiles\Customers")
manager.SetCurrentFile("D:\ApplicationFiles\Customers\contract.pdf", True)
```

After this call:

- `OriginalFile` is the existing contract path;
- `CurrentFile` is the same path;
- calling `Execute` without another change does nothing.

## Add a new file

For a record that does not yet have a file:

```vb
Dim manager As New FileStateManager("D:\ApplicationFiles\Customers")
manager.SetCurrentFile("C:\Users\Public\Documents\new-contract.pdf")
manager.Execute()
```

The source is copied to:

```text
D:\ApplicationFiles\Customers\new-contract.pdf
```

After a successful call, both tracked paths point to the destination file.

## Replace an existing file

Load the original state, then set the newly selected source file:

```vb
Dim manager As New FileStateManager("D:\ApplicationFiles\Customers")
manager.SetCurrentFile("D:\ApplicationFiles\Customers\old-contract.pdf", True)
manager.SetCurrentFile("C:\Users\Public\Documents\replacement-contract.pdf")
manager.Execute()
```

`Execute` deletes `old-contract.pdf` and copies `replacement-contract.pdf` into the target directory.

## Remove an existing file

Clear the current selection by passing `Nothing`, an empty string or whitespace:

```vb
Dim manager As New FileStateManager("D:\ApplicationFiles\Customers")
manager.SetCurrentFile("D:\ApplicationFiles\Customers\contract.pdf", True)
manager.SetCurrentFile(Nothing)
manager.Execute()
```

The original file is deleted and both tracked values become `Nothing`.

## Use an ambient transaction

`FileStateManager` does not create a `TransactionScope` by itself. Wrap `Execute` and related transactional work in the same scope when they must commit or roll back together:

```vb
Dim manager As New FileStateManager("D:\ApplicationFiles\Customers")
manager.SetCurrentFile(existingFilePath, True)
manager.SetCurrentFile(selectedReplacementPath)

Using scope As New TransactionScope()
    manager.Execute()

    SaveCustomerFilePathToDatabase(manager.CurrentFile)

    scope.Complete()
End Using
```

If `scope.Complete()` is not called, enlisted `TxFileManager` operations are rolled back when the scope is disposed.

> [!IMPORTANT]
> `Execute` updates the manager's in-memory `OriginalFile` and `CurrentFile` values before the surrounding transaction outcome is known. If the ambient transaction rolls back, discard or reinitialize that manager instance so its tracked state matches the file system again.

Without an ambient transaction, the delegated file operations are applied normally and there is no application-level rollback boundary coordinating them with database work.

## Configure the temporary directory

`TempDirectory` is shared by every `FileStateManager` instance:

```vb
FileStateManager.TempDirectory = "D:\ApplicationTemp\FileTransactions"
```

When the value is empty or `Nothing`, `TxFileManager` uses its default temporary location. Set this property once during application startup if a custom location is required.

The application identity must be able to read, write and delete content in the selected temporary directory.

## Properties

| Property | Type | Access | Description |
| --- | --- | --- | --- |
| `TempDirectory` | `String` | Shared read/write | Optional temporary directory passed to new `TxFileManager` instances. |
| `TargetDirectory` | `String` | Read-only | Directory that receives copied current files. |
| `OriginalFile` | `String` | Read-only | File path considered to exist before pending changes. |
| `CurrentFile` | `String` | Read-only | File path representing the desired pending state. |

## Methods

| Method | Description |
| --- | --- |
| `New(targetDirectory)` | Creates a manager for one destination directory. |
| `SetCurrentFile(filename, asOriginal)` | Changes the pending file and optionally establishes it as the original. |
| `Execute()` | Applies the add, replace, remove or unchanged state through `TxFileManager`. |
| `Clone()` | Creates another manager from the tracked target and file paths. See the limitation below. |

## Current `Clone` limitation

The current implementation reapplies `OriginalFile` last while constructing the clone. When the source instance has a pending replacement and `CurrentFile` differs from `OriginalFile`, the cloned instance ends with the original path as its current path.

Do not use `Clone` to preserve a pending replacement in version `1.0.0`. Create a new manager and reapply the desired state explicitly when that scenario matters:

```vb
Dim copy As New FileStateManager(manager.TargetDirectory)
copy.SetCurrentFile(manager.OriginalFile, True)
copy.SetCurrentFile(manager.CurrentFile)
```

## File-system behavior

- `TargetDirectory` is not created by `FileStateManager`; create it before calling `Execute`.
- The destination uses only the source file name, not its source directory tree.
- Copy operations pass overwrite as `False`. An existing destination file with the same name can therefore cause an exception.
- Path equality uses the string values tracked by the manager; the class does not normalize paths before comparing them.
- `Execute` is synchronous.
- The service manages one file state per instance.

## Error handling

File-system and transaction errors are allowed to propagate to the caller. Wrap the surrounding transaction or application operation in normal exception handling:

```vb
Try
    Using scope As New TransactionScope()
        manager.Execute()
        SaveCustomerFilePathToDatabase(manager.CurrentFile)
        scope.Complete()
    End Using
Catch ex As IOException
    Console.WriteLine(ex.Message)
Catch ex As UnauthorizedAccessException
    Console.WriteLine(ex.Message)
End Try
```

Common failure causes include a missing source file, missing target directory, destination name collision, insufficient permissions or unavailable temporary storage.

## Recommended usage pattern

1. Create one manager for the record's file field.
2. Call `SetCurrentFile(existingPath, True)` when loading an existing value.
3. Call `SetCurrentFile(selectedPath)` when the user selects a replacement.
4. Call `SetCurrentFile(Nothing)` when the user removes the selection.
5. Start a `TransactionScope` if file and database changes must share a transaction.
6. Call `Execute` only when saving the record.
7. Persist `CurrentFile` after `Execute` succeeds.
8. Complete the ambient transaction.

## License

This package is licensed under the MIT License.
