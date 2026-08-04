# CoreSuite.FileManager

**Asynchronous file and directory copy and deletion with progress reporting and cancellation.**

> [!NOTE]
> `CoreSuite.FileManager` is one of the libraries included in the **CoreSuite** solution. It targets plain .NET 8, works independently of Windows Forms and has no CoreSuite or third-party package dependencies.

## Overview

`CoreSuite.FileManager` provides a focused service for copying and deleting files or complete directory trees without blocking the calling thread during expensive file-system work.

Operations support `CancellationToken`, normalize paths, validate unsafe source and destination combinations and report byte and item progress through dedicated events. Directory traversal deliberately skips reparse points, preventing the service from following symbolic links, junctions and other redirected directory structures.

The package also includes a helper for clearing a directory while preserving selected files and subdirectories.

## Features

- Copies individual files asynchronously.
- Copies one or several directory trees asynchronously.
- Creates missing destination directories automatically during copy operations.
- Preserves empty directories when copying a directory tree.
- Overwrites existing destination files.
- Deletes individual files asynchronously.
- Deletes one or several directory trees asynchronously.
- Optionally preserves the root directory while deleting its contents.
- Clears a directory while preserving selected files and subdirectories.
- Reports progress by bytes and completed file count.
- Captures the caller's synchronization context for progress events.
- Supports cancellation across enumeration, copy and deletion work.
- Removes duplicate file paths from multi-file deletion requests.
- Rejects directory copies whose destination is equal to or inside the source.
- Rejects overlapping directory deletion roots.
- Skips reparse points during recursive enumeration.
- Uses operating-system-appropriate path comparison rules.
- Includes English XML documentation for the public API.
- Requires no third-party packages.

## Requirements

- .NET 8 (`net8.0`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.FileManager
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.FileManager
```

## Namespace

Import the service namespace and `System.IO`:

```vb
Imports System.IO
Imports CoreSuite.Services
```

## Quick start

Copy one file and receive progress updates:

```vb
Dim manager As New FileManager()

AddHandler manager.CopyFileProgressChanged,
    Sub(sender, e)
        Console.WriteLine($"{e.PercentCompleted}% - {e.CurrentPath}")
    End Sub

Dim sourceFile As New FileInfo("C:\Data\large-export.zip")
Dim destinationFile As New FileInfo("D:\Backup\large-export.zip")

Await manager.CopyFileAsync(sourceFile, destinationFile)
```

The destination parent directory is created when necessary. An existing destination file is overwritten.

## Copy a file with cancellation

```vb
Using cancellationSource As New CancellationTokenSource()
    Try
        Await manager.CopyFileAsync(sourceFile, destinationFile, cancellationSource.Token)
    Catch ex As OperationCanceledException
        Console.WriteLine("Copy canceled.")
    End Try
End Using
```

Cancellation does not roll back bytes already written. If the operation is canceled, the destination may contain a partial file and your application should decide whether to keep or remove it.

## Copy a directory

Create a `CopyDirectoryInfo` containing the source and destination:

```vb
Dim manager As New FileManager()
Dim copyRequest As New CopyDirectoryInfo(
    New DirectoryInfo("C:\ApplicationData"),
    New DirectoryInfo("D:\Backup\ApplicationData"))

AddHandler manager.CopyDirectoryProgressChanged,
    Sub(sender, e)
        Console.WriteLine($"{e.PercentCompleted}% - {e.ProcessedItems}/{e.TotalItems}")
    End Sub

Dim copiedBytes As Long = Await manager.CopyDirectoryAsync(copyRequest)
```

The complete source tree is planned before copying starts. This allows the service to calculate the total number of bytes and files. Empty directories are created at the destination.

### Participate in a larger progress calculation

`CopyDirectoryAsync` accepts optional `totalSize` and `handledSize` values when its progress should be combined with other work:

```vb
Dim alreadyCopied As Long = 500000
Dim completeOperationSize As Long = 2000000
Dim copiedNow As Long = Await manager.CopyDirectoryAsync(copyRequest, completeOperationSize, alreadyCopied, cancellationToken)
```

The returned value is the number of bytes copied by the supplied directory request, not the accumulated `handledSize`.

## Copy multiple directories

```vb
Dim requests As CopyDirectoryInfo() = {
    New CopyDirectoryInfo(New DirectoryInfo("C:\Data\Customers"), New DirectoryInfo("D:\Backup\Customers")),
    New CopyDirectoryInfo(New DirectoryInfo("C:\Data\Orders"), New DirectoryInfo("D:\Backup\Orders"))
}

Await manager.CopyDirectoriesAsync(requests, cancellationToken)
```

Progress is calculated across all planned files in all requests.

## Delete files

```vb
Dim filesToDelete As FileInfo() = {
    New FileInfo("C:\Temp\result-1.tmp"),
    New FileInfo("C:\Temp\result-2.tmp")
}

AddHandler manager.DeleteFilesProgressChanged,
    Sub(sender, e)
        Console.WriteLine($"Deleted {e.ProcessedItems} of {e.TotalItems} files.")
    End Sub

Await manager.DeleteFilesAsync(filesToDelete, cancellationToken)
```

Duplicate normalized paths are processed only once. Every requested file must exist when the deletion plan is built; otherwise, `FileNotFoundException` is thrown.

## Delete directories

Use `DeleteDirectoryInfo.DeleteRoot` to choose whether the root is removed after its contents:

```vb
Dim requests As DeleteDirectoryInfo() = {
    New DeleteDirectoryInfo(New DirectoryInfo("C:\Temp\OldImports"), True),
    New DeleteDirectoryInfo(New DirectoryInfo("C:\Temp\Working"), False)
}

AddHandler manager.DeleteDirectoriesProgressChanged,
    Sub(sender, e)
        Console.WriteLine($"{e.PercentCompleted}%")
    End Sub

Await manager.DeleteDirectoriesAsync(requests, cancellationToken)
```

The first request deletes `OldImports` and its contents. The second deletes everything inside `Working` but preserves the `Working` root directory.

Missing directory roots are ignored. Existing roots in the same request may not be equal, nested or otherwise overlap.

## Clear a directory with exclusions

`DeleteDirectoryContentAsync` removes content without deleting the supplied root. It can preserve complete subtrees and individual files:

```vb
Dim rootDirectory As New DirectoryInfo("C:\Application\Cache")
Dim directoriesToKeep As DirectoryInfo() = {
    New DirectoryInfo("C:\Application\Cache\Pinned")
}
Dim filesToKeep As FileInfo() = {
    New FileInfo("C:\Application\Cache\settings.json")
}

Await FileManager.DeleteDirectoryContentAsync(rootDirectory, directoriesToKeep, filesToKeep, cancellationToken)
```

When a directory is excluded, all of its descendants are preserved. The service also preserves the required ancestors of excluded items so that those items remain reachable.

Every exclusion must be located inside the supplied root. Passing the root itself in `exceptDirectories` makes the method return without deleting anything.

> [!NOTE]
> `DeleteDirectoryContentAsync` is shared and does not raise one of the instance progress events.

## Progress events

Each operation has its own event:

| Event | Raised by |
| --- | --- |
| `CopyFileProgressChanged` | `CopyFileAsync` |
| `CopyDirectoryProgressChanged` | `CopyDirectoryAsync` and `CopyDirectoriesAsync` |
| `DeleteFilesProgressChanged` | `DeleteFilesAsync` |
| `DeleteDirectoriesProgressChanged` | `DeleteDirectoriesAsync` |

Progress notifications use `ProgressEventArgs`:

| Property | Type | Description |
| --- | --- | --- |
| `TotalSize` | `Long` | Total bytes represented by the operation. |
| `HandledSize` | `Long` | Bytes copied or deleted so far. |
| `CurrentPath` | `String` | Current file path, or `Nothing` when no single path applies. |
| `ProcessedItems` | `Long` | Number of files completely processed. |
| `TotalItems` | `Long` | Total number of planned files. |
| `PercentCompleted` | `Integer` | Calculated value from `0` through `100`. |

Byte progress is preferred when `TotalSize` is greater than zero. Item progress is used for zero-byte operations. An empty operation is reported as complete.

Progress is throttled during file streaming to avoid raising an event for every buffer. A final completion notification is still sent.

### Windows Forms example

When an operation is started from the UI thread, the event uses the captured synchronization context, so controls can normally be updated directly:

```vb
AddHandler manager.CopyFileProgressChanged,
    Sub(sender, e)
        ProgressBar1.Value = e.PercentCompleted
        StatusLabel.Text = If(e.CurrentPath, "Preparing...")
    End Sub

Await manager.CopyFileAsync(sourceFile, destinationFile, cancellationToken)
```

If no synchronization context exists, notifications may run on a thread-pool thread.

## Public types

| Type | Purpose |
| --- | --- |
| `FileManager` | Executes asynchronous copy and deletion operations. |
| `CopyDirectoryInfo` | Maps one source directory to one destination directory. |
| `DeleteDirectoryInfo` | Describes a directory and whether its root should be deleted. |
| `ProgressEventArgs` | Supplies immutable byte, item and path progress information. |

## `FileManager` API

| Member | Description |
| --- | --- |
| `CopyFileAsync(source, destination, cancellationToken)` | Copies one file and overwrites the destination. |
| `CopyDirectoryAsync(copyInfo, totalSize, handledSize, cancellationToken)` | Copies one directory and returns the copied byte count. |
| `CopyDirectoriesAsync(directories, cancellationToken)` | Copies multiple directory mappings as one progress operation. |
| `DeleteFilesAsync(files, cancellationToken)` | Deletes distinct files. |
| `DeleteDirectoriesAsync(directories, cancellationToken)` | Deletes directory contents and optionally their roots. |
| `DeleteDirectoryContentAsync(directory, exceptDirectories, exceptFiles, cancellationToken)` | Clears one root while preserving selected content. |

## Path and traversal behavior

- Paths are converted to normalized absolute paths.
- Comparisons are case-insensitive on Windows and case-sensitive on other supported operating systems.
- A file cannot be copied onto itself.
- A directory destination cannot equal or be located inside its own source.
- Reparse points are skipped during recursive directory enumeration.
- Inaccessible entries cause an exception; they are not silently ignored.
- The service copies file content but does not explicitly preserve file attributes, timestamps or access-control entries.

## Failure and cancellation behavior

Operations are not transactional. If an exception or cancellation occurs after work has started:

- a copied destination file may be partial;
- some destination directories may already exist;
- some requested files or directories may already have been deleted;
- previously completed items are not restored automatically.

Applications that require all-or-nothing behavior should create their own staging, rollback or backup strategy.

## License

This package is licensed under the MIT License.
