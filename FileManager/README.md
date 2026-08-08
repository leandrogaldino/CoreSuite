# CoreSuite.FileManager

**Asynchronous file and directory copy and deletion with shared operations, percentage progress and cancellation.**

> [!NOTE]
> `CoreSuite.FileManager` is one of the libraries included in the **CoreSuite** solution. It targets plain .NET 8, works independently of Windows Forms and has no CoreSuite or third-party package dependencies.

## Overview

`CoreSuite.FileManager` provides shared asynchronous operations for copying and deleting files or complete directory trees without requiring a `FileManager` instance.

All public operations are exposed through `Shared` methods. Operations that can report progress accept an optional `IProgress(Of Integer)` parameter and report completion values from `0` through `100`. Directory traversal deliberately skips reparse points, preventing the service from following symbolic links, junctions and other redirected directory structures.

The package also includes a helper for clearing a directory while preserving selected files and subdirectories.

## Features

- Fully shared `FileManager` API; no service instance is required.
- Copies individual files asynchronously.
- Copies one or several directory trees asynchronously.
- Creates missing destination directories automatically during copy operations.
- Preserves empty directories when copying a directory tree.
- Overwrites existing destination files.
- Deletes individual files asynchronously.
- Deletes one or several directory trees asynchronously.
- Optionally preserves the root directory while deleting its contents.
- Clears a directory while preserving selected files and subdirectories.
- Reports optional percentage progress through `IProgress(Of Integer)`.
- Supports `Progress(Of Integer)` synchronization-context capture for UI applications.
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

Copy one file and receive percentage progress updates:

```vb
Dim sourceFile As New FileInfo("C:\Data\large-export.zip")
Dim destinationFile As New FileInfo("D:\Backup\large-export.zip")
Dim progress As New Progress(Of Integer)(
    Sub(percent)
        Console.WriteLine($"{percent}%")
    End Sub)

Await FileManager.CopyFileAsync(sourceFile, destinationFile, progress)
```

The destination parent directory is created when necessary. An existing destination file is overwritten.

Progress is optional. When it is not needed, simply omit the parameter:

```vb
Await FileManager.CopyFileAsync(sourceFile, destinationFile)
```

## Copy a file with cancellation

```vb
Using cancellationSource As New CancellationTokenSource()
    Try
        Await FileManager.CopyFileAsync(
            sourceFile,
            destinationFile,
            CancellationToken:=cancellationSource.Token)
    Catch ex As OperationCanceledException
        Console.WriteLine("Copy canceled.")
    End Try
End Using
```

Cancellation does not roll back bytes already written. If the operation is canceled, the destination may contain a partial file and your application should decide whether to keep or remove it.

## Copy a directory

Create a `CopyDirectoryInfo` containing the source and destination:

```vb
Dim copyRequest As New CopyDirectoryInfo(
    New DirectoryInfo("C:\ApplicationData"),
    New DirectoryInfo("D:\Backup\ApplicationData"))

Dim progress As New Progress(Of Integer)(
    Sub(percent)
        Console.WriteLine($"{percent}%")
    End Sub)

Dim copiedBytes As Long = Await FileManager.CopyDirectoryAsync(
    copyRequest,
    Progress:=progress)
```

The complete source tree is planned before copying starts. This allows the service to calculate the total number of bytes. Empty directories are created at the destination.

### Participate in a larger progress calculation

`CopyDirectoryAsync` accepts optional `TotalSize` and `HandledSize` values when its progress should be combined with other work:

```vb
Dim alreadyCopied As Long = 500000
Dim completeOperationSize As Long = 2000000
Dim progress As New Progress(Of Integer)(Sub(percent) Console.WriteLine($"{percent}%"))

Dim copiedNow As Long = Await FileManager.CopyDirectoryAsync(
    copyRequest,
    TotalSize:=completeOperationSize,
    HandledSize:=alreadyCopied,
    Progress:=progress,
    CancellationToken:=cancellationToken)
```

The returned value is the number of bytes copied by the supplied directory request, not the accumulated `HandledSize`.

## Copy multiple directories

```vb
Dim requests As CopyDirectoryInfo() = {
    New CopyDirectoryInfo(New DirectoryInfo("C:\Data\Customers"), New DirectoryInfo("D:\Backup\Customers")),
    New CopyDirectoryInfo(New DirectoryInfo("C:\Data\Orders"), New DirectoryInfo("D:\Backup\Orders"))
}

Dim progress As New Progress(Of Integer)(Sub(percent) Console.WriteLine($"{percent}%"))

Await FileManager.CopyDirectoriesAsync(requests, progress, cancellationToken)
```

Progress is calculated across all planned files in all requests.

## Delete files

```vb
Dim filesToDelete As FileInfo() = {
    New FileInfo("C:\Temp\result-1.tmp"),
    New FileInfo("C:\Temp\result-2.tmp")
}

Dim progress As New Progress(Of Integer)(Sub(percent) Console.WriteLine($"{percent}%"))

Await FileManager.DeleteFilesAsync(filesToDelete, progress, cancellationToken)
```

Duplicate normalized paths are processed only once. Every requested file must exist when the deletion plan is built; otherwise, `FileNotFoundException` is thrown.

## Delete directories

Use `DeleteDirectoryInfo.DeleteRoot` to choose whether the root is removed after its contents:

```vb
Dim requests As DeleteDirectoryInfo() = {
    New DeleteDirectoryInfo(New DirectoryInfo("C:\Temp\OldImports"), True),
    New DeleteDirectoryInfo(New DirectoryInfo("C:\Temp\Working"), False)
}

Dim progress As New Progress(Of Integer)(Sub(percent) Console.WriteLine($"{percent}%"))

Await FileManager.DeleteDirectoriesAsync(requests, progress, cancellationToken)
```

The first request deletes `OldImports` and its contents. The second deletes everything inside `Working` but preserves the `Working` root directory.

Missing directory roots are ignored. Existing roots in the same request may not be equal, nested or otherwise overlap.

## Clear a directory with exclusions

`DeleteDirectoryContentAsync` removes content without deleting the supplied root. It can preserve complete subtrees and individual files and can also report percentage progress:

```vb
Dim rootDirectory As New DirectoryInfo("C:\Application\Cache")
Dim directoriesToKeep As DirectoryInfo() = {
    New DirectoryInfo("C:\Application\Cache\Pinned")
}
Dim filesToKeep As FileInfo() = {
    New FileInfo("C:\Application\Cache\settings.json")
}
Dim progress As New Progress(Of Integer)(Sub(percent) Console.WriteLine($"{percent}%"))

Await FileManager.DeleteDirectoryContentAsync(
    rootDirectory,
    directoriesToKeep,
    filesToKeep,
    progress,
    cancellationToken)
```

When a directory is excluded, all of its descendants are preserved. The service also preserves the required ancestors of excluded items so that those items remain reachable.

Every exclusion must be located inside the supplied root. Passing the root itself in `ExceptDirectories` makes the method return without deleting anything and reports completion when a progress reporter is supplied.

## Progress reporting

Progress is supplied per operation through an optional `IProgress(Of Integer)` parameter. Reported values are always clamped to the range `0` through `100`.

For copy operations and file/directory deletion operations, byte-based progress is preferred when a positive total byte size is available. Item-based progress is used when there is no positive byte total. `DeleteDirectoryContentAsync` uses item-based progress for the files and directories it removes.

Progress updates are throttled during high-throughput work to avoid excessive synchronization-context traffic. A final `100` notification is still reported after successful completion.

### Windows Forms example

Create `Progress(Of Integer)` on the UI thread. The standard .NET implementation captures the current synchronization context, so its callback can normally update controls directly:

```vb
Dim progress As New Progress(Of Integer)(
    Sub(percent)
        ProgressBar1.Value = percent
        StatusLabel.Text = $"{percent}%"
    End Sub)

Await FileManager.CopyFileAsync(
    sourceFile,
    destinationFile,
    progress,
    cancellationToken)
```

No event subscription or `FileManager` instance is required.

## Public types

| Type | Purpose |
| --- | --- |
| `FileManager` | Exposes shared asynchronous copy and deletion operations. |
| `CopyDirectoryInfo` | Maps one source directory to one destination directory. |
| `DeleteDirectoryInfo` | Describes a directory and whether its root should be deleted. |

## `FileManager` API

| Member | Description |
| --- | --- |
| `CopyFileAsync(Source, Destination, Progress, CancellationToken)` | Copies one file and overwrites the destination. |
| `CopyDirectoryAsync(CopyInfo, TotalSize, HandledSize, Progress, CancellationToken)` | Copies one directory and returns the copied byte count. |
| `CopyDirectoriesAsync(Directories, Progress, CancellationToken)` | Copies multiple directory mappings as one progress operation. |
| `DeleteFilesAsync(Files, Progress, CancellationToken)` | Deletes distinct files. |
| `DeleteDirectoriesAsync(Directories, Progress, CancellationToken)` | Deletes directory contents and optionally their roots. |
| `DeleteDirectoryContentAsync(Directory, ExceptDirectories, ExceptFiles, Progress, CancellationToken)` | Clears one root while preserving selected content. |

All parameters after the required operation arguments are optional.

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
