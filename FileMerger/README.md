# CoreSuite.FileMerger

**Package one or more directory trees into a single password-encrypted container file.**

> [!NOTE]
> `CoreSuite.FileMerger` is one of the libraries included in the **CoreSuite** solution. It targets plain .NET 8, works independently of Windows Forms and has no CoreSuite or third-party package dependencies.

## Overview

`CoreSuite.FileMerger` creates a proprietary container that stores every file found below one or more source directories. The original root folder names and relative file paths are recorded so the directory structure can later be restored.

File contents are encrypted with AES using a 256-bit key derived from a password through PBKDF2-HMAC-SHA256. The container also stores its creation date, total file count, total byte size, file paths and root folder names. This metadata can be inspected without extracting the files.

The package supports synchronous and asynchronous merge and extraction calls, format validation and integer progress reporting.

> [!IMPORTANT]
> This is a CoreSuite-specific packaging format, not a ZIP file. It does not compress data and cannot be opened by standard archive applications.

## Features

- Combines files from several directory trees into one output file.
- Preserves each source root folder name and relative file paths.
- Restores the stored directory structure to a target directory.
- Encrypts stored file content with AES.
- Derives a 256-bit encryption key from the supplied password.
- Generates a new AES initialization vector for each container.
- Stores creation date, file count, total size, paths and root folders.
- Reads metadata without extracting file content.
- Validates the container signature and basic header structure.
- Reports progress from `0` through `100` by completed file count.
- Provides synchronous and asynchronous APIs.
- Includes English XML documentation for the public API.
- Requires no third-party packages.

## Requirements

- .NET 8 (`net8.0`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.FileMerger
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.FileMerger
```

## Namespace

Import the service namespace:

```vb
Imports CoreSuite.Services
```

## Quick start

Create a container from two directories:

```vb
Dim sourceDirectories As String() = {
    "C:\Application\Documents",
    "C:\Application\Images"
}

FileMerger.Merge(
    "D:\Backups\application.csfm",
    sourceDirectories,
    "A strong private password")
```

The output contains paths beginning with the source root names:

```text
Documents\...
Images\...
```

Restore the files later:

```vb
FileMerger.UnMerge(
    "D:\Backups\application.csfm",
    "D:\Restored",
    "A strong private password")
```

The target receives `D:\Restored\Documents` and `D:\Restored\Images` with their stored content.

## Merge asynchronously

Use `MergeAsync` in UI applications or services where the calling thread should remain available:

```vb
Dim progress As New Progress(Of Integer)(
    Sub(percent)
        Console.WriteLine($"Creating container: {percent}%")
    End Sub)

Await FileMerger.MergeAsync(
    "D:\Backups\application.csfm",
    sourceDirectories,
    encryptionPassword,
    progress)
```

The asynchronous method runs the synchronous merge operation on a thread-pool thread.

## Extract asynchronously

```vb
Dim progress As New Progress(Of Integer)(
    Sub(percent)
        Console.WriteLine($"Restoring files: {percent}%")
    End Sub)

Await FileMerger.UnMergeAsync(
    "D:\Backups\application.csfm",
    "D:\Restored",
    encryptionPassword,
    progress)
```

Existing files with the same restored paths are overwritten.

## Validate a container

Check the file before reading metadata or extracting it:

```vb
Dim containerPath As String = "D:\Backups\application.csfm"

If Not FileMerger.IsValidFile(containerPath) Then
    Throw New InvalidDataException("The selected file is not a valid CoreSuite FileMerger container.")
End If
```

`IsValidFile` checks the CoreSuite signature, initialization-vector length and available header bytes. It returns `False` instead of propagating ordinary validation errors.

> [!WARNING]
> A `True` result only confirms the expected basic header format. It does not verify the password, decrypt the content or prove that the complete container is intact.

## Read metadata

Metadata can be read without supplying the encryption password:

```vb
Dim metadata As Dictionary(Of String, Object) = FileMerger.GetMetadata(containerPath)

Dim creationDate As DateTime = DirectCast(metadata("CreationDate"), DateTime)
Dim totalFiles As Integer = DirectCast(metadata("TotalFiles"), Integer)
Dim totalSize As Long = DirectCast(metadata("TotalSize"), Long)
Dim filePaths As List(Of String) = DirectCast(metadata("FilePaths"), List(Of String))
Dim rootFolders As List(Of String) = DirectCast(metadata("RootFolders"), List(Of String))
```

The returned dictionary contains these exact keys:

| Key | Runtime type | Description |
| --- | --- | --- |
| `CreationDate` | `DateTime` | Container creation time converted from UTC to local time. |
| `TotalFiles` | `Integer` | Number of stored files. |
| `TotalSize` | `Long` | Sum of the original file sizes in bytes. |
| `FilePaths` | `List(Of String)` | Stored relative path for every file. |
| `RootFolders` | `List(Of String)` | Distinct stored source root names. |

The metadata is written outside the encrypted content section. Do not put sensitive information in source directory or file names if revealing those names is unacceptable.

## Container structure

The generated file contains three logical areas:

1. A header with the CoreSuite format signature and AES initialization vector.
2. Plain-text binary metadata containing the creation time, counts, sizes and path lists.
3. An AES-encrypted sequence containing each relative path, file length and file bytes.

All files are written sequentially. The package does not create an index for random extraction of an individual file.

## Progress reporting

Both merge and extraction accept `IProgress(Of Integer)`:

```vb
Dim progress As IProgress(Of Integer) = New Progress(Of Integer)(AddressOf UpdateProgress)
```

```vb
Private Sub UpdateProgress(percent As Integer)
    ProgressBar1.Value = percent
End Sub
```

Progress advances when an entire file is merged or restored. It does not report partial byte progress inside a single file. An empty input set produces no progress callbacks.

## API reference

| Member | Description |
| --- | --- |
| `Merge(outputFile, directories, password, progress)` | Synchronously creates a container from all files below the supplied directories. |
| `MergeAsync(outputFile, directories, password, progress)` | Runs `Merge` on a thread-pool thread. |
| `UnMerge(inputFile, targetDirectory, password, progress)` | Synchronously restores all stored files. |
| `UnMergeAsync(inputFile, targetDirectory, password, progress)` | Runs `UnMerge` on a thread-pool thread. |
| `IsValidFile(filePath)` | Checks the signature and basic header structure. |
| `GetMetadata(inputFile)` | Reads the unencrypted metadata dictionary without extraction. |

## Encryption details and security limits

The current format uses:

- PBKDF2-HMAC-SHA256 with `100,000` iterations;
- a fixed format salt;
- a 32-byte derived key;
- AES with a new initialization vector per container;
- the platform's default AES mode and padding, normally CBC with PKCS#7 on .NET.

The container does not include a message authentication code or an authenticated-encryption tag. As a result, it does not provide strong cryptographic proof that encrypted content has not been modified.

Use this package for application-controlled packaging and password-based content hiding where those limits are acceptable. For security-critical archives, long-term backup or hostile storage, use a reviewed authenticated archive format with explicit integrity verification.

## Operational limits

- Individual file lengths are stored as signed 32-bit integers. Files of 2 GB or larger are not supported.
- Extraction allocates a byte array for the complete current file before writing it. Available memory therefore limits the practical individual file size.
- The format does not compress file content.
- The asynchronous methods do not currently accept `CancellationToken`.
- Missing source directories are skipped during merge.
- The output file is created or overwritten.
- Source directories should have distinct final folder names to avoid restored path collisions.
- A wrong password or invalid encrypted payload may surface as `UnauthorizedAccessException`; an invalid header can surface as `InvalidDataException`.
- The package does not provide transactional rollback if merge or extraction fails partway through.

## Recommended workflow

For application backups:

1. Select source directories with distinct root names.
2. Write to a temporary output path.
3. Await `MergeAsync`.
4. Confirm `IsValidFile` and inspect `GetMetadata`.
5. Move the completed container into its final backup location.
6. Keep the password through a separate secure recovery process.
7. Test restoration periodically with representative data.

## License

This package is licensed under the MIT License.
