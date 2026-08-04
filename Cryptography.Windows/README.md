# CoreSuite.Cryptography.Windows

**Windows-protected storage for JSON objects and binary application data.**

> [!NOTE]
> `CoreSuite.Cryptography.Windows` is one of the libraries included in the **CoreSuite** solution. It uses the Windows Data Protection API (DPAPI), so protected values can only be saved and loaded on Windows.

## Overview

`CoreSuite.Cryptography.Windows` provides `ProtectedFileStorage`, a simple way to persist sensitive application data without asking the application to create, store or distribute its own encryption key.

The class can serialize an object as JSON or protect an existing byte array. Windows DPAPI performs the cryptographic protection using either the current Windows user or the local computer as its scope. The resulting bytes are stored inside a small versioned CoreSuite file package so the format and selected scope can be identified later.

Common uses include local API tokens, refresh tokens, connection settings, application credentials and other private configuration that must be read again by the same Windows user or computer.

## Features

- Protects JSON-serializable objects through Windows DPAPI.
- Protects arbitrary binary data without JSON conversion.
- Uses `CurrentUser` protection by default.
- Supports `LocalMachine` protection when several accounts on one trusted computer need access.
- Supports optional additional entropy that must be supplied again when loading.
- Provides synchronous and asynchronous save and load methods.
- Writes through a temporary file and atomically replaces the destination.
- Creates missing destination directories automatically.
- Overwrites an existing protected file safely.
- Provides `TryLoad`, `TryLoadBytes` and `LoadOrDefault` for recoverable failures.
- Identifies structurally valid protected files without decrypting them.
- Reads the stored DPAPI scope from a file header.
- Provides configurable `System.Text.Json` serialization.
- Clears temporary plaintext JSON and protected byte buffers after use.
- Includes English XML documentation for the public API.

## Requirements

- .NET 8 for Windows (`net8.0-windows`)
- Windows operating system
- Microsoft `System.Security.Cryptography.ProtectedData` package, resolved automatically by NuGet

This package does not depend on `CoreSuite.Cryptography`. Install `CoreSuite.Cryptography` separately only when the application also needs SHA-256 hashing, password hashing or password-based AES encryption.

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.Cryptography.Windows
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.Cryptography.Windows
```

## Namespace

Import the CoreSuite namespace and the .NET DPAPI types when selecting a scope:

```vb
Imports System.Security.Cryptography
Imports CoreSuite.Services.Cryptography
```

## Quick start

Define an ordinary JSON-serializable class:

```vb
Public Class ApplicationSecrets
    Public Property ApiToken As String
    Public Property ServerName As String
End Class
```

Save and load it with the default `CurrentUser` scope:

```vb
Imports System.IO
Imports CoreSuite.Services.Cryptography

Dim FilePath As String = Path.Combine(
    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
    "CoreSuite.Sample",
    "secrets.dat")

Dim Secrets As New ApplicationSecrets With {
    .ApiToken = "private-token",
    .ServerName = "primary-server"
}

ProtectedFileStorage.Save(FilePath, Secrets)

Dim LoadedSecrets As ApplicationSecrets = ProtectedFileStorage.Load(Of ApplicationSecrets)(FilePath)
```

The destination folder is created automatically when it does not exist.

## How Windows protection works

DPAPI lets Windows manage the cryptographic key material. The application supplies the data and chooses a protection scope; it does not receive or persist the underlying Windows key.

The selected scope determines who can later unprotect the file:

| Scope | Who can normally load the value? | Recommended use |
| --- | --- | --- |
| `DataProtectionScope.CurrentUser` | The same Windows user profile that protected it | Default choice for desktop application secrets and per-user settings |
| `DataProtectionScope.LocalMachine` | Accounts and processes on the same Windows computer, subject to file access and application permissions | Shared machine services or trusted multi-user scenarios |

> [!IMPORTANT]
> `CurrentUser` is the safer default for most applications. `LocalMachine` makes the DPAPI protection available at computer scope, so use file-system permissions to restrict who can read the protected file and use this scope only when the sharing behavior is intentional.

Protected files are not designed as portable encrypted documents:

- A `CurrentUser` file generally depends on the Windows profile that created it.
- A `LocalMachine` file depends on the computer that created it.
- Copying the file to another account or computer does not make it decryptable there.
- Losing the required Windows profile, machine keys or additional entropy can make the content unrecoverable.

## Save and load JSON values

### Synchronous operations

```vb
ProtectedFileStorage.Save(FilePath, Secrets)
Dim LoadedSecrets As ApplicationSecrets = ProtectedFileStorage.Load(Of ApplicationSecrets)(FilePath)
```

`Save` serializes the value directly to UTF-8 JSON, protects those bytes and atomically replaces the destination file. `Load` reverses the process and deserializes the JSON to the requested type.

### Asynchronous operations

```vb
Await ProtectedFileStorage.SaveAsync(FilePath, Secrets, CancellationToken:=CancellationToken)

Dim LoadedSecrets As ApplicationSecrets = Await ProtectedFileStorage.LoadAsync(Of ApplicationSecrets)(
    FilePath,
    CancellationToken:=CancellationToken)
```

The asynchronous methods make file reading and writing asynchronous and support cancellation. The Windows DPAPI protect and unprotect calls themselves remain synchronous operating-system operations.

### Load without exceptions for expected failures

Use `TryLoad` when a missing, unreadable, malformed or undecryptable settings file is an expected condition:

```vb
Dim Secrets As ApplicationSecrets = Nothing

If ProtectedFileStorage.TryLoad(FilePath, Secrets) Then
    Console.WriteLine(Secrets.ServerName)
Else
    Secrets = New ApplicationSecrets()
End If
```

On failure, the method returns `False` and resets the output value to `Nothing`.

If the caller already has an appropriate fallback value, `LoadOrDefault` provides a shorter flow:

```vb
Dim Defaults As New ApplicationSecrets With {
    .ServerName = "localhost"
}

Dim Secrets As ApplicationSecrets = ProtectedFileStorage.LoadOrDefault(FilePath, Defaults)
```

`LoadOrDefault` returns the exact default value supplied by the caller when loading fails.

## Save and load binary data

Use the byte-array methods when the content is already binary or should not pass through JSON serialization:

```vb
Dim OriginalData As Byte() = File.ReadAllBytes("C:\Files\private.bin")

ProtectedFileStorage.SaveBytes(FilePath, OriginalData)

Dim LoadedData As Byte() = ProtectedFileStorage.LoadBytes(FilePath)
```

Asynchronous equivalents are available:

```vb
Await ProtectedFileStorage.SaveBytesAsync(FilePath, OriginalData, CancellationToken:=CancellationToken)
Dim LoadedData As Byte() = Await ProtectedFileStorage.LoadBytesAsync(FilePath, CancellationToken:=CancellationToken)
```

For recoverable failures, use `TryLoadBytes`:

```vb
Dim LoadedData As Byte() = Nothing

If Not ProtectedFileStorage.TryLoadBytes(FilePath, LoadedData) Then
    LoadedData = Array.Empty(Of Byte)()
End If
```

On failure, `TryLoadBytes` returns `False` and assigns an empty byte array to the output parameter.

## Choose the protection scope

The default scope is `CurrentUser`, so it does not need to be passed explicitly:

```vb
ProtectedFileStorage.Save(FilePath, Secrets)
```

Choose `LocalMachine` only when the application intentionally needs computer-wide protection:

```vb
Imports System.Security.Cryptography

ProtectedFileStorage.Save(
    FilePath,
    Secrets,
    Scope:=DataProtectionScope.LocalMachine)
```

The scope is stored in the CoreSuite file header. Loading reads that value automatically, so `Load` does not require a scope argument.

## Use additional entropy

Additional entropy is optional application-provided data included in the DPAPI operation. The exact same bytes must be supplied when the file is loaded:

```vb
Imports System.Text

Dim Entropy As Byte() = Encoding.UTF8.GetBytes("CoreSuite.Sample.Secrets.v1")

ProtectedFileStorage.Save(FilePath, Secrets, Entropy:=Entropy)
Dim LoadedSecrets As ApplicationSecrets = ProtectedFileStorage.Load(Of ApplicationSecrets)(FilePath, Entropy)
```

Important behavior:

- Entropy is not stored in the protected file.
- Different entropy causes loading to fail cryptographic validation.
- Losing randomly generated entropy makes the protected value unrecoverable.
- Changing a purpose string such as the example above requires decrypting and saving the content again with the new entropy.
- Entropy complements the Windows scope; it does not replace file permissions or secret management.

## Customize JSON serialization

The default JSON configuration is case-insensitive during property matching and writes indented JSON before protection. Indentation does not expose JSON because the serialized bytes are protected before they are written to the destination file.

Create an independent copy of the defaults and modify it for one operation:

```vb
Imports System.Text.Json

Dim Options As JsonSerializerOptions = ProtectedFileStorage.CreateDefaultJsonOptions()
Options.PropertyNamingPolicy = JsonNamingPolicy.CamelCase
Options.WriteIndented = False

ProtectedFileStorage.Save(FilePath, Secrets, Options:=Options)
Dim LoadedSecrets As ApplicationSecrets = ProtectedFileStorage.Load(Of ApplicationSecrets)(FilePath, Options:=Options)
```

Each call to `CreateDefaultJsonOptions` returns a separate mutable instance. Changing it does not modify the defaults used by other calls.

## Inspect and delete protected files

### Check the file structure

```vb
If ProtectedFileStorage.IsProtectedFile(FilePath) Then
    Console.WriteLine("Recognized CoreSuite protected file.")
End If
```

`IsProtectedFile` checks the versioned package header and length. It does not decrypt the content, validate optional entropy or prove that the current user can load it. A payload modified without changing its overall structure can still pass this check and then fail during `Load`.

### Read the stored scope

```vb
Dim Scope As DataProtectionScope = ProtectedFileStorage.GetProtectionScope(FilePath)
```

The scope is read from the package header without decrypting the protected payload.

### Delete a file

```vb
Dim WasDeleted As Boolean = ProtectedFileStorage.Delete(FilePath)
```

`Delete` returns `True` when a file existed and was deleted, or `False` when the path did not exist. File-system errors are still propagated.

## Atomic save behavior

Both save flows protect the data before writing it to disk. The completed package is written to a uniquely named temporary file in the destination directory, flushed, and then moved over the destination path.

This design provides these practical benefits:

- Readers do not see a partially written destination file.
- An existing destination remains available until the new package is ready to replace it.
- Missing destination directories are created automatically.
- A temporary file is removed when the save cannot be completed.

Atomic replacement protects against partial writes, but it is not a backup system. Keep a separate backup strategy when the protected data cannot be recreated, and remember that the backup remains tied to its DPAPI scope and entropy.

## API reference

### JSON methods

| Member | Behavior |
| --- | --- |
| `Save(Of T)(filePath, value, scope, entropy, options)` | Serializes, protects and atomically saves a value. |
| `SaveAsync(Of T)(filePath, value, scope, entropy, options, cancellationToken)` | Asynchronous file-save equivalent. |
| `Load(Of T)(filePath, entropy, options)` | Loads, unprotects and deserializes a value. |
| `LoadAsync(Of T)(filePath, entropy, options, cancellationToken)` | Asynchronous file-load equivalent. |
| `TryLoad(Of T)(filePath, value, entropy, options)` | Attempts to load a value and reports failure through `Boolean`. |
| `LoadOrDefault(Of T)(filePath, defaultValue, entropy, options)` | Returns a loaded value or the caller-provided fallback. |
| `CreateDefaultJsonOptions()` | Returns a mutable copy of the default JSON settings. |

### Binary and file methods

| Member | Behavior |
| --- | --- |
| `SaveBytes(filePath, data, scope, entropy)` | Protects and atomically saves binary data. |
| `SaveBytesAsync(filePath, data, scope, entropy, cancellationToken)` | Asynchronous file-save equivalent. |
| `LoadBytes(filePath, entropy)` | Loads and unprotects binary data. |
| `LoadBytesAsync(filePath, entropy, cancellationToken)` | Asynchronous file-load equivalent. |
| `TryLoadBytes(filePath, data, entropy)` | Attempts to load binary data and reports failure through `Boolean`. |
| `IsProtectedFile(filePath)` | Checks whether a file has a recognized package structure. |
| `GetProtectionScope(filePath)` | Reads the stored DPAPI scope from the header. |
| `Delete(filePath)` | Deletes the file if it exists and reports whether deletion occurred. |

## Failure behavior

`TryLoad` and `TryLoadBytes` return `False` for expected recovery conditions such as:

- Missing file or destination directory.
- Access or general file I/O failure.
- Invalid or incomplete CoreSuite package.
- Incorrect or missing additional entropy.
- A DPAPI decryption failure.
- Invalid JSON or a value that cannot be deserialized by the requested JSON configuration (`TryLoad` only).

They deliberately rethrow `PlatformNotSupportedException`, because attempting to use DPAPI outside Windows is a programming or deployment error rather than a damaged settings file.

The direct `Load` methods expose the original exception. This is preferable when the application must distinguish a missing file, malformed package, cryptographic failure and JSON failure.

The asynchronous methods also propagate `OperationCanceledException` when their cancellation token is canceled.

## Security and deployment guidance

- Prefer `DataProtectionScope.CurrentUser` unless several trusted accounts on the same computer must decrypt the file.
- Restrict file-system permissions even though the content is protected. File access control and cryptographic protection serve different purposes.
- Do not assume a copied protected file can be restored on another computer or under another Windows profile.
- Preserve any additional entropy required by the application. Without the same bytes, the data cannot be loaded.
- Do not expose loaded secrets in logs, exception messages or UI diagnostics.
- Delete or overwrite sensitive plaintext byte arrays when the calling code no longer needs them; the package cannot clear buffers owned by the caller.
- Treat `IsProtectedFile` only as a format check. Successful decryption is the actual proof that the protected payload is usable.
- Use `CoreSuite.Cryptography.TextEncryption` instead when data must be portable between systems through a separately managed password.

## Thread safety

The class exposes stateless shared methods and does not keep files open between operations. Independent file paths can be processed concurrently. Coordinate concurrent writers targeting the same file in application code, because the class does not serialize operations for a shared path.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.Cryptography.Windows` |
| Namespace | `CoreSuite.Services.Cryptography` |
| Assembly | `CoreSuite.Cryptography.Windows` |
| Target framework | `net8.0-windows` |
| Supported operating system | Windows |
| CoreSuite dependencies | None |
| Microsoft dependency | `System.Security.Cryptography.ProtectedData` 8.0.0 |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
