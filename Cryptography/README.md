# CoreSuite.Cryptography

**Secure hashing, password storage and authenticated encryption utilities for .NET 8 applications.**

> [!NOTE]
> `CoreSuite.Cryptography` is one of the libraries included in the **CoreSuite** solution. It targets plain .NET 8, works independently of Windows Forms and has no CoreSuite or third-party package dependencies.

## Overview

`CoreSuite.Cryptography` provides three focused classes for common application security tasks:

- `Hashing` creates and verifies SHA-256 fingerprints for text, binary data, streams and files.
- `PasswordHasher` stores passwords safely through versioned PBKDF2-HMAC-SHA256 hashes with a random salt.
- `TextEncryption` encrypts and authenticates text or binary data through AES-256-GCM using a key derived from a password.

The package handles salts, nonces, authentication tags and version metadata automatically. Its generated password hashes and encrypted packages are self-contained, so the values can be stored directly and later processed without maintaining those components in separate database columns.

## Features

- Computes SHA-256 hashes for strings, byte arrays, streams and files.
- Supports synchronous and asynchronous file hashing.
- Verifies hashes with fixed-time comparisons.
- Creates versioned password hashes with PBKDF2-HMAC-SHA256.
- Generates a new cryptographically random salt for every password hash.
- Reports when a valid stored password hash should be upgraded.
- Encrypts text and arbitrary binary data with AES-256-GCM.
- Detects incorrect passwords and modified encrypted content through authentication.
- Generates a new random salt and nonce for every encryption operation.
- Provides non-throwing `Try` methods for expected decryption failures.
- Stores format version and PBKDF2 iteration information in generated values.
- Clears several temporary sensitive byte buffers after use.
- Includes English XML documentation for the public API.
- Requires no third-party packages.

## Requirements

- .NET 8 (`net8.0`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.Cryptography
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.Cryptography
```

## Namespace

Import the namespace containing all public types:

```vb
Imports CoreSuite.Services.Cryptography
```

## Quick start

```vb
Imports CoreSuite.Services.Cryptography

Dim FileHash As String = Hashing.ComputeFileSha256("C:\Files\document.pdf")

Dim StoredPasswordHash As String = PasswordHasher.HashPassword("User password")
Dim PasswordIsValid As Boolean = PasswordHasher.VerifyPassword("User password", StoredPasswordHash)

Dim EncryptedText As String = TextEncryption.Encrypt("Sensitive configuration", "Encryption password")
Dim OriginalText As String = TextEncryption.Decrypt(EncryptedText, "Encryption password")
```

## Choose the correct operation

Hashing and encryption solve different problems. Use the following guide before choosing an API:

| Goal | Type | Can the original value be recovered? | Typical use |
| --- | --- | --- | --- |
| Detect whether ordinary data changed | `Hashing` | No | File integrity, cache keys and content comparison |
| Store a password for later login verification | `PasswordHasher` | No | User authentication |
| Protect data that the application must read again | `TextEncryption` | Yes, with the correct password | Tokens, connection settings and private application values |

> [!IMPORTANT]
> Do not store login passwords with `TextEncryption`. Passwords should normally be processed with `PasswordHasher`, because authentication only needs to verify a password and should not require recovering it.

## SHA-256 hashing

A SHA-256 hash is a one-way, fixed-size fingerprint. The same input always produces the same 64-character hexadecimal value. A hash is useful for comparison and integrity checks, but it does not hide low-entropy or predictable data and it is not encryption.

### Hash text

Text is encoded as UTF-8 before its hash is computed:

```vb
Dim Hash As String = Hashing.ComputeSha256("CoreSuite")
```

The returned text uses uppercase hexadecimal characters. Verification accepts valid hexadecimal hashes in uppercase or lowercase:

```vb
Dim IsMatch As Boolean = Hashing.VerifySha256("CoreSuite", Hash)
```

### Hash binary data

```vb
Dim Data As Byte() = {&H1, &H2, &H3, &H4}
Dim Hash As String = Hashing.ComputeSha256(Data)
Dim IsMatch As Boolean = Hashing.VerifySha256(Data, Hash)
```

### Hash a stream

```vb
Using Input As FileStream = File.OpenRead("C:\Files\archive.zip")
    Dim Hash As String = Hashing.ComputeSha256(Input)
End Using
```

The stream must be readable. Hashing starts at its current position and reads to the end; the method does not dispose the caller-provided stream.

### Hash and verify files

```vb
Dim FilePath As String = "C:\Files\installer.exe"
Dim ExpectedHash As String = Hashing.ComputeFileSha256(FilePath)
Dim FileIsUnchanged As Boolean = Hashing.VerifyFileSha256(FilePath, ExpectedHash)
```

For UI applications, services or large files, use the asynchronous methods so file reading does not block the calling thread:

```vb
Dim Hash As String = Await Hashing.ComputeFileSha256Async(FilePath, CancellationToken)
Dim IsMatch As Boolean = Await Hashing.VerifyFileSha256Async(FilePath, Hash, CancellationToken)
```

The asynchronous methods support cancellation through `CancellationToken`.

### Hashing API reference

| Member | Result |
| --- | --- |
| `ComputeSha256(text)` | SHA-256 of UTF-8 text as 64 uppercase hexadecimal characters. |
| `ComputeSha256(data)` | SHA-256 of a byte array as uppercase hexadecimal text. |
| `ComputeSha256(stream)` | SHA-256 of a readable stream from its current position. |
| `ComputeFileSha256(filePath)` | Synchronously hashes a file. |
| `ComputeFileSha256Async(filePath, cancellationToken)` | Asynchronously hashes a file. |
| `VerifySha256(text, expectedHash)` | Verifies UTF-8 text against an expected SHA-256 hash. |
| `VerifySha256(data, expectedHash)` | Verifies binary data against an expected SHA-256 hash. |
| `VerifyFileSha256(filePath, expectedHash)` | Synchronously verifies a file. |
| `VerifyFileSha256Async(filePath, expectedHash, cancellationToken)` | Asynchronously verifies a file. |

Invalid, empty or incorrectly sized expected hash strings return `False` instead of causing a comparison failure exception.

## Password hashing

`PasswordHasher` uses PBKDF2-HMAC-SHA256 to derive a 32-byte password hash from the supplied password and a new random 16-byte salt. Repeating `HashPassword` with the same password produces different stored values because each operation creates a different salt.

The complete result has this versioned structure:

```text
CSPH$version$algorithm$iterations$salt$hash
```

The salt and hash are Base64 encoded. Store the entire returned string exactly as produced; there is no need to create separate columns for the algorithm, iteration count or salt. The stored value does not contain the original password and cannot be decrypted.

### Create a password hash

When a user creates or changes a password:

```vb
Dim EncodedHash As String = PasswordHasher.HashPassword(PasswordEnteredByUser)

' Save EncodedHash in the password-hash column of the user record.
```

Never store `PasswordEnteredByUser` itself.

### Verify a password

During authentication, read the stored hash and compare it with the password supplied by the user:

```vb
Dim IsValid As Boolean = PasswordHasher.VerifyPassword(PasswordEnteredByUser, StoredHash)

If IsValid Then
    ' Authentication succeeded.
Else
    ' The password is incorrect or the stored hash is invalid.
End If
```

`VerifyPassword` also returns `False` for a malformed or unsupported stored hash.

### Upgrade an older hash after login

Security settings may increase over time. `VerifyPasswordDetailed` distinguishes an ordinary success from a success that should be followed by a new hash:

```vb
Dim Result As PasswordVerificationResult = PasswordHasher.VerifyPasswordDetailed(PasswordEnteredByUser, StoredHash)

Select Case Result
    Case PasswordVerificationResult.Success
        ' The password is valid and the stored settings are current.
    Case PasswordVerificationResult.SuccessRehashNeeded
        ' The password is valid. Replace the old hash while the password is available.
        StoredHash = PasswordHasher.HashPassword(PasswordEnteredByUser)
    Case PasswordVerificationResult.Failed
        ' Authentication failed.
End Select
```

Only replace a hash after a successful verification. `NeedsRehash` can inspect the stored format without a password, but it cannot authenticate the user:

```vb
Dim ShouldUpgrade As Boolean = PasswordHasher.NeedsRehash(StoredHash)
```

### Password hashing API reference

| Member | Behavior |
| --- | --- |
| `HashPassword(password, iterations)` | Creates a new self-contained password hash. |
| `VerifyPassword(password, encodedHash)` | Returns `True` when the password matches. |
| `VerifyPasswordDetailed(password, encodedHash, requiredIterations)` | Returns `Failed`, `Success` or `SuccessRehashNeeded`. |
| `NeedsRehash(encodedHash, requiredIterations)` | Reports malformed, unsupported or lower-iteration hashes. |
| `TryGetIterations(encodedHash, iterations)` | Reads the stored iteration count without authenticating a password. |

### Iteration settings

| Constant | Value | Meaning |
| --- | ---: | --- |
| `DefaultIterations` | `600,000` | Default used for newly generated password hashes. |
| `MinimumIterations` | `100,000` | Lowest accepted value. |
| `MaximumIterations` | `5,000,000` | Highest accepted value. |
| `CurrentFormatVersion` | `1` | Current stored-hash format version. |

An iteration count outside the supported range causes `ArgumentOutOfRangeException`.

## Authenticated encryption

`TextEncryption` uses AES-256-GCM. Encryption hides the content and generates an authentication tag, which also allows decryption to detect an incorrect password or modified package.

The AES key is not stored in the encrypted result. It is derived from the password through PBKDF2-HMAC-SHA256 using a new random 16-byte salt. A new 12-byte nonce is also generated for every operation.

### Encrypt and decrypt text

```vb
Dim EncryptedText As String = TextEncryption.Encrypt("API token: abc123", EncryptionPassword)

' EncryptedText is a self-contained Base64 string suitable for storage.
Dim DecryptedText As String = TextEncryption.Decrypt(EncryptedText, EncryptionPassword)
```

Empty text is supported, but the encryption password cannot be `Nothing` or empty.

Encrypting the same value twice with the same password produces different encrypted strings because fresh random salt and nonce values are generated each time.

### Handle an expected decryption failure

`Decrypt` throws when the package is malformed, the password is incorrect or the encrypted data was changed. Use `TryDecrypt` when failure is an expected branch:

```vb
Dim DecryptedText As String = Nothing

If TextEncryption.TryDecrypt(EncryptedText, EncryptionPassword, DecryptedText) Then
    Console.WriteLine(DecryptedText)
Else
    Console.WriteLine("The value could not be decrypted.")
End If
```

On failure, `TryDecrypt` returns `False` and assigns `String.Empty` to the output parameter.

### Encrypt binary data

```vb
Dim OriginalData As Byte() = File.ReadAllBytes("C:\Files\private.dat")
Dim EncryptedData As Byte() = TextEncryption.EncryptBytes(OriginalData, EncryptionPassword)
Dim DecryptedData As Byte() = TextEncryption.DecryptBytes(EncryptedData, EncryptionPassword)
```

`TryDecryptBytes` provides the equivalent non-throwing flow and returns an empty byte array on failure.

These methods process the complete byte array in memory. For very large files, use a purpose-built streaming file-encryption design instead of loading the entire file into one array.

### Inspect and upgrade encrypted values

```vb
If TextEncryption.IsEncrypted(EncryptedText) Then
    Dim StoredIterations As Integer = TextEncryption.GetIterations(EncryptedText)
    Dim ShouldUpgrade As Boolean = TextEncryption.NeedsReEncryption(EncryptedText)
End If
```

`IsEncrypted` checks whether the package has a recognized and internally consistent structure. It does not verify the password or authentication tag. A structurally valid package that was modified can still return `True`; authenticity is confirmed only by successful decryption.

When `NeedsReEncryption` returns `True`, decrypt the value with the correct password and encrypt it again with the current iteration setting.

### Encryption API reference

| Member | Behavior |
| --- | --- |
| `Encrypt(text, password, iterations)` | Returns a versioned encrypted package as Base64 text. |
| `Decrypt(encryptedText, password)` | Decrypts Base64 text or throws when it cannot be authenticated. |
| `TryDecrypt(encryptedText, password, text)` | Attempts text decryption without throwing for expected format or authentication failures. |
| `EncryptBytes(data, password, iterations)` | Returns a versioned authenticated binary package. |
| `DecryptBytes(encryptedData, password)` | Decrypts and authenticates a binary package. |
| `TryDecryptBytes(encryptedData, password, data)` | Attempts binary decryption without throwing for expected failures. |
| `IsEncrypted(encryptedText)` | Checks the structure of a Base64 package. |
| `IsEncrypted(encryptedData)` | Checks the structure of a binary package. |
| `GetIterations(encryptedText)` | Reads the PBKDF2 iteration count stored in a valid package. |
| `NeedsReEncryption(encryptedText, requiredIterations)` | Reports whether a valid package uses fewer iterations than required. |

`TextEncryption` uses the same default, minimum and maximum PBKDF2 iteration counts as `PasswordHasher`.

## Error behavior

| Situation | Typical result |
| --- | --- |
| Required input is `Nothing` | `ArgumentNullException` |
| A required password or path is empty | `ArgumentException` |
| Iteration count is outside the supported range | `ArgumentOutOfRangeException` |
| Encrypted Base64 or package structure is invalid | `FormatException` |
| Password is incorrect or encrypted content was modified | `CryptographicException` or a derived exception |
| File does not exist or cannot be read | The corresponding file-system exception |
| A `TryDecrypt` operation cannot authenticate or parse data | Returns `False` and resets its output value |

## Security guidance

- Use `PasswordHasher` for authentication passwords and `TextEncryption` only for data that must be recovered.
- Treat encryption passwords as secrets. Do not hard-code them in source code or commit them to a repository.
- Keep a reliable recovery or rotation plan. Encrypted content cannot be recovered if its password is lost.
- Store the complete password hash or encrypted package exactly as returned.
- Do not reuse a SHA-256 hash as a substitute for password hashing; ordinary SHA-256 is intentionally fast and has no per-password work factor.
- Do not lower the PBKDF2 iteration count without measuring and understanding the security tradeoff.
- Prefer the `Try` methods only when failure is expected. Use the throwing methods when malformed or modified data should be treated as an application error.
- Cryptography protects content, but it does not replace authorization, access control, secure transport, backups or safe secret management.

## Thread safety

All APIs are exposed through stateless shared methods. Each operation creates its own cryptographic state and can be called independently from multiple threads.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.Cryptography` |
| Namespace | `CoreSuite.Services.Cryptography` |
| Assembly | `CoreSuite.Cryptography` |
| Target framework | `net8.0` |
| CoreSuite dependencies | None |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
