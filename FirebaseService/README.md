# CoreSuite FirebaseService

**Use Firebase Authentication, Cloud Firestore and Cloud Storage from .NET 8 through one lightweight, REST-based service.**

> [!NOTE]
> `CoreSuite.FirebaseService` is part of the CoreSuite solution. It targets plain .NET 8, does not depend on Windows Forms and has no dependency on another CoreSuite package.

## Overview

`CoreSuite.FirebaseService` provides one configured entry point for three Firebase services:

- `FirebaseAuth` signs in existing users with email and password, keeps the ID and refresh tokens in memory and renews the session before the ID token expires.
- `FirebaseFirestore` creates, reads, lists, queries and deletes documents in Cloud Firestore.
- `FirebaseStorage` uploads, downloads and deletes objects in Cloud Storage for Firebase with byte-based progress events.

The three modules share one `HttpClient` and one authenticated session. Firestore and Storage requests are made on behalf of the signed-in Firebase user, so access is controlled by the Security Rules configured in the Firebase project.

The package communicates directly with the public Firebase REST endpoints. It does not require an official Firebase SDK or service-account credentials.

## Features

- Central `FirebaseService` facade for Authentication, Firestore and Storage.
- Validated configuration through `FirebaseOptions`.
- Support for the default or a named Firestore database.
- Configurable request and file-transfer timeouts.
- Optional caller-provided `HttpClient` for dependency injection, testing and connection reuse.
- Email-and-password sign-in through Firebase Authentication.
- In-memory ID token, refresh token, user ID, email and expiration tracking.
- Thread-safe session renewal with `SemaphoreSlim`.
- Automatic renewal shortly before the ID token expires.
- Correctly paginated Firestore document and collection listing.
- Distinction between HTTP 404 and authorization, quota or server failures.
- `FirestoreDocument` model with metadata separated from user fields.
- Firestore strings, Boolean values, integers, floating-point numbers, timestamps, bytes, references, geographic points, arrays, maps and null values.
- Structured Firestore queries with one or more `AND` filters.
- Safe JSON generation and parsing through `System.Text.Json`.
- Storage upload and download progress with transferred and total byte counts.
- Request-scoped authorization headers that are safe for concurrent operations.
- Atomic downloads that preserve the previous destination file when a transfer fails.
- Cancellation support on every network operation.
- Structured `FirebaseException` errors with service area, HTTP status and Firebase error code.
- Predictable resource cleanup through `IDisposable`.
- No external package dependencies.

## Requirements

- .NET 8 (`net8.0`)
- A Firebase project
- Firebase Authentication with the **Email/Password** provider enabled
- A Cloud Firestore database for Firestore operations
- A Cloud Storage for Firebase bucket for Storage operations
- Security Rules that permit the signed-in user to perform the required operations
- An internet connection

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.FirebaseService
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.FirebaseService
```

## Namespace

```vb
Imports CoreSuite.Services
```

## Changes from the original project implementation

This package preparation also corrects the original service contract. If an application already consumes the source project, review these changes before replacing it:

| Original behavior | Corrected API |
| --- | --- |
| `GetCollectionsAsync` used an incompatible request and response shape. | It now returns `Task(Of List(Of String))`, uses `POST ...:listCollectionIds` and follows every page token. |
| `GetAllDocumentsAsync` returned only the first response page. | It now follows every `nextPageToken`. |
| Firestore reads mixed document metadata into the user dictionary. | Read and query methods now return `FirestoreDocument`, with metadata and `Fields` separated. |
| `UploadFile` and `DownloadFile` omitted the asynchronous suffix. | Use `UploadFileAsync` and `DownloadFileAsync`. |
| Storage progress events exposed untyped arguments. | Events now use `FirebaseTransferProgressChangedEventArgs`. |
| Several failures returned `Nothing`, `False` or an empty string. | Unexpected HTTP and response failures now throw `FirebaseException`; only documented 404 results remain non-exceptional. |
| The service depended on `CoreSuite.Helpers` for partial JSON extraction. | JSON is handled by `System.Text.Json`, so the package has no CoreSuite dependency. |

These changes intentionally make incorrect or ambiguous behavior explicit before the public `1.0.0` API is published.

## Firebase project configuration

The service normally needs these three values:

| Value | Example | Purpose |
| --- | --- | --- |
| API key | `AIza...` | Identifies the Firebase project in Authentication requests. |
| Project ID | `my-application` | Identifies the project in Firestore requests. |
| Storage bucket | `my-application.firebasestorage.app` | Identifies the bucket used for object operations. |

For a registered Firebase web app, these values correspond to `apiKey`, `projectId` and `storageBucket` in the Firebase configuration object. The bucket may be passed with or without the `gs://` prefix.

### Enable email and password sign-in

In the Firebase console, open **Authentication**, choose **Sign-in method** and enable **Email/Password**. This package signs in existing users; it does not create users, send verification messages or reset passwords.

See the official [Firebase Authentication REST documentation](https://firebase.google.com/docs/reference/rest/auth).

### Configure Firestore and Storage rules

An API key identifies a Firebase project but does not grant access to its data. Authentication and Firebase Security Rules decide which documents and objects the current user can read or change.

Review the official [Cloud Firestore Security Rules](https://firebase.google.com/docs/firestore/security/get-started) and [Cloud Storage Security Rules](https://firebase.google.com/docs/storage/security) guides before using the package in production.

## Quick start

The following example signs in, creates a Firestore document and uploads a related file:

```vb
Imports CoreSuite.Services

Dim Options As New FirebaseOptions(
    "YOUR_FIREBASE_API_KEY",
    "YOUR_PROJECT_ID",
    "YOUR_STORAGE_BUCKET")

Using Firebase As New FirebaseService(Options)
    Await Firebase.Auth.LoginAsync("user@example.com", "user-password")

    Dim Customer As New Dictionary(Of String, Object) From {
        {"name", "Maria"},
        {"active", True},
        {"creditLimit", 1500D},
        {"createdAt", DateTime.UtcNow}
    }

    Dim DocumentId As String = Await Firebase.Firestore.SaveDocumentAsync(
        "customers",
        Nothing,
        Customer)

    Dim DownloadToken As String = Await Firebase.Storage.UploadFileAsync(
        "C:\Files\contract.pdf",
        $"contracts/{DocumentId}.pdf",
        "application/pdf")
End Using
```

> [!IMPORTANT]
> Call `LoginAsync` before using Firestore or Storage. Authenticated operations automatically validate and, when possible, renew the current ID token.

## Create and initialize the service

### Use the convenient constructor

```vb
Using Firebase As New FirebaseService(ApiKey, ProjectId, StorageBucket)
    Await Firebase.Auth.LoginAsync(UserEmail, UserPassword)
End Using
```

### Use FirebaseOptions

`FirebaseOptions` is recommended when the application needs a custom database or timeout:

```vb
Dim Options As New FirebaseOptions(ApiKey, ProjectId, StorageBucket) With {
    .DatabaseId = "(default)",
    .RequestTimeout = TimeSpan.FromSeconds(30),
    .TransferTimeout = TimeSpan.FromMinutes(10)
}

Using Firebase As New FirebaseService(Options)
    ' Use Firebase.Auth, Firebase.Firestore and Firebase.Storage.
End Using
```

`RequestTimeout` applies to Authentication and Firestore operations. `TransferTimeout` applies to uploads and downloads and defaults to `Timeout.InfiniteTimeSpan`, because large files should normally be controlled with an explicit `CancellationToken` rather than a short global timeout.

### Initialize later

The parameterless constructor is useful when configuration is loaded after application startup:

```vb
Using Firebase As New FirebaseService()
    Firebase.Initialize(Options)
    Await Firebase.Auth.LoginAsync(UserEmail, UserPassword)
End Using
```

`IsInitialized` indicates whether initialization has completed. Accessing `Auth`, `Firestore`, `Storage` or `Options` before initialization throws `InvalidOperationException` instead of returning `Nothing`.

Calling `Initialize` again replaces the modules and clears the previous in-memory session. The old internally managed `HttpClient` is disposed.

### Supply an HttpClient

Applications that already use dependency injection can provide an `HttpClient`:

```vb
Dim Firebase As New FirebaseService(Options, SharedHttpClient)
```

The caller remains responsible for disposing a supplied client. Its own `HttpClient.Timeout` also remains active, in addition to the operation timeout configured in `FirebaseOptions`.

## Authentication

### Sign in

```vb
Await Firebase.Auth.LoginAsync(UserEmail, UserPassword)

If Firebase.Auth.IsLoggedIn Then
    Console.WriteLine($"Signed in as {Firebase.Auth.Email}")
    Console.WriteLine($"Firebase user ID: {Firebase.Auth.UserId}")
End If
```

`LoginAsync` serializes credentials with `System.Text.Json`, so quotes, backslashes and other valid characters cannot corrupt the request body.

The following session information is available:

| Member | Meaning |
| --- | --- |
| `IsLoggedIn` | `True` only while a non-expired ID token is stored. |
| `CanRefreshSession` | `True` when a refresh token is available in memory. |
| `UserId` | The Firebase user identifier returned by Authentication. |
| `Email` | The email returned by the successful sign-in response. |
| `TokenExpirationUtc` | The current ID token expiration time in UTC. |

The ID token, refresh token and password are intentionally not exposed by the public API. Do not write credentials or tokens to logs, exception reports or plain-text files.

### Automatic session renewal

Firestore and Storage operations call `EnsureValidTokenAsync` internally. When the ID token is missing or within one minute of expiration, one request exchanges the refresh token for a new session. Concurrent callers wait for that renewal instead of sending several refresh requests at the same time.

Most applications do not need to call this method directly:

```vb
Await Firebase.Auth.EnsureValidTokenAsync()
```

### Refresh immediately

```vb
Dim Refreshed As Boolean = Await Firebase.Auth.RefreshSessionAsync()

If Not Refreshed Then
    ' No refresh token is available; sign in again.
End If
```

`False` means that no refresh token was stored. A rejected token, malformed response, network failure or timeout produces a `FirebaseException`, preserving the reason instead of silently returning `False`.

### Log out

```vb
Firebase.Auth.Logout()
```

`Logout` clears the session held by this service instance. It does not delete the Firebase account or revoke sessions on other devices.

## Firestore path rules

Firestore paths alternate between collections and documents:

```text
collection/document/collection/document
```

- A collection path has an odd number of segments: `customers` or `customers/customer-1/orders`.
- A document path has an even number of segments: `customers/customer-1`.
- Methods that accept a separate `DocumentId` require a single identifier without `/` or `\`.
- Leading and trailing `/` characters are ignored, but empty middle segments are rejected.

Invalid paths fail locally with `ArgumentException` before an HTTP request is sent.

## FirestoreDocument

Read and query methods return `FirestoreDocument` objects:

| Property | Meaning |
| --- | --- |
| `Id` | The final document identifier. |
| `Path` | The complete path relative to the database document root. |
| `ResourceName` | The complete Firestore resource name. |
| `Fields` | A read-only dictionary containing only stored user fields. |
| `CreateTimeUtc` | The creation timestamp when returned by Firestore. |
| `UpdateTimeUtc` | The last update timestamp when returned by Firestore. |

Metadata is not inserted into `Fields`. A real field named `firestore_document_id` therefore remains valid and cannot collide with a synthetic package value.

Fields can be accessed through the dictionary or the default property:

```vb
Dim Customer As FirestoreDocument = Await Firebase.Firestore.GetDocumentAsync(
    "customers",
    "customer-1")

If Customer IsNot Nothing Then
    Dim Name As String = CStr(Customer("name"))
    Console.WriteLine($"{Customer.Id}: {Name}")
End If
```

## Firestore document operations

### Create a document with an automatic ID

Pass `Nothing`, `String.Empty` or whitespace as the document ID:

```vb
Dim Fields As New Dictionary(Of String, Object) From {
    {"description", "Keyboard"},
    {"price", 199.9D},
    {"available", True}
}

Dim DocumentId As String = Await Firebase.Firestore.SaveDocumentAsync(
    "products",
    Nothing,
    Fields)
```

### Create or replace a document with a known ID

```vb
Dim DocumentId As String = Await Firebase.Firestore.SaveDocumentAsync(
    "products",
    "keyboard-001",
    Fields)
```

When an ID is supplied, the method sends a Firestore `PATCH` request without an update mask. Treat `Fields` as the complete document field set. Existing fields omitted from the new dictionary are not preserved.

### Read a document

```vb
Dim Product As FirestoreDocument = Await Firebase.Firestore.GetDocumentAsync(
    "products",
    "keyboard-001")
```

`GetDocumentAsync` returns `Nothing` only for HTTP 404. Permission errors, invalid authentication, quota errors and server failures throw `FirebaseException`.

### Check existence

```vb
Dim Exists As Boolean = Await Firebase.Firestore.DocumentExistsAsync(
    "products",
    "keyboard-001")
```

The result is `False` only for HTTP 404. Other failures are not hidden.

### List every document

```vb
Dim Products As List(Of FirestoreDocument) =
    Await Firebase.Firestore.GetAllDocumentsAsync("products")
```

`GetAllDocumentsAsync` follows every `nextPageToken` returned by Firestore. The name therefore reflects the actual behavior; the method does not stop at the first page.

Nested collections are also supported:

```vb
Dim Orders = Await Firebase.Firestore.GetAllDocumentsAsync(
    "customers/customer-1/orders")
```

### Delete a document

```vb
Dim Deleted As Boolean = Await Firebase.Firestore.DeleteDocumentAsync(
    "products",
    "keyboard-001")
```

The method returns `False` when the document does not exist and throws for other Firebase errors.

## List collections

List root collections:

```vb
Dim RootCollections As List(Of String) =
    Await Firebase.Firestore.GetCollectionsAsync()
```

List the collections directly beneath a document:

```vb
Dim Subcollections As List(Of String) =
    Await Firebase.Firestore.GetCollectionsAsync("customers/customer-1")
```

The implementation uses the official `POST ...:listCollectionIds` contract and follows every `nextPageToken` returned by Firestore.

## Firestore queries

Create one or more filters and pass them to `QueryCompositeAsync`:

```vb
Dim Filters As New List(Of FirestoreFilter) From {
    New FirestoreFilter("active", FirestoreOperator.Equal, True),
    New FirestoreFilter("creditLimit", FirestoreOperator.GreaterThanOrEqual, 1000D)
}

Dim Customers As List(Of FirestoreDocument) =
    Await Firebase.Firestore.QueryCompositeAsync("customers", Filters)
```

All supplied filters are combined with logical `AND`.

Available operators:

| Enum value | Firestore operator |
| --- | --- |
| `Equal` | `EQUAL` |
| `NotEqual` | `NOT_EQUAL` |
| `LessThan` | `LESS_THAN` |
| `LessThanOrEqual` | `LESS_THAN_OR_EQUAL` |
| `GreaterThan` | `GREATER_THAN` |
| `GreaterThanOrEqual` | `GREATER_THAN_OR_EQUAL` |
| `ArrayContains` | `ARRAY_CONTAINS` |
| `ArrayContainsAny` | `ARRAY_CONTAINS_ANY` |
| `InList` | `IN` |
| `NotInList` | `NOT_IN` |

`ArrayContainsAny`, `InList` and `NotInList` require an enumerable filter value:

```vb
Dim Filters As New List(Of FirestoreFilter) From {
    New FirestoreFilter(
        "status",
        FirestoreOperator.InList,
        New String() {"pending", "approved"})
}
```

Firestore may require composite indexes for some filter combinations. Index requirements are reported by Firebase through `FirebaseException`.

## Supported Firestore values

The package converts values explicitly instead of silently calling `ToString()` for unknown types.

| .NET input | Firestore value | .NET output |
| --- | --- | --- |
| `Nothing` | `nullValue` | `Nothing` |
| `Boolean` | `booleanValue` | `Boolean` |
| Signed and supported unsigned integers | `integerValue` | `Int64` |
| `Single`, `Double`, `Decimal` | `doubleValue` | `Double` |
| `String` | `stringValue` | `String` |
| `DateTime`, `DateTimeOffset` | `timestampValue` | UTC `DateTime` |
| `Byte()` | `bytesValue` | `Byte()` |
| `FirestoreDocumentReference` | `referenceValue` | `FirestoreDocumentReference` |
| `FirestoreGeoPoint` | `geoPointValue` | `FirestoreGeoPoint` |
| `IEnumerable` | `arrayValue` | `List(Of Object)` |
| Dictionary with string keys | `mapValue` | `Dictionary(Of String, Object)` |

Unsupported types throw `NotSupportedException`. Unsigned integers larger than `Int64.MaxValue` throw `OverflowException`. Non-finite floating-point inputs are rejected. Firestore also does not permit an array to directly contain another array; use a map between the two array levels when that structure is required.

### Store bytes and a geographic point

```vb
Dim Fields As New Dictionary(Of String, Object) From {
    {"thumbnail", File.ReadAllBytes("thumbnail.png")},
    {"location", New FirestoreGeoPoint(-23.5505, -46.6333)}
}
```

### Store a document reference

Create references through the configured Firestore module:

```vb
Dim CustomerReference As FirestoreDocumentReference =
    Firebase.Firestore.CreateDocumentReference("customers/customer-1")

Dim Fields As New Dictionary(Of String, Object) From {
    {"customer", CustomerReference}
}
```

`CreateDocumentReference` validates the relative document path and expands it to the complete resource name for the configured project and database.

## Cloud Storage

Storage paths identify objects inside the configured bucket. They may start with `/`, but they cannot be empty or point only to the bucket root.

### Upload a file

```vb
Dim DownloadToken As String = Await Firebase.Storage.UploadFileAsync(
    "C:\Files\photo.jpg",
    "profiles/user-1/photo.jpg",
    "image/jpeg")
```

The content type defaults to `application/octet-stream` when omitted. The return value is the `downloadTokens` metadata value returned by Firebase Storage. It may be empty when the service does not include a download token.

### Download a file

```vb
Await Firebase.Storage.DownloadFileAsync(
    "profiles/user-1/photo.jpg",
    "C:\Downloads\photo.jpg")
```

The response is first written to a temporary file in the destination directory. The destination is replaced only after the download completes, the stream is flushed and the received byte count matches the declared content length when one was supplied. A failure or cancellation leaves an existing destination file unchanged.

### Delete a file

```vb
Dim Deleted As Boolean = Await Firebase.Storage.DeleteFileAsync(
    "profiles/user-1/photo.jpg")
```

The method returns `False` for HTTP 404 and throws for other errors.

### Track transfer progress

```vb
AddHandler Firebase.Storage.UploadProgressChanged,
    Sub(Sender, e)
        If e.Percentage.HasValue Then
            Console.WriteLine($"Upload: {e.Percentage.Value:F1}%")
        Else
            Console.WriteLine($"Uploaded: {e.BytesTransferred} bytes")
        End If
    End Sub

AddHandler Firebase.Storage.DownloadProgressChanged,
    Sub(Sender, e)
        If e.IsCompleted Then
            Console.WriteLine("Download completed.")
        End If
    End Sub
```

`FirebaseTransferProgressChangedEventArgs` exposes:

- `BytesTransferred`
- `TotalBytes`, when the server or source supplied a total
- `Percentage`, when it can be calculated
- `IsCompleted`

Events are raised on the thread performing the asynchronous operation. A Windows Forms or WPF application must marshal UI updates to its UI thread when necessary.

## Cancellation

Every network method accepts a `CancellationToken`:

```vb
Using CancellationSource As New CancellationTokenSource()
    CancellationSource.CancelAfter(TimeSpan.FromSeconds(15))

    Dim Documents = Await Firebase.Firestore.GetAllDocumentsAsync(
        "products",
        CancellationSource.Token)
End Using
```

Caller cancellation remains an `OperationCanceledException`. Expiration of a configured CoreSuite timeout produces a `FirebaseException` with `ErrorCode` equal to `TIMEOUT`.

## Error handling

Firebase failures are represented by `FirebaseException`:

```vb
Try
    Await Firebase.Firestore.DeleteDocumentAsync("products", "keyboard-001")
Catch ex As FirebaseException
    Console.WriteLine($"Service: {ex.ServiceArea}")
    Console.WriteLine($"HTTP status: {ex.StatusCode}")
    Console.WriteLine($"Firebase code: {ex.ErrorCode}")
    Console.WriteLine(ex.Message)
End Try
```

| Property | Meaning |
| --- | --- |
| `ServiceArea` | `Authentication`, `Firestore` or `Storage`. |
| `StatusCode` | The HTTP response status, or `Nothing` when no response was received. |
| `ErrorCode` | A Firebase code such as `PERMISSION_DENIED`, or a CoreSuite code such as `NETWORK_ERROR`, `TIMEOUT`, `INVALID_RESPONSE` or `AUTHENTICATION_REQUIRED`. |

Argument, path, local file and unsupported-value problems use the normal .NET exception types because they are detected before Firebase is called.

## Concurrency and lifetime

- Normal Authentication, Firestore and Storage operations may run concurrently.
- Authorization is attached to each request; `DefaultRequestHeaders` is never cleared or changed by the package.
- Only one token refresh runs at a time for a service instance.
- Do not call `Initialize` or `Dispose` while operations from the same instance are still running.
- Dispose `FirebaseService` when it created its own `HttpClient`.
- A caller-provided `HttpClient` is never disposed by `FirebaseService`.

## Security considerations

- Store passwords only long enough to call `LoginAsync`.
- Never log passwords, ID tokens or refresh tokens.
- Treat Firebase API keys as project identifiers and restrict them to the APIs your application uses.
- Use restrictive Firestore and Storage Security Rules in production.
- Use HTTPS endpoints only; the package does this by default.
- The session is stored only in memory and is lost when the service is disposed or the application exits.
- This client uses end-user Firebase Authentication and must not be used as an administrative replacement for the Firebase Admin SDK.

## Current scope

Version `1.0.0` intentionally focuses on a small, predictable API:

- Authentication supports existing email-and-password users only.
- Sessions are not persisted or restored across application restarts.
- Firestore queries combine filters with `AND`; ordering, cursors, limits, projections, transactions, batched writes and listeners are outside the current API.
- Storage exposes simple media upload, download and deletion; resumable uploads and object listing are outside the current API.
- Firebase App Check is not currently integrated.

These are scope boundaries rather than silent partial implementations. Document and collection listing methods do paginate until completion.

## API summary

### FirebaseService

- `New()`
- `New(apiKey, projectId, storageBucket)`
- `New(options)`
- `New(options, httpClient)`
- `Initialize(...)`
- `IsInitialized`
- `Options`
- `Auth`
- `Firestore`
- `Storage`
- `Dispose()`

### FirebaseAuth

- `LoginAsync(...)`
- `EnsureValidTokenAsync(...)`
- `RefreshSessionAsync(...)`
- `Logout()`
- `IsLoggedIn`
- `CanRefreshSession`
- `UserId`
- `Email`
- `TokenExpirationUtc`

### FirebaseFirestore

- `GetCollectionsAsync(...)`
- `GetDocumentAsync(...)`
- `GetAllDocumentsAsync(...)`
- `QueryCompositeAsync(...)`
- `SaveDocumentAsync(...)`
- `DeleteDocumentAsync(...)`
- `DocumentExistsAsync(...)`
- `CreateDocumentReference(...)`

### FirebaseStorage

- `UploadFileAsync(...)`
- `DownloadFileAsync(...)`
- `DeleteFileAsync(...)`
- `UploadProgressChanged`
- `DownloadProgressChanged`

## License

CoreSuite is distributed under the MIT license.
