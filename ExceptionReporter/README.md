# ExceptionReporter

**A UI-independent exception reporting service for capturing, serializing, saving, and emailing structured error reports in .NET 8, included in CoreSuite.**

> [!NOTE]
> ExceptionReporter is one of the independent projects that make up the **CoreSuite** solution. The package contains the report models, SMTP configuration, and infrastructure required to collect and deliver diagnostic information without depending on a user interface.

## Overview

`ExceptionReporter` separates exception capture and report delivery from presentation logic. It converts an `Exception` into a structured `ExceptionReport` containing technical diagnostics, the complete inner-exception chain, optional application information, and the steps performed by the user before the failure.

Reports can be serialized as readable JSON, saved atomically to disk, or sent asynchronously through SMTP. The service can be used by Windows Forms controls such as `CoreSuite.CMessageBox`, background services, console applications, and other .NET components that must report failures without owning a user interface.

## Key features

- Captures exception messages, stack traces, sources, HRESULT values, help links, and custom exception data.
- Includes the complete inner-exception chain in the generated report.
- Adds a user-friendly title and message to the technical diagnostics.
- Accepts application-specific information through a custom dictionary.
- Records optional steps supplied by the user.
- Serializes structured reports as readable JSON.
- Saves reports atomically as UTF-8 JSON files.
- Creates destination directories automatically when required.
- Sends reports asynchronously through SMTP.
- Supports authenticated and unauthenticated SMTP servers.
- Supports automatic, STARTTLS, and SSL-on-connect security modes.
- Optionally verifies internet availability before sending.
- Supports cancellation during asynchronous save and email operations.
- Does not depend on Windows Forms or another user-interface framework.

## Requirements

- .NET 8 or a compatible target framework
- A reference to `CoreSuite.ExceptionReporter`

The service uses `MailKit` for SMTP communication and `CoreSuite.Connectivity` to optionally verify internet availability before sending a report.

## Installation

```powershell
dotnet add package CoreSuite.ExceptionReporter
```

Or add `ExceptionReporter/ExceptionReporter.vbproj` as a project reference when working directly with the CoreSuite solution.

## Quick start

Create an `ExceptionReporter`, capture the exception, and save the resulting report.

```vb
Imports CoreSuite.Services
Dim reporter As New ExceptionReporter()
Try
    RunOperation()
Catch ex As Exception
    Dim report As ExceptionReport = reporter.Capture(
        ex,
        "Application error",
        "The operation could not be completed.",
        New Dictionary(Of String, Object) From {
            {"ApplicationVersion", My.Application.Info.Version.ToString()},
            {"MachineName", Environment.MachineName}
        },
        "The error occurred after selecting a file and clicking Import."
    )
    Await reporter.SaveAsync(report, "Reports\latest-error.json")
End Try
```

The resulting report combines the information supplied by the application with the technical details extracted from the exception.

## Capturing reports

Use `Capture` to convert an `Exception` into an `ExceptionReport`.

```vb
Dim report As ExceptionReport = reporter.Capture(
    ex,
    title:="Import failed",
    message:="The selected file could not be imported.",
    userSteps:="Selected a file and clicked Import."
)
```

All descriptive arguments are optional. When available, `additionalInformations` can include application state that helps technical support reproduce or diagnose the failure.

```vb
Dim additionalInformations As New Dictionary(Of String, Object) From {
    {"ApplicationVersion", My.Application.Info.Version.ToString()},
    {"OperatingSystem", Environment.OSVersion.ToString()},
    {"MachineName", Environment.MachineName}
}
Dim report As ExceptionReport = reporter.Capture(
    ex,
    additionalInformations:=additionalInformations
)
```

Avoid adding passwords, authentication tokens, cryptographic keys, or other secrets to the report.

## Serializing reports

Use `Serialize` when the report must be logged, inspected, attached, or transmitted through another mechanism.

```vb
Dim json As String = reporter.Serialize(report)
```

The generated JSON is formatted for readability and contains the structured report together with its captured exception hierarchy.

## Saving reports

Use `Save` to persist a report synchronously as a UTF-8 JSON file.

```vb
reporter.Save(report, "Reports\latest-error.json")
```

Use `SaveAsync` when file persistence should not block the calling thread.

```vb
Await reporter.SaveAsync(report, "Reports\latest-error.json")
```

The asynchronous method accepts an optional `CancellationToken`:

```vb
Await reporter.SaveAsync(report, "Reports\latest-error.json", cancellationToken)
```

The destination directory is created automatically. The complete JSON payload is written to a temporary file beside the destination and then moved into place, preventing a partially written report from replacing an existing file.

## Sending reports by email

Configure `ExceptionEmailOptions` with the SMTP server and recipient information, then call `SendEmailAsync`.

```vb
Dim emailOptions As New ExceptionEmailOptions With {
    .FromName = "Application",
    .FromEmail = "sender@example.com",
    .ToName = "Technical Support",
    .ToEmail = "support@example.com",
    .Host = "smtp.example.com",
    .Port = 587,
    .Password = Environment.GetEnvironmentVariable("APP_SMTP_PASSWORD"),
    .SecureSocket = ExceptionReporterSecureSocket.StartTls
}
Dim sent As Boolean = Await reporter.SendEmailAsync(report, emailOptions)
```

`SendEmailAsync` serializes the supplied report and sends it as the email content. SMTP, authentication, and cancellation errors are propagated to the caller so the application can log or handle them according to its own requirements.

The method also accepts an optional `CancellationToken`:

```vb
Dim sent As Boolean = Await reporter.SendEmailAsync(report, emailOptions, cancellationToken)
```

## Sending custom content

Use the text overload when the content has already been serialized or was produced by another report formatter.

```vb
Dim content As String = reporter.Serialize(report)
Dim sent As Boolean = Await reporter.SendEmailAsync(content, emailOptions)
```

This overload uses the same SMTP configuration, connectivity verification, authentication behavior, and cancellation support as the `ExceptionReport` overload.

## SMTP authentication

By default, configure the sender address, user name, and password required by the SMTP provider.

```vb
Dim emailOptions As New ExceptionEmailOptions With {
    .FromEmail = "sender@example.com",
    .ToEmail = "support@example.com",
    .Host = "smtp.example.com",
    .Port = 587,
    .UserName = "smtp-user",
    .Password = Environment.GetEnvironmentVariable("APP_SMTP_PASSWORD"),
    .UseAuthentication = True,
    .SecureSocket = ExceptionReporterSecureSocket.StartTls
}
```

When `UserName` is not supplied, `FromEmail` can be used as the authentication identity. Set `UseAuthentication` to `False` for an SMTP server that does not require credentials:

```vb
emailOptions.UseAuthentication = False
```

> [!IMPORTANT]
> Do not store SMTP passwords directly in source code. Use environment variables, user secrets, encrypted configuration, or another secure configuration provider.

## Connectivity verification

When `CheckConnectivity` is enabled, `SendEmailAsync` verifies internet availability before opening the SMTP connection.

```vb
emailOptions.CheckConnectivity = True
Dim sent As Boolean = Await reporter.SendEmailAsync(report, emailOptions)
If Not sent Then
    Debug.WriteLine("Internet availability could not be confirmed.")
End If
```

The method returns `False` when connectivity cannot be confirmed. SMTP, authentication, and cancellation failures remain available to the caller as exceptions.

Disable the preliminary check when the application should always attempt the SMTP connection directly:

```vb
emailOptions.CheckConnectivity = False
```

## SMTP security modes

`ExceptionReporterSecureSocket` determines how the SMTP connection is secured.

| Value | Description |
|---|---|
| `Auto` | Lets the service select the appropriate security behavior for the configured SMTP connection. |
| `StartTls` | Connects to the server and upgrades the connection through STARTTLS. |
| `SslOnConnect` | Establishes the encrypted connection immediately when connecting to the server. |

Use the security mode and port required by the SMTP provider. A common configuration uses `StartTls` with port `587`, while SSL-on-connect commonly uses port `465`.

## Email options

| Property | Description |
|---|---|
| `FromName` | Display name used for the sender. |
| `FromEmail` | Email address used as the sender and default authentication identity. |
| `ToName` | Display name used for the report recipient. |
| `ToEmail` | Email address that receives the report. |
| `Host` | SMTP server host name. |
| `Port` | SMTP server port. |
| `UseAuthentication` | Determines whether the SMTP client authenticates with the server. |
| `UserName` | Optional authentication identity when it differs from `FromEmail`. |
| `Password` | Password or credential used for SMTP authentication. |
| `SecureSocket` | Security mode used by the SMTP connection. |
| `CheckConnectivity` | Determines whether internet availability is checked before connecting. |

## Methods

| Method | Description |
|---|---|
| `Capture(exception, title, message, additionalInformations, userSteps)` | Captures an exception and returns a structured `ExceptionReport`. |
| `Serialize(report)` | Serializes an exception report as readable JSON. |
| `Save(report, filePath)` | Saves a report synchronously as an atomic UTF-8 JSON file. |
| `SaveAsync(report, filePath, cancellationToken)` | Saves a report asynchronously with optional cancellation. |
| `SendEmailAsync(report, emailOptions, cancellationToken)` | Serializes and sends an exception report through SMTP. |
| `SendEmailAsync(content, emailOptions, cancellationToken)` | Sends custom text content through SMTP. |

## Package dependencies

| Package | Purpose |
|---|---|
| `MailKit` | Provides SMTP communication, authentication, and transport security. |
| `CoreSuite.Connectivity` | Verifies internet availability before sending when connectivity checking is enabled. |

The package has no dependency on `CoreSuite.CMessageBox`. User-interface components may reference `CoreSuite.ExceptionReporter`, but the reporting service remains independent from them.

## Integration notes

- Use `ExceptionReporter` to keep exception capture and delivery separate from the application UI.
- Capture the original exception before replacing it with a user-friendly message.
- Save a local report before email delivery when diagnostic information must remain available offline.
- Handle SMTP exceptions when report delivery should not interrupt the application flow.
- Reuse `ExceptionEmailOptions` when several reports use the same SMTP configuration.
- Respect user privacy and applicable data-protection requirements before sending diagnostic information.
- Reports and saved JSON files are not encrypted automatically.
- Do not include passwords, access tokens, connection strings, cryptographic keys, or other secrets in `additionalInformations`.

## License

CoreSuite is licensed under the MIT License.