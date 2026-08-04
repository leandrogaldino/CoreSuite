# CMessageBox

**A customizable Windows Forms message box with typed dialogs, exception details and optional email reporting.**

> [!NOTE]
> CMessageBox is part of the **CoreSuite** solution. It provides a consistent alternative to the standard WinForms `MessageBox` for information, success, warning, question and error messages.

## Overview

`CMessageBox` exposes shared `Show` overloads that can be called throughout a Windows Forms application. Its global `Options` property controls fonts, colors, icons, exception-detail visibility, contextual diagnostic data and SMTP reporting. Exception capture, serialization and delivery are delegated to `CoreSuite.ExceptionReporter`, keeping reporting logic independent from the interface.

Question dialogs return `DialogResult.Yes` or `DialogResult.No`. Information, success, warning and error dialogs return `DialogResult.OK` when confirmed.

## Features

* Information, success, warning, question and error message types.
* Multiple `Show` overloads for simple and detailed usage.
* Custom title and message fonts and colors.
* Custom 64 × 64 message icons.
* Optional expandable exception details serialized as JSON.
* Additional contextual information in exception reports.
* Optional asynchronous exception reporting through SMTP.
* Configurable SMTP security mode.
* Designed for .NET 8 Windows Forms applications.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.CMessageBox
```

Or search for `CoreSuite.CMessageBox` in the Visual Studio NuGet Package Manager.

## Namespace

```vb
Imports CoreSuite.Controls
```

## Quick start

```vb
Imports CoreSuite.Controls
CMessageBox.Show("The operation was completed successfully.", "Completed", CMessageBoxType.Success)
```

### Information message

```vb
CMessageBox.Show("Your changes were saved.")
```

### Question message

```vb
Dim Result As DialogResult = CMessageBox.Show("Do you want to continue?", "Confirmation", CMessageBoxType.Question)
If Result = DialogResult.Yes Then
    ContinueOperation()
End If
```

### Error with exception details

```vb
Try
    RunOperation()
Catch Ex As Exception
    CMessageBox.Show("The operation could not be completed.", "Application error", Ex)
End Try
```

## Global configuration

Configure `CMessageBox.Options` once during application startup:

```vb
CMessageBox.Options = New CMessageBoxOptions With {
    .ShowExceptionDetails = True,
    .TitleFont = New Font("Segoe UI", 11.25F, FontStyle.Bold),
    .TitleForeColor = Color.FromArgb(32, 32, 32),
    .MessageFont = New Font("Segoe UI", 9.75F),
    .MessageForeColor = Color.FromArgb(64, 64, 64),
    .AdditionalInformations = New Dictionary(Of String, Object) From {
        {"ApplicationVersion", Application.ProductVersion},
        {"MachineName", Environment.MachineName}
    }
}
```

The configured instance is shared by every subsequent call to `CMessageBox.Show`.

## Custom icons

Use the image properties in `CMessageBoxOptions` to configure the icon for each message type:

```vb
CMessageBox.Options.InformationImage = My.Resources.InformationIcon
CMessageBox.Options.SuccessImage = My.Resources.SuccessIcon
CMessageBox.Options.WarningImage = My.Resources.WarningIcon
CMessageBox.Options.QuestionImage = My.Resources.QuestionIcon
CMessageBox.Options.ErrorImage = My.Resources.ErrorIcon
```

Use 64 × 64 pixel images for the layout expected by the component.

## Exception reporting by email

Set `ExceptionEmail` to send serialized exception details through an SMTP server when an error is displayed:

```vb
CMessageBox.Options.ExceptionEmail = New CMessageBoxExceptionEmail With {
    .FromName = "Application",
    .FromEmail = "sender@example.com",
    .ToName = "Support",
    .ToEmail = "support@example.com",
    .Host = "smtp.example.com",
    .Port = 587,
    .Password = Environment.GetEnvironmentVariable("APP_SMTP_PASSWORD"),
    .SecureSocket = CMessageBoxSecureSocket.StartTls
}
```

> [!IMPORTANT]
> Do not store SMTP passwords directly in source code. Use environment variables, user secrets or another secure configuration provider.

Email sending runs asynchronously and communication failures do not interrupt the message-box flow.

## API reference

### `CMessageBox`

```vb
Public Class CMessageBox
```

#### Options

```vb
Public Shared Property Options As CMessageBoxOptions
```

#### Show overloads

```vb
Public Shared Function Show(Message As String) As DialogResult
Public Shared Function Show(Message As String, Title As String) As DialogResult
Public Shared Function Show(Message As String, MessageType As CMessageBoxType) As DialogResult
Public Shared Function Show(Message As String, Exception As Exception) As DialogResult
Public Shared Function Show(Message As String, Title As String, Exception As Exception) As DialogResult
Public Shared Function Show(Message As String, Title As String, MessageType As CMessageBoxType, Optional Exception As Exception = Nothing) As DialogResult
```

An exception can only be supplied when `MessageType` is `CMessageBoxType.Error`; otherwise, an `ArgumentException` is thrown.

### `CMessageBoxType`

| Value | Buttons | Result |
| --- | --- | --- |
| `Information` | OK | `DialogResult.OK` |
| `Success` | OK | `DialogResult.OK` |
| `Error` | OK; optional Details | `DialogResult.OK` |
| `Warning` | OK | `DialogResult.OK` |
| `Question` | Yes and No | `DialogResult.Yes` or `DialogResult.No` |

### `CMessageBoxOptions`

| Property | Type | Purpose |
| --- | --- | --- |
| `ShowExceptionDetails` | `Boolean` | Shows the expandable details area for errors with exceptions. |
| `ExceptionEmail` | `CMessageBoxExceptionEmail` | Defines optional SMTP reporting. |
| `TitleFont` | `Font` | Controls the title font. |
| `TitleForeColor` | `Color` | Controls the title color. |
| `MessageFont` | `Font` | Controls the message font. |
| `MessageForeColor` | `Color` | Controls the message color. |
| `AdditionalInformations` | `Dictionary(Of String, Object)` | Adds contextual values to exception reports. |
| `ErrorImage` | `Image` | Icon used for error messages. |
| `SuccessImage` | `Image` | Icon used for success messages. |
| `InformationImage` | `Image` | Icon used for information messages. |
| `WarningImage` | `Image` | Icon used for warning messages. |
| `QuestionImage` | `Image` | Icon used for question messages. |

### `CMessageBoxSecureSocket`

| Value | Behavior |
| --- | --- |
| `Auto` | Allows MailKit to select the connection security mode. |
| `StartTls` | Requires an upgrade to TLS with STARTTLS. |
| `StartTlsWhenAvailable` | Uses STARTTLS when the server supports it. |
| `SslOnConnect` | Establishes TLS immediately when connecting. |
| `None` | Uses an unencrypted SMTP connection. |

## Package dependencies

The package uses:

* `CoreSuite.ExceptionReporter`
* `CoreSuite.ControlContainer`
* `CoreSuite.NoFocusCueButton`

NuGet resolves declared package dependencies automatically when all CoreSuite dependencies are available in the configured package source.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.CMessageBox` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.CMessageBox` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| Reporting dependency | CoreSuite.ExceptionReporter |

## License

This package is distributed under the MIT License.
