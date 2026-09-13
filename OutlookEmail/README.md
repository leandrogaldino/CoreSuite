# CoreSuite OutlookEmail

**A lightweight .NET 8 Windows service for composing pre-filled e-mail messages in Microsoft Outlook Classic, included in CoreSuite.**

> [!NOTE]
> OutlookEmail is one of the independent projects that make up the **CoreSuite** solution. The package contains the message model, operation result types, Outlook availability detection, and the infrastructure required to open a composed e-mail in Microsoft Outlook Classic without sending it automatically.

## Overview

`OutlookEmail` provides a small integration layer for composing e-mail messages through Microsoft Outlook Classic. It supports To, Cc, and Bcc recipients, multiple attachments, HTML or plain-text bodies, and optional preservation of the default Outlook signature.

The service uses Outlook COM automation at runtime without requiring a compile-time reference to `Microsoft.Office.Interop.Outlook`. The message editor is always displayed to the user, who remains responsible for reviewing, editing, and sending the message.

## Key features

- Opens the Microsoft Outlook Classic message editor without sending automatically.
- Supports To, Cc, and Bcc recipients.
- Supports multiple file attachments.
- Supports HTML and plain-text message bodies.
- Preserves the default Outlook signature when requested.
- Inserts the supplied body before the generated Outlook signature.
- Detects whether Microsoft Outlook Classic is available.
- Returns a structured `OutlookEmailResult`.
- Exposes specific operation statuses through `OutlookEmailStatus`.
- Validates that at least one recipient is supplied.
- Validates attachment paths before opening Outlook.
- Uses runtime COM automation without a `Microsoft.Office.Interop.Outlook` package dependency.
- Releases COM objects after the editor is opened.
- Has no runtime dependency on another CoreSuite package.

## Requirements

- .NET 8 or a compatible Windows target framework
- Microsoft Windows
- Microsoft Outlook Classic installed and registered for COM automation
- A reference to `CoreSuite.OutlookEmail`

The new Outlook for Windows does not expose the same COM automation model used by this package.

## Installation

```powershell
dotnet add package CoreSuite.OutlookEmail
```

Or add `OutlookEmail/OutlookEmail.vbproj` as a project reference when working directly with the CoreSuite solution.

## Quick start

Create an `OutlookEmailMessage`, add at least one recipient, and call `OutlookEmail.Display`.

```vb
Imports CoreSuite.Services

Dim message As New OutlookEmailMessage With {
    .Subject = "Monthly report",
    .Body = "<p>Please find the report attached.</p>",
    .IsBodyHtml = True,
    .IncludeDefaultSignature = True
}

message.AddTo("customer@example.com")
message.AddAttachment("C:\Reports\Report.pdf")

Dim result As OutlookEmailResult = OutlookEmail.Display(message)

If Not result.Success Then
    Debug.WriteLine(result.Message)
End If
```

`Display` only opens the Outlook editor. It never sends the message automatically.

## Recipients

Use `AddTo`, `AddCc`, and `AddBcc` to add recipients.

```vb
message.AddTo("customer@example.com")
message.AddCc("manager@example.com")
message.AddBcc("archive@example.com")
```

Multiple recipients can be added to the same field:

```vb
message.AddTo("first@example.com")
message.AddTo("second@example.com")
```

Recipient addresses are joined with semicolon delimiters before being supplied to Outlook.

At least one To, Cc, or Bcc recipient must be present. Otherwise, `Display` returns `OutlookEmailStatus.InvalidMessage`.

## Attachments

Use `AddAttachment` with the full path of each file.

```vb
message.AddAttachment("C:\Reports\Report.pdf")
message.AddAttachment("C:\Reports\Data.xlsx")
```

Every attachment path is validated before Outlook is started.

If any file does not exist, the operation returns `OutlookEmailStatus.InvalidMessage` and the editor is not opened.

## HTML body

Set `IsBodyHtml` to `True` when `Body` contains HTML.

```vb
Dim message As New OutlookEmailMessage With {
    .Subject = "Report",
    .Body = "<p>Hello,</p><p>Please find the <strong>report</strong> attached.</p>",
    .IsBodyHtml = True
}
```

HTML content is assigned through Outlook's `HTMLBody` property.

## Plain-text body

Set `IsBodyHtml` to `False` when the message body contains plain text.

```vb
Dim message As New OutlookEmailMessage With {
    .Subject = "Report",
    .Body = "Please find the report attached.",
    .IsBodyHtml = False
}
```

Plain-text content is assigned through Outlook's `Body` property.

## Default Outlook signature

`IncludeDefaultSignature` is enabled by default.

```vb
message.IncludeDefaultSignature = True
```

When signature preservation is enabled, the Outlook editor is displayed before the supplied body is applied. This allows Outlook to generate the signature configured for the current account.

The supplied body is then inserted before the generated signature.

```vb
Dim message As New OutlookEmailMessage With {
    .Body = "<p>Hello,</p><p>Please find the document attached.</p>",
    .IsBodyHtml = True,
    .IncludeDefaultSignature = True
}
```

Disable signature preservation when the message should contain only the supplied body:

```vb
message.IncludeDefaultSignature = False
```

## Outlook availability

Use `OutlookEmail.IsAvailable` to determine whether Microsoft Outlook Classic is registered for COM automation.

```vb
If Not OutlookEmail.IsAvailable Then
    Debug.WriteLine("Microsoft Outlook Classic is not available.")
    Return
End If
```

The check resolves the `Outlook.Application` COM programmatic identifier without starting Outlook.

## Open the message editor

Use `OutlookEmail.Display` to create the Outlook message and display its editor.

```vb
Dim result As OutlookEmailResult = OutlookEmail.Display(message)
```

The method returns immediately after Outlook accepts and displays the message editor.

The user can then review the recipients, subject, body, attachments, signature, and any other Outlook options before sending.

## Complete example

```vb
Imports CoreSuite.Services

Dim message As New OutlookEmailMessage With {
    .Subject = "Customer report",
    .Body = "<p>Hello,</p><p>Please find the requested report attached.</p>",
    .IsBodyHtml = True,
    .IncludeDefaultSignature = True
}

message.AddTo("customer@example.com")
message.AddCc("manager@example.com")
message.AddAttachment("C:\Reports\CustomerReport.pdf")

Dim result As OutlookEmailResult = OutlookEmail.Display(message)

If Not result.Success Then
    Debug.WriteLine(result.Message)

    If result.Exception IsNot Nothing Then
        Debug.WriteLine(result.Exception)
    End If
End If
```

## Operation result

`OutlookEmail.Display` returns an `OutlookEmailResult` describing the outcome of the operation.

```vb
Dim result As OutlookEmailResult = OutlookEmail.Display(message)

If result.Success Then
    Debug.WriteLine("Outlook editor opened.")
Else
    Debug.WriteLine(result.Message)
End If
```

## Main properties

### OutlookEmailMessage

| Property | Default | Description |
|---|---:|---|
| `ToRecipients` | Empty collection | Recipients placed in the To field. |
| `CcRecipients` | Empty collection | Recipients placed in the Cc field. |
| `BccRecipients` | Empty collection | Recipients placed in the Bcc field. |
| `Attachments` | Empty collection | Full paths of files attached to the message. |
| `Subject` | `String.Empty` | Message subject. |
| `Body` | `String.Empty` | Message body. |
| `IsBodyHtml` | `True` | Indicates whether `Body` contains HTML. |
| `IncludeDefaultSignature` | `True` | Preserves the default Outlook signature. |

### OutlookEmailResult

| Property | Description |
|---|---|
| `Status` | Status returned by the operation. |
| `Success` | Indicates whether the Outlook editor was opened successfully. |
| `Message` | User-readable description of the result. |
| `Exception` | Exception raised during the operation, when available. |

## Methods

### OutlookEmailMessage

| Method | Description |
|---|---|
| `AddTo(address)` | Adds a recipient to the To field and returns the current message. |
| `AddCc(address)` | Adds a recipient to the Cc field and returns the current message. |
| `AddBcc(address)` | Adds a recipient to the Bcc field and returns the current message. |
| `AddAttachment(filePath)` | Adds a file attachment and returns the current message. |

### OutlookEmail

| Method | Description |
|---|---|
| `Display(message)` | Creates the Outlook message and opens the editor without sending it. |

## Operation statuses

| Status | Description |
|---|---|
| `Success` | The Outlook editor was opened successfully. |
| `OutlookNotAvailable` | Microsoft Outlook Classic is not available for COM automation. |
| `InvalidMessage` | The supplied message contains invalid data. |
| `Failed` | An unexpected error occurred while preparing or displaying the message. |

## Validation behavior

Before Outlook is started, `Display` validates the supplied message.

The operation is considered invalid when:

- the message instance is `Nothing`;
- no To, Cc, or Bcc recipient is present;
- an attachment path is empty;
- an attachment file does not exist.

Validation failures return `OutlookEmailStatus.InvalidMessage` instead of throwing an exception to the calling application.

Unexpected Outlook or COM failures return `OutlookEmailStatus.Failed` and preserve the original exception in `OutlookEmailResult.Exception`.

## COM behavior

The package resolves Outlook through the following COM programmatic identifier:

```text
Outlook.Application
```

No compile-time reference to `Microsoft.Office.Interop.Outlook` is required.

The service creates the Outlook application object, creates a mail item, fills the requested fields, displays the editor, and releases the COM references it owns.

Releasing the local COM references does not close the displayed editor.

## Outlook compatibility

The package is designed for Microsoft Outlook Classic on Windows.

Microsoft Office desktop installations that expose the traditional Outlook COM automation model can normally be used with this package.

The new Outlook for Windows uses a different application architecture and does not provide the same COM automation API.

## Application integration

The package does not display Windows Forms messages and does not depend on CoreSuite UI controls.

Applications are responsible for deciding how an operation result should be presented.

For example:

```vb
Dim result As OutlookEmailResult = OutlookEmail.Display(message)

If Not result.Success Then
    CMessageBox.Show(result.Message, CMessageBoxType.Warning)
End If
```

This keeps `CoreSuite.OutlookEmail` independent from any specific user-interface implementation.

## Security notes

- The package never sends a message automatically.
- The user remains responsible for the final send action.
- No SMTP server credentials are stored or processed.
- No Outlook account password is required by the package.
- Account authentication remains the responsibility of Microsoft Outlook.
- Attachments are referenced from local file paths.
- Message content and recipient addresses are passed only to the locally installed Outlook application.
- Applications should not place untrusted HTML into `Body` without validating or controlling its source.

## Important behavior

- `Display` always opens the editor instead of sending directly.
- At least one recipient is required.
- Empty recipient strings are ignored by the message helper methods.
- Attachment files must exist before the editor is opened.
- `IsBodyHtml` selects between Outlook's `HTMLBody` and `Body` properties.
- `IncludeDefaultSignature` is enabled by default.
- Signature preservation requires displaying the editor before replacing the body.
- The supplied body is inserted before the generated signature.
- The package has no dependency on Windows Forms.
- The package has no dependency on another CoreSuite package.
- The package has no dependency on `Microsoft.Office.Interop.Outlook`.
- COM failures are returned through `OutlookEmailResult` instead of being silently ignored.
- The new Outlook for Windows is not compatible with the COM automation model used by this package.

## License

CoreSuite is licensed under the MIT License.
