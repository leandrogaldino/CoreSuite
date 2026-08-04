# DateTimeBoxBase

**Abstract base for culture-aware date and time text controls.**

> [!NOTE]
> DateTimeBoxBase is one of the controls or components included in the **CoreSuite** solution. The package has no additional CoreSuite package dependencies.

## Overview

`DateTimeBoxBase` extends `MaskedTextBox` and provides abstract base for culture-aware date and time text controls.

It is designed for .NET 8 Windows Forms applications and can be configured in code or through the Visual Studio designer where designer support applies.

## Features

* Generated input mask (read-only).
* Active culture (read-only).
* Culture exposed to derived classes.
* Format supplied by the derived class.
* Uses the standard Windows Forms event and component model.
* Includes English XML documentation for the public API.
* Requires no third-party packages.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

> [!IMPORTANT]
> This is an infrastructure package used by DateBox and TimeBox. Install those packages for normal application use.

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.DateTimeBoxBase
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.DateTimeBoxBase
```

## Namespace

```vb
Imports CoreSuite.Controls
```

## Quick start

```vb
Imports CoreSuite.Controls

Public Class CustomTemporalBox
    Inherits DateTimeBoxBase

    Protected Overrides ReadOnly Property ValueFormat As String
        Get
            Return "HH:mm"
        End Get
    End Property

    Protected Overrides Function CreateMask(Format As String) As String
        Return "00:00"
    End Function
End Class
```

## Creating a derived control

The base class is not intended to be placed directly on a form. Inherit from it when implementing a specialized CoreSuite input control.

## API reference

### `DateTimeBoxBase`

Represents abstract base for culture-aware date and time text controls.

```vb
Public MustInherit Class DateTimeBoxBase
    Inherits MaskedTextBox
```

### Main members

| Member | Behavior |
| --- | --- |
| `Mask` | generated input mask (read-only). |
| `Culture` | active culture (read-only). |
| `ValueCulture` | culture exposed to derived classes. |
| `ValueFormat` | format supplied by the derived class. |
| `SetTemporalValue()` | updates the internal value. |
| `RefreshValueFormat()` | rebuilds mask and formatted text. |
| `TryParseValue(), FormatValue() and NormalizeValue()` | extension points. |
| `CreateMask()` | required mask-generation implementation. |

## Behavior and validation

* Property changes take effect immediately unless a method explicitly starts or applies an operation.
* Callers should validate user input and referenced objects before performing application-specific work.
* Events use the standard Windows Forms event model and should be handled on the UI thread.

## Standard functionality

Because `DateTimeBoxBase` inherits from `MaskedTextBox`, the standard public members of that base type remain available unless the control intentionally hides or overrides them.

## Resource and lifetime considerations

* Dispose components, controls, images, fonts and menus created dynamically when your application owns them.
* Keep referenced controls alive for as long as a component is attached to them.
* File-based operations should receive valid, accessible paths.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.DateTimeBoxBase` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.DateTimeBoxBase` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| CoreSuite dependencies | `None` |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
