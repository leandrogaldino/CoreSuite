# TimeBox

**Culture-aware time input with a drop-down time selector.**

> [!NOTE]
> TimeBox is one of the controls or components included in the **CoreSuite** solution. The package depends on `CoreSuite.ControlContainer and CoreSuite.DateTimeBoxBase`. NuGet installs the required package or packages automatically.

## Overview

`TimeBox` extends `DateTimeBoxBase` and provides culture-aware time input with a drop-down time selector.

It is designed for .NET 8 Windows Forms applications and can be configured in code or through the Visual Studio designer where designer support applies.

## Features

* Selected TimeSpan.
* Indicates whether a valid time is present.
* Clears the current value.
* Culture used for mask, parsing and display.
* Uses the standard Windows Forms event and component model.
* Includes English XML documentation for the public API.
* Requires no third-party packages.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.TimeBox
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.TimeBox
```

## Namespace

```vb
Imports CoreSuite.Controls
```

## Quick start

```vb
Imports CoreSuite.Controls

Dim StartTimeBox As New TimeBox With {
    .TimeCulture = Globalization.CultureInfo.GetCultureInfo("pt-BR"),
    .ShowSeconds = False,
    .Time = New TimeSpan(14, 30, 0)
}
AddHandler StartTimeBox.TimeValueChanged, Sub() Console.WriteLine(StartTimeBox.Time)
```

## Designer usage

After installing or referencing the package, add `TimeBox` from the Toolbox or create it in code. Properties configured in the Windows Forms Designer are serialized into `InitializeComponent` and remain available at runtime.

## API reference

### `TimeBox`

Represents culture-aware time input with a drop-down time selector.

```vb
Public Class TimeBox
    Inherits DateTimeBoxBase
```

### Main members

| Member | Behavior |
| --- | --- |
| `Time` | selected TimeSpan. |
| `HasTime` | indicates whether a valid time is present. |
| `ClearTime()` | clears the current value. |
| `TimeCulture` | culture used for mask, parsing and display. |
| `ShowSeconds` | includes seconds in the format. |
| `ButtonImage` | optional custom clock image. |
| `TimeValueChanged` | raised when the time changes. |

## Behavior and validation

* Property changes take effect immediately unless a method explicitly starts or applies an operation.
* Callers should validate user input and referenced objects before performing application-specific work.
* Events use the standard Windows Forms event model and should be handled on the UI thread.

## Standard functionality

Because `TimeBox` inherits from `DateTimeBoxBase`, the standard public members of that base type remain available unless the control intentionally hides or overrides them.

## Resource and lifetime considerations

* Dispose components, controls, images, fonts and menus created dynamically when your application owns them.
* Keep referenced controls alive for as long as a component is attached to them.
* File-based operations should receive valid, accessible paths.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.TimeBox` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.TimeBox` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| CoreSuite dependencies | `CoreSuite.ControlContainer and CoreSuite.DateTimeBoxBase` |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
