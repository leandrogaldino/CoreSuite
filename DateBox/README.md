# DateBox

**Culture-aware date input with a drop-down calendar.**

> [!NOTE]
> DateBox is one of the controls or components included in the **CoreSuite** solution. The package depends on `CoreSuite.ControlContainer and CoreSuite.DateTimeBoxBase`. NuGet installs the required package or packages automatically.

## Overview

`DateBox` extends `DateTimeBoxBase` and provides culture-aware date input with a drop-down calendar.

It is designed for .NET 8 Windows Forms applications and can be configured in code or through the Visual Studio designer where designer support applies.

## Features

* Selected date; DateTime.MinValue clears the control.
* Indicates whether a valid date is present.
* Culture used for mask, parsing and display.
* Optional custom calendar image.
* Uses the standard Windows Forms event and component model.
* Includes English XML documentation for the public API.
* Requires no third-party packages.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.DateBox
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.DateBox
```

## Namespace

```vb
Imports CoreSuite.Controls
```

## Quick start

```vb
Imports CoreSuite.Controls

Dim BirthDateBox As New DateBox With {
    .DateCulture = Globalization.CultureInfo.GetCultureInfo("pt-BR"),
    .Date = Date.Today
}
AddHandler BirthDateBox.DateValueChanged, Sub() Console.WriteLine(BirthDateBox.Date)
```

## Designer usage

After installing or referencing the package, add `DateBox` from the Toolbox or create it in code. Properties configured in the Windows Forms Designer are serialized into `InitializeComponent` and remain available at runtime.

## API reference

### `DateBox`

Represents culture-aware date input with a drop-down calendar.

```vb
Public Class DateBox
    Inherits DateTimeBoxBase
```

### Main members

| Member | Behavior |
| --- | --- |
| `Date` | selected date; DateTime.MinValue clears the control. |
| `HasDate` | indicates whether a valid date is present. |
| `DateCulture` | culture used for mask, parsing and display. |
| `ButtonImage` | optional custom calendar image. |
| `DateValueChanged` | raised when the date changes. |

## Behavior and validation

* Property changes take effect immediately unless a method explicitly starts or applies an operation.
* Callers should validate user input and referenced objects before performing application-specific work.
* Events use the standard Windows Forms event model and should be handled on the UI thread.

## Standard functionality

Because `DateBox` inherits from `DateTimeBoxBase`, the standard public members of that base type remain available unless the control intentionally hides or overrides them.

## Resource and lifetime considerations

* Dispose components, controls, images, fonts and menus created dynamically when your application owns them.
* Keep referenced controls alive for as long as a component is attached to them.
* File-based operations should receive valid, accessible paths.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.DateBox` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.DateBox` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| CoreSuite dependencies | `CoreSuite.ControlContainer and CoreSuite.DateTimeBoxBase` |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
