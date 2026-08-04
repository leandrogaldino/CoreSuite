# PercentageBox

**Culture-aware percentage input with percentage and fractional representations.**

> [!NOTE]
> PercentageBox is one of the controls or components included in the **CoreSuite** solution. The package depends on `CoreSuite.NumericBoxBase`. NuGet installs the required package or packages automatically.

## Overview

`PercentageBox` extends `NumericBoxBase` and provides culture-aware percentage input with percentage and fractional representations.

It is designed for .NET 8 Windows Forms applications and can be configured in code or through the Visual Studio designer where designer support applies.

## Features

* Displayed percentage value (15 means 15%).
* Read-only fractional equivalent (0.15).
* Culture used for parsing and formatting.
* Fractional digit count.
* Uses the standard Windows Forms event and component model.
* Includes English XML documentation for the public API.
* Requires no third-party packages.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.PercentageBox
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.PercentageBox
```

## Namespace

```vb
Imports CoreSuite.Controls
```

## Quick start

```vb
Imports CoreSuite.Controls

Dim DiscountBox As New PercentageBox With {
    .PercentageCulture = Globalization.CultureInfo.GetCultureInfo("pt-BR"),
    .PercentageValue = 15D,
    .DecimalPlaces = 2
}
Console.WriteLine(DiscountBox.FractionalValue) ' 0.15
```

## Designer usage

After installing or referencing the package, add `PercentageBox` from the Toolbox or create it in code. Properties configured in the Windows Forms Designer are serialized into `InitializeComponent` and remain available at runtime.

## API reference

### `PercentageBox`

Represents culture-aware percentage input with percentage and fractional representations.

```vb
Public Class PercentageBox
    Inherits NumericBoxBase
```

### Main members

| Member | Behavior |
| --- | --- |
| `PercentageValue` | displayed percentage value (15 means 15%). |
| `FractionalValue` | read-only fractional equivalent (0.15). |
| `PercentageCulture` | culture used for parsing and formatting. |
| `DecimalPlaces` | fractional digit count. |
| `IncludeThousandSeparator` | digit grouping option. |
| `RoundingMode` | midpoint rounding strategy. |
| `PercentageValueChanged` | raised when the percentage changes. |

## Behavior and validation

* Property changes take effect immediately unless a method explicitly starts or applies an operation.
* Callers should validate user input and referenced objects before performing application-specific work.
* Events use the standard Windows Forms event model and should be handled on the UI thread.

## Standard functionality

Because `PercentageBox` inherits from `NumericBoxBase`, the standard public members of that base type remain available unless the control intentionally hides or overrides them.

## Resource and lifetime considerations

* Dispose components, controls, images, fonts and menus created dynamically when your application owns them.
* Keep referenced controls alive for as long as a component is attached to them.
* File-based operations should receive valid, accessible paths.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.PercentageBox` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.PercentageBox` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| CoreSuite dependencies | `CoreSuite.NumericBoxBase` |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
