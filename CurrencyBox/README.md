# CurrencyBox

**Currency-formatted numeric input for Windows Forms.**

> [!NOTE]
> CurrencyBox is one of the controls or components included in the **CoreSuite** solution. The package depends on `CoreSuite.NumericBoxBase`. NuGet installs the required package or packages automatically.

## Overview

`CurrencyBox` extends `NumericBoxBase` and provides currency-formatted numeric input for Windows Forms.

It is designed for .NET 8 Windows Forms applications and can be configured in code or through the Visual Studio designer where designer support applies.

## Features

* Current monetary value.
* Culture used for the currency symbol and formatting.
* Number of fractional digits.
* Enables digit grouping.
* Uses the standard Windows Forms event and component model.
* Includes English XML documentation for the public API.
* Requires no third-party packages.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.CurrencyBox
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.CurrencyBox
```

## Namespace

```vb
Imports CoreSuite.Controls
```

## Quick start

```vb
Imports CoreSuite.Controls

Dim PriceBox As New CurrencyBox With {
    .CurrencyCulture = Globalization.CultureInfo.GetCultureInfo("pt-BR"),
    .CurrencyValue = 1250.5D,
    .DecimalPlaces = 2,
    .IncludeThousandSeparator = True
}
AddHandler PriceBox.CurrencyValueChanged, Sub() Console.WriteLine(PriceBox.CurrencyValue)
```

## Designer usage

After installing or referencing the package, add `CurrencyBox` from the Toolbox or create it in code. Properties configured in the Windows Forms Designer are serialized into `InitializeComponent` and remain available at runtime.

## API reference

### `CurrencyBox`

Represents currency-formatted numeric input for Windows Forms.

```vb
Public Class CurrencyBox
    Inherits NumericBoxBase
```

### Main members

| Member | Behavior |
| --- | --- |
| `CurrencyValue` | current monetary value. |
| `CurrencyCulture` | culture used for the currency symbol and formatting. |
| `DecimalPlaces` | number of fractional digits. |
| `IncludeThousandSeparator` | enables digit grouping. |
| `RoundingMode` | midpoint rounding strategy. |
| `CurrencyValueChanged` | raised when the value changes. |

## Behavior and validation

* Property changes take effect immediately unless a method explicitly starts or applies an operation.
* Callers should validate user input and referenced objects before performing application-specific work.
* Events use the standard Windows Forms event model and should be handled on the UI thread.

## Standard functionality

Because `CurrencyBox` inherits from `NumericBoxBase`, the standard public members of that base type remain available unless the control intentionally hides or overrides them.

## Resource and lifetime considerations

* Dispose components, controls, images, fonts and menus created dynamically when your application owns them.
* Keep referenced controls alive for as long as a component is attached to them.
* File-based operations should receive valid, accessible paths.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.CurrencyBox` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.CurrencyBox` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| CoreSuite dependencies | `CoreSuite.NumericBoxBase` |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
