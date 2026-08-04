# NumericBoxBase

**Abstract foundation for culture-aware numeric text controls.**

> [!NOTE]
> NumericBoxBase is one of the controls or components included in the **CoreSuite** solution. The package has no additional CoreSuite package dependencies.

## Overview

`NumericBoxBase` extends `TextBox` and provides abstract foundation for culture-aware numeric text controls.

It is designed for .NET 8 Windows Forms applications and can be configured in code or through the Visual Studio designer where designer support applies.

## Features

* Fractional digit count.
* Digit grouping option.
* Parsing and formatting culture.
* Midpoint rounding strategy.
* Uses the standard Windows Forms event and component model.
* Includes English XML documentation for the public API.
* Requires no third-party packages.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

> [!IMPORTANT]
> This is an infrastructure package used by DecimalBox, CurrencyBox and PercentageBox. Install a derived package unless you are creating a custom numeric control.

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.NumericBoxBase
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.NumericBoxBase
```

## Namespace

```vb
Imports CoreSuite.Controls
```

## Quick start

```vb
Imports CoreSuite.Controls

Public Class MeasurementBox
    Inherits NumericBoxBase

    Protected Overrides Function FormatNumericValueForDisplay(Value As Decimal) As String
        Return Value.ToString($"N{DecimalPlaces}", NumericCulture) & " cm"
    End Function
End Class
```

## Creating a derived control

The base class is not intended to be placed directly on a form. Inherit from it when implementing a specialized CoreSuite input control.

## API reference

### `NumericBoxBase`

Represents abstract foundation for culture-aware numeric text controls.

```vb
Public MustInherit Class NumericBoxBase
    Inherits TextBox
```

### Main members

| Member | Behavior |
| --- | --- |
| `DecimalPlaces` | fractional digit count. |
| `IncludeThousandSeparator` | digit grouping option. |
| `NumericCulture` | parsing and formatting culture. |
| `RoundingMode` | midpoint rounding strategy. |
| `SetNumericValue()` | changes the value from a derived control. |
| `RefreshNumericText()` | reapplies display formatting. |
| `FormatNumericValueForDisplay()` | required display implementation. |
| `TryParseNumericValue(), NormalizeNumericValue() and OnNumericValueChanged()` | extension points. |

## Behavior and validation

* Property changes take effect immediately unless a method explicitly starts or applies an operation.
* Callers should validate user input and referenced objects before performing application-specific work.
* Events use the standard Windows Forms event model and should be handled on the UI thread.

## Standard functionality

Because `NumericBoxBase` inherits from `TextBox`, the standard public members of that base type remain available unless the control intentionally hides or overrides them.

## Resource and lifetime considerations

* Dispose components, controls, images, fonts and menus created dynamically when your application owns them.
* Keep referenced controls alive for as long as a component is attached to them.
* File-based operations should receive valid, accessible paths.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.NumericBoxBase` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.NumericBoxBase` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| CoreSuite dependencies | `None` |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
