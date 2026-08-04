# Separator

**Lightweight horizontal or vertical separator control.**

> [!NOTE]
> Separator is one of the controls or components included in the **CoreSuite** solution. The package has no additional CoreSuite package dependencies.

## Overview

`Separator` extends `Control` and provides lightweight horizontal or vertical separator control.

It is designed for .NET 8 Windows Forms applications and can be configured in code or through the Visual Studio designer where designer support applies.

## Features

* Line color.
* Line thickness in pixels.
* Horizontal or vertical layout.
* Line placement within the control.
* Uses the standard Windows Forms event and component model.
* Includes English XML documentation for the public API.
* Requires no third-party packages.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.Separator
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.Separator
```

## Namespace

```vb
Imports CoreSuite.Controls
```

## Quick start

```vb
Imports CoreSuite.Controls

Dim Line As New Separator With {
    .Orientation = Orientation.Horizontal,
    .SeparatorAlignment = SeparatorAlignment.Center,
    .SeparatorColor = SystemColors.ControlDark,
    .Thickness = 1,
    .Dock = DockStyle.Top
}
```

## Designer usage

After installing or referencing the package, add `Separator` from the Toolbox or create it in code. Properties configured in the Windows Forms Designer are serialized into `InitializeComponent` and remain available at runtime.

## API reference

### `Separator`

Represents lightweight horizontal or vertical separator control.

```vb
Public Class Separator
    Inherits Control
```

### Main members

| Member | Behavior |
| --- | --- |
| `SeparatorColor` | line color. |
| `Thickness` | line thickness in pixels. |
| `Orientation` | horizontal or vertical layout. |
| `SeparatorAlignment` | line placement within the control. |
| `Text and TextChanged are intentionally hidden because the control does not render text.` | See the source XML documentation for detailed behavior. |

## Behavior and validation

* Property changes take effect immediately unless a method explicitly starts or applies an operation.
* Callers should validate user input and referenced objects before performing application-specific work.
* Events use the standard Windows Forms event model and should be handled on the UI thread.

## Standard functionality

Because `Separator` inherits from `Control`, the standard public members of that base type remain available unless the control intentionally hides or overrides them.

## Resource and lifetime considerations

* Dispose components, controls, images, fonts and menus created dynamically when your application owns them.
* Keep referenced controls alive for as long as a component is attached to them.
* File-based operations should receive valid, accessible paths.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.Separator` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.Separator` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| CoreSuite dependencies | `None` |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
