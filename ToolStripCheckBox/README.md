# ToolStripCheckBox

**Hosts a standard CheckBox inside ToolStrip-based controls.**

> [!NOTE]
> ToolStripCheckBox is one of the controls or components included in the **CoreSuite** solution. The package has no additional CoreSuite package dependencies.

## Overview

`ToolStripCheckBox` extends `ToolStripControlHost` and provides hosts a standard CheckBox inside ToolStrip-based controls.

It is designed for .NET 8 Windows Forms applications and can be configured in code or through the Visual Studio designer where designer support applies.

## Features

* Current checked state.
* Hosted CheckBox instance for advanced customization.
* Raised when the checked state changes.
* See the source XML documentation for detailed behavior.
* Uses the standard Windows Forms event and component model.
* Includes English XML documentation for the public API.
* Requires no third-party packages.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.ToolStripCheckBox
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.ToolStripCheckBox
```

## Namespace

```vb
Imports CoreSuite.Controls
```

## Quick start

```vb
Imports CoreSuite.Controls

Dim RememberItem As New ToolStripCheckBox With {
    .Text = "Remember selection",
    .Checked = True
}
OptionsToolStrip.Items.Add(RememberItem)
AddHandler RememberItem.CheckedChanged, Sub() Console.WriteLine(RememberItem.Checked)
```

## Designer usage

After installing or referencing the package, add `ToolStripCheckBox` from the Toolbox or create it in code. Properties configured in the Windows Forms Designer are serialized into `InitializeComponent` and remain available at runtime.

## API reference

### `ToolStripCheckBox`

Represents hosts a standard CheckBox inside ToolStrip-based controls.

```vb
Public Class ToolStripCheckBox
    Inherits ToolStripControlHost
```

### Main members

| Member | Behavior |
| --- | --- |
| `Checked` | current checked state. |
| `CheckBoxControl` | hosted CheckBox instance for advanced customization. |
| `CheckedChanged` | raised when the checked state changes. |
| `Standard ToolStripItem properties such as Text, Enabled and Alignment remain available.` | See the source XML documentation for detailed behavior. |

## Behavior and validation

* Property changes take effect immediately unless a method explicitly starts or applies an operation.
* Callers should validate user input and referenced objects before performing application-specific work.
* Events use the standard Windows Forms event model and should be handled on the UI thread.

## Standard functionality

Because `ToolStripCheckBox` inherits from `ToolStripControlHost`, the standard public members of that base type remain available unless the control intentionally hides or overrides them.

## Resource and lifetime considerations

* Dispose components, controls, images, fonts and menus created dynamically when your application owns them.
* Keep referenced controls alive for as long as a component is attached to them.
* File-based operations should receive valid, accessible paths.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.ToolStripCheckBox` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.ToolStripCheckBox` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| CoreSuite dependencies | `None` |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
