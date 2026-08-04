# SplitButton

**Button with a configurable drop-down area and ContextMenuStrip.**

> [!NOTE]
> SplitButton is one of the controls or components included in the **CoreSuite** solution. The package depends on `CoreSuite.NoFocusCueButton`. NuGet installs the required package or packages automatically.

## Overview

`SplitButton` extends `NoFocusCueButton` and provides button with a configurable drop-down area and ContextMenuStrip.

It is designed for .NET 8 Windows Forms applications and can be configured in code or through the Visual Studio designer where designer support applies.

## Features

* Supports separate main-button and drop-down actions.
* Supports split and drop-down-only modes.
* Provides keyboard access and right-to-left layout.
* Allows the arrow, separator and drop-down area to be customized.
* Uses the standard Windows Forms event and component model.
* Includes English XML documentation for the public API.
* Requires no third-party packages.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.SplitButton
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.SplitButton
```

## Namespace

```vb
Imports CoreSuite.Controls
```

## Quick start

```vb
Imports CoreSuite.Controls

Dim Menu As New ContextMenuStrip()
Menu.Items.Add("Open")
Menu.Items.Add("Open read-only")

Dim OpenButton As New SplitButton With {
    .Text = "Open",
    .DropDownMenu = Menu,
    .Mode = SplitButtonMode.Split,
    .DropDownAreaWidth = 24
}
AddHandler OpenButton.DropDownOpening, Sub(Sender, E) E.Cancel = False
```

## Designer usage

After installing or referencing the package, add `SplitButton` from the Toolbox or create it in code. Properties configured in the Windows Forms Designer are serialized into `InitializeComponent` and remain available at runtime.

## API reference

### `SplitButton`

Represents button with a configurable drop-down area and ContextMenuStrip.

```vb
Public Class SplitButton
    Inherits NoFocusCueButton
```

### Main members

| Member | Behavior |
| --- | --- |
| `DropDownMenu` | menu displayed by the arrow area. |
| `Mode` | split or drop-down behavior. |
| `DropDownAreaWidth` | arrow-region width. |
| `ArrowColor / SeparatorColor / ShowSeparator` | appearance. |
| `DropDownRectangle` | current arrow-region bounds. |
| `ShowDropDown()` | opens the menu programmatically. |
| `DropDownOpening, DropDownOpened and DropDownClosed` | menu lifecycle events. |

## Behavior and validation

* Property changes take effect immediately unless a method explicitly starts or applies an operation.
* Callers should validate user input and referenced objects before performing application-specific work.
* Events use the standard Windows Forms event model and should be handled on the UI thread.

## Standard functionality

Because `SplitButton` inherits from `NoFocusCueButton`, the standard public members of that base type remain available unless the control intentionally hides or overrides them.

## Resource and lifetime considerations

* Dispose components, controls, images, fonts and menus created dynamically when your application owns them.
* Keep referenced controls alive for as long as a component is attached to them.
* File-based operations should receive valid, accessible paths.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.SplitButton` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.SplitButton` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| CoreSuite dependencies | `CoreSuite.NoFocusCueButton` |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
