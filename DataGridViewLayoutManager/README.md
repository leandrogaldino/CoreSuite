# DataGridViewLayoutManager

**Component for capturing, restoring and persisting DataGridView layouts.**

> [!NOTE]
> DataGridViewLayoutManager is one of the controls or components included in the **CoreSuite** solution. The package has no additional CoreSuite package dependencies.

## Overview

`DataGridViewLayoutManager` extends `Component` and provides component for capturing, restoring and persisting DataGridView layouts.

It is designed for .NET 8 Windows Forms applications and can be configured in code or through the Visual Studio designer where designer support applies.

## Features

* Captures column visibility, width, order, alignment and formatting.
* Restores saved layouts and applies a versioned default layout.
* Preserves grid sorting information.
* Exposes explicit `Load` and `Save` operations.
* Uses the standard Windows Forms event and component model.
* Includes English XML documentation for the public API.
* Requires no third-party packages.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.DataGridViewLayoutManager
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.DataGridViewLayoutManager
```

## Namespace

```vb
Imports CoreSuite.Controls
```

## Quick start

```vb
Imports CoreSuite.Controls

Dim Manager As New DataGridViewLayoutManager With {
    .DataGridView = ResultsGrid,
    .DefaultLayout = New DataGridViewLayout("Results", 1)
}
Manager.Load()
' Call after the user changes the layout:
Manager.Save()
```

## Designer usage

After installing or referencing the package, add `DataGridViewLayoutManager` from the Toolbox or create it in code. Properties configured in the Windows Forms Designer are serialized into `InitializeComponent` and remain available at runtime.

## API reference

### `DataGridViewLayoutManager`

Represents component for capturing, restoring and persisting DataGridView layouts.

```vb
Public Class DataGridViewLayoutManager
    Inherits Component
```

### Main members

| Member | Behavior |
| --- | --- |
| `DataGridView` | grid whose layout is managed. |
| `DefaultLayout` | fallback layout and persistence identity. |
| `Load()` | restores a saved layout or applies the default. |
| `Save()` | captures and persists the current layout. |
| `Loaded` | raised after a layout is applied. |
| `DataGridViewLayout` | stores name, version, sort state and columns. |
| `DataGridViewLayoutColumn` | stores visibility, width, order, alignment and format. |

## Behavior and validation

* Property changes take effect immediately unless a method explicitly starts or applies an operation.
* Callers should validate user input and referenced objects before performing application-specific work.
* Events use the standard Windows Forms event model and should be handled on the UI thread.

## Standard functionality

Because `DataGridViewLayoutManager` inherits from `Component`, the standard public members of that base type remain available unless the control intentionally hides or overrides them.

## Resource and lifetime considerations

* Dispose components, controls, images, fonts and menus created dynamically when your application owns them.
* Keep referenced controls alive for as long as a component is attached to them.
* File-based operations should receive valid, accessible paths.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.DataGridViewLayoutManager` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.DataGridViewLayoutManager` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| CoreSuite dependencies | `None` |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
