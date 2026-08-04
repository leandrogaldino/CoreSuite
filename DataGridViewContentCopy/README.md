# DataGridViewContentCopy

**Component that adds cell and row copy commands to a DataGridView context menu.**

> [!NOTE]
> DataGridViewContentCopy is one of the controls or components included in the **CoreSuite** solution. The package has no additional CoreSuite package dependencies.

## Overview

`DataGridViewContentCopy` extends `Component` and provides component that adds cell and row copy commands to a DataGridView context menu.

It is designed for .NET 8 Windows Forms applications and can be configured in code or through the Visual Studio designer where designer support applies.

## Features

* Grid managed by the component.
* Menu used for the copy commands.
* Controls whether icons are displayed beside the copy commands.
* Prefixes copied cells with the column header.
* Uses the standard Windows Forms event and component model.
* Includes English XML documentation for the public API.
* Requires no third-party packages.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.DataGridViewContentCopy
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.DataGridViewContentCopy
```

## Namespace

```vb
Imports CoreSuite.Controls
```

## Quick start

```vb
Imports CoreSuite.Controls

Dim Copier As New DataGridViewContentCopy With {
    .DataGridView = ResultsGrid,
    .IncludeHeaderTextInCellCopy = True,
    .IncludeHeaderTextInRowCopy = True,
    .ShowImages = False
}
```

## Designer usage

After installing or referencing the package, add `DataGridViewContentCopy` from the Toolbox or create it in code. Properties configured in the Windows Forms Designer are serialized into `InitializeComponent` and remain available at runtime.

## API reference

### `DataGridViewContentCopy`

Represents component that adds cell and row copy commands to a DataGridView context menu.

```vb
Public Class DataGridViewContentCopy
    Inherits Component
```

### Main members

| Member | Behavior |
| --- | --- |
| `DataGridView` | grid managed by the component. |
| `ContextMenuStrip` | menu used for the copy commands. |
| `ShowImages` | controls whether icons are displayed beside the copy commands. |
| `IncludeHeaderTextInCellCopy` | prefixes copied cells with the column header. |
| `IncludeHeaderTextInRowCopy` | includes headers when copying a row. |
| `CopyCellButtonText / CopyRowButtonText` | configure the captions of the two copy commands. |

## Behavior and validation

* Property changes take effect immediately unless a method explicitly starts or applies an operation.
* Callers should validate user input and referenced objects before performing application-specific work.
* Events use the standard Windows Forms event model and should be handled on the UI thread.

## Standard functionality

Because `DataGridViewContentCopy` inherits from `Component`, the standard public members of that base type remain available unless the control intentionally hides or overrides them.

## Resource and lifetime considerations

* Dispose components, controls, images, fonts and menus created dynamically when your application owns them.
* Keep referenced controls alive for as long as a component is attached to them.
* File-based operations should receive valid, accessible paths.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.DataGridViewContentCopy` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.DataGridViewContentCopy` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| CoreSuite dependencies | `None` |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
