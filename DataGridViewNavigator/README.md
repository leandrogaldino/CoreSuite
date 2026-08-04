# DataGridViewNavigator

**Navigation component that connects ToolStrip buttons to a DataGridView.**

> [!NOTE]
> DataGridViewNavigator is one of the controls or components included in the **CoreSuite** solution. The package has no additional CoreSuite package dependencies.

## Overview

`DataGridViewNavigator` extends `Component` and provides navigation component that connects ToolStrip buttons to a DataGridView.

It is designed for .NET 8 Windows Forms applications and can be configured in code or through the Visual Studio designer where designer support applies.

## Features

* Target grid.
* Navigation buttons.
* Callbacks surrounding navigation.
* Cancels the next requested move.
* Uses the standard Windows Forms event and component model.
* Includes English XML documentation for the public API.
* Requires no third-party packages.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.DataGridViewNavigator
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.DataGridViewNavigator
```

## Namespace

```vb
Imports CoreSuite.Controls
```

## Quick start

```vb
Imports CoreSuite.Controls

Dim Navigator As New DataGridViewNavigator With {
    .DataGridView = ResultsGrid,
    .FirstButton = BtnFirst,
    .PreviousButton = BtnPrevious,
    .NextButton = BtnNext,
    .LastButton = BtnLast
}
Navigator.RefreshButtons()
```

## Designer usage

After installing or referencing the package, add `DataGridViewNavigator` from the Toolbox or create it in code. Properties configured in the Windows Forms Designer are serialized into `InitializeComponent` and remain available at runtime.

## API reference

### `DataGridViewNavigator`

Represents navigation component that connects ToolStrip buttons to a DataGridView.

```vb
Public Class DataGridViewNavigator
    Inherits Component
```

### Main members

| Member | Behavior |
| --- | --- |
| `DataGridView` | target grid. |
| `FirstButton / PreviousButton / NextButton / LastButton` | navigation buttons. |
| `ActionBeforeMove / ActionAfterMove` | callbacks surrounding navigation. |
| `CancelNextMove` | cancels the next requested move. |
| `MoveToFirst(), MoveToPrevious(), MoveToNext(), MoveToLast()` | programmatic navigation. |
| `EnsureVisibleRow()` | scrolls a row into view. |
| `RefreshButtons()` | updates enabled states. |

## Behavior and validation

* Property changes take effect immediately unless a method explicitly starts or applies an operation.
* Callers should validate user input and referenced objects before performing application-specific work.
* Events use the standard Windows Forms event model and should be handled on the UI thread.

## Standard functionality

Because `DataGridViewNavigator` inherits from `Component`, the standard public members of that base type remain available unless the control intentionally hides or overrides them.

## Resource and lifetime considerations

* Dispose components, controls, images, fonts and menus created dynamically when your application owns them.
* Keep referenced controls alive for as long as a component is attached to them.
* File-based operations should receive valid, accessible paths.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.DataGridViewNavigator` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.DataGridViewNavigator` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| CoreSuite dependencies | `None` |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
