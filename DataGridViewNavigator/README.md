# DataGridViewNavigator

**Navigation component that connects ToolStrip buttons to a DataGridView.**

> [!NOTE]
> DataGridViewNavigator is one of the controls or components included in the **CoreSuite** solution. The package has no additional CoreSuite package dependencies.

## Overview

`DataGridViewNavigator` extends `Component` and provides navigation functionality for a `DataGridView` using ToolStrip buttons.

It supports first, previous, next and last row navigation, synchronous and asynchronous callbacks, cancellable before-move operations, navigation concurrency protection, and an optional position label with customizable formatting.

It is designed for .NET 8 Windows Forms applications and can be configured in code or through the Visual Studio designer where designer support applies.

## Features

* Target grid.
* First, previous, next and last row navigation.
* Synchronous and asynchronous callbacks before and after navigation.
* Cancellable before-move operations.
* Navigation concurrency protection.
* Optional navigation position display.
* Customizable position formatting.
* Automatic navigation button state updates.
* Automatic synchronization with grid selection and data source changes.
* Ignores hidden rows and the new-row placeholder during navigation.
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
    .FirstButton = BtnFirst,
    .PreviousButton = BtnPrevious,
    .NextButton = BtnNext,
    .LastButton = BtnLast,
    .PositionLabel = LblPosition,
    .PositionFormat = "{0} of {1}",
    .DataGridView = ResultsGrid
}
```

`DataGridView` is assigned last in the initializer so the initial button states and position display are refreshed after the related controls and formatting have been configured.

The default position format is:

```text
{0} of {1}
```

Where `{0}` represents the current position and `{1}` represents the total number of navigable rows.

For example:

```vb
Navigator.PositionFormat = "Record {0} of {1}"
```

or:

```vb
Navigator.PositionFormat = "{0}/{1}"
```

## Navigation callbacks

Synchronous callbacks can be assigned before and after navigation:

```vb
Navigator.ActionBeforeMove = Sub(E)
    If Not CanLeaveCurrentRecord() Then E.Cancel = True
End Sub

Navigator.ActionAfterMove = Sub()
    LoadSelectedRecord()
End Sub
```

Asynchronous callbacks are also supported:

```vb
Navigator.ActionBeforeMoveAsync = Async Function(E)
    If Not Await CanLeaveCurrentRecordAsync() Then E.Cancel = True
End Function

Navigator.ActionAfterMoveAsync = Async Function()
    Await LoadSelectedRecordAsync()
End Function
```

The before-move callback receives a `CancelEventArgs`. Setting `Cancel` to `True` prevents the requested navigation.

## Programmatic navigation

Navigation can be performed synchronously:

```vb
Navigator.MoveToFirst()
Navigator.MoveToPrevious()
Navigator.MoveToNext()
Navigator.MoveToLast()
```

Or asynchronously:

```vb
Await Navigator.MoveToFirstAsync()
Await Navigator.MoveToPreviousAsync()
Await Navigator.MoveToNextAsync()
Await Navigator.MoveToLastAsync()
```

Synchronous navigation methods cannot be used when asynchronous navigation callbacks are configured.

## Designer usage

After installing or referencing the package, add `DataGridViewNavigator` from the Toolbox or create it in code.

Properties configured in the Windows Forms Designer are serialized into `InitializeComponent` and remain available at runtime.

## API reference

### `DataGridViewNavigator`

Represents a navigation component that connects ToolStrip buttons to a DataGridView.

```vb
Public Class DataGridViewNavigator
    Inherits Component
```

### Main members

| Member | Behavior |
| --- | --- |
| `DataGridView` | Target grid used for navigation. |
| `FirstButton` | Button used to navigate to the first row. |
| `PreviousButton` | Button used to navigate to the previous row. |
| `NextButton` | Button used to navigate to the next row. |
| `LastButton` | Button used to navigate to the last row. |
| `PositionLabel` | Optional `ToolStripLabel` used to display the current navigation position. |
| `PositionFormat` | Format used to display the current position and total number of navigable rows. |
| `ActionBeforeMove` | Synchronous callback executed before navigation. |
| `ActionAfterMove` | Synchronous callback executed after successful navigation. |
| `ActionBeforeMoveAsync` | Asynchronous callback executed before navigation. |
| `ActionAfterMoveAsync` | Asynchronous callback executed after successful navigation. |
| `IsNavigating` | Indicates whether a navigation operation is currently being executed. |
| `MoveToFirst()` | Moves synchronously to the first navigable row. |
| `MoveToPrevious()` | Moves synchronously to the previous navigable row. |
| `MoveToNext()` | Moves synchronously to the next navigable row. |
| `MoveToLast()` | Moves synchronously to the last navigable row. |
| `MoveToFirstAsync()` | Moves asynchronously to the first navigable row. |
| `MoveToPreviousAsync()` | Moves asynchronously to the previous navigable row. |
| `MoveToNextAsync()` | Moves asynchronously to the next navigable row. |
| `MoveToLastAsync()` | Moves asynchronously to the last navigable row. |
| `EnsureVisibleRow()` | Ensures that the specified row is visible within the grid. |
| `RefreshButtons()` | Updates navigation button states and the current position display. |

## Position display

The position display is optional and is enabled by assigning a `ToolStripLabel` to `PositionLabel`.

```vb
Navigator.PositionLabel = LblPosition
```

The default format is:

```vb
"{0} of {1}"
```

The format can be customized:

```vb
Navigator.PositionFormat = "Record {0} of {1}"
```

Only navigable rows are included in the position count. Hidden rows and the DataGridView new-row placeholder are ignored.

## Navigation behavior

The component automatically configures the associated DataGridView with:

```vb
MultiSelect = False
SelectionMode = DataGridViewSelectionMode.FullRowSelect
```

Navigation automatically skips rows that:

* Are hidden.
* Represent the DataGridView new-row placeholder.

Navigation buttons are automatically enabled or disabled according to the current row and available navigation direction.

The component refreshes its state when:

* The DataGridView data source changes.
* The DataGridView selection changes.
* A navigation operation starts.
* A navigation operation finishes.

While an asynchronous navigation operation is in progress, additional navigation requests are ignored until the current operation completes.

## Behavior and validation

* Property changes take effect immediately unless a method explicitly starts or applies an operation.
* Before-move callbacks may cancel navigation through `CancelEventArgs`.
* Synchronous navigation cannot be used while asynchronous callbacks are configured.
* The position display reflects only navigable rows.
* Callers should validate user input and referenced objects before performing application-specific work.
* Events and callbacks that interact with Windows Forms controls should execute on the UI thread.

## Resource and lifetime considerations

* Dispose the component when it is no longer required.
* Event handlers attached to the DataGridView and navigation buttons are removed when the component is disposed.
* Keep referenced controls alive for as long as the component is attached to them.

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
