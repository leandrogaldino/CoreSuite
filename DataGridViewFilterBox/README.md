# DataGridViewFilterBox

**A debounced local and remote filtering text box for .NET 8 Windows Forms, included in CoreSuite.**

> [!NOTE]
> DataGridViewFilterBox is one of the independent projects that make up the **CoreSuite** solution. The package contains the control, filtering configuration types, event data, and Windows Forms designer support.

## Overview

`DataGridViewFilterBox` extends the standard WinForms `TextBox` with a complete filtering workflow for a `DataGridView`. It waits for a configurable debounce interval, resolves searchable columns, safely escapes user text, and applies a `DataView.RowFilter` expression without discarding a filter that was already active.

When the source cannot be filtered through a `DataView`, the control can raise `FilterRequested` instead. This makes the same control useful for database queries, APIs, paged data, virtual grids, and application-specific collections.

## Key features

- Inherits from the standard WinForms `TextBox`.
- Filters `DataTable`, `DataView`, and compatible `BindingSource` data locally.
- Preserves and restores an existing `RowFilter`.
- Supports automatic, local-only, and custom filtering modes.
- Debounces filtering with `FilterInterval`.
- Delays filtering until `MinimumCharacters` is reached.
- Searches visible compatible columns automatically.
- Allows explicit included and ignored column collections.
- Supports contains, starts-with, ends-with, and exact matching.
- Supports case-sensitive or case-insensitive local matching.
- Raises cancellable custom requests for asynchronous remote searches.
- Includes a built-in clear glyph and accepts a custom clear-button image.
- Clears with `Escape` and applies immediately with `Enter`.
- Includes smart-tag actions for common designer settings.
- Includes XML documentation and NuGet symbol generation.

## Requirements

- Windows Forms
- .NET 8 for Windows (`net8.0-windows`)
- A reference to `CoreSuite.DataGridViewFilterBox`

The control has no runtime dependency on another CoreSuite package.

## Installation

```powershell
dotnet add package CoreSuite.DataGridViewFilterBox
```

Or add `DataGridViewFilterBox/DataGridViewFilterBox.vbproj` as a project reference when working directly with the CoreSuite solution.

## Quick start: local filtering

Bind a `DataTable`, `DataView`, or `BindingSource` backed by either type to the grid, then associate the filter box with the grid.

```vb
Imports CoreSuite.Controls
Private Sub ConfigureProducts()
    ProductsGrid.DataSource = ProductsTable
    ProductsFilter.DataGridView = ProductsGrid
    ProductsFilter.PlaceholderText = "Filter products..."
    ProductsFilter.MinimumCharacters = 1
    ProductsFilter.FilterInterval = 300
    ProductsFilter.SearchMode = DataGridViewFilterSearchMode.Contains
End Sub
```

When `FilterColumns` is empty, the control automatically includes compatible visible grid columns and excludes images and byte arrays.

### Select explicit columns

Use data column names, `DataGridViewColumn.Name` values, or `DataGridViewColumn.DataPropertyName` values.

```vb
ProductsFilter.FilterColumns.Add("Code")
ProductsFilter.FilterColumns.Add("Description")
ProductsFilter.FilterColumns.Add("CategoryName")
```

Only the resolved columns participate in the generated local expression.

### Ignore selected columns

`IgnoredColumns` always takes precedence over `FilterColumns` and automatic selection.

```vb
ProductsFilter.IgnoredColumns.Add("InternalNotes")
ProductsFilter.IgnoredColumns.Add("SupplierCost")
```

## Existing filters are preserved

If the data view already contains a filter, DataGridViewFilterBox combines it with its generated expression through `AND`.

```vb
ProductsTable.DefaultView.RowFilter = "Active = True"
ProductsFilter.DataGridView = ProductsGrid
```

Typing `keyboard` conceptually produces:

```text
(Active = True) AND (
    CONVERT([Code], 'System.String') LIKE '%keyboard%'
    OR CONVERT([Description], 'System.String') LIKE '%keyboard%'
)
```

Calling `ClearFilter` restores `Active = True`. If another part of the application replaces `RowFilter` while the CoreSuite filter is active, the control does not overwrite that newer external expression when clearing.

## Custom and remote filtering

Set `FilterMode` to `Custom` when the rows must come from a database, web service, paged source, or another application-defined process.

```vb
Private Async Sub ProductsFilter_FilterRequested(sender As Object, e As FilterRequestedEventArgs) Handles ProductsFilter.FilterRequested
    Try
        Dim Results As DataTable = Await ProductService.SearchAsync(e.FilterText, e.CancellationToken)
        If e.CancellationToken.IsCancellationRequested Then Return
        ProductsGrid.DataSource = Results
    Catch OperationCanceled As OperationCanceledException
    Catch Failure As Exception
        MessageBox.Show(Failure.Message)
    End Try
End Sub
```

Every newer request cancels the token supplied with the preceding event. Clearing or disposing the control also cancels the active request.

Handle `FilterCleared` to restore the unfiltered remote source:

```vb
Private Async Sub ProductsFilter_FilterCleared(sender As Object, e As EventArgs) Handles ProductsFilter.FilterCleared
    ProductsGrid.DataSource = Await ProductService.GetAllAsync()
End Sub
```

`Automatic` uses local filtering when a `DataView` is available and raises `FilterRequested` otherwise. `Local` never falls back to a custom request and raises `FilterFailed` for unsupported sources.

## Filter modes

| Value | Behavior |
|---|---|
| `Automatic` | Uses local `DataView` filtering when possible; otherwise raises `FilterRequested`. |
| `Local` | Requires a compatible local source and raises `FilterFailed` when it cannot filter it. |
| `Custom` | Never changes the source and always raises `FilterRequested` after debounce. |

## Search modes

| Value | Example for `abc` |
|---|---|
| `Contains` | `%abc%` |
| `StartsWith` | `abc%` |
| `EndsWith` | `%abc` |
| `ExactMatch` | `abc` |

User text is escaped before it is inserted into a local `RowFilter` expression. Apostrophes and the `LIKE` wildcard characters `[`, `]`, `%`, and `*` are treated as literal input.

## Main properties

| Property | Default | Description |
|---|---:|---|
| `DataGridView` | `Nothing` | Grid whose current data source is inspected and filtered. |
| `BindingSource` | `Nothing` | Optional source that takes precedence over the grid source. |
| `FilterColumns` | Empty | Explicit columns to include; empty enables automatic selection. |
| `IgnoredColumns` | Empty | Columns that must never participate in filtering. |
| `FilterMode` | `Automatic` | Chooses automatic, local-only, or custom processing. |
| `SearchMode` | `Contains` | Defines the local text-matching behavior. |
| `FilterInterval` | `300` | Debounce interval in milliseconds. |
| `MinimumCharacters` | `1` | Minimum text length required before filtering. |
| `FilterEnabled` | `True` | Enables or disables processing without disabling the text box. |
| `CaseSensitive` | `False` | Controls local `DataTable` case sensitivity while the filter is active. |
| `IncludeHiddenColumns` | `False` | Includes hidden grid columns in automatic selection. |
| `ShowClearButton` | `True` | Displays the embedded clear button while text is present. |
| `ClearButtonImage` | `Nothing` | Replaces the built-in clear glyph with a custom image. |
| `IsFilterApplied` | Read-only | Indicates whether a local filter owned by this control is active. |
| `LastFilterExpression` | Read-only | Exposes the complete active local expression. |

Inherited `TextBox` properties such as `PlaceholderText`, `CharacterCasing`, `MaxLength`, `ReadOnly`, `Font`, `ForeColor`, and data binding remain available.

## Methods

| Method | Description |
|---|---|
| `ApplyFilter()` | Processes the current text immediately and returns whether a filter was applied or requested. |
| `RefreshFilter()` | Reprocesses the current text after a source or configuration change. |
| `ClearFilter()` | Clears the text, cancels custom work, and restores the original local filter. |

## Events

| Event | Description |
|---|---|
| `FilterRequested` | Requests custom or remote filtering and supplies text, columns, and a cancellation token. |
| `FilterApplied` | Reports a successful local filter, its expression, columns, and matched row count. |
| `FilterCleared` | Reports that an active local or custom filter was cleared. |
| `FilterFailed` | Reports an unsupported local source, unresolved columns, or an invalid expression. |

```vb
Private Sub ProductsFilter_FilterApplied(sender As Object, e As FilterAppliedEventArgs) Handles ProductsFilter.FilterApplied
    ResultsLabel.Text = $"{e.MatchedRowCount} result(s)"
End Sub
Private Sub ProductsFilter_FilterFailed(sender As Object, e As FilterFailedEventArgs) Handles ProductsFilter.FilterFailed
    ErrorProvider1.SetError(ProductsFilter, e.Exception.Message)
End Sub
```

## Keyboard and clear-button behavior

| Input | Behavior |
|---|---|
| Typing | Restarts the debounce interval. |
| `Enter` | Applies the current filter immediately in single-line mode. |
| `Escape` | Clears the text and active filter. |
| Clear button | Performs the same operation as `ClearFilter()` and returns focus to the text box. |

The clear button supports right-to-left layout. Assign `ClearButtonImage` when the application's visual language requires a custom icon; otherwise, the control draws a DPI-independent close glyph.

## Local filtering notes

- Local filtering requires a `DataView`, directly or through a `DataTable`/`BindingSource`.
- Values are converted to `System.String` in the generated expression so text, numbers, dates, Boolean values, GUIDs, and other convertible columns can be searched together.
- Image and `Byte()` columns are excluded from automatic selection.
- `CaseSensitive` temporarily changes the target table's case-sensitivity setting and restores it when the owned filter is removed.
- Explicit column names that cannot be resolved are skipped. `FilterFailed` is raised when no configured column can be used.
- `FilterApplied` refers only to filtering performed locally by the control. Completion of custom or remote work remains owned by the application.

## License

CoreSuite is licensed under the MIT License.
