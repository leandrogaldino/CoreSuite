# AsyncLookupBox

**A data-source-independent asynchronous lookup and selection control for .NET 8 Windows Forms, included in CoreSuite.**

> [!NOTE]
> AsyncLookupBox is one of the independent projects that make up the **CoreSuite** solution. The package contains the control, result-column configuration, event data, drop-down interface, and Windows Forms designer support.

## Overview

`AsyncLookupBox` is a text box that waits while the user types, requests results from application code, and displays selectable objects in a tabular drop-down. It does not create SQL and does not depend on a database, HTTP client, ORM, or business-object library.

The application supplies a `Task` in the `SearchRequested` event. The control manages the rest of the interaction:

1. Waits for `MinimumCharacters` and the debounce interval.
2. Cancels the preceding request when the text changes.
3. Displays loading, empty-result, and error states.
4. Ignores an old result if a newer request has already started.
5. Builds the result grid from `Columns`.
6. Supports mouse and keyboard selection.
7. Exposes the selected object and selected value.

This makes the same control suitable for REST APIs, Entity Framework, Dapper, ADO.NET, local services, files, caches, and in-memory collections.

## Key features

- Inherits from the standard WinForms `TextBox`.
- Accepts any application-defined enumerable result type.
- Uses real asynchronous tasks without blocking the UI thread.
- Debounces searches with `SearchInterval`.
- Cancels replaced requests through `CancellationToken`.
- Prevents late results from overwriting newer searches.
- Supports properties, `DataRow`, `DataRowView`, and dictionaries.
- Supports nested property paths such as `Category.Name`.
- Provides configurable result columns, formatting, widths, and fill modes.
- Tracks `SelectedItem`, `SelectedValue`, and `HasSelection`.
- Clearly distinguishes selected and editable states through configurable colors and a built-in check mark.
- Supports automatic selection of a single result.
- Includes loading animation and a combined cancel/clear button.
- Supports `Up`, `Down`, `Enter`, `Escape`, `F4`, and `Alt+Down`.
- Includes localized status-message properties.
- Includes a smart tag and collection editor for the modern .NET WinForms designer.
- Has no runtime dependency on another CoreSuite package.

## Requirements

- Windows Forms
- .NET 8 for Windows (`net8.0-windows`)
- A reference to `CoreSuite.AsyncLookupBox`

## Installation

```powershell
dotnet add package CoreSuite.AsyncLookupBox
```

Or add `AsyncLookupBox/AsyncLookupBox.vbproj` as a project reference when working directly with the CoreSuite solution.

## Quick start

Assume that the application searches this business class:

```vb
Public Class Product
    Public Property Id As Integer
    Public Property Code As String
    Public Property Name As String
    Public Property Category As String
    Public Property Price As Decimal
End Class
```

Configure the lookup once, usually in the Designer or the form's `Load` event:

```vb
Imports CoreSuite.Controls
Private Sub ConfigureProductLookup()
    ProductLookup.DisplayMember = NameOf(Product.Name)
    ProductLookup.ValueMember = NameOf(Product.Id)
    ProductLookup.MinimumCharacters = 2
    ProductLookup.SearchInterval = 300
    ProductLookup.DropDownWidth = 520
    ProductLookup.PlaceholderText = "Search by code, name, or category..."
    ProductLookup.Columns.Add(NameOf(Product.Code)).HeaderText = "Code"
    ProductLookup.Columns.Add(NameOf(Product.Name)).HeaderText = "Product"
    ProductLookup.Columns.Add(NameOf(Product.Category)).HeaderText = "Category"
    Dim PriceColumn As AsyncLookupColumn = ProductLookup.Columns.Add(NameOf(Product.Price))
    PriceColumn.HeaderText = "Price"
    PriceColumn.Format = "C2"
End Sub
```

Handle `SearchRequested` and give the event arguments the task returned by your service:

```vb
Private Sub ProductLookup_SearchRequested(Sender As Object, E As AsyncLookupSearchRequestedEventArgs) Handles ProductLookup.SearchRequested
    E.SetSearchTask(ProductService.SearchAsync(E.SearchText, E.CancellationToken))
End Sub
```

The handler itself is intentionally not `Async`. `SetSearchTask` registers the running operation, and the control awaits it internally. The service can return `Task(Of List(Of Product))`, `Task(Of Product())`, or another task whose result implements `IEnumerable`.

Read the selected values in `SelectionChanged` or when saving the form:

```vb
Private Sub ProductLookup_SelectionChanged(Sender As Object, E As AsyncLookupSelectionChangedEventArgs) Handles ProductLookup.SelectionChanged
    If Not ProductLookup.HasSelection Then
        SelectedProductLabel.Text = "No product selected"
        Return
    End If
    Dim Product As Product = DirectCast(ProductLookup.SelectedItem, Product)
    SelectedProductLabel.Text = $"{Product.Code} - {Product.Name}"
    Dim ProductId As Integer = CInt(ProductLookup.SelectedValue)
End Sub
```

## Complete test without a database

The CoreSuite `Test` project included with the source contains a ready-to-run example with 50 products. The following is the essential search method used by that example:

```vb
Private Async Function SearchProductsAsync(SearchText As String, CancellationToken As CancellationToken) As Task(Of List(Of Product))
    Await Task.Delay(700, CancellationToken)
    Return Products.Where(Function(Product) Product.Code.Contains(SearchText, StringComparison.CurrentCultureIgnoreCase) OrElse Product.Name.Contains(SearchText, StringComparison.CurrentCultureIgnoreCase) OrElse Product.Category.Contains(SearchText, StringComparison.CurrentCultureIgnoreCase)).ToList()
End Function
```

The delay makes the loading indicator and cancellation behavior visible. Type one term and immediately replace it with another; the first token is canceled and only the newest result can appear.

## Why `SetSearchTask` is used

A normal .NET event is raised synchronously. If the event handler were an `Async Sub`, the control could not reliably know when that handler had completed or catch its exceptions.

This pattern keeps event handling safe:

```vb
Private Sub CustomerLookup_SearchRequested(Sender As Object, E As AsyncLookupSearchRequestedEventArgs) Handles CustomerLookup.SearchRequested
    E.SetSearchTask(SearchCustomersAsync(E.SearchText, E.CancellationToken))
End Sub
```

`SetSearchTask` accepts the task immediately. `AsyncLookupBox` then awaits it, catches failures, checks cancellation, updates its state, and displays the results.

Call `SetResults` instead when the objects are already available synchronously:

```vb
Private Sub CityLookup_SearchRequested(Sender As Object, E As AsyncLookupSearchRequestedEventArgs) Handles CityLookup.SearchRequested
    Dim Matches = Cities.Where(Function(City) City.Name.Contains(E.SearchText, StringComparison.CurrentCultureIgnoreCase)).ToList()
    E.SetResults(Matches)
End Sub
```

Only one of `SetSearchTask` or `SetResults` may be called for each request.

## Search cancellation

Every request receives its own token:

```vb
E.CancellationToken
```

The token is canceled when:

- a newer search starts;
- the text is cleared;
- the clear/cancel button is clicked while searching;
- `CancelSearch()` is called;
- `SearchEnabled` becomes `False`;
- the control is disabled or disposed.

Pass the token through every operation that supports it:

```vb
Private Async Function SearchCustomersAsync(SearchText As String, CancellationToken As CancellationToken) As Task(Of List(Of Customer))
    Return Await CustomerApi.SearchAsync(SearchText, CancellationToken)
End Function
```

Even when an external service ignores cancellation, AsyncLookupBox compares request versions and refuses to display a result belonging to an older search.

## Configuring result columns

When `Columns` is empty, the drop-down displays one fill-width column containing `DisplayMember`. If `DisplayMember` is empty, the control displays each object's `ToString()` result.

Each configured `AsyncLookupColumn` supports:

| Property | Default | Description |
|---|---:|---|
| `PropertyName` | Empty | Property, data-column, dictionary-key, or nested path to display. |
| `HeaderText` | Empty | Column header; empty uses `PropertyName`. |
| `Width` | `120` | Width used when automatic sizing is disabled. |
| `MinimumWidth` | `5` | Smallest permitted width. |
| `AutoSizeMode` | `None` | Standard `DataGridViewAutoSizeColumnMode`. |
| `FillWeight` | `100` | Relative width when `AutoSizeMode` is `Fill`. |
| `Format` | Empty | Standard or custom value format, such as `C2` or `dd/MM/yyyy`. |
| `NullValue` | Empty | Text shown for `Nothing` or `DBNull.Value`. |
| `Visible` | `True` | Determines whether the column is displayed. |

Example with a nested object:

```vb
Dim CategoryColumn As AsyncLookupColumn = ProductLookup.Columns.Add("Category.Name")
CategoryColumn.HeaderText = "Category"
CategoryColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
```

The collection can also be edited through the `Columns` property in the Windows Forms Designer.

## DataTable and dictionary results

The same property names work with `DataRow` and `DataRowView` results:

```vb
Private Sub SupplierLookup_SearchRequested(Sender As Object, E As AsyncLookupSearchRequestedEventArgs) Handles SupplierLookup.SearchRequested
    E.SetSearchTask(SupplierRepository.SearchRowsAsync(E.SearchText, E.CancellationToken))
End Sub
```

Configure columns using the `DataColumn.ColumnName` values. For dictionary results, use dictionary keys. String dictionary keys are resolved case-insensitively when an exact key is unavailable.

## Selection behavior

Selecting a row performs these operations:

1. Stores the complete object in `SelectedItem`.
2. Resolves `SelectedValue` from `ValueMember`, or uses the complete object when `ValueMember` is empty.
3. Resolves the text from `DisplayMember`, or uses the object's text representation when it is empty.
4. Cancels pending searching.
5. Closes the result drop-down.
6. Raises `SelectionChanged`.

If the user edits the selected text, the stale selection is cleared before a new search begins. `SelectionChanged` then contains the old item and `Nothing` as the new item.

The selected state is visible by default:

- the text box uses `SelectedItemBackColor` and `SelectedItemForeColor`;
- the embedded action button displays a check mark;
- changing or deleting any character immediately restores the normal `BackColor` and `ForeColor`, replaces the check mark with the clear glyph, clears `SelectedItem` and `SelectedValue`, and starts a new search.

The color change is only visual. It does not set `ReadOnly` and does not prevent the user from editing the selected text.

Customize the selected state in the Designer or in code:

```vb
ProductLookup.HighlightSelectedItem = True
ProductLookup.SelectedItemBackColor = Color.Honeydew
ProductLookup.SelectedItemForeColor = Color.DarkGreen
```

Set `HighlightSelectedItem` to `False` to retain the normal text-box colors. The check mark remains available as long as `ShowClearButton` is enabled. Assign `SelectedButtonImage` to replace the built-in check mark with a custom image.

Select an object programmatically:

```vb
ProductLookup.SelectItem(Product)
```

Clear everything:

```vb
ProductLookup.ClearSelection()
```

## Localizing status text

All user-facing drop-down messages are configurable:

```vb
ProductLookup.LoadingText = "Pesquisando..."
ProductLookup.NoResultsText = "Nenhum produto encontrado."
ProductLookup.SearchErrorText = "Não foi possível carregar os produtos."
ProductLookup.SearchNotConfiguredText = "A pesquisa não foi configurada."
ProductLookup.CharactersRemainingText = "Digite mais {0} caractere(s)."
ProductLookup.ResultColumnHeaderText = "Resultado"
```

`CharactersRemainingText` receives the missing character count as format argument `{0}`.

## Keyboard and mouse behavior

| Input | Behavior |
|---|---|
| Typing | Clears a stale selection and restarts the debounce interval. |
| `Down` / `Up` | Moves through visible results. |
| `Enter` | Selects the current row or starts an immediate search. |
| `Escape` | Closes the drop-down; when already closed, clears the lookup. |
| `F4` / `Alt+Down` | Reopens retained results. |
| Result-row click | Selects that result. |
| Action button while searching | Cancels the active search. |
| Action button while idle | Clears text and selection. |

## Main properties

| Property | Default | Description |
|---|---:|---|
| `DisplayMember` | Empty | Property path used as selected text. |
| `ValueMember` | Empty | Property path used for `SelectedValue`. |
| `Columns` | Empty | Columns displayed by the result grid. |
| `SearchInterval` | `300` | Debounce interval in milliseconds. |
| `MinimumCharacters` | `2` | Character count required before searching. |
| `MaximumResults` | `100` | Maximum retained results; `0` removes this limit. |
| `SearchEnabled` | `True` | Enables or disables search requests. |
| `AutoSelectSingleResult` | `False` | Automatically selects exactly one result. |
| `DropDownWidth` | `0` | Width; `0` uses at least the control width. |
| `DropDownHeight` | `220` | Drop-down height in pixels. |
| `ShowColumnHeaders` | `True` | Displays result headers. |
| `ShowClearButton` | `True` | Displays the embedded clear/cancel button. |
| `HighlightSelectedItem` | `True` | Uses distinct text-box colors while an item is selected. |
| `SelectedItemBackColor` | `AliceBlue` | Background color used while an item is selected. |
| `SelectedItemForeColor` | `RoyalBlue` | Text color used while an item is selected. |
| `SelectedButtonImage` | `Nothing` | Image used for the selected state; `Nothing` draws a check mark. |
| `SelectedItem` | Read-only | Complete selected result object. |
| `SelectedValue` | Read-only | Value resolved through `ValueMember`. |
| `HasSelection` | Read-only | Indicates whether an object is selected. |
| `IsSearching` | Read-only | Indicates whether a current request is running. |
| `IsDropDownOpen` | Read-only | Indicates whether the result list is open. |
| `Results` | Read-only | Results retained from the latest successful search. |

Inherited `TextBox` properties such as `PlaceholderText`, `CharacterCasing`, `MaxLength`, `ReadOnly`, `Font`, `BackColor`, `ForeColor`, and data binding remain available.

## Methods

| Method | Description |
|---|---|
| `RefreshResultsAsync()` | Immediately searches the current text and returns retained results. |
| `CancelSearch()` | Cancels the active request and closes the drop-down. |
| `SelectItem(item)` | Selects an application object programmatically. |
| `ClearSelection()` | Clears text, selection, results, and pending searching. |
| `CloseDropDown()` | Closes only the result drop-down. |

## Events

| Event | Description |
|---|---|
| `SearchRequested` | Requests a task or immediate collection from the application. |
| `SearchCompleted` | Reports a successful current search, results, duration, and truncation. |
| `SearchFailed` | Reports an exception from the event handler or supplied task. |
| `IsSearchingChanged` | Reports transitions into and out of active searching. |
| `SelectionChanged` | Reports selected and cleared items and values. |
| `DropDownOpened` | Reports that the result interface opened. |
| `DropDownClosed` | Reports that the result interface closed. |

Do not use an `Async Sub` handler for `SearchRequested`. Supply the task synchronously with `SetSearchTask`, as shown throughout this README.

## AsyncLookupBox and QueriedBox

| Requirement | Use |
|---|---|
| Build SQL from tables, joins, conditions, parameters, and supported dialects | `QueriedBox` |
| Search any application-defined service or source asynchronously | `AsyncLookupBox` |
| Configure database querying inside the control | `QueriedBox` |
| Keep data access completely outside the control | `AsyncLookupBox` |
| Require cancellation and protection against late results | `AsyncLookupBox` |

Both controls provide a lookup experience, but they solve different integration problems and can coexist in the same application.

## License

CoreSuite is licensed under the MIT License.
