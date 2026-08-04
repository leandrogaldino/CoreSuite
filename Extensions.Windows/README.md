# CoreSuite.Extensions.Windows

**Windows Forms extensions for converting collections to tables and populating `DataGridView` controls.**

> [!NOTE]
> `CoreSuite.Extensions.Windows` is one of the libraries included in the **CoreSuite** solution. It contains the Windows-specific extensions and depends on `CoreSuite.Helpers`.

## Overview

`CoreSuite.Extensions.Windows` converts object and dictionary collections to `DataTable` instances and provides a convenient method for filling a Windows Forms `DataGridView`. Object conversion supports property exclusion through `IgnoreInToTableAttribute`, while grid population can preserve the selected row, scroll position and sort direction and can optionally display an order column.

This package is intended for .NET 8 Windows Forms applications. General-purpose string utilities remain in the separate `CoreSuite.Extensions` package.

## Features

- Converts object collections to `DataTable` instances through public properties.
- Converts dictionary collections to tables using dictionary keys as columns.
- Excludes selected model properties through `IgnoreInToTableAttribute`.
- Represents `Lazy(Of T)` properties without evaluating their values.
- Populates an existing `DataGridView` from any supported collection.
- Optionally preserves the selected row and first displayed row.
- Optionally reapplies the current grid sort.
- Optionally inserts a one-based `Order` column as the first column.
- Includes English XML documentation for the public API.
- Requires no third-party packages.

## Requirements

- .NET 8 for Windows (`net8.0-windows`)
- Windows Forms
- `CoreSuite.Helpers`

The NuGet dependency on `CoreSuite.Helpers` is resolved automatically when the package is installed.

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.Extensions.Windows
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.Extensions.Windows
```

## Namespaces

The extension methods and the exclusion attribute use these namespaces:

```vb
Imports CoreSuite.Extensions
Imports CoreSuite.Extensions.Extensions
```

## Quick start

```vb
Imports System.Data
Imports CoreSuite.Extensions.Extensions

Dim Customers As New List(Of Customer) From {
    New Customer With {.Id = 1, .Name = "Ana"},
    New Customer With {.Id = 2, .Name = "Bruno"}
}
Dim CustomerTable As DataTable = Customers.ToTable()
DataGridView1.Fill(Customers)
```

## API reference

### Collection extensions

| Member | Behavior |
| --- | --- |
| `ToTable(source As IEnumerable)` | Converts the public properties of each collection item to columns and rows. |
| `ToTable(source As IEnumerable(Of Dictionary(Of String, Object)))` | Creates columns from all discovered dictionary keys and rows from their values. |

### DataGridView extension

| Member | Default | Behavior |
| --- | --- | --- |
| `Fill(dgv, collection, keepSelection, keepSort, showOrder)` | `True`, `True`, `True` | Converts the collection to a table, assigns it to `DataSource` and restores the requested grid state. |

### Attribute

| Type | Behavior |
| --- | --- |
| `IgnoreInToTableAttribute` | Excludes a public property from object-to-table conversion. |

## Convert an object collection to a DataTable

Each public property of the collection item type becomes a `DataColumn`, and each item becomes a `DataRow`.

```vb
Imports System.Data
Imports CoreSuite.Extensions.Extensions

Dim Products As New List(Of Product) From {
    New Product With {.Id = 1, .Name = "Keyboard", .Price = 120D},
    New Product With {.Id = 2, .Name = "Mouse", .Price = 80D}
}
Dim ProductTable As DataTable = Products.ToTable()
```

Properties of type `Lazy(Of T)` are represented by text in the form `Lazy<TypeName>` and are not evaluated.

## Ignore a property

Apply `IgnoreInToTableAttribute` to a model property that must not become a table column:

```vb
Imports CoreSuite.Extensions

Public Class Customer
    Public Property Id As Integer
    Public Property Name As String
    <IgnoreInToTable>
    Public Property InternalNotes As String
End Class
```

The attribute is useful for calculated values, complex objects, internal references or any other property that should not appear in the resulting `DataTable`.

## Convert dictionaries to a DataTable

Every distinct key found across the dictionaries becomes a column. A missing or `Nothing` value is stored as `DBNull.Value`.

```vb
Imports System.Data
Imports CoreSuite.Extensions.Extensions

Dim Rows As New List(Of Dictionary(Of String, Object)) From {
    New Dictionary(Of String, Object) From {{"Id", 1}, {"Name", "Ana"}},
    New Dictionary(Of String, Object) From {{"Id", 2}, {"Name", "Bruno"}, {"Active", True}}
}
Dim Table As DataTable = Rows.ToTable()
```

Column names are collected case-insensitively. The dictionary collection overload returns `Nothing` when its source is `Nothing`.

## Fill a DataGridView

```vb
Imports CoreSuite.Extensions.Extensions

DataGridView1.Fill(
    Products,
    KeepSelection:=True,
    KeepSort:=True,
    ShowOrder:=True)
```

### Parameters

| Parameter | Default | Description |
| --- | --- | --- |
| `Collection` | Required | Collection converted through `ToTable`. |
| `KeepSelection` | `True` | Preserves the selected row index and first displayed row when exactly one row is selected. |
| `KeepSort` | `True` | Reapplies the current sort column and direction after rebinding. |
| `ShowOrder` | `True` | Adds a one-based `Order` column and places it at index zero. |

`Fill` replaces the grid's current `DataSource` with the generated `DataTable`.

## Important behavior

- Object conversion reads public properties through reflection, so ordinary property getters are executed.
- A string is not treated as a collection of items and causes `InvalidOperationException` in the object collection overload.
- The non-generic object collection must not be `Nothing`.
- A model property named `Order` conflicts with the optional order column. Set `ShowOrder:=False` when the source already contains that column name.
- Selection is preserved only when the grid has exactly one selected row before it is filled.
- Grid sorting depends on the generated columns supporting the previous sort operation.
- `IgnoreInToTableAttribute` affects object conversion only; dictionary keys are not filtered by the attribute.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.Extensions.Windows` |
| Namespaces | `CoreSuite.Extensions`, `CoreSuite.Extensions.Extensions` |
| Assembly | `CoreSuite.Extensions.Windows` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| CoreSuite dependency | `CoreSuite.Helpers` |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
