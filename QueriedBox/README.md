# QueriedBox

**A database-backed lookup and selection control for WinForms, included in CoreSuite.**

> [!NOTE]
> QueriedBox is one of the projects that make up the **CoreSuite** solution. This page documents the QueriedBox control and its related query model.

## Overview

QueriedBox extends the standard WinForms `TextBox` with a complete database lookup workflow. As the user types, the control builds and executes a parameterized search, displays matching records in a floating `DataGridView`, and allows one record to be selected and retained as a frozen value.

The control keeps both sides of a lookup together:

- the text shown to the user;
- the primary key required by the application;
- the raw values returned by the selected row.

This makes QueriedBox useful for customer selectors, product searches, employee lookups, address finders, foreign-key editors, and other data-entry screens that would otherwise require a custom popup, timer, query builder, keyboard handling, and selection state management.

## Key features

- Search automatically after a configurable number of characters.
- Debounce queries with a configurable interval.
- Display results in a non-activating floating window.
- Navigate results with the keyboard or mouse.
- Store the selected primary key independently from the displayed text.
- Build the frozen text from one or more returned columns.
- Retrieve raw values from the selected row by column alias.
- Load and freeze a record programmatically by primary key.
- Configure columns, joins, conditions, parameters, ordering, distinct results, limits, and offsets through objects.
- Generate dialect-specific parameter prefixes, null-handling expressions, and pagination syntax.
- Customize the popup, grid, headers, selection, messages, and frozen text color.
- Configure the query model through Visual Studio collection editors.
- Use designer smart-tag actions for the most common search settings.

## Requirements

- Windows Forms
- .NET 8 for Windows (`net8.0-windows`)
- An ADO.NET provider that supplies a `DbConnection` and `DbProviderFactory`
- A reference to the `CoreSuite.QueriedBox` project or assembly

When using the project directly from the CoreSuite repository, add `QueriedBox/QueriedBox.vbproj` as a project reference. Its CoreSuite project dependencies are resolved by the solution.

## Supported SQL dialects

The query generator currently includes dedicated behavior for the following dialects:

| Dialect | Parameter prefix | Null replacement | Pagination |
|---|---:|---|---|
| MySQL | `@` | `IFNULL` | `LIMIT` / `OFFSET` |
| SQL Server | `@` | `ISNULL` | `OFFSET` / `FETCH` |
| PostgreSQL | `@` | `COALESCE` | `LIMIT` / `OFFSET` |
| SQLite | `@` | `IFNULL` | `LIMIT` / `OFFSET` |
| Oracle | `:` | `COALESCE` | `OFFSET` / `FETCH` |
| Firebird | `@` | `COALESCE` | `OFFSET` / `FETCH` |

The selected database provider is still responsible for opening connections, creating commands, binding parameters, and filling the result table.

> [!IMPORTANT]
> Use an `OrderBy` definition when the selected database requires ordering for paginated queries. SQL Server, for example, requires `ORDER BY` when `OFFSET` and `FETCH` are generated.

The database provider must be installed separately by the consuming application.
QueriedBox does not include or require a specific ADO.NET provider.

## Quick start

The following example configures a SQLite-backed person lookup. Every query column is assigned a unique alias because QueriedBox uses aliases to identify grid cells and frozen raw values.

### 1. Configure the shared connection factory

Configure the connection factory once during application startup. The factory must return a new connection every time it is invoked; QueriedBox manages opening, closing, and disposing it.

```vb
Imports System.Data.SQLite
Imports CoreSuite.Controls
Private Sub ConfigureDatabase()
    QueriedBox.ConnectionFactory = Function()
                                       Return New SQLiteConnection("Data Source=app.db;Version=3;")
                                   End Function
    QueriedBox.SqlDialect = SqlDialect.Sqlite
End Sub
```

`ConnectionFactory` and `SqlDialect` are shared settings, so they apply to all QueriedBox instances in the application.

### 2. Configure the lookup query

```vb
Imports System.Collections.ObjectModel
Imports CoreSuite.Controls
Private Sub ConfigurePersonBox()
    PersonBox.Query = New Query With {
        .Table = New QueryTable("Person", "p"),
        .PrimaryKeyColumnName = "p.Id",
        .Columns = New Collection(Of QueryColumn) From {
            New QueryColumn With {
                .ColumnName = "p.Name",
                .ColumnAlias = "Name",
                .Options = New QueryColumnOptions With {
                    .Display = True,
                    .Freeze = True,
                    .Searchable = True,
                    .SizeColumnMode = DataGridViewAutoSizeColumnMode.Fill
                }
            },
            New QueryColumn With {
                .ColumnName = "p.Email",
                .ColumnAlias = "Email",
                .IfNull = "''",
                .Options = New QueryColumnOptions With {
                    .Display = True,
                    .Freeze = True,
                    .Searchable = True,
                    .Prefix = " <",
                    .Suffix = ">",
                    .SizeColumnMode = DataGridViewAutoSizeColumnMode.AllCells
                }
            }
        },
        .OrderBy = New Collection(Of QueryOrderBy) From {
            New QueryOrderBy(New QueryColumnReference("p.Name"))
        },
        .Limit = 50
    }
    PersonBox.CharactersToQuery = 2
    PersonBox.QueryInterval = 300
End Sub
```

When the selected row contains `Alice` and `alice@example.com`, the control freezes the following display value:

```text
Alice <alice@example.com>
```

At the same time, `FrozenPrimaryKey` stores the value returned from `p.Id`.

### 3. Read the selected record

```vb
Private Sub PersonBox_FrozenPrimaryKeyChanged(sender As Object, e As FrozenPrimaryKeyEventArgs) Handles PersonBox.FrozenPrimaryKeyChanged
    If e.NewPrimaryKey Is Nothing Then
        SelectedPersonLabel.Text = "No person selected"
        Return
    End If
    SelectedPersonLabel.Text = $"Selected person ID: {e.NewPrimaryKey}"
    Dim Email = PersonBox.GetRawFrozenValueOf("Email")
End Sub
```

The alias passed to `GetRawFrozenValueOf` must match the configured `ColumnAlias`.

## Search behavior

After the minimum character count is reached, QueriedBox creates one `LIKE` condition for each column whose `Searchable` option is enabled. These conditions are grouped with `OR` and combined with the query's configured conditions.

Conceptually, the generated search resembles this:

```sql
SELECT p.Id AS <internal_primary_key_alias>, p.Name AS Name, IFNULL(p.Email, '') AS Email
FROM Person p
WHERE (LOWER(p.Name) LIKE LOWER(@search) OR LOWER(p.Email) LIKE LOWER(@search))
ORDER BY p.Name ASC
LIMIT 50;
```

The search text is passed through a command parameter with `%` added to both sides. Structural values such as table names, column names, aliases, joins, and raw condition expressions come from the developer-defined query configuration.

> [!TIP]
> Mark only textual columns as searchable unless the selected database safely supports applying `LOWER` and `LIKE` to the column type.

## Selecting and freezing records

A frozen record has four related values:

| Member | Purpose |
|---|---|
| `IsFrozen` | Indicates whether a record is currently selected. |
| `FrozenPrimaryKey` | Contains the selected record's primary key. |
| `FrozenValue` | Contains the final text assembled from frozen columns. |
| `GetRawFrozenValueOf(alias)` | Returns the unformatted value stored for a configured column alias. |

The frozen text is assembled from every column whose `Options.Freeze` property is `True`. `Prefix` and `Suffix` can be used to create a readable composite value without changing the raw data.

```vb
New QueryColumn With {
    .ColumnName = "p.Code",
    .ColumnAlias = "Code",
    .Options = New QueryColumnOptions With {
        .Display = True,
        .Freeze = True,
        .Searchable = True,
        .Prefix = "[",
        .Suffix = "] "
    }
}
```

A code and name combination can therefore be displayed as:

```text
[00042] Alice
```

### Programmatic selection

Use `Freeze` to load a record and select it without requiring the user to search first.

```vb
PersonBox.Freeze(42, "p.Id")
```

Use `Unfreeze` to clear the selected primary key and return the control to its editable state.

```vb
PersonBox.Unfreeze()
```

When the user edits text that no longer matches the frozen value, QueriedBox automatically removes the frozen state. `ClearOnUnfreeze` determines whether the current text is also cleared during this automatic transition.

## Keyboard and mouse interaction

| Input | Behavior |
|---|---|
| Type | Opens the popup and schedules a query after the configured interval. |
| `Down` / `Up` | Moves through the returned rows. |
| `Enter` | Freezes the selected row and closes the popup. |
| `Escape` | Closes the popup. |
| Double-click | Freezes the clicked row and closes the popup. |
| `Tab`, focus loss, or outside click | Freezes the selected row only when the control text exactly matches one of its configured column values; otherwise, closes the popup. |

The popup is implemented as a borderless, non-activating tool window, allowing the text box to retain the expected lookup-control behavior.

## Query model

QueriedBox uses a small object model to define SQL statements.

### `Query`

| Property | Description |
|---|---|
| `Table` | Main table and optional alias. |
| `PrimaryKeyColumnName` | Primary key expression used to identify the selected row. |
| `Columns` | Columns selected, displayed, searched, and frozen. |
| `Joins` | Join definitions and their column relationships. |
| `Conditions` | Additional `WHERE` conditions. |
| `Parameters` | Values bound to configured conditions. |
| `OrderBy` | Sorting definitions. |
| `Limit` | Maximum number of returned rows; defaults to `500`. |
| `Offset` | Number of rows skipped before results are returned. |
| `Distinct` | Enables `SELECT DISTINCT`. |
| `Dialect` | Optional dialect value exposed by the individual query model. |

The query model can also generate SQL text directly through `ToString`, `GetSelect`, `GetJoins`, `GetWhere`, and `GetOrder`.

### `QueryColumn`

| Property | Description |
|---|---|
| `ColumnName` | SQL column or expression. |
| `ColumnAlias` | Unique alias used in the result grid and frozen-value storage. |
| `IfNull` | SQL replacement expression used when the column is null. |
| `Options` | QueriedBox-specific display, search, freeze, formatting, and sizing options. |

`IfNull` is treated as a SQL expression. To replace a null text value with an empty string, use a quoted expression such as `"''"`.

### `QueryColumnOptions`

| Property | Description |
|---|---|
| `Display` | Shows or hides the column in the popup grid. |
| `Freeze` | Includes the value when assembling the final frozen text. |
| `Searchable` | Includes the column in the generated text-search conditions. |
| `Prefix` | Text inserted before the frozen value. |
| `Suffix` | Text inserted after the frozen value. |
| `SizeColumnMode` | Applies a `DataGridViewAutoSizeColumnMode` to the result column. |

A hidden column can still participate in searching or frozen-value storage by setting `Display` to `False` while keeping the other options enabled.

## Joins

Joins are defined with a join type, a table, and one or more relationships between column references.

```vb
Dim AddressJoin As New QueryJoin With {
    .Type = QueryJoinType.Left,
    .Table = New QueryTable("PersonAddress", "pa")
}
AddressJoin.Conditions.Add(New QueryJoinCondition(
    New QueryColumnReference("p.Id"),
    QueryJoinConditionOperator.Equal,
    New QueryColumnReference("pa.PersonId"),
    QueryRelation.And))
PersonBox.Query.Joins.Add(AddressJoin)
PersonBox.Query.Columns.Add(New QueryColumn With {
    .ColumnName = "pa.City",
    .ColumnAlias = "City",
    .Options = New QueryColumnOptions With {
        .Display = True,
        .Freeze = False,
        .Searchable = True,
        .SizeColumnMode = DataGridViewAutoSizeColumnMode.AllCells
    }
})
```

Available join types are `Inner`, `Left`, `Right`, and `Full`. The selected database must support the generated join type.

## Conditions and parameters

Additional conditions are appended to the generated lookup search. Use `QueryParameter` for values that should be bound through the ADO.NET command.

```vb
Dim ActiveParameter = QueriedBox.SqlDialect.GetParameterPrefix() & "active"
PersonBox.Query.Conditions.Add(New QueryCondition(
    New QueryColumnReference("p.Active"),
    QueryConditionOperator.Equal,
    {ActiveParameter},
    QueryRelation.And))
PersonBox.Query.Parameters.Add(New QueryParameter With {
    .ParameterName = ActiveParameter,
    .ParameterValue = "1"
})
```

Available condition operators are:

- `Equal`
- `NotEqual`
- `GreaterThan`
- `LessThan`
- `GreaterThanOrEqual`
- `LessThanOrEqual`
- `Like`
- `Between`
- `In`
- `NotIn`

A `Between` condition requires exactly two values. `In` and `NotIn` generate a parenthesized value list. The `Relation` stored on a condition determines how it is connected to the next configured condition.

> [!WARNING]
> Condition values are SQL expressions in the query model. Do not concatenate untrusted user input into them. Use `QueryParameter` for dynamic values. The normal text entered into QueriedBox is parameterized automatically.

## Events

| Event | Description |
|---|---|
| `FrozenPrimaryKeyChanging` | Raised before the stored primary key changes. |
| `FrozenPrimaryKeyChanged` | Raised after the stored primary key changes. |
| `HyperlinkClicked` | Raised when a frozen value is clicked while hyperlink behavior is active. |

`FrozenPrimaryKeyEventArgs` exposes `OldPrimaryKey` and `NewPrimaryKey`. The hyperlink event provides the current key through `NewPrimaryKey`.

### Hyperlink behavior

When `AllowHyperlink` is enabled, a frozen value can temporarily behave like a link. Hold `Ctrl` while the control is focused and click the underlined text to raise `HyperlinkClicked`.

```vb
Private Sub PersonBox_HyperlinkClicked(sender As Object, e As FrozenPrimaryKeyEventArgs) Handles PersonBox.HyperlinkClicked
    OpenPersonDetails(e.NewPrimaryKey)
End Sub
```

## Appearance and layout

QueriedBox exposes visual settings for both the text box state and its popup:

- `FreezeColor`
- `GridBackColor`
- `GridForeColor`
- `GridSelectionBackColor`
- `GridSelectionForeColor`
- `GridHeaderBackColor`
- `GridHeaderForeColor`
- `GridHeadersBold`
- `GridHeaderVisible`
- `ShowVerticalGridLines`
- `LabelBackColor`
- `LabelForeColor`
- `DropDownBorderColor`

Popup dimensions can be adjusted with:

- `DropDownStretchLeft`
- `DropDownStretchRight`
- `DropDownStretchDown`
- `DropDownAutoStretchRight`

When automatic right stretching is enabled, QueriedBox expands the popup until the horizontal scrollbar is no longer required.

## User-facing messages

The messages displayed by the popup can be customized without changing the control source:

```vb
PersonBox.NoResultsText = "No people found."
PersonBox.CharactersRemainingSingularText = "Type {0} more character to search."
PersonBox.CharactersRemainingPluralText = "Type {0} more characters to search."
```

The `{0}` placeholder is replaced with the remaining character count.

## Design-time support

The `Query` property is expandable in the Visual Studio Property Grid. Collections such as `Columns`, `Joins`, `Conditions`, `Parameters`, and `OrderBy` use collection editors and are serialized by the WinForms designer.

The designer smart tag provides direct access to:

- `QueryEnabled`
- `QueryInterval`
- `CharactersToQuery`

The custom designer also preserves standard horizontal resizing for a single-line QueriedBox and enables full resizing when multiline mode is selected. Enabling multiline mode disables query functionality because QueriedBox is primarily designed as a single-line lookup control.

## Configuration rules

QueriedBox validates its configuration before running a typed search. A valid setup must satisfy the following rules:

1. `QueryInterval` must be at least `300` milliseconds.
2. `ConnectionFactory` must be configured.
3. `Query` and `Query.Table.TableName` must be configured.
4. `Query.PrimaryKeyColumnName` must be configured.
5. At least one query column must exist.
6. At least one column must be displayed.
7. At least one column must participate in the frozen text.
8. Every column must have a non-empty `ColumnName` and `ColumnAlias`.
9. Column aliases must be unique.
10. At least one column should be searchable for typed lookup behavior.
11. Parameter names must be valid and unique.
12. Every join must define a table and at least one valid join condition.
13. `Between` conditions must contain exactly two values.

Invalid configurations close the popup and raise an `InvalidOperationException` with a descriptive message.

## Debugging

Enable `DebugOnTextChanged` to print the generated command and parameters through the CoreSuite database debugging helper whenever a lookup query runs.

```vb
PersonBox.DebugOnTextChanged = True
```

Use `DebugBaseQuery` to print the base `SELECT` and `JOIN` portion of the configured query:

```vb
PersonBox.DebugBaseQuery()
```

The query model can also be inspected directly:

```vb
Debug.Print(PersonBox.Query.ToString())
```

## Performance recommendations

- Keep `Limit` at a practical value for lookup scenarios.
- Increase `CharactersToQuery` for large tables.
- Increase `QueryInterval` when users are likely to type long terms.
- Mark only useful columns as searchable.
- Index frequently searched database columns.
- Avoid expensive expressions and unnecessary joins in interactive lookups.
- Keep the connection factory lightweight and return a fresh connection instance.

Query execution is synchronous in the current implementation, so lookup queries should remain fast enough for interactive UI use.

## Typical use cases

- Selecting a customer while retaining the customer ID.
- Searching products by code, name, or barcode.
- Choosing employees, suppliers, cities, accounts, or categories.
- Editing foreign-key fields without exposing numeric IDs to users.
- Displaying composite selected values assembled from multiple columns.
- Opening a detailed record through the optional hyperlink event.

## Project information

- **Project:** QueriedBox
- **Assembly:** `CoreSuite.QueriedBox`
- **Namespace:** `CoreSuite.Controls`
- **Framework:** `.NET 8 for Windows`
- **UI technology:** Windows Forms
- **Repository:** Part of the CoreSuite solution

QueriedBox is designed to remove repetitive lookup-control code while keeping SQL configuration, result presentation, primary-key state, and WinForms interaction in one reusable CoreSuite component.
