# CentralizedComboBox

**A Windows Forms ComboBox that keeps its selected text and drop-down items horizontally centered.**

> [!NOTE]
> CentralizedComboBox is one of the controls included in the **CoreSuite** solution. It preserves the standard WinForms `ComboBox` API and behavior while adding centered text rendering.

## Overview

`CentralizedComboBox` extends the standard Windows Forms `ComboBox`. It uses owner drawing to center every item in the drop-down list and applies the native Windows `ES_CENTER` style to the internal editable text field.

Because it inherits from `ComboBox`, existing properties, data binding, events and designer support remain available.

## Features

* Centers the selected value.
* Centers all items displayed in the drop-down list.
* Supports editable and selection-only ComboBox styles.
* Preserves standard data binding and selection behavior.
* Works in the Visual Studio Windows Forms designer.
* Requires no external NuGet dependencies.
* Designed for .NET 8 Windows Forms applications.

## Requirements

* Windows
* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

The control uses native Windows APIs and is intended exclusively for Windows applications.

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.CentralizedComboBox
```

Or search for `CoreSuite.CentralizedComboBox` in the Visual Studio NuGet Package Manager.

## Namespace

```vb
Imports CoreSuite.Controls
```

## Quick start

```vb
Imports CoreSuite.Controls
Public Class MainForm
    Private Sub MainForm_Load(Sender As Object, E As EventArgs) Handles MyBase.Load
        Dim StatusComboBox As New CentralizedComboBox With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Location = New Point(20, 20),
            .Width = 220
        }
        StatusComboBox.Items.AddRange({"Pending", "In progress", "Completed"})
        StatusComboBox.SelectedIndex = 0
        Controls.Add(StatusComboBox)
    End Sub
End Class
```

## Data binding

The control supports the same binding properties as a standard `ComboBox`:

```vb
Dim StatusComboBox As New CentralizedComboBox With {
    .DataSource = StatusItems,
    .DisplayMember = "Name",
    .ValueMember = "Id",
    .DropDownStyle = ComboBoxStyle.DropDownList
}
```

## Designer usage

After installing or referencing the package:

1. Open a Windows Forms form in the Visual Studio designer.
2. Add `CentralizedComboBox` from the Toolbox.
3. Configure it with the standard `ComboBox` properties.
4. Add items through `Items` or configure `DataSource`, `DisplayMember` and `ValueMember`.

No additional property is required to enable centering.

## API reference

### `CentralizedComboBox`

```vb
Public Class CentralizedComboBox
    Inherits ComboBox
```

The control does not introduce new public properties or events. All public configuration is inherited from `ComboBox`.

Common inherited members include:

* `Items`
* `SelectedIndex`
* `SelectedItem`
* `SelectedValue`
* `DataSource`
* `DisplayMember`
* `ValueMember`
* `DropDownStyle`
* `AutoCompleteMode`
* `AutoCompleteSource`
* `SelectedIndexChanged`
* `SelectionChangeCommitted`

## Rendering behavior

The control sets `DrawMode` to `OwnerDrawFixed`. Item text is rendered with horizontal and vertical centering. When the native handle is created, the editable portion also receives the Windows centered-text style.

If application code changes `DrawMode` or implements separate owner-drawing behavior, that code may replace the built-in centered item rendering.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.CentralizedComboBox` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.CentralizedComboBox` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| External dependencies | None |

## License

This package is distributed under the MIT License.
