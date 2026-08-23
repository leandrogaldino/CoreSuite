# StatusSelector

**A reusable and highly customizable status/state selector for WinForms ToolStrip interfaces, included in CoreSuite.**

> [!NOTE]
> StatusSelector is one of the projects that make up the **CoreSuite** solution. It is designed for screens that need to select a state such as Active/Inactive, Pending/Approved/Cancelled, workflow stages, task states, license states, or any other finite application status.

## Overview

`StatusSelector` inherits from `ToolStripDropDownButton` and turns it into a dedicated state selector. Each dropdown item combines presentation information with an arbitrary application value, so business logic does not need to depend on display strings.

A selector can represent two states:

```text
Ativo ▼
  ✓ Ativo
    Inativo
```

or any larger workflow:

```text
Em andamento ▼
    Pendente
    Aprovado
  ✓ Em andamento
    Concluído
    Cancelado
```

The control automatically updates the button text, button color, selected indicator and menu item state when a new value is selected.

## Key features

- Inherits from the standard WinForms `ToolStripDropDownButton`.
- Supports any number of states.
- Associates each item with any `Object`, Boolean, numeric value, string or enum.
- Exposes `SelectedItem`, `SelectedIndex`, `SelectedValue` and `SelectedText`.
- Supports `Items.Add(...)` and convenience `Add(...)` overloads.
- Allows each item to define its own foreground and background colors.
- Allows each item to override global hover foreground and background colors.
- Supports optional item images, custom fonts and tooltips.
- Marks the currently selected item with a configurable check mark.
- Automatically reflects the selected item text and color on the dropdown button.
- Supports optional selected-item background propagation to the button.
- Customizable menu background, menu border, image/check margin, hover background, hover foreground and hover border.
- Customizable menu item padding, menu padding and minimum dropdown width.
- Optional automatic dropdown width based on visible items.
- Raises `SelectedItemChanged`, `SelectedIndexChanged`, `SelectedValueChanged` and `StatusItemClick` events.
- Reacts automatically when an existing `StatusSelectorItem` is modified after being added.
- Does not require application logic to compare status display strings.

## Requirements

- Windows Forms
- .NET 8 for Windows (`net8.0-windows`)
- A reference to `CoreSuite.StatusSelector`

## Installation

```powershell
dotnet add package CoreSuite.StatusSelector
```

Or add `StatusSelector/StatusSelector.vbproj` as a project reference when working directly with the CoreSuite solution.

## Quick start

Add a `StatusSelector` to a `ToolStrip` and configure its values after `InitializeComponent()`.

```vb
Private Sub ConfigureStatus()
    DdbStatusValue.Items.Add("Ativo", True, Color.DodgerBlue)
    DdbStatusValue.Items.Add("Inativo", False, Color.Firebrick)
End Sub
```

Load the current value:

```vb
DdbStatusValue.SelectedValue = Company.IsActive
```

Read the selected value:

```vb
Company.IsActive = CBool(DdbStatusValue.SelectedValue)
```

Track changes:

```vb
Private Sub DdbStatusValue_SelectedValueChanged(sender As Object, e As EventArgs) Handles DdbStatusValue.SelectedValueChanged
    UpdateSaveButton()
End Sub
```

## Using enums

The item value is not limited to Boolean values. Enums are recommended for screens with more than two states.

```vb
Public Enum OrderStatus
    Pending
    Approved
    InProgress
    Completed
    Cancelled
End Enum
```

Configure the selector:

```vb
DdbStatusValue.Items.Add("Pendente", OrderStatus.Pending, Color.DarkOrange)
DdbStatusValue.Items.Add("Aprovado", OrderStatus.Approved, Color.RoyalBlue)
DdbStatusValue.Items.Add("Em andamento", OrderStatus.InProgress, Color.DodgerBlue)
DdbStatusValue.Items.Add("Concluído", OrderStatus.Completed, Color.SeaGreen)
DdbStatusValue.Items.Add("Cancelado", OrderStatus.Cancelled, Color.Firebrick)
```

Select and retrieve the enum value:

```vb
DdbStatusValue.SelectedValue = Order.Status
Order.Status = DirectCast(DdbStatusValue.SelectedValue, OrderStatus)
```

## Per-item customization

`Items.Add` returns the created `StatusSelectorItem`, allowing additional customization immediately.

```vb
Dim Pending = DdbStatusValue.Items.Add("Pendente", OrderStatus.Pending, Color.DarkOrange)
Pending.BackColor = Color.FromArgb(255, 250, 235)
Pending.HoverBackColor = Color.FromArgb(255, 245, 215)
Pending.HoverForeColor = Color.DarkOrange
Pending.ToolTipText = "Aguardando processamento"
```

An item supports:

| Property | Description |
|---|---|
| `Text` | Display text. |
| `Value` | Application value represented by the item. |
| `ForeColor` | Normal text color. |
| `BackColor` | Normal background color. `Color.Empty` inherits `MenuBackColor`. |
| `HoverForeColor` | Hover text color. `Color.Empty` inherits the selector setting. |
| `HoverBackColor` | Hover background color. `Color.Empty` inherits `HoverBackColor`. |
| `Enabled` | Enables or disables selection. |
| `Visible` | Shows or hides the item. |
| `Image` | Optional image displayed beside the item and on the selected button. |
| `ToolTipText` | Optional tooltip. |
| `Font` | Optional custom item font. |

Changing these properties after the item has been added automatically rebuilds the dropdown presentation.

## Menu appearance

The default menu uses a white background and a subtle light-blue hover:

```vb
DdbStatusValue.MenuBackColor = Color.White
DdbStatusValue.MenuBorderColor = Color.FromArgb(210, 210, 210)
DdbStatusValue.HoverBackColor = Color.FromArgb(240, 247, 255)
DdbStatusValue.HoverBorderColor = Color.FromArgb(210, 230, 250)
DdbStatusValue.ImageMarginBackColor = Color.White
```

To preserve each item's own text color during hover, leave `HoverForeColor` as `Color.Empty`:

```vb
DdbStatusValue.HoverForeColor = Color.Empty
```

Or force a shared hover text color:

```vb
DdbStatusValue.HoverForeColor = Color.Navy
```

## Selection indicator

The selected menu item is marked with a configurable check mark by default.

```vb
DdbStatusValue.ShowSelectedCheckMark = True
DdbStatusValue.SelectionIndicatorColor = Color.DodgerBlue
DdbStatusValue.SelectionIndicatorThickness = 2.0F
```

Disable it when the selected text on the button is enough:

```vb
DdbStatusValue.ShowSelectedCheckMark = False
```

## Button appearance

By default, the button adopts the selected item's `ForeColor` but not its background color.

```vb
DdbStatusValue.UseSelectedItemForeColor = True
DdbStatusValue.UseSelectedItemBackColor = False
```

The unselected state can also be customized:

```vb
DdbStatusValue.UnselectedText = "Select..."
DdbStatusValue.UnselectedForeColor = SystemColors.ControlText
DdbStatusValue.UnselectedBackColor = Color.Empty
```

## Layout customization

```vb
DdbStatusValue.MenuItemPadding = New Padding(8, 5, 8, 5)
DdbStatusValue.MenuPadding = New Padding(2)
DdbStatusValue.MenuMinimumWidth = 150
DdbStatusValue.AutoSizeDropDownWidth = True
DdbStatusValue.MenuRoundedEdges = False
```

When `AutoSizeDropDownWidth` is enabled, the menu expands to fit the widest visible item and never becomes narrower than the selector button or `MenuMinimumWidth`.

## Selecting values programmatically

Assigning a value selects the first matching item:

```vb
DdbStatusValue.SelectedValue = OrderStatus.Completed
```

Use `SelectValue` when you need to know whether a matching value exists:

```vb
If Not DdbStatusValue.SelectValue(OrderStatus.Completed) Then
    Throw New InvalidOperationException("The requested status is not available.")
End If
```

Use `FindByValue` to obtain the item without changing the selection:

```vb
Dim Item = DdbStatusValue.FindByValue(OrderStatus.Pending)
```

Clear the current selection while keeping all items:

```vb
DdbStatusValue.ClearSelection()
```

## Complete company example

```vb
Private Sub ConfigureStatus()
    DdbStatusValue.Items.Add("Ativo", True, Color.DodgerBlue)
    DdbStatusValue.Items.Add("Inativo", False, Color.Firebrick)
End Sub

Private Sub FillFormWithModel()
    _Loading = True
    DdbStatusValue.SelectedValue = _Company.IsActive
    _Loading = False
End Sub

Private Sub FillModelWithForm()
    _Company.IsActive = CBool(DdbStatusValue.SelectedValue)
End Sub

Private Function HasChanges() As Boolean
    Return _Company.IsActive <> CBool(DdbStatusValue.SelectedValue)
End Function

Private Sub DdbStatusValue_SelectedValueChanged(sender As Object, e As EventArgs) Handles DdbStatusValue.SelectedValueChanged
    If _Loading Then Return
    UpdateSaveButton()
End Sub
```

## Design notes

`StatusSelector` deliberately keeps the status value separate from its display text. The text can therefore be translated or renamed without changing application logic.

The collection is intended to be configured in application code, usually in a form or control constructor after `InitializeComponent()`. This matches scenarios where available states vary by screen, workflow, permissions or business rules.

## Package

- Package ID: `CoreSuite.StatusSelector`
- Target framework: `net8.0-windows`
- Namespace: `CoreSuite.Controls`
- License: MIT
