# TileView

**Designer-editable tile lists and grids for Windows Forms.**

> [!NOTE]
> TileView is one of the controls included in the **CoreSuite** solution. It has no CoreSuite package dependencies and requires no third-party packages.

## Overview

`TileView` provides a reusable list or grid of visual tiles while keeping the programming model close to standard Windows Forms controls.

The container inherits directly from `FlowLayoutPanel`, so standard layout properties such as `FlowDirection`, `WrapContents`, `AutoScroll`, `Padding`, `Margin`, `Dock`, and `Anchor` remain available. Each tile inherits from `UserControl` through `TileViewItem`, which allows application-specific tile classes to continue using the normal Visual Studio Windows Forms Designer.

The control centralizes behavior that would otherwise have to be repeated in every screen: single selection, selected-item tracking, keyboard navigation, scrolling, activation, filtering, and item events.

## Features

* `TileView` inherits from the standard Windows Forms `FlowLayoutPanel`.
* `TileViewItem` inherits from the standard Windows Forms `UserControl`.
* Application-specific tiles can be edited visually in the Windows Forms Designer.
* Single selection is managed by the container instead of by sibling tiles.
* `SelectedItem` and `SelectedIndex` expose the current selection.
* Clicks and double-clicks on child labels, panels, pictures, and other nested controls are forwarded to the tile.
* Arrow keys navigate using the visual position of neighboring tiles.
* Home and End select the first and last visible items.
* Enter activates the selected item.
* Space selects the focused item without forcing an activation.
* Selected items are automatically scrolled into view.
* Optional navigation wrapping.
* Optional first-item automatic selection.
* Optional selection clearing when blank view space is clicked.
* Built-in filtering through a `Predicate(Of TileViewItem)`.
* Generic `SetItems` helper for building tiles from application data.
* `BeginUpdate` and `EndUpdate` helpers reduce intermediate layout work during bulk changes.
* Native Windows colors are used by default for borders, selection, background, and focus cues.
* Uses the standard Windows Forms event model.
* Includes English XML documentation for the public API and internal behavior.
* Requires no third-party packages.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.TileView
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.TileView
```

## Namespace

```vb
Imports CoreSuite.Controls
```

## Designer usage

### Adding TileView to a form or UserControl

After installing or referencing the package, add `TileView` from the Visual Studio Toolbox exactly as you would add a `FlowLayoutPanel`.

Because `TileView` inherits from `FlowLayoutPanel`, the standard layout properties remain available in the Properties window. A vertical list can be configured with:

```text
FlowDirection = TopDown
WrapContents = False
AutoScroll = True
```

A wrapping tile grid can use:

```text
FlowDirection = LeftToRight
WrapContents = True
AutoScroll = True
```

### Creating a tile visually

Create a normal Windows Forms UserControl in the consuming application and change its base class from `UserControl` to `TileViewItem`:

```vb
Imports CoreSuite.Controls

Public Class UcLicenseTile
    Inherits TileViewItem
End Class
```

The control remains a designer-editable UserControl. Labels, panels, picture boxes, and other controls can be positioned visually in the `.Designer.vb` file exactly as with an ordinary UserControl.

For data-driven tiles, keep a public parameterless constructor for the Visual Studio Designer and add a second constructor for runtime data:

```vb
Public Class UcLicenseTile
    Inherits TileViewItem

    Public Property CustomerLicense As CustomerLicenseModel

    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(customerLicense As CustomerLicenseModel)
        Me.New()
        Me.CustomerLicense = customerLicense
        Reload()
    End Sub

    Public Sub Reload()
        If CustomerLicense Is Nothing Then Return
        LblCustomerNameValue.Text = CustomerLicense.Name
        LblCustomerDocumentValue.Text = CustomerLicense.Document
        LblExpirationDateValue.Text = CustomerLicense.Expiration.ToLongDateString()
        LblLicenseKeyValue.Text = CustomerLicense.LicenseKey
    End Sub
End Class
```

> [!IMPORTANT]
> Keep the parameterless constructor. The Windows Forms Designer uses it when opening the inherited UserControl at design time.

## Quick start

Create tiles from a data source with `SetItems`:

```vb
Dim licenses = Await _Service.GetAll()
TileLicenses.SetItems(licenses, Function(license) New UcLicenseTile(license))
```

Open the selected license when an item is activated:

```vb
Private Sub TileLicenses_ItemActivated(sender As Object, e As TileViewItemEventArgs) Handles TileLicenses.ItemActivated
    Dim tile = DirectCast(e.Item, UcLicenseTile)
    Using form As New FrmLicense(_Service, tile.CustomerLicense)
        form.ShowDialog()
    End Using
    tile.Reload()
End Sub
```

A toolbar Edit button can activate the same selected tile without searching the container controls:

```vb
Private Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnEdit.Click
    TileLicenses.ActivateSelected()
End Sub
```

## Filtering

Filtering changes only item visibility and keeps layout management inside the TileView:

```vb
Private Sub TxtFilter_TextChanged(sender As Object, e As EventArgs) Handles TxtFilter.TextChanged
    Dim filter = TxtFilter.Text.Trim()
    If filter.Length = 0 Then
        TileLicenses.ClearFilter()
        Return
    End If

    TileLicenses.Filter(
        Function(item)
            Dim tile = DirectCast(item, UcLicenseTile)
            Return tile.CustomerLicense.Name.Contains(filter, StringComparison.CurrentCultureIgnoreCase) OrElse
                   tile.CustomerLicense.Document.Contains(filter, StringComparison.CurrentCultureIgnoreCase)
        End Function)
End Sub
```

If the currently selected tile becomes hidden by a filter, selection is cleared automatically.

## Selection

Selection belongs to the `TileView`, not to individual sibling tiles.

```vb
Dim selectedTile = TryCast(TileLicenses.SelectedItem, UcLicenseTile)
If selectedTile Is Nothing Then Return

Dim license = selectedTile.CustomerLicense
```

Select an item programmatically:

```vb
TileLicenses.SelectedItem = tile
```

Or by index:

```vb
TileLicenses.SelectedIndex = 0
```

Clear selection:

```vb
TileLicenses.ClearSelection()
```

This design intentionally avoids having each tile search its parent for sibling tiles to deselect.

## Keyboard behavior

When a tile has focus, the TileView processes common list-navigation keys:

| Key | Behavior |
| --- | --- |
| `Up` | Selects the closest enabled and visible tile above the current tile. |
| `Down` | Selects the closest enabled and visible tile below the current tile. |
| `Left` | Selects the closest enabled and visible tile to the left. |
| `Right` | Selects the closest enabled and visible tile to the right. |
| `Home` | Selects the first enabled and visible tile. |
| `End` | Selects the last enabled and visible tile. |
| `Enter` | Activates the selected tile. |
| `Space` | Selects the focused tile. |

Directional navigation is based on actual tile positions, so it works with both vertical lists and wrapped grids.

## Mouse behavior

A click anywhere inside a tile selects it. Child-control clicks are forwarded to the tile, so labels and nested layout panels do not need manual click handlers.

A double-click raises `ItemDoubleClick` and, when `ActivateOnDoubleClick` is enabled, also raises `ItemActivated`.

Clicking unused TileView space clears the selection by default. Set `ClearSelectionOnBlankClick` to `False` to retain it.

## Appearance

`TileViewItem` uses native Windows system colors by default:

| Property | Default |
| --- | --- |
| `BorderColor` | `SystemColors.ControlDark` |
| `SelectedBorderColor` | `SystemColors.Highlight` |
| `BorderWidth` | `1` |
| `SelectedBorderWidth` | `2` |
| `ShowFocusRectangle` | `True` |

`TileView` uses `SystemColors.Window` as its default background and a standard fixed border.

The tile itself remains a normal UserControl, so its internal appearance is completely application-defined in the Designer.

## API reference

### `TileView`

Represents the tile container and selection manager.

```vb
Public Class TileView
    Inherits FlowLayoutPanel
```

### Main TileView members

| Member | Behavior |
| --- | --- |
| `Items` | Read-only snapshot of the current `TileViewItem` controls. |
| `ItemCount` | Number of tile items. |
| `SelectedItem` | Gets or sets the current selected tile. |
| `SelectedIndex` | Gets or sets the selected tile index. |
| `ActivateOnDoubleClick` | Controls whether double-click activates a tile. |
| `ClearSelectionOnBlankClick` | Controls whether blank-space clicks clear selection. |
| `AutoSelectFirst` | Selects the first item when appropriate. |
| `WrapNavigation` | Enables keyboard navigation wrapping. |
| `Add(item)` | Adds a tile. |
| `AddRange(items)` | Adds multiple tiles. |
| `SetItems(source, factory)` | Rebuilds the view from application data. |
| `Remove(item)` | Removes a tile. |
| `ClearItems()` | Removes all items. |
| `SelectItem(item)` | Selects a specific item. |
| `ClearSelection()` | Clears selection. |
| `SelectFirst()` | Selects the first visible enabled item. |
| `SelectLast()` | Selects the last visible enabled item. |
| `SelectNext()` | Moves to the next item in control order. |
| `SelectPrevious()` | Moves to the previous item in control order. |
| `EnsureVisible(item)` | Scrolls an item into view. |
| `ActivateSelected()` | Activates the selected tile. |
| `ActivateItem(item)` | Selects and activates a specific tile. |
| `Filter(predicate)` | Shows only tiles accepted by the predicate. |
| `ClearFilter()` | Makes all tiles visible. |
| `BeginUpdate()` | Suspends layout during bulk changes. |
| `EndUpdate()` | Resumes layout after bulk changes. |

### TileView events

| Event | Behavior |
| --- | --- |
| `SelectedItemChanged` | Raised whenever the selected item changes. |
| `SelectionChanged` | Raised with previous and current selected items. |
| `ItemClick` | Raised when a tile is clicked. |
| `ItemDoubleClick` | Raised when a tile is double-clicked. |
| `ItemActivated` | Raised after tile activation. |
| `ItemAdded` | Raised when a tile is added. |
| `ItemRemoved` | Raised when a tile is removed. |

### `TileViewItem`

Represents the designer-editable base class for application-specific tiles.

```vb
Public Class TileViewItem
    Inherits UserControl
```

### Main TileViewItem members

| Member | Behavior |
| --- | --- |
| `IsSelected` | Indicates whether the item is selected. |
| `OwnerView` | Returns the owning TileView. |
| `BorderColor` | Border color used while unselected. |
| `SelectedBorderColor` | Border color used while selected. |
| `BorderWidth` | Unselected border width. |
| `SelectedBorderWidth` | Selected border width. |
| `ShowFocusRectangle` | Shows the standard Windows focus rectangle. |
| `SelectOnFocus` | Selects the item when it receives focus. |
| `FocusOnClick` | Moves keyboard focus to the tile when clicked. |
| `SelectItem()` | Requests selection through the owner. |
| `ActivateItem()` | Requests activation through the owner. |
| `Activated` | Item-level activation event. |

## Recommended licensing screen

With TileView, the screen that previously had to search through `FlowLayoutPanel.Controls`, manually find the selected tile, and assign one action delegate to every tile can be reduced to:

```vb
Imports CoreSuite.Infrastructure

Public Class UcLicensing
    Private ReadOnly _Service As CustomerLicenseService

    Public Sub New()
        InitializeComponent()
        _Service = Locator.GetInstance(Of CustomerLicenseService)()
    End Sub

    Private Async Function LoadLicenses() As Task
        Await LoadingOverlay.RunAsync(
        Async Function()
            Dim watch = Stopwatch.StartNew()
            Dim customerLicenses = Await _Service.GetAll()
            TileLicenses.SetItems(customerLicenses, Function(license) New UcLicenseTile(license))
            watch.Stop()
            Dim remainingTime = TimeSpan.FromSeconds(2) - watch.Elapsed
            If remainingTime > TimeSpan.Zero Then Await Task.Delay(remainingTime)
        End Function)
    End Function

    Private Async Sub UcLicensing_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Await LoadLicenses()
    End Sub

    Private Sub TileLicenses_ItemActivated(sender As Object, e As TileViewItemEventArgs) Handles TileLicenses.ItemActivated
        Dim tile = DirectCast(e.Item, UcLicenseTile)
        Using form As New FrmLicense(_Service, tile.CustomerLicense)
            form.ShowDialog()
        End Using
        tile.Reload()
    End Sub

    Private Sub BtnInclude_Click(sender As Object, e As EventArgs) Handles BtnInclude.Click
        Using form As New FrmLicense(_Service, Nothing)
            form.ShowDialog()
        End Using
    End Sub

    Private Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnEdit.Click
        TileLicenses.ActivateSelected()
    End Sub

    Private Async Sub BtnRefresh_Click(sender As Object, e As EventArgs) Handles BtnRefresh.Click
        Await LoadLicenses()
    End Sub

    Private Sub TxtFilter_TextChanged(sender As Object, e As EventArgs) Handles TxtFilter.TextChanged
        Dim filter = TxtFilter.Text.Trim()
        If filter.Length = 0 Then
            TileLicenses.ClearFilter()
            Return
        End If
        TileLicenses.Filter(
            Function(item)
                Dim tile = DirectCast(item, UcLicenseTile)
                Return tile.CustomerLicense.Name.Contains(filter, StringComparison.CurrentCultureIgnoreCase) OrElse
                       tile.CustomerLicense.Document.Contains(filter, StringComparison.CurrentCultureIgnoreCase)
            End Function)
    End Sub
End Class
```

## Standard functionality

Because `TileView` inherits from `FlowLayoutPanel`, standard FlowLayoutPanel and Panel members remain available unless the control intentionally overrides their behavior.

Because `TileViewItem` inherits from `UserControl`, standard UserControl members, layout, resources, accessibility, anchoring, docking, child controls, and Visual Studio Designer serialization remain available.

## Behavior and validation

* A `TileViewItem` can belong to only one `TileView` at a time.
* Non-`TileViewItem` controls can still be hosted because `TileView` remains a real `FlowLayoutPanel`, but they do not participate in selection, filtering, navigation, or item events.
* Disabled or hidden tiles are ignored by keyboard navigation and activation.
* Filtering clears selection when the selected tile becomes hidden.
* `SetItems` clears the current controls before creating new tiles.
* Event handlers use the normal Windows Forms UI thread model.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.TileView` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.TileView` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| CoreSuite dependencies | None |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
