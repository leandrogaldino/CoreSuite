# NavigationView

**A Windows Forms navigation control that replaces tab-style page selection with a configurable navigation pane and lazily created `UserControl` pages, included in CoreSuite.**

> [!NOTE]
> NavigationView is one of the independent projects that make up the **CoreSuite** solution. It has no dependency on another CoreSuite package.

## Overview

`NavigationView` combines a navigation pane, a vertical `FlowLayoutPanel` of page buttons, and a content panel that displays one `UserControl` at a time. It is intended for administrative systems, dashboards, settings windows, registrations, and other applications whose screens are better represented by named navigation destinations than by traditional tabs.

Pages are created only when first requested. Each page can preserve its instance and visual state with `KeepAlive`, or dispose its control whenever navigation leaves it with `Recreate`. Pages that have parameterless constructors can be configured in the Windows Forms Designer, while controls that require services or other constructor arguments can be registered with a run-time factory.

## Key features

- Provides an integrated navigation pane and content area.
- Uses an automatically scrolling vertical `FlowLayoutPanel` for navigation buttons.
- Provides a designer-serializable `Pages` collection.
- Supports page lookup by numeric index or case-insensitive string key.
- Creates each `UserControl` lazily on first navigation.
- Supports `KeepAlive` and `Recreate` page caching modes.
- Accepts a `ControlType` for designer-configured pages.
- Accepts a run-time `Func(Of UserControl)` factory for dependency injection and custom constructors.
- Provides cancelable navigation through `Navigating`.
- Provides cancelable explicit page closing through `PageClosing`.
- Provides page creation, disposal, selection, completion, and failure events.
- Supports programmatic navigation, reload, close, cache clearing, and cached-control lookup.
- Supports hidden pages that remain reachable programmatically.
- Supports left and right navigation-pane placement.
- Supports configurable widths, heights, spacing, padding, image sizing, colors, tooltips, and selection indicator.
- Supports keyboard focus, accessible names, text ellipsis, right-to-left text, and high-quality image scaling.
- Does not clone or dispose images assigned to page definitions.

## Requirements

- .NET 8 for Windows
- Windows Forms
- Windows

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.NavigationView
```

## Quick start

Import the control namespace:

```vbnet
Imports CoreSuite.Controls
```

Add pages whose controls have public parameterless constructors:

```vbnet
NavigationView1.Pages.Add(Of CustomersControl)("customers", "Customers", My.Resources.Customers20)
NavigationView1.Pages.Add(Of ProductsControl)("products", "Products", My.Resources.Products20)
NavigationView1.Pages.Add(Of SettingsControl)("settings", "Settings", My.Resources.Settings20)
```

The first visible enabled page opens automatically when `AutoNavigateFirstPage` is `True`. Navigate explicitly when another page should open first:

```vbnet
NavigationView1.Navigate("products")
```

## Designer usage

After installing the package and adding `NavigationView` to the Visual Studio Toolbox:

1. Place `NavigationView` on the form.
2. Set `Dock` to `Fill` when it should occupy the complete client area.
3. Open the `Pages` collection editor from the Property Grid or smart tag.
4. Add one `NavigationPage` for each destination.
5. Configure `Key`, `Text`, `Image`, `ToolTipText`, `ControlType`, `CacheMode`, `Visible`, and `Enabled`.
6. Choose a `ControlType` derived from `UserControl` with a public parameterless constructor.
7. Configure layout and appearance properties on the `NavigationView` itself.

The Designer displays the navigation buttons but does not instantiate page controls. Page instances are created only at run time.

The generated form code follows the normal collection-content pattern:

```vbnet
Dim CustomersPage As New NavigationPage With {
    .Key = "customers",
    .Text = "Customers",
    .ControlType = GetType(CustomersControl),
    .CacheMode = NavigationPageCacheMode.KeepAlive
}
NavigationView1.Pages.Add(CustomersPage)
```

## Creating pages in code

### Generic overload

Use the generic overload for a `UserControl` with a public parameterless constructor:

```vbnet
Dim CustomersPage As NavigationPage = NavigationView1.Pages.Add(Of CustomersControl)(
    "customers",
    "Customers",
    My.Resources.Customers20)
```

### Type overload

The non-generic equivalent accepts a `Type`:

```vbnet
NavigationView1.Pages.Add(
    "products",
    "Products",
    GetType(ProductsControl),
    My.Resources.Products20)
```

`ControlType` must derive from `UserControl`, must not be abstract, and must provide a public parameterless constructor.

### Factory overload

Use a factory when a page needs constructor arguments or services:

```vbnet
NavigationView1.Pages.Add(
    "orders",
    "Orders",
    My.Resources.Orders20,
    Function() New OrdersControl(_OrderService, _CurrentUser))
```

The factory is stored only at run time and is not serialized by the Windows Forms Designer. It must return a new, non-disposed, unparented `UserControl`. When both `Factory` and `ControlType` are assigned, `Factory` takes precedence.

## Accessing pages

### By numeric index

```vbnet
Dim FirstPage As NavigationPage = NavigationView1.Pages(0)
```

### By key

```vbnet
Dim ProductsPage As NavigationPage = NavigationView1.Pages("products")
ProductsPage.Enabled = False
```

Key comparison is case-insensitive. `Pages("Products")` and `Pages("products")` identify the same page.

The string indexer validates as follows:

| Condition | Result |
|---|---|
| Key is `Nothing`, empty, or whitespace | Throws `ArgumentException`. |
| No page has the key | Throws `KeyNotFoundException`. |
| A page already has the key when adding or renaming | Throws `ArgumentException`. |

Use `FindByKey` when a missing page should return `Nothing`:

```vbnet
Dim SettingsPage As NavigationPage = NavigationView1.Pages.FindByKey("settings")
If SettingsPage IsNot Nothing Then SettingsPage.Visible = True
```

The collection also provides `Contains(key)`, `IndexOf(key)`, and `Remove(key)` overloads.

## Navigation lifecycle

When `Navigate` succeeds, the control performs these operations:

1. Resolves and validates the target page.
2. Raises the cancelable `Navigating` event.
3. Creates the target `UserControl` if necessary.
4. Adds it to the content area with `Dock = Fill`.
5. Hides the previous control and displays the target.
6. Updates the selected navigation button.
7. Disposes the previous control when its cache mode is `Recreate`.
8. Raises `Navigated`.

A page with `Visible = False` does not have a button but can still be opened through `Navigate`. A page with `Enabled = False` cannot be opened through its button or programmatically.

## Canceling navigation

Use `Navigating` to protect unsaved changes:

```vbnet
Private Sub NavigationView1_Navigating(Sender As Object, E As NavigationCancelEventArgs) Handles NavigationView1.Navigating
    If E.CurrentPage IsNot Nothing AndAlso HasUnsavedChanges(E.CurrentPage) Then
        E.Cancel = MessageBox.Show(
            "Discard the current changes?",
            "Navigation",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question) = DialogResult.No
    End If
End Sub
```

`CurrentPage` is `Nothing` during the first navigation. `TargetPage` always identifies the requested destination.

## Cache modes

### KeepAlive

`KeepAlive` is the default. The page control remains inside the content area while hidden, preserving text, selection, scroll position, loaded data, and other visual state.

```vbnet
NavigationView1.Pages("customers").CacheMode = NavigationPageCacheMode.KeepAlive
```

### Recreate

`Recreate` disposes the page control after successful navigation to another page. Returning to the page creates a new control.

```vbnet
NavigationView1.Pages("reports").CacheMode = NavigationPageCacheMode.Recreate
```

Use `Recreate` for expensive screens whose state does not need to remain in memory. The target page is created successfully before the previous `Recreate` control is disposed, so a creation failure does not destroy the page currently on screen.

## Reloading a page

`ReloadCurrentPage` creates a replacement before disposing the currently displayed instance:

```vbnet
If Not NavigationView1.ReloadCurrentPage() Then
    MessageBox.Show("The page was not reloaded.")
End If
```

The method raises `Navigating` with the same page as both `CurrentPage` and `TargetPage`, allowing the application to cancel the reload. A successful reload raises `PageClosed`, `PageCreated`, and `Navigated`.

## Closing pages and clearing the cache

`ClosePage` disposes a page control without removing its `NavigationPage` definition:

```vbnet
NavigationView1.ClosePage("customers")
```

If the page is currently selected, the content area becomes empty. Call `Navigate` to display another page. Explicit closing raises the cancelable `PageClosing` event followed by `PageClosed` when disposal occurs.

Clear every cached control except the current page:

```vbnet
Dim ClosedCount As Integer = NavigationView1.ClearCache()
```

Include the current page:

```vbnet
Dim ClosedCount As Integer = NavigationView1.ClearCache(True)
```

Removing a page definition from `Pages` automatically disposes its created control. Clearing the collection disposes every created page.

## Accessing created controls

Check whether a page has already been created:

```vbnet
If NavigationView1.IsPageCreated("customers") Then
    Dim Customers As CustomersControl = TryCast(
        NavigationView1.GetCachedControl("customers"),
        CustomersControl)
End If
```

`GetCachedControl` never creates a page. It returns `Nothing` when the key is missing, the page has not been opened, its cache was cleared, or its control was externally disposed.

The current values are also available through `SelectedPage`, `SelectedPageKey`, and `SelectedControl`.

## Handling creation failures

By default, an exception raised by a constructor, factory, or display operation is rethrown after `NavigationFailed` is raised. Set `Handled = True` when the application has reported the failure and wants `Navigate` to return `False` instead:

```vbnet
Private Sub NavigationView1_NavigationFailed(Sender As Object, E As NavigationFailedEventArgs) Handles NavigationView1.NavigationFailed
    MessageBox.Show(E.Exception.Message, $"Unable to open {E.TargetPage.Text}", MessageBoxButtons.OK, MessageBoxIcon.Error)
    E.Handled = True
End Sub
```

## Updating pages at run time

Changes to a page definition rebuild or update its navigation button automatically:

```vbnet
Dim ReportsPage As NavigationPage = NavigationView1.Pages("reports")
ReportsPage.Text = "Updated reports"
ReportsPage.Image = My.Resources.ReportsUpdated20
ReportsPage.ToolTipText = "Open the updated reporting center."
ReportsPage.Visible = True
ReportsPage.Enabled = True
```

`ControlType` and `Factory` cannot be changed while a page owns a created control. Call `ClosePage` first so the existing instance is disposed deliberately.

## API reference

### NavigationPage properties

| Property | Type | Default | Description |
|---|---|---|---|
| `Key` | `String` | Generated when added empty | Unique case-insensitive page identifier. |
| `Text` | `String` | Empty | Navigation button text; the key is used when empty. |
| `Image` | `Image` | `Nothing` | Navigation button image. |
| `ToolTipText` | `String` | Empty | Navigation button tooltip. |
| `AccessibleName` | `String` | Empty | Accessible button name; text or key is used when empty. |
| `ControlType` | `Type` | `Nothing` | Designer-serializable `UserControl` type with a parameterless constructor. |
| `Factory` | `Func(Of UserControl)` | `Nothing` | Run-time control factory; hidden from Designer serialization. |
| `CacheMode` | `NavigationPageCacheMode` | `KeepAlive` | Determines whether the control remains cached. |
| `Visible` | `Boolean` | `True` | Determines whether the navigation button is displayed. |
| `Enabled` | `Boolean` | `True` | Determines whether navigation to the page is allowed. |
| `Tag` | `Object` | `Nothing` | Stores application-defined data. |
| `IsCreated` | `Boolean` | Read-only | Indicates whether a non-disposed control is cached. |
| `CachedControl` | `UserControl` | Read-only | Returns the cached control without creating it. |

### NavigationView behavior and layout properties

| Property | Type | Default | Valid range | Description |
|---|---|---|---|---|
| `Pages` | `NavigationPageCollection` | Empty | — | Ordered page definitions. |
| `AutoNavigateFirstPage` | `Boolean` | `True` | — | Opens the first visible enabled page after loading. |
| `NavigationPosition` | `NavigationPanePosition` | `Left` | — | Places the navigation pane on the left or right. |
| `NavigationWidth` | `Integer` | `220` | 80–600 | Navigation pane width. |
| `ButtonHeight` | `Integer` | `44` | 24–128 | Navigation button height. |
| `ButtonSpacing` | `Integer` | `2` | 0–32 | Vertical space between buttons. |
| `NavigationPadding` | `Padding` | `8` | — | Space around the button list. |
| `ButtonPadding` | `Padding` | `12, 0, 12, 0` | — | Space inside each button. |
| `ContentPadding` | `Padding` | `0` | — | Space around the displayed page. |
| `ImageSize` | `Size` | `20, 20` | 0–64 per dimension | Drawn page-image size. |
| `SelectedIndicatorWidth` | `Integer` | `4` | 0–16 | Selected-page indicator width. |
| `ShowImages` | `Boolean` | `True` | — | Displays page images. |
| `ShowToolTips` | `Boolean` | `True` | — | Displays page tooltips. |

Values outside documented numeric ranges throw `ArgumentOutOfRangeException`.

### Appearance properties

| Property | Default |
|---|---|
| `NavigationBackColor` | `SystemColors.Control` |
| `ContentBackColor` | `SystemColors.Window` |
| `ButtonBackColor` | `SystemColors.Control` |
| `ButtonHoverBackColor` | `SystemColors.ControlLight` |
| `ButtonForeColor` | `SystemColors.ControlText` |
| `SelectedButtonBackColor` | `SystemColors.Highlight` |
| `SelectedButtonForeColor` | `SystemColors.HighlightText` |
| `SelectedIndicatorColor` | `SystemColors.Highlight` |

The control also uses its inherited `Font`, `Enabled`, and `RightToLeft` properties when rendering buttons.

### Methods

| Method | Description |
|---|---|
| `Navigate(key)` | Displays the page identified by a key. |
| `Navigate(page)` | Displays a page object owned by the control. |
| `ReloadCurrentPage()` | Replaces the selected page control with a fresh instance. |
| `ClosePage(key)` | Disposes a created page control without removing its definition. |
| `ClosePage(page)` | Disposes the specified page control without removing its definition. |
| `ClearCache()` | Closes every cached control except the current page and returns the count. |
| `ClearCache(includeCurrentPage)` | Closes cached controls and optionally includes the selected page. |
| `IsPageCreated(key)` | Indicates whether a page owns a created control. |
| `GetCachedControl(key)` | Returns a cached control without creating it. |

### Events

| Event | Event arguments | Description |
|---|---|---|
| `Navigating` | `NavigationCancelEventArgs` | Occurs before navigation and can cancel it. |
| `Navigated` | `NavigationEventArgs` | Occurs after successful navigation. |
| `SelectedPageChanged` | `EventArgs` | Occurs whenever the selected page reference changes. |
| `PageCreated` | `NavigationPageEventArgs` | Occurs after a page control is created. |
| `PageClosing` | `NavigationPageCancelEventArgs` | Occurs before an explicit close and can cancel it. |
| `PageClosed` | `NavigationPageEventArgs` | Occurs after a page control is disposed. |
| `NavigationFailed` | `NavigationFailedEventArgs` | Occurs when page creation or display fails. |

## Page and image ownership

`NavigationView` owns every `UserControl` returned by `ControlType` activation or a page factory. It disposes those controls when required by cache mode, explicit close, page removal, collection clearing, or disposal of the `NavigationView` itself.

The control references page images but does not clone or dispose them. Images stored in application resources remain owned by those resources. Images created dynamically should be disposed by the application only after no page uses them.

## Package information

| Item | Value |
|---|---|
| Package | `CoreSuite.NavigationView` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.NavigationView` |
| Target framework | `.NET 8 for Windows` |
| UI framework | Windows Forms |
| Version | `1.0.0` |
