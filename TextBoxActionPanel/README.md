# TextBoxActionPanel

**A non-visual Windows Forms component that adds a configurable image-action panel to an existing `TextBoxBase` control, included in CoreSuite.**

> [!NOTE]
> TextBoxActionPanel is one of the independent projects that make up the **CoreSuite** solution. It enhances an existing text box rather than replacing it, and its required `CoreSuite.NoFocusCueButton` dependency is installed automatically with the NuGet package.

## Overview

`TextBoxActionPanel` attaches a compact row of image buttons to a `TextBoxBase` control already placed on a form. The panel appears above or below the target while the field has focus and can expose actions such as viewing, searching, creating, clearing, or opening related content.

The component appears in the Visual Studio component tray, like `ToolTip`, `ErrorProvider`, and `ContextMenuStrip`. Its floating window does not activate, allowing the target text box to retain its keyboard focus, caret, and selection while an action button is clicked.

## Key features

- Attaches to any existing control derived from `TextBoxBase`.
- Works with standard `TextBox` and `RichTextBox` controls.
- Works with CoreSuite controls derived from `TextBoxBase`, including `QueriedBox`.
- Appears in the Windows Forms component tray instead of occupying the form surface.
- Provides a designer-serializable `Actions` collection.
- Supports action lookup by numeric index or case-insensitive string key.
- Provides `Contains`, `IndexOf`, and `Remove` operations for action objects.
- Displays the first action at the right edge and adds subsequent actions toward the left.
- Uses `NoFocusCueButton` internally for action buttons.
- Shows and hides automatically as the target receives and loses focus.
- Restores the panel when focus returns after a modal window closes.
- Automatically follows target movement, resizing, parent movement, scrolling, and form movement.
- Supports automatic, above-only, and below-only placement.
- Supports a general `ActionClicked` event and an optional delegate for each action.
- Supports run-time changes to images, tooltips, visibility, enabled state, ordering, sizing, spacing, and colors.

## Requirements

- .NET 8 for Windows
- Windows Forms
- Windows

The package depends on `CoreSuite.NoFocusCueButton`. NuGet resolves this dependency automatically.

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.TextBoxActionPanel
```

## Quick start

Import the control namespace:

```vbnet
Imports CoreSuite.Controls
```

Configure the target and add actions:

```vbnet
CustomerActions.TargetControl = CustomerTextBox
CustomerActions.Actions.Add("View", My.Resources.View16, "View the selected customer.")
CustomerActions.Actions.Add("Search", My.Resources.Search16, "Search for a customer.")
CustomerActions.Actions.Add("Create", My.Resources.Create16, "Create a customer.")
```

Handle every action through the general event:

```vbnet
Private Sub CustomerActions_ActionClicked(Sender As Object, E As TextBoxActionClickEventArgs) Handles CustomerActions.ActionClicked
    Select Case E.Action.Key
        Case "View"
            ViewCustomer(E.TargetControl.Text)
        Case "Search"
            SearchCustomer(E.TargetControl.Text)
        Case "Create"
            CreateCustomer(E.TargetControl.Text)
    End Select
End Sub
```

`E.TargetControl` is the existing control assigned to `TargetControl`.

## Designer usage

After installing the package and adding `TextBoxActionPanel` to the Visual Studio Toolbox:

1. Place the desired `TextBox`, `RichTextBox`, `QueriedBox`, or another `TextBoxBase`-derived control on the form.
2. Add `TextBoxActionPanel` from the Toolbox.
3. Select the component in the component tray.
4. Set `TargetControl` to the existing text box.
5. Open the `Actions` collection editor.
6. Add, remove, or reorder `TextBoxAction` items.
7. Configure `Key`, `Image`, `ToolTipText`, `AccessibleName`, `Visible`, and `Enabled` for each action.
8. Handle the component's `ActionClicked` event when a shared event handler is required.

The component intentionally does not draw buttons over the text box in the Visual Studio Designer. The floating panel is created only at run time.

One component per target control is recommended because each component owns its own target, actions, appearance, behavior, and events.

## Creating actions in code

Use `Actions.Add` to create and append an action:

```vbnet
Dim ViewAction As TextBoxAction = CustomerActions.Actions.Add("View", My.Resources.View16, "View the selected customer.")
Dim SearchAction As TextBoxAction = CustomerActions.Actions.Add("Search", My.Resources.Search16, "Search for a customer.")
Dim CreateAction As TextBoxAction = CustomerActions.Actions.Add("Create", My.Resources.Create16, "Create a customer.")
```

A 16-by-16 pixel image is recommended. The default button is 24 by 24 pixels and centers the image without modifying it.

## Accessing actions

### By numeric index

Use the inherited numeric indexer when the collection position is known:

```vbnet
CustomerActions.Actions(0).Image = My.Resources.View16
```

### By key

Use the string indexer to access an action directly by `Key`:

```vbnet
CustomerActions.Actions("Search").Image = My.Resources.Search16
CustomerActions.Actions("Search").ToolTipText = "Search for another customer."
```

Key comparison is case-insensitive, so `Actions("search")` and `Actions("Search")` resolve the same first matching action.

The string indexer has the following validation behavior:

| Condition | Result |
|---|---|
| Key is `Nothing`, empty, or whitespace | Throws `ArgumentException`. |
| No action has the specified key | Throws `KeyNotFoundException`. |
| More than one action has the key | Returns the first case-insensitive match. |

Use `FindByKey` when a missing action should return `Nothing` instead of throwing an exception:

```vbnet
Dim SearchAction As TextBoxAction = CustomerActions.Actions.FindByKey("Search")
If SearchAction IsNot Nothing Then SearchAction.Enabled = True
```

Keys are not required to be unique, but unique keys are strongly recommended.

### By action object

When the `TextBoxAction` instance is already available, edit it directly:

```vbnet
Dim SearchAction As TextBoxAction = CustomerActions.Actions.FindByKey("Search")
If SearchAction IsNot Nothing Then SearchAction.Image = My.Resources.Search16
```

Use the standard collection operations to work with the object itself:

```vbnet
If CustomerActions.Actions.Contains(SearchAction) Then
    Dim ActionIndex As Integer = CustomerActions.Actions.IndexOf(SearchAction)
End If

CustomerActions.Actions.Remove(SearchAction)
```

An `Actions(action)` overload is unnecessary because the existing object can already be modified directly.

## Per-action delegate

Assign `ClickHandler` at run time when an action should own its handler. Function references are not serialized by the Windows Forms Designer.

```vbnet
Private Sub MainForm_Load(Sender As Object, E As EventArgs) Handles MyBase.Load
    Dim CreateAction As TextBoxAction = CustomerActions.Actions.FindByKey("Create")
    If CreateAction IsNot Nothing Then CreateAction.ClickHandler = AddressOf CreateCustomerAction
End Sub
Private Sub CreateCustomerAction(E As TextBoxActionClickEventArgs)
    MessageBox.Show($"Create a customer from: {E.TargetControl.Text}")
End Sub
```

When an action executes, `ActionClicked` is raised first and its `ClickHandler` is invoked afterward. Applications normally use one mechanism for each action to avoid executing the same command twice.

## Updating actions at run time

Changes to an action are reflected by its corresponding button while the application is running:

```vbnet
Dim SearchAction As TextBoxAction = CustomerActions.Actions("Search")
SearchAction.Image = My.Resources.SearchActive16
SearchAction.ToolTipText = "Search using the current text."
SearchAction.Visible = True
SearchAction.Enabled = True
```

`ToolTipText` is applied to the internal button and is updated when the action changes. Image, tooltip, visibility, enabled state, and appearance changes request the required panel refresh automatically.

## Action order

Collection order is interpreted from right to left:

| Collection index | Visual position |
|---:|---|
| `0` | Rightmost button. |
| `1` | Immediately left of index `0`. |
| `2` | Immediately left of index `1`. |

Invisible actions do not reserve space. Reordering the collection changes the visual order when the panel is rebuilt.

## Performing actions programmatically

`PerformAction` executes the first enabled action matching a key and returns whether execution occurred:

```vbnet
If Not CustomerActions.PerformAction("Search") Then MessageBox.Show("The Search action is unavailable.")
```

Key comparison is case-insensitive. A hidden but enabled action can still execute programmatically: `Visible` controls presentation, while `Enabled` controls execution.

## Focus behavior

The panel is hosted in a borderless, non-activating window. Its buttons are removed from keyboard selection and have `TabStop` set to `False`.

Consequently:

- Typing continues in the target while the panel is visible.
- Clicking an action button does not focus that button.
- The target's caret and selection remain controlled by the target.
- Normal focus changes hide the panel when `HideOnLeave` is enabled.
- Closing a modal dialog restores the panel when keyboard focus returns to the target.
- Deactivating or closing the owner form hides the panel.

An action can intentionally move focus or open another form. That application behavior is not overridden by the component.

## Placement behavior

`Placement = Auto` prefers the area above the target. When the panel would extend beyond the current screen's working area, it uses the area below the target.

Forced `Above` or `Below` placement is constrained to the working area so the popup remains reachable. Its right edge is aligned with the right edge of the target whenever screen space permits.

## Compatibility with CoreSuite controls

Any CoreSuite control derived from `TextBoxBase` can be assigned directly. For example, `QueriedBox` derives from `TextBox`:

```vbnet
ProductActions.TargetControl = ProductQueriedBox
```

The action panel does not read or modify query configuration, frozen values, primary keys, text, or selection state. Cast `E.TargetControl` when derived members are required:

```vbnet
Dim ProductBox As QueriedBox = DirectCast(E.TargetControl, QueriedBox)
```

## API reference

### TextBoxAction properties

| Property | Type | Default | Description |
|---|---|---|---|
| `Key` | `String` | Generated when empty | Identifies the action in events, lookup, and `PerformAction`. |
| `Image` | `Image` | `Nothing` | Gets or sets the button image. A 16-by-16 image is recommended. |
| `ToolTipText` | `String` | Empty | Gets or sets the text displayed when the pointer rests over the button. |
| `AccessibleName` | `String` | Empty | Gets or sets the name announced by accessibility clients. |
| `Visible` | `Boolean` | `True` | Gets or sets whether the button is included in the panel. |
| `Enabled` | `Boolean` | `True` | Gets or sets whether the action can be executed. |
| `ClickHandler` | Delegate | `Nothing` | Gets or sets the optional run-time handler invoked after `ActionClicked`. |

`ClickHandler` is hidden from the Property Grid and excluded from Designer serialization.

### Component properties

#### Association and behavior

| Property | Type | Default | Description |
|---|---|---|---|
| `TargetControl` | `TextBoxBase` | `Nothing` | Gets or sets the existing control enhanced by the component. |
| `Actions` | `TextBoxActionCollection` | Empty | Gets the ordered collection of action definitions. |
| `Enabled` | `Boolean` | `True` | Gets or sets whether the entire component is enabled. |
| `ShowOnFocus` | `Boolean` | `True` | Gets or sets whether the panel appears when the target receives focus. |
| `HideOnLeave` | `Boolean` | `True` | Gets or sets whether the panel hides when the target loses focus. |
| `Placement` | `TextBoxActionPanelPlacement` | `Auto` | Gets or sets automatic, above-only, or below-only placement. |

#### Layout

| Property | Type | Default | Valid range | Description |
|---|---|---|---|---|
| `ButtonSize` | `Integer` | `24` | `16` to `64` | Gets or sets the width and height of each square button. |
| `ButtonSpacing` | `Integer` | `0` | `0` to `16` | Gets or sets the space between adjacent buttons. |
| `PanelPadding` | `Integer` | `0` | `0` to `16` | Gets or sets the internal space around the buttons. |
| `PanelOffset` | `Integer` | `0` | `0` to `32` | Gets or sets the distance between the panel and target. |

Values outside these ranges throw `ArgumentOutOfRangeException`.

#### Appearance

| Property | Type | Default | Description |
|---|---|---|---|
| `TransparentBackground` | `Boolean` | `True` | Gets or sets whether the area surrounding the buttons is transparent. |
| `ShowBorder` | `Boolean` | `False` | Gets or sets whether a one-pixel panel border is drawn. |
| `PanelBackColor` | `Color` | `SystemColors.Window` | Gets or sets the panel background used when transparency is disabled. |
| `BorderColor` | `Color` | `SystemColors.ControlDark` | Gets or sets the border color used when `ShowBorder` is enabled. |
| `ButtonBackColor` | `Color` | `SystemColors.Window` | Gets or sets the normal button background. |
| `ButtonHoverBackColor` | `Color` | `SystemColors.ControlLight` | Gets or sets the button background while the pointer is over it. |
| `ButtonPressedBackColor` | `Color` | `SystemColors.ControlDark` | Gets or sets the button background while it is pressed. |

### Collection members

| Member | Description |
|---|---|
| `Item(index)` | Gets or sets the action at a numeric collection index. |
| `Item(key)` | Gets the first case-insensitive key match or throws when the key is invalid or missing. |
| `Add(key, image, toolTipText)` | Creates, appends, and returns a new action. |
| `FindByKey(key)` | Returns the first case-insensitive match, or `Nothing` when it is not found. |
| `Contains(action)` | Indicates whether the collection contains the specified action object. |
| `IndexOf(action)` | Returns the numeric index of the specified action object. |
| `Remove(action)` | Removes the specified action object from the collection. |

### Component methods

| Method | Description |
|---|---|
| `ShowPanel()` | Displays the panel when the target, owner form, component state, and visible actions permit it. |
| `HidePanel()` | Hides the panel without clearing the target or actions. |
| `RefreshPanel()` | Rebuilds visible buttons and recalculates the position while the panel is open. |
| `PerformAction(key)` | Executes the first enabled action matching the key and returns whether execution occurred. |

### Events

| Event | Event arguments | Description |
|---|---|---|
| `ActionClicked` | `TextBoxActionClickEventArgs` | Occurs when an enabled action executes. |
| `TargetControlChanged` | `EventArgs` | Occurs after the target changes or is disposed. |
| `PanelShown` | `EventArgs` | Occurs after the floating panel becomes visible. |
| `PanelHidden` | `EventArgs` | Occurs after the floating panel is hidden. |

## Image lifetime

The component references action images but does not clone or dispose them. Images stored in form resources remain owned by those resources. Images created dynamically should be disposed by the code that created them after the panel no longer uses them.

## Accessibility considerations

Every button receives an accessible name using this priority:

1. `AccessibleName`
2. `ToolTipText`
3. `Key`

Because the floating buttons intentionally do not participate in keyboard focus traversal, provide an equivalent keyboard command when an action is essential. `PerformAction` can be called from a shortcut handler.

## Package information

| Item | Value |
|---|---|
| Package | `CoreSuite.TextBoxActionPanel` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.TextBoxActionPanel` |
| Target framework | `.NET 8 for Windows` |
| UI framework | `Windows Forms` |
| CoreSuite dependency | `CoreSuite.NoFocusCueButton` |

## License

MIT License.
