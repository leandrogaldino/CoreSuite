# ControlContainer

**A reusable Windows Forms component for displaying any control inside a floating drop-down anchored to another control, included in CoreSuite.**

> [!NOTE]
> ControlContainer is one of the independent projects that make up the **CoreSuite** solution. It can be installed and used separately and also serves as drop-down infrastructure for other CoreSuite controls.

## Overview

`ControlContainer` displays a Windows Forms `Control` inside a floating window positioned relative to a host control. It provides the infrastructure required to build calendar selectors, lookup panels, option pickers, custom menus, and other drop-down interfaces without implementing a separate floating form for each use case.

The component handles positioning, screen boundaries, focus, outside clicks, the Escape key, parent movement, and the complete opening and closing lifecycle.

## Key features

- Hosts any Windows Forms `Control`.
- Anchors the drop-down to a button, text box, or another control.
- Supports automatic opening when the host is clicked.
- Supports explicit opening and closing through methods.
- Keeps the floating window inside the current screen's working area.
- Closes on outside clicks or when Escape is pressed.
- Closes when the host control's parent moves.
- Uses the hosted control to determine the drop-down size.
- Supports a configurable border color.
- Exposes opening, opened, closing, closed, and state-change events.
- Provides the current drop-down lifecycle state.
- Supports Visual Studio Designer configuration.

## Requirements

- .NET 8 for Windows
- Windows Forms
- Windows

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.ControlContainer
```

## Quick start

Import the control namespace:

```vbnet
Imports CoreSuite.Controls
```

The following example displays a `MonthCalendar` below a button and copies the selected date to a text box:

```vbnet
Public Class MainForm
    Private ReadOnly dropDownContent As New MonthCalendar()
    Private ReadOnly dateDropDown As New ControlContainer()
    Public Sub New()
        InitializeComponent()
        dateDropDown.HostControl = BtnChooseDate
        dateDropDown.HostedControl = dropDownContent
        dateDropDown.DropDownBorderColor = Color.DodgerBlue
        AddHandler dropDownContent.DateSelected, AddressOf DropDownContent_DateSelected
    End Sub
    Private Sub DropDownContent_DateSelected(sender As Object, e As DateRangeEventArgs)
        TxtDate.Text = e.Start.ToShortDateString()
        dateDropDown.CloseDropDown()
    End Sub
End Class
```

With `DropDownEnabled` set to `True`, clicking `HostControl` opens the configured drop-down automatically.

## Host and content controls

`HostControl` identifies the control used as the drop-down anchor and automatic click trigger:

```vbnet
ControlContainer1.HostControl = BtnShowOptions
```

`HostedControl` identifies the content displayed inside the floating window:

```vbnet
ControlContainer1.HostedControl = OptionsPanel
```

The drop-down follows the size of `HostedControl`. While displayed, the hosted control is temporarily placed inside the internal floating window.

Any Windows Forms `Control` can be hosted, including:

- `MonthCalendar`
- `Panel`
- `UserControl`
- `ListBox`
- `DataGridView`
- A custom CoreSuite control

## Automatic opening

Automatic opening is enabled by default:

```vbnet
ControlContainer1.DropDownEnabled = True
```

When enabled, clicking `HostControl` requests the drop-down to open. Both `HostControl` and `HostedControl` must be assigned before it can be displayed.

Disable automatic opening when application logic must decide when the content is shown:

```vbnet
ControlContainer1.DropDownEnabled = False
```

## Manual control

Use `CanDrop` to verify that the container can currently open, then call `ShowDropDown`:

```vbnet
If ControlContainer1.CanDrop Then ControlContainer1.ShowDropDown()
```

Close the floating window explicitly when the user completes an action:

```vbnet
ControlContainer1.CloseDropDown()
```

`ShowDropDown` requires both the host and hosted controls. `CloseDropDown` closes the container when it is open.

## Drop-down lifecycle

The component exposes lifecycle events in the order in which the drop-down changes state:

1. `Dropping` is raised when opening begins.
2. `Dropped` is raised when opening completes.
3. `Closing` is raised when closing begins.
4. `Closed` is raised when closing completes.

`DropStateChanged` is raised whenever `DropState` changes.

```vbnet
Private Sub ControlContainer1_DropStateChanged(sender As Object, e As EventArgs) Handles ControlContainer1.DropStateChanged
    StatusLabel.Text = ControlContainer1.DropState.ToString()
End Sub
```

The possible `ControlContainerDropDownState` values are:

| Value | Description |
|---|---|
| `Closed` | The drop-down is not displayed. |
| `Closing` | The drop-down is being closed. |
| `Dropping` | The drop-down is being opened. |
| `Dropped` | The drop-down is open and displayed. |

## Positioning and closing behavior

When the container opens, it positions the floating window relative to `HostControl` and adjusts the result to keep it inside the working area of the current screen.

The drop-down closes automatically when:

- The user clicks outside the floating window.
- The user presses Escape.
- The parent of `HostControl` moves.
- Application code calls `CloseDropDown`.

This behavior allows the hosted content to act like a native drop-down while remaining fully customizable.

## Appearance

Use `DropDownBorderColor` to match the floating window to the visual style of the host control or application:

```vbnet
ControlContainer1.DropDownBorderColor = Color.DodgerBlue
```

The default border color is `SystemColors.HotTrack`.

## Designer usage

After installing the package and adding `ControlContainer` to the Visual Studio Toolbox:

1. Add a `ControlContainer` component to the form.
2. Assign `HostControl` to the control that anchors the drop-down.
3. Assign `HostedControl` to the content that should appear in the floating window.
4. Keep `DropDownEnabled` enabled for automatic opening, or disable it for manual control.
5. Configure `DropDownBorderColor` as needed.
6. Create handlers for the lifecycle events when the form must react to opening or closing.

The hosted control can be designed normally on the form or inside another design-time container before being assigned to `HostedControl`.

## API reference

### Properties

| Property | Type | Default | Description |
|---|---|---|---|
| `HostControl` | `Control` | `Nothing` | Gets or sets the anchor control and automatic click trigger. |
| `HostedControl` | `Control` | `Nothing` | Gets or sets the control displayed inside the floating window. |
| `DropDownEnabled` | `Boolean` | `True` | Gets or sets whether clicking the host can open the drop-down automatically. |
| `DropDownBorderColor` | `Color` | `SystemColors.HotTrack` | Gets or sets the floating window border color. |
| `CanDrop` | `Boolean` | Read-only | Indicates whether the drop-down can currently be opened. |
| `DropState` | `ControlContainerDropDownState` | Read-only | Gets the current drop-down lifecycle state. |

### Methods

| Method | Description |
|---|---|
| `ShowDropDown()` | Opens the drop-down and adjusts its position to the current screen. |
| `CloseDropDown()` | Closes the drop-down when it is open. |

### Events

| Event | Description |
|---|---|
| `Dropping` | Occurs when opening begins. |
| `Dropped` | Occurs when opening completes. |
| `Closing` | Occurs when closing begins. |
| `Closed` | Occurs when closing completes. |
| `DropStateChanged` | Occurs whenever the lifecycle state changes. |

## Behavior notes

- The drop-down size follows `HostedControl`.
- The hosted control is moved into the internal floating window while the drop-down is displayed.
- Opening requires valid `HostControl` and `HostedControl` references.
- Screen-boundary adjustment uses the working area of the screen containing the host.
- Outside-click and Escape handling are active while the drop-down is displayed.
- Opening and closing operations should run on the Windows Forms UI thread.

## Integration notes

- Close the drop-down after the user confirms or selects a value when the workflow represents a single selection.
- Use lifecycle events to synchronize icons, button states, or related form content.
- Avoid disposing `HostedControl` while the drop-down is open.
- Dispose the containing form or component normally so the internal floating-window resources can be released.

## Package information

| Item | Value |
|---|---|
| Package | `CoreSuite.ControlContainer` |
| Namespace | `CoreSuite.Controls` |
| Target framework | `.NET 8 for Windows` |
| UI framework | `Windows Forms` |

## License

MIT License.
