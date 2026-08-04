# ControlSlider

**A non-visual Windows Forms component that makes a control draggable while keeping it inside the client area of a parent container, included in CoreSuite.**

> [!NOTE]
> ControlSlider is one of the independent projects that make up the **CoreSuite** solution. It can be installed and used separately without requiring the other CoreSuite controls.

## Overview

`ControlSlider` adds mouse-based dragging behavior to an existing Windows Forms control. The component observes the configured child control, calculates its new position while the left mouse button is pressed, and constrains movement to the client area of a configured parent control.

Because it is non-visual, `ControlSlider` does not draw its own interface. It can be configured through the Visual Studio component tray or directly in VB.NET code.

## Key features

- Adds drag behavior to an existing Windows Forms control.
- Uses the left mouse button for interactive movement.
- Constrains the child to the parent's client area.
- Supports panels, labels, picture boxes, user controls, and other controls.
- Automatically manages the required mouse event handlers.
- Updates event subscriptions when `Parent` or `Child` changes.
- Supports configuration through the Visual Studio Designer.
- Has no dependency on other CoreSuite packages.

## Requirements

- .NET 8 for Windows
- Windows Forms
- Windows

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.ControlSlider
```

## Quick start

Import the control namespace:

```vbnet
Imports CoreSuite.Controls
```

Create a `ControlSlider`, assign the container to `Parent`, and assign the draggable control to `Child`:

```vbnet
Dim slider As New ControlSlider With {
    .Parent = CanvasPanel,
    .Child = MovablePanel
}
```

The user can now drag `MovablePanel` with the left mouse button. Its location is automatically constrained so the entire control remains inside `CanvasPanel`.

## Designer usage

After installing the package and adding `ControlSlider` to the Visual Studio Toolbox:

1. Drag `ControlSlider` onto the form.
2. Select the component in the component tray.
3. Assign `Parent` to the container that defines the movement area.
4. Assign `Child` to the control that should be draggable.
5. Run the application and drag the child with the left mouse button.

The control assigned to `Child` should be contained by the control assigned to `Parent`.

## Parent and child controls

`Parent` defines the coordinate system and movement boundaries:

```vbnet
ControlSlider1.Parent = CanvasPanel
```

`Child` identifies the control that receives the dragging behavior:

```vbnet
ControlSlider1.Child = MovablePanel
```

A typical relationship is:

```vbnet
CanvasPanel.Controls.Add(MovablePanel)
ControlSlider1.Parent = CanvasPanel
ControlSlider1.Child = MovablePanel
```

This ensures that the child's `Location` uses the same client coordinates that define the movement limits.

## Movement behavior

Dragging begins when the user presses the left mouse button over `Child`. As the pointer moves, `ControlSlider` updates the child's location relative to the configured parent.

Movement is constrained on every edge:

- The left edge cannot move before the parent's left client boundary.
- The top edge cannot move before the parent's top client boundary.
- The right edge cannot move beyond the parent's client width.
- The bottom edge cannot move beyond the parent's client height.

Dragging ends when the left mouse button is released.

## Example: draggable tool panel

The component can be used to create a movable panel inside a larger editing surface:

```vbnet
Public Class EditorForm
    Private ReadOnly toolPanelSlider As New ControlSlider()
    Public Sub New()
        InitializeComponent()
        toolPanelSlider.Parent = WorkspacePanel
        toolPanelSlider.Child = ToolPanel
    End Sub
End Class
```

The `ToolPanel` can contain buttons, labels, or other controls. Dragging any area that belongs to the configured child and receives its mouse events moves the panel while preserving the workspace boundaries.

## Changing controls at runtime

The component can be reassigned when a different control must become draggable:

```vbnet
ControlSlider1.Child = SecondaryPanel
```

The required mouse handlers are updated automatically when the property changes. The previous child is no longer managed by that `ControlSlider` instance.

The parent can also be changed:

```vbnet
ControlSlider1.Parent = SecondaryCanvas
```

After reassignment, the new parent's client area defines the movement boundaries.

## API reference

### Properties

| Property | Type | Default | Description |
|---|---|---|---|
| `Parent` | `Control` | `Nothing` | Gets or sets the container whose client area defines the movement boundaries. |
| `Child` | `Control` | `Nothing` | Gets or sets the control that can be dragged interactively. |

## Behavior notes

- `Parent` should contain the control assigned to `Child`.
- Movement uses the parent control's client coordinates and client size.
- The entire child is kept inside the configured movement area.
- Dragging uses the left mouse button.
- The component manages its mouse event subscriptions when `Child` changes.
- Assigning a different parent changes the active movement boundaries.
- The component is non-visual and appears in the Visual Studio component tray.

## Integration notes

- Configure both `Parent` and `Child` before the user interacts with the form.
- Ensure the child is smaller than the available parent client area.
- Use `Anchor` and `Dock` carefully on the child because layout processing may override locations assigned during dragging.
- Avoid simultaneously changing the child's `Location` from another workflow while it is being dragged.
- Perform runtime property changes on the Windows Forms UI thread.

## Package information

| Item | Value |
|---|---|
| Package | `CoreSuite.ControlSlider` |
| Namespace | `CoreSuite.Controls` |
| Target framework | `.NET 8 for Windows` |
| UI framework | `Windows Forms` |
| CoreSuite dependencies | None |

## License

MIT License.
