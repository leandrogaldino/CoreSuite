# CToolStripRender

**A custom Windows Forms system renderer that removes the standard ToolStrip border and controls the checked-button highlight, included in CoreSuite.**

> [!NOTE]
> CToolStripRender is one of the independent projects that make up the **CoreSuite** solution. It can be installed and used separately without requiring the other CoreSuite controls.

## Overview

`CToolStripRender` extends `ToolStripSystemRenderer` with two focused visual changes: the standard border around a `ToolStrip` is always suppressed, and the default background rectangle of checked `ToolStripButton` items can be shown or hidden.

Unchecked buttons and other ToolStrip items continue to use the normal system renderer behavior, allowing the component to make a small visual adjustment without replacing the complete Windows Forms rendering pipeline.

## Key features

- Inherits the standard Windows Forms `ToolStripSystemRenderer`.
- Removes the standard ToolStrip border.
- Optionally preserves the default checked-button background.
- Can hide the checked rectangle for cleaner custom button designs.
- Preserves normal system rendering for unchecked buttons and other items.
- Can be assigned to `ToolStrip`, `MenuStrip`, `StatusStrip`, and compatible ToolStrip-derived controls.
- Requires only one property to configure its checked-state behavior.
- Has no dependency on other CoreSuite packages.

## Requirements

- .NET 8 for Windows
- Windows Forms
- Windows

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.CToolStripRender
```

## Quick start

Import the control namespace:

```vbnet
Imports CoreSuite.Controls
```

Assign the renderer to a `ToolStrip`:

```vbnet
ToolStrip1.Renderer = New CToolStripRender()
```

The standard ToolStrip border is removed automatically. Because `ShowButtonCheckedRectangle` defaults to `False`, checked ToolStrip buttons are displayed without the default checked-state background rectangle.

## Showing the checked-button rectangle

Set `ShowButtonCheckedRectangle` to `True` when checked buttons should retain the normal system background:

```vbnet
ToolStrip1.Renderer = New CToolStripRender With {
    .ShowButtonCheckedRectangle = True
}
```

This is useful when the standard Windows checked-state indication should remain visible while only the ToolStrip border is removed.

## Hiding the checked-button rectangle

Set `ShowButtonCheckedRectangle` to `False` to suppress the default background rectangle of checked `ToolStripButton` items:

```vbnet
ToolStrip1.Renderer = New CToolStripRender With {
    .ShowButtonCheckedRectangle = False
}
```

The checked state itself is not changed. Only the default checked-button background rendering is suppressed, allowing an image, text style, or another application-defined visual cue to represent the state.

## Using checkable buttons

A ToolStrip button can manage its own checked state while using `CToolStripRender`:

```vbnet
ViewButton.CheckOnClick = True
ToolStrip1.Renderer = New CToolStripRender With {
    .ShowButtonCheckedRectangle = False
}
```

`CheckOnClick` and `Checked` remain standard `ToolStripButton` properties. The renderer only determines whether the normal checked background is painted.

## Reusing a renderer

One configured renderer can be shared by multiple ToolStrip controls:

```vbnet
Dim renderer As New CToolStripRender With {
    .ShowButtonCheckedRectangle = True
}

MainToolStrip.Renderer = renderer
NavigationToolStrip.Renderer = renderer
```

Changing `ShowButtonCheckedRectangle` on the shared instance affects subsequent rendering performed by every ToolStrip that uses it.

## Form configuration

The renderer is normally assigned after `InitializeComponent`, ensuring that the Designer-generated ToolStrip configuration already exists:

```vbnet
Public Class MainForm
    Public Sub New()
        InitializeComponent()
        MainToolStrip.Renderer = New CToolStripRender With {
            .ShowButtonCheckedRectangle = False
        }
    End Sub
End Class
```

The assignment can also be made in the form's `Load` event when renderer selection depends on runtime settings.

## Rendering behavior

### ToolStrip border

The standard ToolStrip border is always suppressed. This behavior is built into the renderer and does not require a separate property.

### Checked ToolStrip buttons

| `ShowButtonCheckedRectangle` | Result |
|---|---|
| `False` | The default checked-button background rectangle is not rendered. |
| `True` | The normal system renderer paints the checked-button background. |

### Other items

Unchecked buttons and other ToolStrip items continue through the standard `ToolStripSystemRenderer` behavior.

## API reference

### Properties

| Property | Type | Default | Description |
|---|---|---|---|
| `ShowButtonCheckedRectangle` | `Boolean` | `False` | Gets or sets whether checked ToolStrip buttons use the default system checked-state background. |

## Behavior notes

- The renderer removes the standard border from every ToolStrip to which it is assigned.
- Suppressing the checked rectangle does not change a button's `Checked` value.
- The property affects checked `ToolStripButton` background rendering.
- Unchecked items continue to use normal system rendering.
- Assigning another renderer replaces `CToolStripRender` for that ToolStrip.
- Rendering occurs as part of the normal Windows Forms painting lifecycle.

## Accessibility considerations

- When the checked rectangle is hidden, provide another clear visual indication of the checked state.
- Do not rely exclusively on a subtle color difference to distinguish checked and unchecked buttons.
- Keep meaningful button text, accessible names, or ToolTips for image-only ToolStrip buttons.
- Verify sufficient contrast for any custom checked-state images or colors.

## Integration notes

- Assign the renderer on the Windows Forms UI thread.
- Reuse one instance when several ToolStrip controls should follow the same configuration.
- Use separate instances when individual ToolStrip controls require different checked-state behavior.
- If a theme system replaces ToolStrip renderers at runtime, reapply `CToolStripRender` when the custom behavior must be preserved.

## Package information

| Item | Value |
|---|---|
| Package | `CoreSuite.CToolStripRender` |
| Namespace | `CoreSuite.Controls` |
| Target framework | `.NET 8 for Windows` |
| UI framework | `Windows Forms` |
| Base renderer | `ToolStripSystemRenderer` |
| CoreSuite dependencies | None |

## License

MIT License.
