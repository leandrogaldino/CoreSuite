# ColorPicker

**An embedded color selection control that exposes the standard Windows Web and System color palettes directly inside a Windows Forms interface, included in CoreSuite.**

> [!NOTE]
> ColorPicker is one of the independent projects that make up the **CoreSuite** solution. It can be installed and used separately without requiring the other CoreSuite controls.

## Overview

`ColorPicker` hosts the standard Windows color editor directly inside a form. Applications can therefore provide the familiar Web and System color tabs without opening a separate modal color dialog.

The control exposes the selected color, raises an event when the selection changes, supports keyboard navigation outside the embedded editor, and can paint color previews in other parts of the interface.

## Key features

- Embedded standard Windows color editor.
- Web and System color tabs.
- Direct access to the currently selected color.
- `ColorChanged` event for live interface updates.
- Optional Tab and Shift+Tab navigation outside the editor.
- Automatic minimum-size management.
- Methods for opening and closing the embedded editor.
- Built-in color preview painting.
- Visual Studio Designer support.

## Requirements

- .NET 8 for Windows
- Windows Forms
- Windows

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.ColorPicker
```

## Quick start

Import the control namespace:

```vbnet
Imports CoreSuite.Controls
```

Add a `ColorPicker` to a form using the Visual Studio Toolbox or create it in code:

```vbnet
Imports System.Drawing
Imports CoreSuite.Controls

Dim picker As New ColorPicker With {
    .Color = Color.CornflowerBlue,
    .AllowTabOut = True,
    .Location = New Point(20, 20)
}

AddHandler picker.ColorChanged, AddressOf Picker_ColorChanged
Controls.Add(picker)
```

Respond to changes in the selected color:

```vbnet
Private Sub Picker_ColorChanged(sender As Object, e As EventArgs)
    Dim picker As ColorPicker = DirectCast(sender, ColorPicker)
    PreviewPanel.BackColor = picker.Color
End Sub
```

## Reading and changing the color

Use `Color` to read the current selection:

```vbnet
Dim selectedColor As Color = ColorPicker1.Color
```

Assign a `Color` value to update the selection programmatically:

```vbnet
ColorPicker1.Color = Color.MediumSeaGreen
```

The `ColorChanged` event is raised when the selected color changes, allowing previews and related controls to remain synchronized.

## Editor lifecycle

The embedded editor is normally displayed as part of the control. It can also be managed explicitly when required by the form lifecycle.

Display or recreate the editor:

```vbnet
ColorPicker1.ShowEditor()
```

Close the editor and release the resources associated with it:

```vbnet
ColorPicker1.CloseEditor()
```

`ShowEditor` can be called again after `CloseEditor` when the embedded editor must be restored.

## Keyboard navigation

Set `AllowTabOut` to `True` when keyboard users should be able to leave the embedded editor with Tab or Shift+Tab:

```vbnet
ColorPicker1.AllowTabOut = True
```

With this option enabled:

- Tab moves focus to the next control.
- Shift+Tab moves focus to the previous control.

When it is disabled, Tab remains available to the embedded editor according to its standard behavior.

## Painting a color preview

Use `PaintValue` when another control or custom drawing surface needs to display a preview of a color:

```vbnet
Private Sub PreviewPanel_Paint(sender As Object, e As PaintEventArgs) Handles PreviewPanel.Paint
    Dim previewBounds As New Rectangle(4, 4, PreviewPanel.ClientSize.Width - 8, PreviewPanel.ClientSize.Height - 8)
    ColorPicker1.PaintValue(ColorPicker1.Color, e.Graphics, previewBounds)
End Sub
```

Invalidate the preview whenever the selected color changes:

```vbnet
Private Sub ColorPicker1_ColorChanged(sender As Object, e As EventArgs) Handles ColorPicker1.ColorChanged
    PreviewPanel.Invalidate()
End Sub
```

## Designer usage

After installing the package and adding `ColorPicker` to the Visual Studio Toolbox:

1. Drag the control onto a Windows Form.
2. Set `Color` to define the initial selection.
3. Enable `AllowTabOut` when the editor should participate in the form's normal Tab order.
4. Create a handler for `ColorChanged` when another part of the interface must react to the selection.
5. Adjust inherited layout properties such as `Location`, `Anchor`, and `Dock` as needed.

The control manages the minimum size required by the embedded Windows editor automatically.

## API reference

### Properties

| Property | Type | Description |
|---|---|---|
| `Color` | `Color` | Gets or sets the currently selected color. |
| `AllowTabOut` | `Boolean` | Gets or sets whether Tab and Shift+Tab can move focus outside the embedded editor. |

### Methods

| Method | Description |
|---|---|
| `ShowEditor()` | Creates and displays the embedded Windows color editor. |
| `CloseEditor()` | Closes the embedded editor and releases its resources. |
| `PaintValue(color, graphics, bounds)` | Paints a preview of the specified color inside a rectangle. |

### Events

| Event | Event arguments | Description |
|---|---|---|
| `ColorChanged` | `EventArgs` | Occurs when the selected color changes. |

## Behavior notes

- The control embeds the standard Windows color editor instead of opening a separate dialog.
- The Web and System color tabs are available directly inside the form.
- The minimum control size is managed to accommodate the embedded editor.
- `AllowTabOut` affects keyboard focus navigation but does not change the selected color.
- The selected value is represented by `System.Drawing.Color`.
- The editor can be closed and displayed again during the control lifecycle.

## Accessibility considerations

- Enable `AllowTabOut` so keyboard users are not trapped inside the embedded editor.
- Keep the control in a logical `TabIndex` order with the surrounding form controls.
- Do not use color as the only way to communicate status or meaning.
- Provide a text label or other contextual description when the selected color represents an important application setting.

## Package information

| Item | Value |
|---|---|
| Package | `CoreSuite.ColorPicker` |
| Namespace | `CoreSuite.Controls` |
| Target framework | `.NET 8 for Windows` |
| UI framework | `Windows Forms` |

## License

MIT License.
