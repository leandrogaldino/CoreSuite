# NoFocusCueButton

**A lightweight Windows Forms button that suppresses the default focus rectangle and provides built-in tooltip support.**

> [!NOTE]
> NoFocusCueButton is one of the controls included in the **CoreSuite** solution. It preserves the standard WinForms `Button` behavior while removing the visual focus cue rendered around the button text.

## Overview

`NoFocusCueButton` extends the standard Windows Forms `Button` control and always suppresses its default focus rectangle.

The control remains focusable and continues to support normal keyboard, mouse, command, image, text, layout, and designer behavior. Only the standard focus cue is hidden.

It also provides a `TooltipText` property, allowing a tooltip to be configured directly on the button without creating and managing a separate `ToolTip` component.

## Features

* Inherits all standard WinForms `Button` functionality.
* Suppresses the default dotted focus rectangle.
* Preserves keyboard and mouse interaction.
* Provides built-in tooltip support through `TooltipText`.
* Supports Visual Studio designer configuration.
* Requires no external dependencies.
* Designed for .NET 8 Windows Forms applications.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.NoFocusCueButton
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.NoFocusCueButton
```

## Namespace

```vb
Imports CoreSuite.Controls
```

## Quick start

Add the control to a form and configure it like a standard WinForms button:

```vb
Imports CoreSuite.Controls
Public Class MainForm
    Private Sub MainForm_Load(Sender As Object, E As EventArgs) Handles MyBase.Load
        Dim SaveButton As New NoFocusCueButton With {
            .Text = "Save",
            .TooltipText = "Save the current changes.",
            .AutoSize = True,
            .Location = New Point(20, 20)
        }
        AddHandler SaveButton.Click, AddressOf SaveButton_Click
        Controls.Add(SaveButton)
    End Sub
    Private Sub SaveButton_Click(Sender As Object, E As EventArgs)
        MessageBox.Show("Changes saved.")
    End Sub
End Class
```

The button behaves normally when clicked or focused, but the default dotted focus rectangle is not displayed.

## Designer usage

After installing or referencing the package:

1. Open a Windows Forms form in the Visual Studio designer.
2. Add `NoFocusCueButton` from the Toolbox.
3. Configure the standard `Button` properties normally.
4. Set `TooltipText` in the **NoFocusCueButton** property category.

The tooltip is displayed automatically when the mouse pointer remains over the button.

## API reference

### `NoFocusCueButton`

Represents a standard WinForms button that suppresses the default focus cue and supports an associated tooltip.

```vb
Public Class NoFocusCueButton
    Inherits Button
```

### `TooltipText`

Gets or sets the text displayed in the tooltip when the mouse pointer hovers over the button.

```vb
Public Overridable Property TooltipText As String
```

| Value            | Behavior                                                                              |
| ---------------- | ------------------------------------------------------------------------------------- |
| Empty string     | No tooltip text is displayed.                                                         |
| Non-empty string | The configured text is displayed when the pointer hovers over the button.             |
| `Nothing`        | Treated as no configured tooltip text by the underlying WinForms `ToolTip` component. |

Example:

```vb
SaveButton.TooltipText = "Save the current document."
```

### Focus cue behavior

The control overrides the protected `ShowFocusCues` property and always returns `False`:

```vb
Protected Overrides ReadOnly Property ShowFocusCues As Boolean
```

This affects only the standard visual focus rectangle. The button can still:

* receive focus;
* be reached through keyboard navigation;
* respond to `Enter` and `Space` according to normal WinForms behavior;
* raise standard focus, keyboard, mouse, and click events.

## Standard button functionality

Because `NoFocusCueButton` inherits directly from `Button`, it supports standard properties and events such as:

* `Text`
* `Image`
* `ImageAlign`
* `TextAlign`
* `TextImageRelation`
* `FlatStyle`
* `DialogResult`
* `AutoSize`
* `Enabled`
* `Click`
* `MouseEnter`
* `MouseLeave`
* `GotFocus`
* `LostFocus`

No special setup is required for these members.

## Using the control as a base class

`NoFocusCueButton` can also be used as the base class for custom CoreSuite controls that require standard button behavior without the native focus rectangle.

```vb
Imports CoreSuite.Controls
Public Class ActionButton
    Inherits NoFocusCueButton
    Public Sub New()
        Text = "Action"
        TooltipText = "Execute the action."
    End Sub
End Class
```

Derived controls inherit the suppressed focus cue and tooltip functionality automatically.

## Accessibility considerations

The focus rectangle is a visual indication of keyboard focus. Since this control intentionally suppresses that indication, applications should provide another clear focus state when keyboard navigation is important, such as a border, background change, or other visible style.

Suppressing the focus cue does not remove keyboard support, but it can make the focused control less obvious to users when no alternative visual indication is provided.

## Package information

| Item                  | Value                        |
| --------------------- | ---------------------------- |
| Package               | `CoreSuite.NoFocusCueButton` |
| Namespace             | `CoreSuite.Controls`         |
| Assembly              | `CoreSuite.NoFocusCueButton` |
| Target framework      | `net8.0-windows`             |
| UI framework          | Windows Forms                |
| External dependencies | None                         |

## License

This package is distributed under the license defined by the CoreSuite repository.
