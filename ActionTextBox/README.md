# ActionTextBox

**Standard Windows Forms text input with an integrated configurable action button.**

> [!NOTE]
> ActionTextBox is one of the controls or components included in the **CoreSuite** solution. The package has no CoreSuite runtime dependencies.

## Overview

`ActionTextBox` extends the standard Windows Forms `TextBox` and adds a real child action button that can be displayed on either side of the input.

The button follows the same structural approach used by CoreSuite controls such as `TimeBox`: a dedicated child control is positioned inside the text box while native edit margins reserve the corresponding text area.

It is designed for .NET 8 Windows Forms applications and can be configured in code or through the Visual Studio designer where designer support applies.

## Features

* Real integrated action button hosted inside the text box.
* Button position on the left or right side.
* Configurable image, width and image padding.
* Independent visibility and enabled state.
* Tooltip support.
* Hover, pressed and disabled visual states.
* Programmatic action invocation.
* DPI-aware button dimensions.
* Native text margins prevent text from being drawn below the button.
* Uses the standard Windows Forms event and component model.
* Includes English XML documentation for the public API.
* Requires no third-party runtime packages.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.ActionTextBox
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.ActionTextBox
```

## Namespace

```vb
Imports CoreSuite.Controls
```

## Quick start

```vb
Imports CoreSuite.Controls

Dim SearchBox As New ActionTextBox With {
    .ActionButtonImage = My.Resources.Search,
    .ActionButtonPosition = ActionButtonPosition.Right,
    .ActionButtonToolTipText = "Search"
}
AddHandler SearchBox.ActionButtonClick, Sub() PerformSearch(SearchBox.Text)
```

## Designer usage

After installing or referencing the package, add `ActionTextBox` from the Toolbox or create it in code. Properties configured in the Windows Forms Designer are serialized into `InitializeComponent` and remain available at runtime.

## API reference

### `ActionTextBox`

Represents a standard Windows Forms text box with an integrated configurable action button.

```vb
Public Class ActionTextBox
    Inherits TextBox
```

### Main members

| Member | Behavior |
| --- | --- |
| `ActionButtonImage` | image displayed by the action button. |
| `ActionButtonPosition` | side on which the action button is displayed. |
| `ActionButtonVisible` | shows or hides the action button. |
| `ActionButtonEnabled` | enables or disables action button interaction. |
| `ActionButtonWidth` | button width in logical pixels. |
| `ActionButtonPadding` | image padding in logical pixels. |
| `ActionButtonToolTipText` | tooltip displayed for the action button. |
| `ActionButtonClick` | raised when the action button is clicked. |
| `PerformActionButtonClick()` | programmatically invokes the action button when available. |

### `ActionButtonPosition`

Defines the side on which the action button is displayed.

| Value | Behavior |
| --- | --- |
| `Left` | displays the button on the left side. |
| `Right` | displays the button on the right side. |

## Button implementation

The action button is a dedicated child `PictureBox`, not a region painted directly over the native text box window. The control repositions that child whenever the text box size, border style or DPI changes and uses the native `EM_SETMARGINS` edit message to keep text out of the button area.

The image is painted only inside the child button, preserving its aspect ratio and respecting `ActionButtonPadding`.

## Standard functionality

Because `ActionTextBox` inherits from `TextBox`, standard members such as `Text`, `ReadOnly`, `Multiline`, `PasswordChar`, `CharacterCasing`, `MaxLength`, `TextAlign`, `TextChanged` and data binding remain available.

## Resource and lifetime considerations

* Images assigned to `ActionButtonImage` remain owned by the caller and are not disposed by the control.
* The internal button and tooltip are disposed with the control.
* Events use the standard Windows Forms event model and should be handled on the UI thread.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.ActionTextBox` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.ActionTextBox` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| CoreSuite dependencies | None |
| External runtime dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
