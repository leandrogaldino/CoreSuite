# RichTextButton

**Button that renders multiple text segments with independent fonts and colors.**

> [!NOTE]
> RichTextButton is one of the controls or components included in the **CoreSuite** solution. The package depends on `CoreSuite.NoFocusCueButton`. NuGet installs the required package or packages automatically.

## Overview

`RichTextButton` extends `NoFocusCueButton` and provides button that renders multiple text segments with independent fonts and colors.

It is designed for .NET 8 Windows Forms applications and can be configured in code or through the Visual Studio designer where designer support applies.

## Features

* Bindable collection of RichTextPart objects.
* Segment content.
* Segment font.
* Segment color.
* Uses the standard Windows Forms event and component model.
* Includes English XML documentation for the public API.
* Requires no third-party packages.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.RichTextButton
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.RichTextButton
```

## Namespace

```vb
Imports CoreSuite.Controls
```

## Quick start

```vb
Imports CoreSuite.Controls

Dim Button As New RichTextButton()
Button.TextParts.Add(New RichTextPart With {
    .Text = "Delete ",
    .Color = Color.Black
})
Button.TextParts.Add(New RichTextPart With {
    .Text = "item",
    .Font = New Font(Button.Font, FontStyle.Bold),
    .Color = Color.Firebrick
})
```

## Designer usage

After installing or referencing the package, add `RichTextButton` from the Toolbox or create it in code. Properties configured in the Windows Forms Designer are serialized into `InitializeComponent` and remain available at runtime.

## API reference

### `RichTextButton`

Represents button that renders multiple text segments with independent fonts and colors.

```vb
Public Class RichTextButton
    Inherits NoFocusCueButton
```

### Main members

| Member | Behavior |
| --- | --- |
| `TextParts` | bindable collection of RichTextPart objects. |
| `RichTextPart.Text` | segment content. |
| `RichTextPart.Font` | segment font. |
| `RichTextPart.Color` | segment color. |
| `TooltipText` | tooltip inherited from NoFocusCueButton. |
| `Text` | concatenated read-only-style representation of all parts. |

## Behavior and validation

* Property changes take effect immediately unless a method explicitly starts or applies an operation.
* Callers should validate user input and referenced objects before performing application-specific work.
* Events use the standard Windows Forms event model and should be handled on the UI thread.

## Standard functionality

Because `RichTextButton` inherits from `NoFocusCueButton`, the standard public members of that base type remain available unless the control intentionally hides or overrides them.

## Resource and lifetime considerations

* Dispose components, controls, images, fonts and menus created dynamically when your application owns them.
* Keep referenced controls alive for as long as a component is attached to them.
* File-based operations should receive valid, accessible paths.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.RichTextButton` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.RichTextButton` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| CoreSuite dependencies | `CoreSuite.NoFocusCueButton` |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
