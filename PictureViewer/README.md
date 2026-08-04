# PictureViewer

**Image-path viewer with navigation, add, remove and save commands.**

> [!NOTE]
> PictureViewer is one of the controls or components included in the **CoreSuite** solution. The package depends on `CoreSuite.NoFocusCueButton`. NuGet installs the required package or packages automatically.

## Overview

`PictureViewer` extends `UserControl` and provides image-path viewer with navigation, add, remove and save commands.

It is designed for .NET 8 Windows Forms applications and can be configured in code or through the Visual Studio designer where designer support applies.

## Features

* Maintains an ordered collection of image file paths.
* Provides first, previous, next and last navigation.
* Includes add, remove and save commands.
* Supports custom command images, colors and tooltip captions.
* Uses the standard Windows Forms event and component model.
* Includes English XML documentation for the public API.
* Requires no third-party packages.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.PictureViewer
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.PictureViewer
```

## Namespace

```vb
Imports CoreSuite.Controls
```

## Quick start

```vb
Imports CoreSuite.Controls

Dim Viewer As New PictureViewer With {
    .MaximumPictures = 10,
    .CounterMask = "{0} of {1}",
    .ShowControlBar = True,
    .ShowCounterBar = True
}
Viewer.AddPictures({"C:\Images\one.png", "C:\Images\two.jpg"})
Viewer.MoveFirst()
```

## Designer usage

After installing or referencing the package, add `PictureViewer` from the Toolbox or create it in code. Properties configured in the Windows Forms Designer are serialized into `InitializeComponent` and remain available at runtime.

## API reference

### `PictureViewer`

Represents image-path viewer with navigation, add, remove and save commands.

```vb
Public Class PictureViewer
    Inherits UserControl
```

### Main members

| Member | Behavior |
| --- | --- |
| `Pictures` | read-only list of paths. |
| `SelectedPicture / SelectedIndex` | current selection. |
| `MaximumPictures` | optional collection limit. |
| `CounterMask` | counter text format. |
| `ShowControlBar / ShowCounterBar` | visibility options. |
| `ToolTips` | expandable command captions. |
| `AddPicture(s), RemovePicture(), Clear()` | collection operations. |
| `MoveFirst/Previous/Next/Last()` | navigation. |
| `SaveSelectedPicture()` | copies the selected image. |
| `PictureAdded, PictureRemoved and SelectedPictureChanged` | lifecycle events. |

## Behavior and validation

* Property changes take effect immediately unless a method explicitly starts or applies an operation.
* Callers should validate user input and referenced objects before performing application-specific work.
* Events use the standard Windows Forms event model and should be handled on the UI thread.

## Standard functionality

Because `PictureViewer` inherits from `UserControl`, the standard public members of that base type remain available unless the control intentionally hides or overrides them.

## Resource and lifetime considerations

* Dispose components, controls, images, fonts and menus created dynamically when your application owns them.
* Keep referenced controls alive for as long as a component is attached to them.
* File-based operations should receive valid, accessible paths.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.PictureViewer` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.PictureViewer` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| CoreSuite dependencies | `CoreSuite.NoFocusCueButton` |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
