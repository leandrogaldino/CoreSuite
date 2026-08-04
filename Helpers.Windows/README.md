# Helpers.Windows

**Windows Forms helpers for working with controls and images.**

> [!NOTE]
> Helpers.Windows is one of the libraries included in the **CoreSuite** solution. It has no CoreSuite or third-party package dependencies.

## Overview

`Helpers.Windows` provides small, reusable utilities for Windows Forms applications. `ControlHelper` helps traverse control hierarchies, control double buffering and verify whether a child control is fully visible. `ImageHelper` provides image loading, cloning, recoloring, color analysis and proportional resizing.

It is designed for .NET 8 Windows Forms applications and can be used from any form, user control or Windows Forms component.

## Features

* Enumerates every descendant control from a root control.
* Optionally includes the root control in control enumeration.
* Enables or disables the protected `DoubleBuffered` property on controls and forms.
* Checks whether a child control is fully visible inside its parent.
* Loads an independent copy of a file-based image.
* Supplies a built-in error image when a file cannot be loaded.
* Recolors images by replacing an exact color or preserving per-pixel alpha.
* Obtains all distinct colors used by an image.
* Resizes images proportionally within optional width and height limits.
* Includes English XML documentation for the public API.
* Requires no third-party packages.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.Helpers.Windows
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.Helpers.Windows
```

## Namespace

```vb
Imports CoreSuite.Helpers
```

## Quick start

```vb
Imports CoreSuite.Helpers

Dim AllControls As IEnumerable(Of Control) = ControlHelper.GetAllControls(Me, True)
ControlHelper.EnableFormDoubleBuffer(Me, True)

Dim PreviewImage As Image = ImageHelper.GetCopyImage("C:\Images\logo.png")
PictureBox1.Image = ImageHelper.GetResizedImage(PreviewImage, 128, 128)
```

## API reference

### `ControlHelper`

Provides utility methods for traversing, configuring and inspecting Windows Forms controls and forms.

| Member | Behavior |
| --- | --- |
| `GetAllControls(root, includeRoot)` | Enumerates child controls recursively, optionally including the root control. |
| `EnableControlDoubleBuffer(control, setting)` | Enables or disables double buffering on a control through reflection. |
| `EnableFormDoubleBuffer(form, setting)` | Enables or disables double buffering on a form through reflection. |
| `IsControlFullyVisible(parent, child)` | Returns whether the complete child bounds are visible inside the parent client area. |

### `ImageHelper`

Provides helper methods for common image operations.

| Member | Behavior |
| --- | --- |
| `GetCopyImage(imagePath)` | Loads and returns an independent image copy; returns the built-in error image when loading fails. |
| `GetRecoloredImage(image, fromColor, toColor)` | Returns a bitmap replacing one exact ARGB color with another. |
| `GetRecoloredImage(image, baseColor)` | Recolors all image colors with the supplied base color while preserving alpha values. |
| `GetImageColors(image)` | Returns the distinct colors used by the image. |
| `GetResizedImage(image, maxWidth, maxHeight)` | Resizes proportionally only when the image exceeds a supplied dimension limit. |

## Examples

### Enable double buffering on a DataGridView

```vb
ControlHelper.EnableControlDoubleBuffer(DataGridView1, True)
```

### Recolor an icon while preserving transparency

```vb
Dim AccentIcon As Image = ImageHelper.GetRecoloredImage(My.Resources.Icon, Color.DodgerBlue)
Button1.Image = AccentIcon
```

### Resize an image before displaying it

```vb
Dim Preview As Image = ImageHelper.GetCopyImage(FileName)
PictureBox1.Image = ImageHelper.GetResizedImage(Preview, 320, 180)
```

## Behavior and lifetime considerations

* `EnableControlDoubleBuffer` and `EnableFormDoubleBuffer` use reflection because `DoubleBuffered` is not publicly exposed by Windows Forms.
* `GetResizedImage` disposes the supplied image when a resized replacement is created. Do not use the original image reference after that call.
* An image returned by `GetCopyImage` or either `GetRecoloredImage` overload is owned by the caller and should be disposed when no longer needed, unless it is assigned to a control that owns its lifetime.
* Image methods operate on `Bitmap` pixel data and are best suited to icons and reasonably sized images.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.Helpers.Windows` |
| Namespace | `CoreSuite.Helpers` |
| Assembly | `CoreSuite.Helpers.Windows` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| CoreSuite dependencies | None |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
