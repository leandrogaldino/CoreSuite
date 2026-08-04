# FluidResizer

**Smoothly animates a Windows Forms control to a target size.**

> [!NOTE]
> FluidResizer is one of the controls or components included in the **CoreSuite** solution. The package has no additional CoreSuite package dependencies.

## Overview

`FluidResizer` extends `Object` and provides smoothly animates a Windows Forms control to a target size.

It is designed for .NET 8 Windows Forms applications and can be configured in code or through the Visual Studio designer where designer support applies.

## Features

* Associates the resizer with a control.
* Starts resizing toward the requested size.
* Raised when the animation finishes.
* Original and final sizes.
* Uses the standard Windows Forms event and component model.
* Includes English XML documentation for the public API.
* Requires no third-party packages.

## Requirements

* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.FluidResizer
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.FluidResizer
```

## Namespace

```vb
Imports CoreSuite.Controls
```

## Quick start

```vb
Imports CoreSuite.Controls

Dim Resizer As New FluidResizer(DetailsPanel)
AddHandler Resizer.ResizeEnd, Sub(Sender, E)
    Console.WriteLine($"{E.StartSize} -> {E.EndSize}")
End Sub
Resizer.SetSize(New Size(600, 400))
```

## Designer usage

After installing or referencing the package, add `FluidResizer` from the Toolbox or create it in code. Properties configured in the Windows Forms Designer are serialized into `InitializeComponent` and remain available at runtime.

## API reference

### `FluidResizer`

Represents smoothly animates a Windows Forms control to a target size.

```vb
Public Class FluidResizer
    Inherits Object
```

### Main members

| Member | Behavior |
| --- | --- |
| `New(Control)` | associates the resizer with a control. |
| `SetSize(TargetSize)` | starts resizing toward the requested size. |
| `ResizeEnd` | raised when the animation finishes. |
| `ResizeEndEventArgs.StartSize / EndSize` | original and final sizes. |

## Behavior and validation

* Property changes take effect immediately unless a method explicitly starts or applies an operation.
* Callers should validate user input and referenced objects before performing application-specific work.
* Events use the standard Windows Forms event model and should be handled on the UI thread.

## Standard functionality

Because `FluidResizer` inherits from `Object`, the standard public members of that base type remain available unless the control intentionally hides or overrides them.

## Resource and lifetime considerations

* Dispose components, controls, images, fonts and menus created dynamically when your application owns them.
* Keep referenced controls alive for as long as a component is attached to them.
* File-based operations should receive valid, accessible paths.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.FluidResizer` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.FluidResizer` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| CoreSuite dependencies | `None` |
| External dependencies | None |

## License

This package is distributed under the [MIT License](https://github.com/leandrogaldino/CoreSuite/blob/main/LICENSE) defined by the CoreSuite repository.
