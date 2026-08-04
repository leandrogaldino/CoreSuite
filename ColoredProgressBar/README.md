# ColoredProgressBar

**A lightweight and customizable Windows Forms progress bar with configurable ranges and vertical gradient colors, included in CoreSuite.**

> [!NOTE]
> ColoredProgressBar is one of the independent projects that make up the **CoreSuite** solution. It can be installed and used separately without requiring the other CoreSuite controls.

## Overview

`ColoredProgressBar` provides a simple progress indicator for .NET 8 Windows Forms applications. It displays the current value relative to a configurable range and fills the completed area with a vertical color gradient.

The control can be configured through the Visual Studio Designer or directly in VB.NET code.

## Key features

- Configurable minimum, maximum, and current values.
- Automatic value constraint within the configured range.
- Customizable start and end colors.
- Vertically rendered color gradient.
- Visual Studio Designer support.
- Lightweight implementation for Windows Forms.
- No external package dependencies.

## Requirements

- .NET 8 for Windows
- Windows Forms
- Windows

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.ColoredProgressBar
```

## Quick start

Import the control namespace:

```vbnet
Imports CoreSuite.Controls
```

Add a `ColoredProgressBar` to a form using the Visual Studio Toolbox or create it in code:

```vbnet
Imports System.Drawing
Imports CoreSuite.Controls

Dim progressBar As New ColoredProgressBar With {
    .Minimum = 0,
    .Maximum = 100,
    .Value = 65,
    .ProgressStartColor = Color.LimeGreen,
    .ProgressEndColor = Color.ForestGreen,
    .Location = New Point(20, 20),
    .Size = New Size(250, 24)
}

Controls.Add(progressBar)
```

## Updating the progress

Set `Value` whenever the represented operation advances:

```vbnet
progressBar.Value = 80
```

For an asynchronous operation, update the value after each completed step:

```vbnet
For index As Integer = 1 To 100
    Await Task.Delay(25)
    progressBar.Value = index
Next
```

When progress is reset, assign the configured minimum value:

```vbnet
progressBar.Value = progressBar.Minimum
```

## Range behavior

The displayed progress is calculated from `Minimum`, `Maximum`, and `Value`.

```vbnet
progressBar.Minimum = 0
progressBar.Maximum = 250
progressBar.Value = 125
```

In this example, the control displays 50 percent of its available progress area.

Values outside the configured range are automatically constrained:

- A value lower than `Minimum` becomes `Minimum`.
- A value higher than `Maximum` becomes `Maximum`.

This keeps the control in a valid visual state without requiring the caller to constrain each assigned value manually.

## Gradient colors

Use `ProgressStartColor` and `ProgressEndColor` to define the vertical gradient used in the completed portion of the control:

```vbnet
progressBar.ProgressStartColor = Color.DeepSkyBlue
progressBar.ProgressEndColor = Color.RoyalBlue
```

Use the same color for both properties when a solid-looking fill is preferred:

```vbnet
progressBar.ProgressStartColor = Color.ForestGreen
progressBar.ProgressEndColor = Color.ForestGreen
```

## Designer usage

After installing the package and adding `ColoredProgressBar` to the Visual Studio Toolbox:

1. Drag the control onto a Windows Form.
2. Set `Minimum` and `Maximum` to define the progress range.
3. Set `Value` to preview the current progress.
4. Configure `ProgressStartColor` and `ProgressEndColor` in the Properties window.
5. Adjust the inherited layout properties such as `Location`, `Size`, `Anchor`, and `Dock` as needed.

Changes to the range, value, and colors are reflected by the control without requiring additional setup.

## API reference

### Properties

| Property | Type | Default | Description |
|---|---|---|---|
| `Minimum` | `Integer` | `0` | Gets or sets the lowest value in the progress range. |
| `Maximum` | `Integer` | `100` | Gets or sets the highest value in the progress range. |
| `Value` | `Integer` | `0` | Gets or sets the current progress value. Values outside the range are automatically constrained. |
| `ProgressStartColor` | `Color` | `ForestGreen` | Gets or sets the color at the top of the vertical progress gradient. |
| `ProgressEndColor` | `Color` | `ForestGreen` | Gets or sets the color at the bottom of the vertical progress gradient. |

## Rendering behavior

- The filled width is calculated proportionally from the current range and value.
- The completed portion is painted using the configured vertical gradient.
- Changing a range, value, or color property refreshes the control.
- A zero-length range does not produce invalid drawing calculations.
- The control uses its current size when calculating the progress area.

## Integration notes

- Update `Value` on the UI thread when progress originates from background work.
- Use `Invoke`, `BeginInvoke`, or an awaited UI workflow when a worker thread needs to report progress.
- Use `Anchor` or `Dock` when the control should resize with its parent container.
- Choose gradient colors with sufficient contrast against the control background.

## Package information

| Item | Value |
|---|---|
| Package | `CoreSuite.ColoredProgressBar` |
| Namespace | `CoreSuite.Controls` |
| Target framework | `.NET 8 for Windows` |
| UI framework | `Windows Forms` |

## License

MIT License.
