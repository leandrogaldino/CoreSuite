# DateTimeBox

**A culture-aware Windows Forms control for entering and selecting a complete date and time in a single field, included in CoreSuite.**

> [!NOTE]
> DateTimeBox is one of the independent projects that make up the **CoreSuite** solution. Its required CoreSuite dependencies are installed automatically with the NuGet package.

## Overview

`DateTimeBox` combines masked keyboard input with a drop-down editor containing a calendar, a time selector, and explicit confirmation and cancellation actions.

The displayed mask, parsing, formatting, calendar, and time selector follow the configured culture. The control supports optional seconds, localized drop-down text, data binding, empty values, custom button images, and Visual Studio Designer integration.

## Key features

- Culture-aware date and time parsing and formatting.
- Short date pattern obtained from the configured culture.
- Short or long time pattern depending on `ShowSeconds`.
- Masked keyboard input.
- Combined calendar and time drop-down editor.
- Explicit OK and Cancel actions.
- Customizable time label and button text.
- Optional custom drop-down button image.
- Built-in calendar and clock glyph when no image is assigned.
- Data binding through the `DateTime` property.
- Windows Forms Designer and smart tag support.
- Empty-value support through `DateTime.MinValue` and `ClearDateTime`.
- Change notification through `DateTimeValueChanged`.

## Requirements

- .NET 8 for Windows
- Windows Forms
- Windows

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.DateTimeBox
```

The package automatically installs its required CoreSuite dependencies.

## Quick start

Import the control namespace:

```vbnet
Imports CoreSuite.Controls
```

Create and configure a `DateTimeBox` in code:

```vbnet
Imports System.Globalization
Imports CoreSuite.Controls

Dim appointmentBox As New DateTimeBox With {
    .DateTimeCulture = CultureInfo.GetCultureInfo("pt-BR"),
    .ShowSeconds = False,
    .DateTime = New DateTime(2026, 8, 2, 14, 30, 0),
    .Width = 180
}

AddHandler appointmentBox.DateTimeValueChanged, Sub(sender As Object, e As EventArgs) If appointmentBox.HasDateTime Then Console.WriteLine(appointmentBox.DateTime)

Controls.Add(appointmentBox)
```

The field displays and parses its value using the short date and time patterns of `DateTimeCulture`.

## Reading and changing the value

Use `DateTime` to read or assign the complete represented value:

```vbnet
AppointmentBox.DateTime = New DateTime(2026, 8, 2, 14, 30, 0)

If AppointmentBox.HasDateTime Then
    Dim appointment As DateTime = AppointmentBox.DateTime
End If
```

`HasDateTime` indicates whether the control currently contains a complete and valid date and time.

## Clearing the value

Call `ClearDateTime` to return the control to its empty state:

```vbnet
AppointmentBox.ClearDateTime()
```

Assigning `DateTime.MinValue` has the same effect:

```vbnet
AppointmentBox.DateTime = DateTime.MinValue
```

When empty:

- `HasDateTime` returns `False`.
- `DateTime` returns `DateTime.MinValue`.
- The masked field remains available for a new value.

## Culture and formatting

Set `DateTimeCulture` to control the date and time patterns used by the field and drop-down editor:

```vbnet
AppointmentBox.DateTimeCulture = CultureInfo.GetCultureInfo("pt-BR")
' Example: 02/08/2026 14:30
```

```vbnet
AppointmentBox.DateTimeCulture = CultureInfo.GetCultureInfo("en-US")
' Example: 08/02/2026 02:30 PM
```

The control obtains its date and time patterns from the configured `CultureInfo`. Single-character components are normalized internally so they can be represented consistently by the masked input field.

Changing the culture updates the input format while preserving the represented `DateTime` value whenever one is available.

## Showing seconds

Enable `ShowSeconds` when the field should display and accept seconds:

```vbnet
AppointmentBox.ShowSeconds = True
AppointmentBox.DateTime = New DateTime(2026, 8, 2, 14, 30, 45)
```

When `ShowSeconds` is `False`, seconds remain part of a value assigned programmatically or selected through the drop-down, but they are not displayed in the field.

## Drop-down editor

The drop-down combines a calendar with a time selector. It starts with the current control value; if the control is empty, it starts with the current local date and time.

- OK combines the selected calendar date and time and updates the control.
- Cancel closes the drop-down without changing the existing value.
- Clicking outside the drop-down closes it without applying pending changes.
- Pressing Enter while the field is focused opens the editor.

The selected value is committed only through the confirmation action.

## Localizing the drop-down

Use the text properties to adapt the drop-down interface to the application's language:

```vbnet
AppointmentBox.TimeLabelText = "Horário:"
AppointmentBox.OKButtonText = "Confirmar"
AppointmentBox.CancelButtonText = "Cancelar"
```

These properties change only the user-facing text. Date and time formatting continue to be controlled by `DateTimeCulture`.

## Drop-down button image

Assign `ButtonImage` to use an application-specific image:

```vbnet
AppointmentBox.ButtonImage = My.Resources.CalendarClock
```

When `ButtonImage` is `Nothing`, the control draws its built-in calendar and clock glyph.

## DateTimeValueChanged event

`DateTimeValueChanged` occurs whenever the represented value changes, including when a valid value is entered, selected from the drop-down, assigned programmatically, or cleared.

```vbnet
Private Sub AppointmentBox_DateTimeValueChanged(sender As Object, e As EventArgs) Handles AppointmentBox.DateTimeValueChanged
    If AppointmentBox.HasDateTime Then
        Text = AppointmentBox.DateTime.ToString("F", AppointmentBox.DateTimeCulture)
    Else
        Text = "No appointment selected"
    End If
End Sub
```

## Data binding

`DateTime` is the control's default binding property:

```vbnet
AppointmentBox.DataBindings.Add(NameOf(AppointmentBox.DateTime), ViewModel, NameOf(ViewModel.AppointmentDateTime), formattingEnabled:=True, updateMode:=DataSourceUpdateMode.OnPropertyChanged)
```

An empty control exposes `DateTime.MinValue`. Convert that sentinel value when the bound model uses `Nullable(Of DateTime)` to represent the absence of a value.

## Designer usage

After installing the package and adding `DateTimeBox` to the Visual Studio Toolbox:

1. Drag the control onto a Windows Form.
2. Set `DateTimeCulture` to the intended input culture.
3. Enable `ShowSeconds` when seconds are required.
4. Configure `TimeLabelText`, `OKButtonText`, and `CancelButtonText` for the interface language.
5. Assign `ButtonImage` when a custom drop-down image is desired.
6. Create a handler for `DateTimeValueChanged` when the form must react to changes.
7. Use the smart tag for the commonly configured DateTimeBox properties.

The control supports horizontal resizing in the Designer while retaining the height required by its masked input and button layout.

## API reference

### Properties

| Property | Type | Default | Description |
|---|---|---|---|
| `DateTime` | `DateTime` | `DateTime.MinValue` | Gets or sets the represented date and time. Assigning `DateTime.MinValue` clears the control. |
| `HasDateTime` | `Boolean` | Read-only | Indicates whether the control currently contains a valid date and time. |
| `DateTimeCulture` | `CultureInfo` | `CultureInfo.CurrentCulture` | Gets or sets the culture used for parsing, formatting, and the drop-down editor. |
| `ShowSeconds` | `Boolean` | `False` | Gets or sets whether seconds are displayed and accepted. |
| `ButtonImage` | `Image` | `Nothing` | Gets or sets a custom image for the drop-down button. |
| `TimeLabelText` | `String` | `"Time:"` | Gets or sets the label displayed beside the time selector. |
| `OKButtonText` | `String` | `"OK"` | Gets or sets the confirmation button text. |
| `CancelButtonText` | `String` | `"Cancel"` | Gets or sets the cancellation button text. |

### Methods

| Method | Description |
|---|---|
| `ClearDateTime()` | Clears the current value and returns the control to its empty state. |

### Events

| Event | Event arguments | Description |
|---|---|---|
| `DateTimeValueChanged` | `EventArgs` | Occurs whenever the represented date and time changes or is cleared. |

## Behavior notes

- `DateTime.MinValue` represents an empty value.
- The mask is rebuilt from the configured culture and `ShowSeconds` setting.
- The drop-down applies pending changes only when the user confirms them.
- The current value is preserved when the drop-down is canceled or dismissed.
- A built-in glyph is drawn when no custom button image is assigned.
- Value changes should be performed on the Windows Forms UI thread.

## Package information

| Item | Value |
|---|---|
| Package | `CoreSuite.DateTimeBox` |
| Namespace | `CoreSuite.Controls` |
| Target framework | `.NET 8 for Windows` |
| UI framework | `Windows Forms` |

## License

MIT License.
