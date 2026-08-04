# ValidationProvider

**Centralized, designer-friendly validation for .NET 8 Windows Forms, included in CoreSuite.**

> [!NOTE]
> ValidationProvider is one of the independent projects that make up the **CoreSuite** solution. The package contains the non-visual component, validation result types, event data, and Windows Forms designer support.

## Overview

`ValidationProvider` centralizes form validation in one non-visual component. After the component is added to a form, every Windows Forms control receives validation properties in the designer, including `Required`, `MinimumLength`, `MaximumLength`, `RegularExpression`, `CompareWith`, and `ValidationGroup`.

The provider evaluates the configured rules, displays errors through the familiar `ErrorProvider` icon, and returns a single Boolean result. This removes repeated `If` blocks from Save and Confirm buttons while keeping each field's rules close to that field in the designer.

Because `ValidationProvider` inherits from `ErrorProvider`, inherited settings such as `Icon`, `BlinkStyle`, `IconAlignment`, `IconPadding`, and `ContainerControl` remain available.

## Key features

- Non-visual component with designer extender properties on every control.
- Required-value validation for text, selection, date, numeric, Boolean, and collection values.
- Minimum- and maximum-length rules.
- Regular-expression validation with a one-second safety timeout.
- `MaskedTextBox.MaskCompleted` validation.
- Value comparison between two controls, useful for password and e-mail confirmation.
- Named validation groups for multi-step forms and tabbed screens.
- Automatic validation during each control's `Validating` event.
- Optional focus cancellation for invalid fields.
- Optional focus on the first invalid control after batch validation.
- Automatic clearing of stale feedback as values change.
- Configurable participation of disabled and hidden controls.
- Custom display names and fully configurable default messages.
- Application-defined value resolution and custom validation rules through events.
- Explicit inclusion of controls that use custom-only validation rules.
- Nested value paths for custom controls, such as `Frozen.IsFrozen`.
- Per-control validation results and complete batch results.
- Smart-tag actions for the most common provider settings.
- XML documentation and NuGet symbol generation.

## Requirements

- Windows Forms
- .NET 8 for Windows (`net8.0-windows`)
- A reference to `CoreSuite.ValidationProvider`

The component has no runtime dependency on another CoreSuite package.

## Installation

```powershell
dotnet add package CoreSuite.ValidationProvider
```

Or add `ValidationProvider/ValidationProvider.vbproj` as a project reference when working directly with the CoreSuite solution.

## Quick start

Add `ValidationProvider` to the form through the toolbox. The designer associates its inherited `ContainerControl` property with the form and adds validation properties to the other controls.

Configure the fields in the designer or in code:

```vb
Imports CoreSuite.Controls
Private Sub ConfigureValidation()
    ValidationProvider1.SetRequired(NameTextBox, True)
    ValidationProvider1.SetValidationDisplayName(NameTextBox, "Name")
    ValidationProvider1.SetMinimumLength(NameTextBox, 3)
    ValidationProvider1.SetRequired(EmailTextBox, True)
    ValidationProvider1.SetValidationDisplayName(EmailTextBox, "E-mail")
    ValidationProvider1.SetRegularExpression(EmailTextBox, "^[^@\s]+@[^@\s]+\.[^@\s]+$")
End Sub
```

Validate before saving:

```vb
Private Sub SaveButton_Click(Sender As Object, E As EventArgs) Handles SaveButton.Click
    If Not ValidationProvider1.Validate() Then Return
    SaveCustomer()
End Sub
```

When validation fails, the provider focuses the first invalid control by default and displays its message using the inherited ErrorProvider icon.

## Designer extender properties

These properties appear under the `ValidationProvider` category of every control on the form.

| Property | Default | Description |
|---|---:|---|
| `Required` | `False` | Requires the control to represent a non-empty value. |
| `CustomValidationEnabled` | `False` | Includes a control that uses only an application-defined validation rule. |
| `ValidationGroup` | Empty | Assigns the control to a named validation group. |
| `ValidationMessage` | Empty | Replaces the generated error message for this control. |
| `ValidationDisplayName` | Empty | Defines the friendly field name inserted into generated messages. |
| `MinimumLength` | `0` | Requires at least this many characters; zero disables the rule. |
| `MaximumLength` | `0` | Allows no more than this many characters; zero disables the rule. |
| `RegularExpression` | Empty | Requires the represented text to match a regular expression. |
| `CompareWith` | `Nothing` | Requires this control's value to match another control's value. |
| `ValuePropertyName` | Empty | Uses a specific property or nested property path as the value. |

The designer serializes these settings through methods such as:

```vb
ValidationProvider1.SetRequired(NameTextBox, True)
ValidationProvider1.SetValidationGroup(NameTextBox, "Customer")
```

## Validation groups

Groups allow one provider to validate independent sections of the same form.

```vb
ValidationProvider1.SetValidationGroup(NameTextBox, "Customer")
ValidationProvider1.SetValidationGroup(EmailTextBox, "Customer")
ValidationProvider1.SetValidationGroup(CardNumberTextBox, "Payment")
```

Validate only the current section:

```vb
If Not ValidationProvider1.ValidateGroup("Customer") Then Return
```

Group names are compared without case sensitivity. Calling `ValidateGroup(String.Empty)` validates configured controls that do not belong to a named group.

## Compare two fields

`CompareWith` compares the represented text of both controls. It is useful for confirmation fields:

```vb
ValidationProvider1.SetRequired(PasswordTextBox, True)
ValidationProvider1.SetRequired(ConfirmPasswordTextBox, True)
ValidationProvider1.SetCompareWith(ConfirmPasswordTextBox, PasswordTextBox)
ValidationProvider1.SetValidationDisplayName(ConfirmPasswordTextBox, "Password confirmation")
```

Comparison is case-sensitive by default. Set `CaseSensitiveComparison` to `False` for values where character casing is not significant.

## Custom messages and localization

Set `ValidationMessage` when one control needs a specific message:

```vb
ValidationProvider1.SetValidationMessage(NameTextBox, "Enter the customer's full name.")
```

The provider also exposes default message templates:

| Property | Placeholders |
|---|---|
| `RequiredErrorMessage` | `{0}` = display name |
| `MinimumLengthErrorMessage` | `{0}` = display name, `{1}` = minimum length |
| `MaximumLengthErrorMessage` | `{0}` = display name, `{1}` = maximum length |
| `InvalidFormatErrorMessage` | `{0}` = display name |
| `ComparisonErrorMessage` | `{0}` = current display name, `{1}` = comparison display name |
| `CustomErrorMessage` | `{0}` = display name |

This allows the component to be fully localized without changing its source code:

```vb
ValidationProvider1.RequiredErrorMessage = "{0} é obrigatório."
ValidationProvider1.InvalidFormatErrorMessage = "O formato de {0} é inválido."
ValidationProvider1.ComparisonErrorMessage = "{0} deve ser igual a {1}."
```

## Automatic validation

`AutomaticValidation` is enabled by default. Each configured control is validated when its `Validating` event occurs, normally when focus moves to another control.

| Property | Default | Behavior |
|---|---:|---|
| `AutomaticValidation` | `True` | Validates a configured control during its `Validating` event. |
| `CancelValidationOnError` | `False` | Sets `CancelEventArgs.Cancel` when automatic validation fails. |
| `ClearErrorOnValueChanged` | `True` | Clears stale feedback while the user edits the value. |

`CancelValidationOnError` is disabled by default so users are not trapped in a field. Enable it only when the form workflow should prevent focus from leaving invalid input.

## Custom validation rules

Handle `ValidatingControl` for business rules that cannot be represented by the built-in properties:

```vb
Private Sub ConfigureUserNameValidation()
    ValidationProvider1.SetCustomValidationEnabled(UserNameTextBox, True)
End Sub
Private Sub ValidationProvider1_ValidatingControl(Sender As Object, E As ValidatingControlEventArgs) Handles ValidationProvider1.ValidatingControl
    If E.TargetControl Is UserNameTextBox AndAlso String.Equals(CStr(E.Value), "admin", StringComparison.OrdinalIgnoreCase) Then
        E.IsValid = False
        E.FailureReason = ValidationFailureReason.Custom
        E.ErrorMessage = "This user name is reserved."
    End If
End Sub
```

The event receives the result of the built-in rules. A handler may reject an otherwise valid value, replace the error message, or accept a built-in failure when application logic explicitly permits it.

## Custom value resolution

The provider automatically recognizes common WinForms values:

- `Text` for text-based controls.
- `SelectedValue` or selection state for list controls.
- `Checked` for check boxes and radio buttons.
- `Value` for `DateTimePicker`, `NumericUpDown`, and compatible custom controls.
- `HasDate`, `HasTime`, `HasDateTime`, or `HasValue` when exposed by a custom control.
- Checked item collections for `CheckedListBox`.

Use `ValuePropertyName` when the desired value is exposed through another property:

```vb
ValidationProvider1.SetRequired(CustomerQueriedBox, True)
ValidationProvider1.SetValuePropertyName(CustomerQueriedBox, "Frozen.IsFrozen")
```

Dot-separated nested paths are supported. A missing property is treated as a configuration error and raises `InvalidOperationException`, making misspelled paths visible during testing.

For controls that need code-based resolution, handle `ValidationValueRequested`:

```vb
Private Sub ValidationProvider1_ValidationValueRequested(Sender As Object, E As ValidationValueRequestedEventArgs) Handles ValidationProvider1.ValidationValueRequested
    If E.TargetControl Is TagsControl Then
        E.Value = TagsControl.SelectedTags
        E.Handled = True
    End If
End Sub
```

## CoreSuite control integration

The package remains independent from other CoreSuite projects, but its convention-based value resolution supports them naturally.

| Control | Default validation value |
|---|---|
| `DateBox` | Empty when `HasDate` is false; otherwise its `Date` value. |
| `TimeBox` | Empty when `HasTime` is false; otherwise its `Time` value. |
| `DateTimeBox` | Empty when its availability property is false; otherwise its value. |
| Numeric boxes | Their represented value or displayed text. Zero remains a valid numeric value. |
| `QueriedBox` | Its text by default; use `Frozen.IsFrozen` when a frozen selection is required. |

## Methods

| Method | Description |
|---|---|
| `Validate()` | Validates all eligible configured controls. |
| `ValidateGroup(groupName)` | Validates eligible controls in one case-insensitive group. |
| `ValidateControl(control)` | Validates one configured control. |
| `ClearValidation()` | Clears every message displayed by the provider. |
| `ClearValidation(control)` | Clears one control's message. |
| `ClearValidationGroup(groupName)` | Clears messages for one group. |

Disabled and hidden controls are excluded by default. Change `ValidateDisabledControls` or `ValidateHiddenControls` when those controls must participate.

## Events and results

| Event | Description |
|---|---|
| `ValidationValueRequested` | Allows the application to replace the automatically resolved value. |
| `ValidatingControl` | Allows custom validation after the built-in rules run. |
| `ControlValidated` | Reports the result of one completed control validation. |
| `ValidationCompleted` | Reports every result produced by `Validate` or `ValidateGroup`. |

```vb
Private Sub ValidationProvider1_ValidationCompleted(Sender As Object, E As ValidationCompletedEventArgs) Handles ValidationProvider1.ValidationCompleted
    StatusLabel.Text = If(E.IsValid, "Ready to save", $"{E.InvalidControlCount} invalid field(s)")
End Sub
```

Each `ValidationResult` contains the target control, group name, final message, validity, and a `ValidationFailureReason` value.

## ErrorProvider appearance

Because the component inherits `ErrorProvider`, standard appearance configuration remains available:

```vb
ValidationProvider1.BlinkStyle = ErrorBlinkStyle.NeverBlink
ValidationProvider1.SetIconAlignment(NameTextBox, ErrorIconAlignment.MiddleRight)
ValidationProvider1.SetIconPadding(NameTextBox, 4)
```

## License

CoreSuite is licensed under the MIT License.
