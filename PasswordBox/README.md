# CoreSuite PasswordBox

`CoreSuite.PasswordBox` is a reusable Windows Forms password editor for .NET 8. It combines a borderless `TextBox` with a right-aligned action button and supports a password-replacement workflow that does not require an already stored password to be loaded into the control.

## Features

- Password entry with system password masking.
- Embedded action button positioned on the right side.
- Customizable `ShowPasswordImage`, `HidePasswordImage`, and `ChangePasswordImage`.
- No built-in images or GDI+ glyphs; the application defines the initial visual identity.
- Existing-password state without loading the real stored password.
- Fixed-length placeholder for already stored passwords.
- Context-sensitive action button that switches between change, show, and hide actions.
- Password replacement tracking through `PasswordChanged`.
- Cancelable replacement workflow.
- `AcceptPasswordChange` helper to clear the plaintext replacement value after persistence.
- Read-only mode.
- Configurable reveal behavior, button width, image alignment, and mask length.
- Customizable border colors and localizable tooltips.
- `F2` starts password replacement and `Esc` cancels it.
- Accessibility metadata for the embedded action button.
- Complete XML documentation.

## Installation

```powershell
Install-Package CoreSuite.PasswordBox
```

Or:

```bash
dotnet add package CoreSuite.PasswordBox
```

## Configuring action images

No image is defined by default. Configure them in the Windows Forms Designer or in code:

```vb
PasswordBox1.ShowPasswordImage = My.Resources.ShowPassword
PasswordBox1.HidePasswordImage = My.Resources.HidePassword
PasswordBox1.ChangePasswordImage = My.Resources.ChangePassword
```

The control changes the image automatically according to its current state.

| State | Property |
|---|---|
| Password hidden and editable | `ShowPasswordImage` |
| Password visible and editable | `HidePasswordImage` |
| Existing password defined | `ChangePasswordImage` |

The supplied images are not disposed by `PasswordBox`; ownership remains with the application or designer resources.

## Basic usage

```vb
PasswordBox1.PasswordDefined = False
PasswordBox1.PlaceholderText = "Enter password"
```

When saving:

```vb
If PasswordBox1.PasswordChanged Then
    EmailCredentialsService.SavePassword(PasswordBox1.Password)
    PasswordBox1.AcceptPasswordChange()
End If
```

## Existing password workflow

```vb
PasswordBox1.PasswordDefined = True
```

The real stored password does not need to be loaded into the control.

```text
Stored:      [ ••••••••                 ][ change ]
Editing:     [ new password             ][ show   ]
Visible:     [ new password             ][ hide   ]
```

After successful persistence:

```vb
PasswordBox1.AcceptPasswordChange(passwordDefined:=True)
```

To cancel:

```vb
PasswordBox1.CancelPasswordChange()
```

## Main properties

| Property | Description |
|---|---|
| `PasswordDefined` | Indicates whether the application already has a stored password. |
| `Password` | Returns the replacement password currently entered by the user. |
| `PasswordChanged` | Indicates whether the replacement password has been modified. |
| `IsEditingPassword` | Indicates whether replacement mode is active. |
| `IsPasswordVisible` | Indicates whether the replacement password is visible. |
| `ReadOnly` | Disables editing and replacement. |
| `AllowPasswordReveal` | Enables or disables password reveal. |
| `ShowActionButton` | Shows or hides the action button. |
| `ActionButtonWidth` | Sets the button width. |
| `ActionButtonImageAlign` | Sets the image alignment in the button. |
| `ShowPasswordImage` | Image used for the show action. |
| `HidePasswordImage` | Image used for the hide action. |
| `ChangePasswordImage` | Image used for the change action. |
| `DefinedPasswordMaskLength` | Sets the fixed stored-password placeholder length. |
| `MaxLength` | Sets maximum password length. |
| `TextAlign` | Sets password text alignment. |
| `ShortcutsEnabled` | Enables standard text editing shortcuts. |
| `PlaceholderText` | Sets the empty-field placeholder. |
| `BorderColor` | Sets normal border color. |
| `FocusedBorderColor` | Sets focused border color. |
| `DisabledBorderColor` | Sets disabled border color. |
| `ChangePasswordToolTipText` | Sets the change action tooltip. |
| `ShowPasswordToolTipText` | Sets the show action tooltip. |
| `HidePasswordToolTipText` | Sets the hide action tooltip. |

## Main methods

| Method | Description |
|---|---|
| `BeginPasswordChange()` | Starts password replacement mode. |
| `CancelPasswordChange()` | Cancels replacement. |
| `AcceptPasswordChange()` | Marks the replacement as persisted and clears plaintext. |
| `SetPassword()` | Sets a replacement password programmatically. |
| `ClearPassword()` | Clears the replacement editor. |
| `SelectAllPassword()` | Selects the active replacement password. |
| `FocusPassword()` | Focuses the password editor. |

## Events

| Event | Description |
|---|---|
| `PasswordValueChanged` | Raised when the replacement password changes. |
| `PasswordEditStarted` | Raised when replacement begins. |
| `PasswordEditCanceled` | Raised when replacement is canceled. |
| `PasswordVisibilityChanged` | Raised when password visibility changes. |

## Security notes

`PasswordBox` can represent an existing password without receiving the existing plaintext value. The fixed placeholder length does not expose the actual password length.

A password mask is visual protection only and does not encrypt a password while it is being edited.

## Requirements

- .NET 8
- Windows Forms
- `net8.0-windows`

## License

MIT
