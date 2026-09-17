# CoreSuite.ChangeTracker

A non-visual Windows Forms component that tracks control properties against an accepted baseline and exposes a centralized Boolean change state.

Targets **.NET 8 for Windows**, uses the **CoreSuite.Controls** namespace, and follows the CoreSuite component pattern: designer extender properties, English component metadata, XML documentation, PascalCase identifiers, a designer smart tag, and NuGet package metadata.

## Why a separate component?

`ValidationProvider` answers whether values are valid. `ChangeTracker` answers whether values differ from the accepted baseline. Neither component depends on the other. ChangeTracker does not validate, save, restore, or modify control values.

## Requirements and project layout

- Windows and the .NET 8 SDK for building and running.
- Visual Studio with Windows Forms tooling for visual design.
- NuGet access to restore `Microsoft.WinForms.Designer.SDK` version `1.6.0`, matching the ValidationProvider template. The designer package is marked `PrivateAssets="all"`.
- Create, configure, use, and dispose the tracker on the same UI thread as its controls.

```text
ChangeTracker/
    ChangeTracker.sln
    ChangeTracker.vbproj
    README.md
    LICENSE
    VALIDATION.md
    Source/
        ChangeTracker.vb
        ChangeTrackerDesigner.vb
        ChangeTrackerDesignerActionList.vb
        ChangeTrackingResult.vb
        ControlTrackingSettings.vb
        HasChangesChangedEventArgs.vb
        TrackedPropertyChangedEventArgs.vb
        TrackingSuspension.vb
    Tests/
        ChangeTracker.Tests.vbproj
        Program.vb
```

The test sources are explicitly excluded from the library project. Open the supplied solution for a standalone build, or add `ChangeTracker.vbproj` to the existing CoreSuite solution. No changes to an existing CoreSuite solution are required by this delivery.

## Build, test, and package

Run these commands from the `ChangeTracker` directory on Windows:

```powershell
dotnet restore ChangeTracker.sln
dotnet build ChangeTracker.sln --configuration Release
dotnet run --project Tests/ChangeTracker.Tests.vbproj --configuration Release
dotnet pack ChangeTracker.vbproj --configuration Release --output Artifacts
```

`GeneratePackageOnBuild` is enabled, as in the template. The package is `CoreSuite.ChangeTracker`, version `1.0.0`; the assembly is `CoreSuite.ChangeTracker.dll`. The package includes the README and XML documentation, and packing also produces a `.snupkg` symbol package. Source delivery does not imply that this version has been published to a remote feed.

To consume a local package after packing:

```powershell
dotnet add YourApplication.vbproj package CoreSuite.ChangeTracker --version 1.0.0 --source C:\Packages
```

Replace `C:\Packages` with the directory containing the generated package, or use a project reference during development.

## Designer configuration

1. Build and reference the project, then add `ChangeTracker` to the form's component tray. If necessary, use the toolbox's component selection to locate the assembly.
2. Select each participating control and find its `ChangeTracker` category.
3. Set `TrackChanges` to `True`.
4. Set `TrackedProperties` to one or more property names, separated by commas. Leave it empty only when a built-in default is appropriate.
5. Handle `HasChangesChanged` to react to the centralized Boolean state.
6. Call `AcceptChanges()` after the initial data and bindings have finished loading.

The smart tag displays configuration guidance. It does not execute runtime tracking operations in the designer. The component provides both parameterless and `IContainer` constructors; using the form's component container ensures automatic disposal.

| Extender property | Default | Purpose |
| --- | --- | --- |
| `TrackChanges` | `False` | Includes the control in tracking. |
| `TrackedProperties` | Empty | Comma-separated readable scalar property names. Empty selects a known default. |
| `TrackedEvents` | Empty | Optional comma-separated `System.EventHandler` event names, replacing automatic notification discovery for that control. |
| `AutomaticTracking` | `True` | Attaches notifications. When `False`, call `RefreshChanges()` explicitly. |

Names are trimmed and resolved case-insensitively; duplicate names are removed. Empty list elements are rejected. Only direct property names are supported, not nested paths or indexers. Configuration is independent of setter order in generated designer code. Runtime subscriptions and property validation begin when tracking starts, not while the designer is assigning extender properties.

### Built-in property defaults

| Control | Default property | Notes |
| --- | --- | --- |
| `TextBoxBase` descendants | `Text` | Includes `TextBox`, `MaskedTextBox`, and `RichTextBox`; formatting is not included. |
| `CheckBox` | `CheckState` | Preserves the indeterminate state. |
| `RadioButton` | `Checked` | Tracks the Boolean selection. |
| `ComboBox` | `SelectedIndex` | Use `SelectedValue` for bound identifiers; add `Text` if typed text matters. |
| `DateTimePicker` | `Value` | If its checkbox matters, configure additional observation explicitly as described below. |
| `NumericUpDown` | `Value` | Tracks the decimal value. |
| `TrackBar` | `Value` | Tracks the numeric position. |
| Other controls | No implicit default | Configure `TrackedProperties` explicitly. |

Derived controls inherit these defaults. For a CoreSuite control whose meaningful property is not `Text`, explicitly configure that property, for example `DecimalValue` or the actual selected-key property exposed by that control. There is no dependency on other CoreSuite packages and no guesswork about custom control APIs.

## Minimal usage

If configuration is supplied by the designer, only the event handler and baseline call are required. This complete form example configures standard controls in code:

```vb
Imports System.ComponentModel
Imports CoreSuite.Controls

Public Class ExampleForm
    Inherits Form
    Private ReadOnly _Components As New Container
    Private WithEvents _ChangeTracker As New ChangeTracker(_Components)
    Private ReadOnly _NameInput As New TextBox With {.Dock = DockStyle.Top}
    Private ReadOnly _ActiveInput As New CheckBox With {.Text = "Active", .Dock = DockStyle.Top}
    Private ReadOnly _SaveButton As New Button With {.Text = "Save", .Dock = DockStyle.Bottom, .Enabled = False}
    Public Sub New()
        Controls.AddRange({_NameInput, _ActiveInput, _SaveButton})
        _ChangeTracker.SetTrackChanges(_NameInput, True)
        _ChangeTracker.SetTrackedProperties(_NameInput, "Text")
        _ChangeTracker.SetTrackChanges(_ActiveInput, True)
        _ChangeTracker.SetTrackedProperties(_ActiveInput, "Checked")
        _NameInput.Text = "Initial name"
        _ActiveInput.Checked = True
        _ChangeTracker.AcceptChanges()
    End Sub
    Private Sub TrackerHasChangesChanged(Sender As Object, E As HasChangesChangedEventArgs) Handles _ChangeTracker.HasChangesChanged
        _SaveButton.Enabled = E.HasChanges
    End Sub
    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing Then _Components.Dispose()
        MyBase.Dispose(Disposing)
    End Sub
End Class
```

The example requires a Windows Forms project and its standard `System.Windows.Forms` import. It illustrates tracking only; the save button intentionally has no persistence implementation.

## Baseline semantics

Given a baseline of `"Original"`:

| Operation | `HasChanges` | Boolean event |
| --- | --- | --- |
| `AcceptChanges()` | `False` | None when already clean. |
| Set text to `"Edited"` | `True` | `True`. |
| Set text to `"Edited again"` | `True` | None; use `TrackedPropertyChanged` for each edit. |
| Set text back to `"Original"` | `False` | `False`. |
| Edit and then successfully save; call `AcceptChanges()` | `False` | `False` if previously dirty. |

`HasChanges` remains `True` while **any** participating property differs from its baseline. Returning one field to its initial value does not clear another field's change. Repeated notifications with equal values do not produce duplicate property-change events.

An initial `AcceptChanges()` establishes the baseline and starts tracking. There is no automatic baseline based on `Load`, `Shown`, focus, or handle creation: the application decides when its asynchronous loading and data binding are actually complete. Initialize save-button state explicitly because the initial clean state does not raise a transition event.

## Public runtime API

| Member | Behavior |
| --- | --- |
| `HasChanges` | Cached aggregate state; no property reads or events are triggered by getting it. |
| `ChangedPropertyCount` | Number of observed properties differing from their baselines, not number of controls or edits. |
| `IsTracking` | Whether tracking has successfully started. |
| `IsSuspended` | Whether one or more suspension scopes remain open. |
| `StartTracking()` | Validates configuration, captures initial values, and subscribes. When already active, refreshes without accepting edits. |
| `AcceptChanges()` | Captures current values as the new baseline; starts tracking if stopped. |
| `StopTracking()` | Unsubscribes, discards the baseline, and clears the state; retains configuration and never changes control values. |
| `SuspendTracking()` | Returns a nestable `IDisposable` scope. The outermost disposal refreshes against the existing baseline. |
| `RefreshChanges()` | Reads all participating properties and reconciles state; does nothing when stopped or suspended. |
| `GetChanges()` | Returns a read-only snapshot of cached differences, including control, property, original value, and current value. |

There is deliberately no ambiguous `ResetChanges()` method: use `AcceptChanges()` to accept current values, or reload values yourself to undo edits. There is no automatic undo because control setters can have application-specific side effects, dependencies, and binding behavior.

## Events

### `HasChangesChanged`

Raised only when the aggregate Boolean changes. Its arguments expose `HasChanges` and `ChangedPropertyCount`. Accepting changes, stopping, or removing a disposed control can also cause a transition. Disposing the tracker releases resources without raising user events.

```vb
Private Sub TrackerHasChangesChanged(Sender As Object, E As HasChangesChangedEventArgs) Handles ChangeTracker1.HasChangesChanged
    BtnSave.Enabled = E.HasChanges
End Sub
```

### `TrackedPropertyChanged`

Raised for each distinct observed value change, including a return to the baseline. Use this event when you want the centralized Boolean **on every observed edit**, even if it remains `True`:

```vb
Private Sub TrackerPropertyChanged(Sender As Object, E As TrackedPropertyChangedEventArgs) Handles ChangeTracker1.TrackedPropertyChanged
    BtnSave.Enabled = E.HasChanges
    Debug.WriteLine($"{E.TargetControl.Name}.{E.PropertyName}: {E.PreviousValue} -> {E.CurrentValue}")
End Sub
```

Arguments include `TargetControl`, `PropertyName`, `OriginalValue`, `PreviousValue`, `CurrentValue`, `IsChanged`, `HasChanges`, and `ChangedPropertyCount`. `OriginalValue` is the accepted baseline; `PreviousValue` is the preceding observation. `IsChanged` refers to the single property; `HasChanges` refers to the entire tracker.

Each refresh reads all values, computes the complete aggregate, commits the cache, raises a Boolean transition if needed, and then raises individual property events. Events contain snapshots from that refresh. If a handler changes another tracked control, another refresh is queued after the current batch. Do not write handlers that repeatedly oscillate values. Configuration and lifecycle calls inside tracker handlers are rejected; defer them using `BeginInvoke` if required. Event-handler exceptions propagate to the caller.

Accepting a baseline or stopping tracking does not emit synthetic property-edit events. Changes made and reverted entirely inside a suspension are not observed as intermediate edits. This component is a state tracker, not an audit log.

## Loading another record

Suspension postpones observation; it does not itself mark changes as accepted. Keep `AcceptChanges()` inside the scope to avoid an intermediate dirty transition:

```vb
Using ChangeTracker1.SuspendTracking()
    TxtName.Text = Model.Name
    ChkActive.Checked = Model.IsActive
    CmbPerson.SelectedValue = Model.PersonID
    ChangeTracker1.AcceptChanges()
End Using
```

For asynchronous loading, await the data first, then enter a short synchronous suspension to apply the result on the UI thread. Do not hold a suspension across `Await`, because other UI interactions could then occur while observation is suspended.

```vb
Dim Model = Await Service.LoadAsync(ModelID)
Using ChangeTracker1.SuspendTracking()
    TxtName.Text = Model.Name
    ChangeTracker1.AcceptChanges()
End Using
```

These integration snippets assume the application's own model, service, and controls. If applying values fails, the scope still resumes observation and any partially applied edits remain visible as changes; it does not silently accept them.

## Saving and validation

Keep validation and persistence in the save workflow. Avoid calling `Validate()` from each tracking event, because validation may display errors or move focus during normal editing.

```vb
If Not ValidationProvider1.Validate() Then Return
Await SaveCurrentRecordAsync()
ChangeTracker1.AcceptChanges()
```

Disable editing during an asynchronous save, or otherwise ensure that `AcceptChanges()` corresponds to the values actually saved. If the user edits again while the save is running, blindly accepting the current controls would incorrectly accept unsaved edits. Do not call `AcceptChanges()` when persistence fails.

## Multiple properties and bound selections

```vb
ChangeTracker1.SetTrackChanges(CmbPerson, True)
ChangeTracker1.SetTrackedProperties(CmbPerson, "SelectedIndex, SelectedValue")
```

For a bound combo, tracking only `SelectedValue` is usually preferable: reordering the data source should not count as a different business identifier. Include `SelectedIndex` only if the position itself matters. For editable combos, add `Text` if free typing matters. Establish the baseline after the data source, `ValueMember`, and selected value have settled.

## Custom controls and notification discovery

Automatic discovery uses `PropertyDescriptor.SupportsChangeEvents` and paired `AddValueChanged` / `RemoveValueChanged` subscriptions. This supports conventional property notifications exposed by the component model, including supported `INotifyPropertyChanged` properties. The control must actually raise a notification after updating its value.

If a computed property uses a differently named event, configure `TrackedEvents`. All listed events trigger a full refresh, and replace automatic discovery for that control:

```vb
ChangeTracker1.SetTrackChanges(AmountControl, True)
ChangeTracker1.SetTrackedProperties(AmountControl, "DecimalValue")
ChangeTracker1.SetTrackedEvents(AmountControl, "TextChanged")
```

This example is appropriate only if that control updates `DecimalValue` before raising `TextChanged`. If it raises a later event, use that event instead. The same principle applies to a custom selected-key property. Multiple event names can be separated with commas.

Explicit events must use the exact `System.EventHandler` delegate type. For events such as `KeyEventHandler`, `EventHandler(Of TEventArgs)`, or other custom delegates, set `AutomaticTracking` to `False` and call `RefreshChanges()` from an application handler after the value has changed. No timers, polling, message hooks, dynamic code generation, or background threads are used.

### Manual observation

```vb
ChangeTracker1.SetTrackChanges(CustomControl, True)
ChangeTracker1.SetTrackedProperties(CustomControl, "SelectedID")
ChangeTracker1.SetAutomaticTracking(CustomControl, False)
ChangeTracker1.AcceptChanges()

' Call after an application action or custom event has updated SelectedID.
ChangeTracker1.RefreshChanges()
```

Every refresh reads **all** participating controls, including those configured for manual observation. `AutomaticTracking=False` suppresses subscriptions for that control; it does not exclude its values from refreshes triggered by another control. Cached state can be stale until a refresh occurs.

### RichTextBox formatting

The default is `Text`, not `Rtf`. To include bold, font, color, bullets, and other formatting, explicitly track `Rtf` and call `RefreshChanges()` after each editor command, since text-change notifications alone are not a complete contract for formatting-only edits:

```vb
ChangeTracker1.SetTrackChanges(RtbBody, True)
ChangeTracker1.SetTrackedProperties(RtbBody, "Rtf")
ChangeTracker1.SetAutomaticTracking(RtbBody, False)
ChangeTracker1.AcceptChanges()

' After an editor command changes SelectionFont, SelectionColor, SelectionBullet, or Rtf:
ChangeTracker1.RefreshChanges()
```

Also call `RefreshChanges()` from the editor's `TextChanged` handler to observe typing, and cover keyboard formatting shortcuts, paste, undo, and redo in the application's editor workflow. Comparison is exact serialized RTF string equality, not semantic document equality.

### DateTimePicker checkbox

Tracking `Value` alone does not promise to track a checkbox toggle. When `ShowCheckBox=True` and `Checked` matters, configure `"Value, Checked"` and ensure an appropriate notification is available, or select manual observation and refresh from the application's checkbox/date workflow. A missing notification is reported at startup; it is never silently treated as supported.

## Supported values and comparisons

- Supported snapshots: `Nothing`, `DBNull.Value`, strings, primitive values, enums, `Decimal`, `DateTime`, `DateTimeOffset`, `TimeSpan`, `Guid`, `DateOnly`, and `TimeOnly`. Nullable values are boxed as their underlying scalar or `Nothing`.
- Equality uses `Object.Equals`: strings are exact and case-sensitive, whitespace matters, and numeric types are not coerced.
- `Nothing`, an empty string, and `DBNull.Value` are distinct.
- Arbitrary objects, lists, arrays, images, custom structures, and mutable object graphs are rejected. Track a stable scalar ID or serialized string instead.
- `Tag` can be tracked manually when its value is a supported scalar. A later unsupported value also fails explicitly.
- Read-only computed properties are supported if their values are supported scalars and notifications are supplied or manual refresh is used.
- Hidden and disabled controls remain included. Whether a value changed does not depend on whether the user can currently see or edit it.

There is no automatic DataGridView row/cell snapshot, checked-list collection tracking, or data-model transaction tracking. Those require a separate collection-aware design. Tracking `DataGridView.Text` would not represent its data and is not a built-in default.

## Runtime reconfiguration and lifetime

Extender configuration is locked while tracking is active so that changing settings cannot silently accept or discard an existing baseline. Finish any save/discard decision before reconfiguration:

```vb
ChangeTracker1.StopTracking()
ChangeTracker1.SetTrackedProperties(TxtName, "Text")
ChangeTracker1.SetTrackChanges(TxtName, True)
ChangeTracker1.AcceptChanges()
```

Dispose all suspension scopes before calling `StopTracking()`. Restarting establishes a new baseline for the current values. Disabled tracking configurations are retained but not evaluated. Configured controls are strongly referenced until they or the tracker are disposed. Use the form's component container to release the tracker with the form. A disposed control is removed and the aggregate is refreshed; during suspension this reconciliation waits until resume.

## Failure behavior

- Unknown properties, unknown or unsupported events, missing automatic notifications, and unsupported value types throw descriptive exceptions when tracking starts or values are read.
- A failed start releases previously attached runtime subscriptions and leaves tracking stopped. Correct the configuration and start again.
- Refresh and acceptance read all participating values before updating cached values or baselines. A getter failure leaves the previous cached state and baseline intact; the actual controls may already have changed.
- Getter failures identify the control and property and retain the original exception as `InnerException`. Getters must be side-effect free.
- All mutation and observation operations reject access from a different thread. Read-only cached properties are not a cross-thread synchronization API.
- Subscriber exceptions are not swallowed. Keep handlers short and handle application failures at the appropriate UI boundary.

## Verification

The included STA executable contains 20 dependency-free regression checks and exits with an unhandled exception/nonzero status if a check fails. It covers baseline reversions, event transitions, multiple properties, repeated start, acceptance, nested suspension, suspended loads, manual refresh, alternate events, `INotifyPropertyChanged`, rollback after failed startup, unsupported values, disposal, reentrant notifications, reconfiguration guards, thread affinity, built-in defaults, hidden/disabled controls, invalid events, atomic acceptance, and explicit RTF formatting observation.

See `VALIDATION.md` for the verification actually performed in the delivery environment and the remaining Windows build and designer checks.

## License

MIT. See `LICENSE`.
