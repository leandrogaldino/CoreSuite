# BusyOverlay

**A designer-friendly Windows Forms component that blocks any existing control with a customizable busy surface while work is in progress.**

> [!NOTE]
> BusyOverlay is one of the independent projects that make up the **CoreSuite** solution. The package contains the non-visual component, its animated run-time surface, operation scope, progress and cancellation event data, and Windows Forms designer support.

## Overview

`BusyOverlay` provides clear feedback while a Windows Forms application is loading, saving, searching, processing, or waiting for another asynchronous operation. Assign a form, panel, `DataGridView`, or another existing control to `TargetControl`; the component creates a temporary surface over exactly that area and prevents mouse and keyboard interaction with the covered content.

The component does not replace the target, move its child controls, hide them, or require a separate loader form. It can manage an asynchronous delegate for you through `RunAsync`, participate in an existing workflow through a disposable operation scope, or be controlled manually with `ShowOverlay` and `HideOverlay`.

Short asynchronous operations do not need to flash a loading interface. `OperationDisplayDelay` waits before showing the overlay, while `MinimumOperationDisplayTime` keeps an overlay that was shown visible long enough to be perceived comfortably.

## Key features

- Non-visual component displayed in the Visual Studio component tray.
- Covers a complete form or one specific control.
- Supports panels, user controls, tab pages, group boxes, `DataGridView`, and other standard Windows Forms controls.
- Blocks pointer interaction without changing the target's `Enabled` or `Visible` state.
- Optionally moves keyboard focus to the overlay and restores the previously focused control afterward.
- Rotating spinner, animated marquee bar, determinate progress bar, or no indicator.
- Primary and secondary text with independent fonts and colors.
- Optional centered content panel with configurable color, opacity, border, radius, padding, spacing, and maximum width.
- Configurable overlay color and opacity.
- Optional target snapshot beneath translucent overlays.
- Automatic positioning and resizing when the target moves, resizes, changes parent, or scrolls.
- Four `RunAsync` overloads for operations with results and cancellation tokens.
- Delayed display for fast operations.
- Minimum display time to avoid a brief visual flash.
- Optional cancellation button and cancelable `CancellationRequested` event.
- Determinate progress range, value, percentage, and detail updates.
- Reference-counted `BusyOverlayScope` for nested or externally managed operations.
- Idempotent manual `ShowOverlay` and `HideOverlay` methods.
- Separate logical busy and visual shown/hidden events.
- Designer smart tag for common association, content, appearance, and timing properties.
- Complete XML documentation and NuGet symbol generation.
- No run-time dependency on another CoreSuite package.

## Requirements

- Windows Forms
- .NET 8 for Windows (`net8.0-windows`)
- A reference to `CoreSuite.BusyOverlay`

## Installation

Install the package with the .NET CLI:

```powershell
dotnet add package CoreSuite.BusyOverlay
```

Or search for `CoreSuite.BusyOverlay` in the Visual Studio NuGet Package Manager.

When working directly with the CoreSuite solution, add `BusyOverlay/BusyOverlay.vbproj` as a project reference.

## Namespace

```vb
Imports CoreSuite.Controls
```

## Designer setup

1. Add `BusyOverlay` from the Toolbox to the form.
2. Select the component in the component tray.
3. Set `TargetControl` to the form or control that should be blocked.
4. Set `MessageText` and, when useful, `DetailText`.
5. Choose an `IndicatorStyle`.
6. Customize overlay and content panel appearance as needed.
7. Enable `AllowCancellation` only for operations that observe a cancellation token.

When the component is first added, its designer assigns the root form to `TargetControl` automatically. Change the property to a child control when only one region should be blocked.

The overlay is intentionally not shown on the design surface. It is created only at run time, so the component cannot interfere with control selection and layout in the Windows Forms Designer.

## Quick start

Assume `BusyOverlay1.TargetControl` is assigned to the current form:

```vb
Private Async Sub LoadButton_Click(sender As Object, e As EventArgs) Handles LoadButton.Click
    BusyOverlay1.MessageText = "Loading customers..."
    BusyOverlay1.DetailText = "Contacting the server"
    Await BusyOverlay1.RunAsync(
        Async Function()
            Dim customers As IReadOnlyList(Of Customer) = Await CustomerService.GetAllAsync()
            CustomersDataGridView.DataSource = customers
        End Function)
End Sub
```

`RunAsync` starts the delegate immediately. The overlay appears only when the operation lasts longer than `OperationDisplayDelay`, which is 150 milliseconds by default. It is hidden in guaranteed cleanup after success, cancellation, or failure.

Exceptions are not swallowed. The returned task completes with the original exception after the overlay state is restored.

## Cover one control

Set the target to a `DataGridView` when the rest of the form should remain available:

```vb
Private Sub ConfigureOrdersOverlay()
    OrdersBusyOverlay.TargetControl = OrdersDataGridView
    OrdersBusyOverlay.MessageText = "Refreshing orders..."
    OrdersBusyOverlay.IndicatorStyle = BusyOverlayIndicatorStyle.MarqueeBar
End Sub
```

Run the operation normally:

```vb
Private Async Sub RefreshButton_Click(sender As Object, e As EventArgs) Handles RefreshButton.Click
    Await OrdersBusyOverlay.RunAsync(
        Async Function()
            OrdersDataGridView.DataSource = Await OrderService.GetRecentAsync()
        End Function)
End Sub
```

The overlay is hosted above the target in its current parent. It follows the target when layout, docking, anchoring, or scrolling changes its bounds.

## Cancellation

Use the token-aware overload and enable the cancellation button:

```vb
Private Async Sub SynchronizeButton_Click(sender As Object, e As EventArgs) Handles SynchronizeButton.Click
    BusyOverlay1.AllowCancellation = True
    BusyOverlay1.CancelButtonText = "Cancel"
    BusyOverlay1.MessageText = "Synchronizing records..."
    Try
        Await BusyOverlay1.RunAsync(
            Async Function(cancellationToken As CancellationToken)
                Await SynchronizationService.RunAsync(cancellationToken)
            End Function)
    Catch ex As OperationCanceledException
        StatusLabel.Text = "Synchronization canceled."
    End Try
End Sub
```

Clicking the button raises `CancellationRequested` before any token is canceled. A handler can reject the request:

```vb
Private Sub BusyOverlay1_CancellationRequested(sender As Object, e As BusyOverlayCancellationEventArgs) Handles BusyOverlay1.CancellationRequested
    If MessageBox.Show("Cancel the current operation?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then e.Cancel = True
End Sub
```

The event exposes `CancellableOperationCount`. When `e.Cancel` remains `False`, every token-aware `RunAsync` operation currently owned by the component receives cancellation.

Cancellation is cooperative. The delegate must observe the token by passing it to asynchronous APIs, calling `ThrowIfCancellationRequested`, or checking `IsCancellationRequested`.

Cancellation can also be requested in code:

```vb
Dim requestAccepted As Boolean = BusyOverlay1.RequestCancellation()
```

`RequestCancellation` works even when `AllowCancellation` is `False`; that property controls presentation of the button, not programmatic cancellation.

## Determinate progress

Select `ProgressBar` and report values within the configured range:

```vb
Private Async Sub ImportButton_Click(sender As Object, e As EventArgs) Handles ImportButton.Click
    BusyOverlay1.IndicatorStyle = BusyOverlayIndicatorStyle.ProgressBar
    BusyOverlay1.ProgressMinimum = 0
    BusyOverlay1.ProgressMaximum = 100
    BusyOverlay1.ProgressValue = 0
    BusyOverlay1.ShowProgressPercentage = True
    BusyOverlay1.MessageText = "Importing products..."
    Await BusyOverlay1.RunAsync(
        Async Function(cancellationToken As CancellationToken)
            For progressValue As Integer = 0 To 100
                cancellationToken.ThrowIfCancellationRequested()
                BusyOverlay1.ReportProgress(progressValue, $"Processing batch {progressValue} of 100")
                Await Task.Delay(40, cancellationToken)
            Next
        End Function)
End Sub
```

`ReportProgress` updates `ProgressValue`, optionally replaces `DetailText`, redraws the overlay, and raises `ProgressChanged`.

The default range is 0 through 100. `ProgressMinimum` must remain below `ProgressMaximum`, and `ProgressValue` must remain within the range. Changing a bound automatically moves the current value to that bound when necessary.

## Returning a result

The generic overload returns the original result:

```vb
Private Async Function LoadCustomerAsync(customerId As Integer) As Task
    Dim customer As Customer = Await BusyOverlay1.RunAsync(
        Async Function(cancellationToken As CancellationToken) As Task(Of Customer)
            Return Await CustomerService.GetAsync(customerId, cancellationToken)
        End Function)
    DisplayCustomer(customer)
End Function
```

An equivalent generic overload is available for delegates that do not receive a cancellation token.

## Disposable operation scope

Use `BeginOperation` when another part of the application already owns execution and only needs a reliable visible scope:

```vb
Private Async Function SaveAsync() As Task
    Using busyScope As BusyOverlayScope = BusyOverlay1.BeginOperation()
        Await Repository.SaveAsync()
    End Using
End Function
```

`BeginOperation` shows immediately and increments `ActiveOperationCount`. Disposing the returned scope decrements the count. The overlay hides only after the last active scope or `RunAsync` operation has completed.

Nested scopes are supported:

```vb
Using outerScope As BusyOverlayScope = BusyOverlay1.BeginOperation()
    Await LoadHeaderAsync()
    Using innerScope As BusyOverlayScope = BusyOverlay1.BeginOperation()
        Await LoadLinesAsync()
    End Using
    Await CalculateTotalsAsync()
End Using
```

Disposing the same scope more than once has no effect.

## Manual control

Use manual control when the start and finish points cannot share a task or scope:

```vb
Private Sub StartManualWork()
    BusyOverlay1.MessageText = "Waiting for external completion..."
    BusyOverlay1.ShowOverlay()
End Sub
Private Sub FinishManualWork()
    BusyOverlay1.HideOverlay()
End Sub
```

`ShowOverlay` is idempotent: repeated calls create only one manual busy state. `HideOverlay` clears that state.

If a `RunAsync` operation or `BusyOverlayScope` is still active, calling `HideOverlay` does not hide the surface prematurely. The remaining operation continues to own the busy state.

## Indicator styles

| Value | Behavior |
| --- | --- |
| `Spinner` | Draws a rotating circular arc. This is the default. |
| `MarqueeBar` | Draws an animated segment moving across a horizontal track. |
| `ProgressBar` | Draws a determinate bar based on the configured range and value. |
| `None` | Displays text, content panel, and cancellation button without an indicator. |

The spinner uses `IndicatorSize` and `IndicatorThickness`. Both bar modes use `ProgressBarWidth` and `ProgressBarHeight`. `IndicatorColor` is the active color, while `IndicatorTrackColor` is the inactive bar track.

## Association and state properties

| Property | Type | Default | Description |
| --- | --- | --- | --- |
| `TargetControl` | `Control` | `Nothing` | Form or child control covered by the overlay. |
| `Enabled` | `Boolean` | `True` | Permits the visual overlay to be displayed. Active logical operations remain active while disabled. |
| `IsBusy` | `Boolean` | Read-only | Whether a manual, scoped, or asynchronous operation is active. |
| `IsOverlayVisible` | `Boolean` | Read-only | Whether the run-time surface is currently visible. |
| `ActiveOperationCount` | `Integer` | Read-only | Number of active scopes and `RunAsync` operations. Manual state is not included. |
| `CancellableOperationCount` | `Integer` | Read-only | Number of active token-aware `RunAsync` operations. |
| `CanCancel` | `Boolean` | Read-only | Whether the user-facing cancellation button can currently be used. |

`IsBusy` and `IsOverlayVisible` intentionally represent different states. During `OperationDisplayDelay`, an operation is busy but its surface may not yet be visible.

## Text properties

| Property | Default | Description |
| --- | --- | --- |
| `MessageText` | `Please wait...` | Primary centered message. |
| `DetailText` | Empty | Optional secondary message. |
| `MessageFont` | System message-box font | Font used for the primary message. |
| `DetailFont` | System message-box font | Font used for detail text and progress percentage. |
| `MessageForeColor` | `SystemColors.WindowText` | Primary message color. |
| `DetailForeColor` | `SystemColors.GrayText` | Detail and percentage color. |

Long messages wrap automatically within `ContentMaximumWidth` and the available target width.

## Overlay appearance

| Property | Default | Valid range | Description |
| --- | ---: | ---: | --- |
| `OverlayColor` | `SystemColors.Control` | Any non-empty color | Tint applied across the blocked target. |
| `OverlayOpacity` | `190` | 0 to 255 | Alpha used for the target tint. |
| `CaptureTarget` | `True` | Boolean | Captures the target beneath translucent tinting. |
| `UseWaitCursor` | `True` | Boolean | Displays the wait cursor over the blocked area. |
| `BlockKeyboardInput` | `True` | Boolean | Moves focus to the overlay and restores the previous control afterward. |

Target capture is best effort. Some native or hardware-rendered controls may not support `DrawToBitmap`; the component catches those failures and still displays the overlay using the configured tint.

Set `CaptureTarget` to `False` when the target is very large, changes continuously while busy, or does not need to remain visually recognizable below the tint.

## Content panel appearance

| Property | Default | Valid range | Description |
| --- | ---: | ---: | --- |
| `ShowContentPanel` | `True` | Boolean | Draws content on a centered panel. |
| `ContentBackColor` | `SystemColors.Window` | Any non-empty color | Panel background color. |
| `ContentOpacity` | `245` | 0 to 255 | Panel background alpha. |
| `ContentBorderColor` | `SystemColors.ControlDark` | Any non-empty color | Optional panel border color. |
| `ContentBorderThickness` | `0` | 0 to 10 | Panel border width; zero disables it. |
| `ContentCornerRadius` | `8` | 0 to 64 | Rounded corner radius; zero uses square corners. |
| `ContentPadding` | `20` | 0 to 64 | Space between panel edges and content. |
| `ContentSpacing` | `10` | 0 to 64 | Space between indicator, text, percentage, and button. |
| `ContentMaximumWidth` | `420` | 120 to 2000 | Maximum panel width before text wraps. |

Set `ShowContentPanel` to `False` for a minimal design that draws the indicator and text directly over the tint.

## Indicator and progress properties

| Property | Default | Valid range | Description |
| --- | ---: | ---: | --- |
| `IndicatorStyle` | `Spinner` | Enum | Selects the visual indicator. |
| `IndicatorColor` | `SystemColors.Highlight` | Any non-empty color | Spinner, marquee, and progress fill color. |
| `IndicatorTrackColor` | `SystemColors.ControlDark` | Any non-empty color | Bar track color. |
| `IndicatorSize` | `32` | 16 to 128 | Spinner diameter in pixels. |
| `IndicatorThickness` | `4` | 1 to 32 | Spinner stroke width; cannot exceed half the diameter. |
| `AnimationInterval` | `75` | 15 to 1000 | Timer interval in milliseconds. |
| `ProgressBarWidth` | `220` | 60 to 1000 | Preferred width of marquee and progress bars. |
| `ProgressBarHeight` | `8` | 2 to 64 | Height of marquee and progress bars. |
| `ProgressMinimum` | `0` | Less than maximum | Lower determinate range bound. |
| `ProgressMaximum` | `100` | Greater than minimum | Upper determinate range bound. |
| `ProgressValue` | `0` | Within range | Current determinate value. |
| `ProgressPercentage` | Calculated | Read-only | Normalized value from 0 through 100. |
| `ShowProgressPercentage` | `True` | Boolean | Draws the percentage below a determinate bar. |

## Cancellation properties

| Property | Default | Description |
| --- | --- | --- |
| `AllowCancellation` | `False` | Displays a cancel button when at least one token-aware operation is active. |
| `CancelButtonText` | `Cancel` | Text displayed by the button. |
| `CancelButtonSize` | `90, 30` | Button size; minimum 40 by 20 pixels. |

The button is hidden automatically when no cancellable operation is active, even when `AllowCancellation` remains `True`.

## Timing properties

| Property | Default | Valid range | Description |
| --- | ---: | ---: | --- |
| `OperationDisplayDelay` | `150` | 0 to 60000 ms | How long `RunAsync` waits before displaying the overlay. |
| `MinimumOperationDisplayTime` | `300` | 0 to 60000 ms | Minimum visible time after a `RunAsync` overlay appears. |

These properties apply to `RunAsync`. `ShowOverlay` and `BeginOperation` display immediately because they explicitly request a visible busy state.

## Methods

| Method | Description |
| --- | --- |
| `ShowOverlay()` | Creates an idempotent manual busy state and shows immediately. |
| `HideOverlay()` | Clears the manual state without ending other operations. |
| `BeginOperation()` | Starts a visible reference-counted operation and returns a disposable scope. |
| `RunAsync(Func(Of Task))` | Runs an asynchronous operation without a result or cancellation token. |
| `RunAsync(Func(Of CancellationToken, Task), token)` | Runs a cancellable asynchronous operation. |
| `RunAsync(Of TResult)(Func(Of Task(Of TResult)))` | Runs an asynchronous operation and returns its result. |
| `RunAsync(Of TResult)(Func(Of CancellationToken, Task(Of TResult)), token)` | Runs a cancellable asynchronous operation and returns its result. |
| `ReportProgress(value, detailText)` | Updates determinate progress and optional detail text. |
| `RequestCancellation()` | Requests cancellation of all active token-aware operations. |

## Events

| Event | Raised when |
| --- | --- |
| `BusyStarted` | Logical state changes from idle to busy. |
| `BusyEnded` | The final manual, scoped, or asynchronous operation ends. |
| `OverlayShown` | The visual surface becomes visible. |
| `OverlayHidden` | The visual surface is hidden. |
| `CancellationRequested` | Before active cancellable operations receive cancellation. |
| `ProgressChanged` | `ReportProgress` or `ProgressValue` changes the current progress. |
| `TargetControlChanged` | Target reference changes or the assigned target is disposed. |

`BusyStarted` can occur without `OverlayShown` when a `RunAsync` operation completes inside the display delay. This makes the events suitable for both state tracking and visual tracking without conflating the two.

## Focus and interaction behavior

The overlay is a real control positioned above the target. It intercepts pointer input without disabling or hiding the target.

With `BlockKeyboardInput = True`, the component records the deepest focused control, moves focus to the blocking surface, and restores the previous focus when the overlay hides. A disposed, hidden, or disabled previous control is not focused again.

Set `BlockKeyboardInput` to `False` only when keyboard input should remain available to the previously focused target while pointer interaction is blocked.

## Threading behavior

`BusyOverlay` is a Windows Forms component and its properties and methods must be accessed from the UI thread that owns `TargetControl`. Calls made from another thread after the target handle exists raise `InvalidOperationException`.

`RunAsync` does not move work to a background thread. It awaits the supplied task and keeps the UI responsive when that task is genuinely asynchronous.

For CPU-bound work, explicitly use `Task.Run` and keep UI updates on the UI thread:

```vb
Dim result As ReportResult = Await BusyOverlay1.RunAsync(
    Function(cancellationToken As CancellationToken)
        Return Task.Run(Function() ReportBuilder.Build(cancellationToken), cancellationToken)
    End Function)
DisplayReport(result)
```

`ReportProgress` should normally be called from the captured UI context. `Progress(Of T)` can be used to marshal background progress reports back to the form.

## Error and disposal behavior

- A missing target raises `InvalidOperationException` when a run-time method is called.
- A non-form target without a parent raises `InvalidOperationException`.
- A disposed target cannot be assigned.
- `TargetControl` cannot change while the component is busy.
- Invalid numeric ranges raise `ArgumentOutOfRangeException`.
- Empty colors raise `ArgumentException`.
- A delegate that returns `Nothing` instead of a task raises `InvalidOperationException`.
- Operation exceptions and `OperationCanceledException` propagate to the caller.
- Disposing the component requests cancellation for active cancellable operations, detaches target events, releases the target snapshot, and destroys the overlay surface.

## BusyOverlay and AsyncLoader

Both components communicate that work is in progress, but they solve different interface problems.

| Capability | `BusyOverlay` | `AsyncLoader` |
| --- | --- | --- |
| Visual content | Built-in spinner, bars, text, panel, and cancel button | Separate application-provided loader form |
| Target | Any form or child control | Parent form |
| Covers one panel or grid | Yes | No; form-oriented |
| Executes/awaits operation | Built-in `RunAsync` overloads | Start and stop are controlled separately |
| Determinate progress | Built in | Implemented by the custom loader form |
| Cancellation | Built in | Implemented by the custom loader form/application |
| Nested operation count | Built in | Not its primary model |
| Best use | Consistent lightweight busy feedback | Fully custom loader windows with arbitrary controls |

Use `BusyOverlay` when the application needs a consistent blocking surface with minimal setup. Use `AsyncLoader` when the loading experience itself is a custom form with specialized controls or animation.

## Package information

| Item | Value |
| --- | --- |
| Package | `CoreSuite.BusyOverlay` |
| Namespace | `CoreSuite.Controls` |
| Assembly | `CoreSuite.BusyOverlay` |
| Target framework | `net8.0-windows` |
| UI framework | Windows Forms |
| Run-time CoreSuite dependencies | None |

## License

This package is distributed under the MIT license used by the CoreSuite repository.

## Repository

[CoreSuite on GitHub](https://github.com/leandrogaldino/CoreSuite)
