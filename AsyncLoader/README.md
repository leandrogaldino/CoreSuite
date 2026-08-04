# CoreSuite.AsyncLoader

**A Windows Forms loading overlay that keeps the interface responsive while asynchronous work is running.**

> [!NOTE]
> `CoreSuite.AsyncLoader` is one of the libraries included in the **CoreSuite** solution. It targets .NET 8 for Windows, uses Windows Forms and depends on `CoreSuite.Helpers.Windows`.

## Overview

`CoreSuite.AsyncLoader` displays a custom loader form centered over a parent form while an operation is in progress. It can either hide the parent's controls behind a solid overlay or leave them visible and temporarily disable them.

The service also temporarily removes the parent's minimize, maximize and close commands, keeps the loader centered when the parent moves or resizes, constrains an oversized loader to the available client area and restores the original parent state when loading ends.

The package does not execute the application operation itself. Your code starts the loader, awaits the real asynchronous work and stops the loader in a `Finally` block.

## Features

- Displays any custom Windows Form as the loading interface.
- Centers the loader over its parent form.
- Repositions the loader when the parent moves or resizes.
- Offers covered and visible-but-disabled parent modes.
- Preserves and restores the original visibility or enabled state of child controls.
- Temporarily disables the parent's control box and resizing commands.
- Supports a configurable solid overlay color.
- Supports rounded loader corners.
- Constrains the loader to the parent form's available client area.
- Prevents the loader form from being closed independently while it is active.
- Exposes the current state through `IsRunning`.
- Provides optional start and stop delays.
- Includes English XML documentation for the public API.

## Requirements

- .NET 8 for Windows (`net8.0-windows`)
- Windows Forms
- Windows operating system
- `CoreSuite.Helpers.Windows` dependency, installed automatically by NuGet

## Installation

Install the package from NuGet:

```powershell
dotnet add package CoreSuite.AsyncLoader
```

Or use the Visual Studio NuGet Package Manager and search for:

```text
CoreSuite.AsyncLoader
```

## Namespace

Import the service namespace:

```vb
Imports CoreSuite.Services
```

## Quick start

Create a borderless-looking loader form in the Windows Forms designer. It can contain an animated image, progress text or any other visual content. `AsyncLoader` applies the required border and ownership settings at runtime.

The following button handler shows the loader while a background operation runs:

```vb
Imports CoreSuite.Services

Private Async Sub ImportButton_Click(sender As Object, e As EventArgs) Handles ImportButton.Click
    Dim loaderForm As New LoadingForm()
    Dim loadingService As New AsyncLoader(Me, loaderForm, 20, True, Color.White)

    Try
        Await loadingService.Start(0)
        Await Task.Run(AddressOf ImportLargeFile)
    Finally
        If loadingService.IsRunning Then
            Await loadingService.Stop()
        End If
    End Try
End Sub
```

The `Finally` block is important: it restores the parent form even when the operation fails.

> [!IMPORTANT]
> Keep UI work on the Windows Forms thread. Move CPU-intensive or synchronous blocking work to a background task, or call a genuinely asynchronous API. Showing a loader does not make synchronous UI-thread work asynchronous.

## Constructor

```vb
Dim loadingService As New AsyncLoader(container, loader, borderRadius, coverParent, backColor)
```

| Parameter | Description |
| --- | --- |
| `container` | Parent form whose interface is temporarily blocked. |
| `loader` | Custom form displayed over the parent. |
| `borderRadius` | Diameter used to construct the loader's rounded corner arcs. Negative values are normalized to `0`. |
| `coverParent` | `True` to hide the parent's child controls behind a solid overlay; `False` to keep the controls visible but disabled. |
| `backColor` | Color applied to the parent and overlay while covered mode is active. |

## Display modes

### Cover the parent

Use `CoverParent = True` when the parent content should disappear completely while loading:

```vb
Dim loadingService As New AsyncLoader(Me, New LoadingForm(), 24, True, Color.WhiteSmoke)
```

In this mode, the service:

- records the `Visible` state of every child control;
- hides those controls;
- adds a dock-filled overlay panel;
- temporarily changes the parent background to `BackColor`;
- restores every recorded visibility value when the loader closes.

`BackColor` must be an opaque or translucent concrete color. `Color.Transparent` is rejected while covered mode is enabled.

### Keep the parent visible

Use `CoverParent = False` when users should still see the existing interface:

```vb
Dim loadingService As New AsyncLoader(Me, New LoadingForm(), 16, False, Color.Transparent)
```

In this mode, the service records and disables the parent controls instead of hiding them. Their original `Enabled` states are restored when loading ends.

## Starting and stopping

### Start

```vb
Await loadingService.Start()
```

The default start delay is `1000` milliseconds. This delay keeps the method incomplete for the selected interval after the loader is displayed; it is useful when a transition should remain visible for a minimum time.

To display it immediately:

```vb
Await loadingService.Start(0)
```

### Stop

```vb
Await loadingService.Stop()
```

The default stop delay is `0`. A delay can keep the loader visible briefly before closing it:

```vb
Await loadingService.Stop(300)
```

`Stop` closes and disposes the loader form, restores the parent interface and sets the service's `Loader` property to `Nothing`.

> [!NOTE]
> Because the loader form is disposed by `Stop`, create a new loader form before starting a later operation. An `AsyncLoader` instance is intended for one active loader lifecycle at a time.

## Properties

| Property | Type | Description |
| --- | --- | --- |
| `Container` | `Form` | Parent form that will be blocked while loading. |
| `Loader` | `Form` | Child form used as the loading interface. It becomes `Nothing` after `Stop`. |
| `CoverParent` | `Boolean` | Selects covered mode or visible-but-disabled mode. |
| `BackColor` | `Color` | Background used by the overlay and parent in covered mode. Default: `Color.White`. |
| `BorderRadius` | `Integer` | Rounded corner setting for the loader. Negative values become `0`. |
| `IsRunning` | `Boolean` | Read-only value indicating whether the loader has been started and not yet stopped. |

## Loader form behavior

When `Start` runs, the loader form is configured as follows:

- the parent form becomes its owner;
- `ShowInTaskbar` is set to `False`;
- `ShowIcon` is set to `False`;
- `FormBorderStyle` is set to `None`;
- double buffering is enabled;
- its size is limited to the parent's client size minus a 40-pixel margin;
- its current constrained size becomes its minimum size;
- attempts to close it directly are canceled.

Call `Stop` to close it through the supported lifecycle.

## Complete example with cancellation

`AsyncLoader` does not own the operation's cancellation. Supply a `CancellationToken` to the operation itself:

```vb
Private operationCancellation As CancellationTokenSource

Private Async Sub SynchronizeButton_Click(sender As Object, e As EventArgs) Handles SynchronizeButton.Click
    operationCancellation = New CancellationTokenSource()
    Dim loadingService As New AsyncLoader(Me, New LoadingForm(), 20, True, Color.White)

    Try
        Await loadingService.Start(0)
        Await SynchronizeAsync(operationCancellation.Token)
    Catch ex As OperationCanceledException
        StatusLabel.Text = "Synchronization canceled."
    Finally
        If loadingService.IsRunning Then
            Await loadingService.Stop()
        End If
        operationCancellation.Dispose()
        operationCancellation = Nothing
    End Try
End Sub

Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CancelButton.Click
    operationCancellation?.Cancel()
End Sub
```

## API reference

| Member | Description |
| --- | --- |
| `New(container, loader, borderRadius, coverParent, backColor)` | Creates a loader service for a parent and loader form. |
| `Start(delay)` | Displays the loader and blocks or covers the parent. Default delay: `1000` ms. |
| `Stop(delay)` | Closes and disposes the loader, then restores the parent. Default delay: `0` ms. |

## Usage notes

- Call `Start` and `Stop` from the Windows Forms UI thread.
- Always stop the loader in a `Finally` block.
- Do not call `Start` more than once concurrently on the same instance.
- The service restores the child-control state captured when `Start` runs.
- Changes made directly to those same `Visible` or `Enabled` values while loading may be replaced by the recorded state during restoration.
- The service is a visual and interaction layer; it does not report operation progress by itself.
- The package is Windows-specific because it uses Windows Forms.

## License

This package is licensed under the MIT License.
