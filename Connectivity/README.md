# Connectivity

**A lightweight internet availability checker and connection monitor for .NET 8, included in CoreSuite.**

> [!NOTE]
> Connectivity is one of the independent projects that make up the **CoreSuite** solution. The package contains the connectivity service and the event data used to report status changes.

## Overview

`Connectivity` checks whether at least one of several well-known HTTPS endpoints can be reached. It can perform a single synchronous or asynchronous check, or remain active in the background and notify the application whenever the detected connection state changes.

The service is suitable for availability indicators, offline-mode transitions, retry coordination, synchronization scheduling, and other scenarios that need a practical internet reachability signal. A positive result confirms that one tested endpoint responded successfully; it does not guarantee that every remote service used by the application is available.

## Key features

- Performs synchronous and asynchronous availability checks.
- Monitors connectivity continuously in the background.
- Uses several independent HTTPS endpoints instead of relying on a single host.
- Considers the internet available as soon as one endpoint responds successfully.
- Supports a configurable interval between monitoring checks.
- Raises general status-change, connected, and disconnected events.
- Exposes the current monitoring state and last stored connection state.
- Uses a shared `HttpClient` with a three-second request timeout.
- Falls back from `HEAD` to `GET` when an endpoint rejects `HEAD` requests.
- Stops active monitoring through `IDisposable`.
- Has no external package dependencies.
- Includes XML documentation and NuGet symbol generation.

## Requirements

- .NET 8 (`net8.0`) or a compatible target framework
- A reference to `CoreSuite.Connectivity`

The service has no runtime dependency on another CoreSuite package and can be used by console, desktop, web, worker, and service applications.

## Installation

```powershell
dotnet add package CoreSuite.Connectivity
```

Or add `Connectivity/Connectivity.vbproj` as a project reference when working directly with the CoreSuite solution.

## Quick start: check once

Import the service namespace and prefer the asynchronous API when the caller can use `Await`.

```vb
Imports CoreSuite.Services
Public Async Function CheckInternetAsync() As Task
    Using connectivityService As New Connectivity()
        Dim isOnline As Boolean = Await connectivityService.IsAvailableAsync()
        Console.WriteLine($"Internet available: {isOnline}")
    End Using
End Function
```

`IsAvailableAsync()` performs a fresh check each time it is called and returns `True` as soon as one configured endpoint responds successfully.

### Synchronous check

Use `IsAvailable()` only when blocking the current thread is acceptable.

```vb
Imports CoreSuite.Services
Using connectivityService As New Connectivity()
    Dim isOnline As Boolean = connectivityService.IsAvailable()
End Using
```

In UI applications, prefer `IsAvailableAsync()` so the interface remains responsive while network requests are in progress.

## Continuous monitoring

Configure `MonitoringInterval`, subscribe to the desired events, and call `StartMonitoring()`.

```vb
Imports CoreSuite.Services
Public Class MainForm
    Private ReadOnly connectivityService As New Connectivity()
    Public Sub New()
        InitializeComponent()
        connectivityService.MonitoringInterval = TimeSpan.FromSeconds(2)
        AddHandler connectivityService.ConnectivityChanged, AddressOf ConnectivityService_ConnectivityChanged
        AddHandler connectivityService.Connected, AddressOf ConnectivityService_Connected
        AddHandler connectivityService.Disconnected, AddressOf ConnectivityService_Disconnected
        connectivityService.StartMonitoring()
    End Sub
    Private Sub ConnectivityService_ConnectivityChanged(sender As Object, e As ConnectivityEventArgs)
        Debug.WriteLine($"Internet available: {e.InternetAvailable}")
    End Sub
    Private Sub ConnectivityService_Connected(sender As Object, e As EventArgs)
        Debug.WriteLine("Connected")
    End Sub
    Private Sub ConnectivityService_Disconnected(sender As Object, e As EventArgs)
        Debug.WriteLine("Disconnected")
    End Sub
    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        connectivityService.Dispose()
        MyBase.OnFormClosed(e)
    End Sub
End Class
```

Calling `StartMonitoring()` while monitoring is already active has no effect. The first check establishes the comparison baseline; events are raised only when a later check detects a different state.

## Updating a Windows Forms interface

Monitoring continues outside the UI thread, so event handlers must marshal interface changes back to the form when required.

```vb
Private Sub ConnectivityService_ConnectivityChanged(sender As Object, e As ConnectivityEventArgs)
    If InvokeRequired Then
        BeginInvoke(Sub() UpdateConnectivityStatus(e.InternetAvailable))
        Return
    End If
    UpdateConnectivityStatus(e.InternetAvailable)
End Sub
Private Sub UpdateConnectivityStatus(isOnline As Boolean)
    StatusLabel.Text = If(isOnline, "Online", "Offline")
    StatusLabel.ForeColor = If(isOnline, Color.Green, Color.Firebrick)
End Sub
```

The service does not capture a Windows Forms, WPF, ASP.NET, or other synchronization context.

## Main properties

| Property | Default | Description |
|---|---:|---|
| `MonitoringInterval` | `1 second` | Delay between checks while continuous monitoring is active. |
| `IsMonitoring` | Read-only | Indicates whether the monitoring loop is currently active. |
| `IsConnected` | Read-only | Returns the connection state retained by monitoring after a status has been recorded. |

`MonitoringInterval` should be a positive `TimeSpan`. Very short intervals can generate unnecessary requests and should be avoided.

## Methods

| Method | Description |
|---|---|
| `IsAvailableAsync()` | Performs a fresh asynchronous availability check. |
| `IsAvailable()` | Performs a fresh synchronous availability check and blocks until it completes. |
| `StartMonitoring()` | Starts the background monitoring loop if it is not already active. |
| `StopMonitoring()` | Stops monitoring and cancels its current delay or request. |
| `Dispose()` | Stops monitoring and suppresses finalization for the instance. |

## Events

| Event | Description |
|---|---|
| `ConnectivityChanged` | Raised when the monitored state changes and supplies the new state through `ConnectivityEventArgs`. |
| `Connected` | Raised after the monitored state changes from unavailable to available. |
| `Disconnected` | Raised after the monitored state changes from available to unavailable. |

All three events describe transitions detected by continuous monitoring. One-time calls to `IsAvailable()` and `IsAvailableAsync()` return their results directly and do not raise these events.

## ConnectivityEventArgs

| Property | Description |
|---|---|
| `InternetAvailable` | Indicates whether internet access was available for the detected transition. |

```vb
Private Sub ConnectivityService_ConnectivityChanged(sender As Object, e As ConnectivityEventArgs)
    If e.InternetAvailable Then
        Debug.WriteLine("The connection is available again.")
    Else
        Debug.WriteLine("The connection is unavailable.")
    End If
End Sub
```

## Endpoint-checking behavior

For each configured HTTPS address, the service:

1. sends a `HEAD` request;
2. returns `True` immediately when the response is successful;
3. retries the same address with `GET` when the server returns `405 Method Not Allowed`;
4. continues with the next address after an unsuccessful response or network error;
5. returns `False` when no address responds successfully.

The shared HTTP client applies a three-second timeout to each request. Because the service tests endpoints sequentially, a completely unavailable connection can take longer than three seconds to report when several requests must time out.

## Starting and stopping monitoring

Monitoring can be paused and started again with the same instance.

```vb
connectivityService.StartMonitoring()
' Application continues running...
connectivityService.StopMonitoring()
' Monitoring can be started again later.
connectivityService.StartMonitoring()
```

`StopMonitoring()` cancels the current monitoring operation and releases its cancellation source. Dispose the instance when its owning application component shuts down.

## Monitoring notes

- Connectivity means that at least one tested public endpoint responded successfully.
- A positive result does not confirm DNS, authentication, authorization, or availability of the application's own API.
- The first monitoring check establishes a baseline and does not raise `ConnectivityChanged`, `Connected`, or `Disconnected`.
- Later events are raised only when the result differs from the preceding monitored result.
- Event handlers may execute on a background thread.
- `IsConnected` represents stored monitoring state; use `IsAvailableAsync()` when a fresh result is required.
- Network and HTTP failures are treated as unsuccessful endpoint checks instead of being exposed as connectivity exceptions.
- Stopping or disposing the service cancels the active monitoring operation.
- Long-lived instances should be disposed by their owner.

## License

CoreSuite is licensed under the MIT License.
