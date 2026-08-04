Imports System.ComponentModel
Partial Public Class BusyOverlay
    ''' <summary>
    ''' Occurs when the component changes from idle to busy.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Occurs when the component changes from idle to busy.")>
    Public Event BusyStarted As EventHandler
    ''' <summary>
    ''' Occurs when the last manual, scoped, or asynchronous operation completes.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Occurs when the component changes from busy to idle.")>
    Public Event BusyEnded As EventHandler
    ''' <summary>
    ''' Occurs after the run-time overlay surface becomes visible.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Occurs after the run-time overlay surface becomes visible.")>
    Public Event OverlayShown As EventHandler
    ''' <summary>
    ''' Occurs after the run-time overlay surface is hidden.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Occurs after the run-time overlay surface is hidden.")>
    Public Event OverlayHidden As EventHandler
    ''' <summary>
    ''' Occurs before active cancellable operations receive cancellation.
    ''' </summary>
    ''' <remarks>Set <see cref="BusyOverlayCancellationEventArgs.Cancel"/> to <see langword="True"/> to reject the request.</remarks>
    <Category(CategoryName)>
    <Description("Occurs before active cancellable operations receive cancellation.")>
    Public Event CancellationRequested As EventHandler(Of BusyOverlayCancellationEventArgs)
    ''' <summary>
    ''' Occurs after <see cref="ReportProgress"/> changes the progress value or detail text.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Occurs after ReportProgress changes the progress value or detail text.")>
    Public Event ProgressChanged As EventHandler(Of BusyOverlayProgressChangedEventArgs)
    ''' <summary>
    ''' Occurs after the <see cref="TargetControl"/> reference changes or the target is disposed.
    ''' </summary>
    <Category(CategoryName)>
    <Description("Occurs after the target control reference changes.")>
    Public Event TargetControlChanged As EventHandler
    Private Sub OnBusyStarted()
        RaiseEvent BusyStarted(Me, EventArgs.Empty)
    End Sub
    Private Sub OnBusyEnded()
        RaiseEvent BusyEnded(Me, EventArgs.Empty)
    End Sub
    Private Sub OnOverlayShown()
        RaiseEvent OverlayShown(Me, EventArgs.Empty)
    End Sub
    Private Sub OnOverlayHidden()
        RaiseEvent OverlayHidden(Me, EventArgs.Empty)
    End Sub
    Private Sub OnProgressChanged()
        RaiseEvent ProgressChanged(Me, New BusyOverlayProgressChangedEventArgs(_ProgressValue, ProgressPercentage, _DetailText))
    End Sub
    Private Sub OnTargetControlChanged()
        RaiseEvent TargetControlChanged(Me, EventArgs.Empty)
    End Sub
End Class
