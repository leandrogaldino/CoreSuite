Imports System.Runtime.InteropServices
Partial Public Class BusyOverlay
    Private Sub SetTargetControl(Value As Control)
        If ReferenceEquals(_TargetControl, Value) Then Return
        If Value IsNot Nothing AndAlso Value.IsDisposed Then Throw New ObjectDisposedException(Value.Name, "A disposed control cannot be assigned as the busy overlay target.")
        If IsBusy Then Throw New InvalidOperationException("TargetControl cannot be changed while the component is busy.")
        DetachTargetControl()
        _TargetControl = Value
        AttachTargetControl()
        OnTargetControlChanged()
    End Sub
    Private Sub AttachTargetControl()
        If _TargetControl Is Nothing Then Return
        AddHandler _TargetControl.ParentChanged, AddressOf TargetControl_ParentChanged
        AddHandler _TargetControl.LocationChanged, AddressOf TargetControl_BoundsChanged
        AddHandler _TargetControl.SizeChanged, AddressOf TargetControl_BoundsChanged
        AddHandler _TargetControl.VisibleChanged, AddressOf TargetControl_StateChanged
        AddHandler _TargetControl.Disposed, AddressOf TargetControl_Disposed
        ObserveScrollableParent()
    End Sub
    Private Sub DetachTargetControl()
        HideOverlayView()
        StopObservingScrollableParent()
        If _TargetControl Is Nothing Then Return
        RemoveHandler _TargetControl.ParentChanged, AddressOf TargetControl_ParentChanged
        RemoveHandler _TargetControl.LocationChanged, AddressOf TargetControl_BoundsChanged
        RemoveHandler _TargetControl.SizeChanged, AddressOf TargetControl_BoundsChanged
        RemoveHandler _TargetControl.VisibleChanged, AddressOf TargetControl_StateChanged
        RemoveHandler _TargetControl.Disposed, AddressOf TargetControl_Disposed
        _TargetControl = Nothing
    End Sub
    Private Sub ObserveScrollableParent()
        StopObservingScrollableParent()
        If _TargetControl Is Nothing OrElse TypeOf _TargetControl Is Form Then Return
        _ObservedScrollableParent = TryCast(_TargetControl.Parent, ScrollableControl)
        If _ObservedScrollableParent IsNot Nothing Then AddHandler _ObservedScrollableParent.Scroll, AddressOf TargetParent_Scroll
    End Sub
    Private Sub StopObservingScrollableParent()
        If _ObservedScrollableParent Is Nothing Then Return
        RemoveHandler _ObservedScrollableParent.Scroll, AddressOf TargetParent_Scroll
        _ObservedScrollableParent = Nothing
    End Sub
    Private Sub TargetControl_ParentChanged(sender As Object, e As EventArgs)
        ObserveScrollableParent()
        AttachOverlayViewToTarget()
        SynchronizeOverlayVisibility()
    End Sub
    Private Sub TargetControl_BoundsChanged(sender As Object, e As EventArgs)
        UpdateOverlayBounds()
        If IsOverlayVisible Then CaptureTargetSnapshot()
    End Sub
    Private Sub TargetControl_StateChanged(sender As Object, e As EventArgs)
        SynchronizeOverlayVisibility()
    End Sub
    Private Sub TargetControl_Disposed(sender As Object, e As EventArgs)
        DetachTargetControl()
        DisposeOverlayView()
        OnTargetControlChanged()
    End Sub
    Private Sub TargetParent_Scroll(sender As Object, e As ScrollEventArgs)
        UpdateOverlayBounds()
    End Sub
    Private Sub EnsureOverlayView()
        If _View IsNot Nothing AndAlso Not _View.IsDisposed Then Return
        _View = New BusyOverlayView(Me)
        AddHandler _View.CancellationClick, AddressOf OverlayView_CancellationClick
        AttachOverlayViewToTarget()
    End Sub
    Private Sub AttachOverlayViewToTarget()
        If _View Is Nothing OrElse _View.IsDisposed Then Return
        Dim HostControl As Control = GetOverlayHost()
        If HostControl Is Nothing OrElse HostControl.IsDisposed Then
            If _View.Parent IsNot Nothing Then _View.Parent.Controls.Remove(_View)
            Return
        End If
        If Not ReferenceEquals(_View.Parent, HostControl) Then
            If _View.Parent IsNot Nothing Then _View.Parent.Controls.Remove(_View)
            HostControl.Controls.Add(_View)
        End If
        UpdateOverlayBounds()
    End Sub
    Private Function GetOverlayHost() As Control
        If _TargetControl Is Nothing OrElse _TargetControl.IsDisposed Then Return Nothing
        If TypeOf _TargetControl Is Form Then Return _TargetControl
        Return _TargetControl.Parent
    End Function
    Private Sub UpdateOverlayBounds()
        If _View Is Nothing OrElse _View.IsDisposed OrElse _TargetControl Is Nothing OrElse _TargetControl.IsDisposed Then Return
        If TypeOf _TargetControl Is Form Then
            _View.Bounds = _TargetControl.ClientRectangle
        ElseIf _TargetControl.Parent IsNot Nothing AndAlso ReferenceEquals(_View.Parent, _TargetControl.Parent) Then
            _View.Bounds = _TargetControl.Bounds
        End If
    End Sub
    Private Sub SynchronizeOverlayVisibility()
        If _IsDisposed OrElse IsInDesignMode Then Return
        If _Enabled AndAlso IsBusy AndAlso (_ManualBusy OrElse _OperationDisplayRequested) AndAlso CanDisplayOverlay() Then
            ShowOverlayView()
        Else
            HideOverlayView()
        End If
    End Sub
    Private Function CanDisplayOverlay() As Boolean
        If _TargetControl Is Nothing OrElse _TargetControl.IsDisposed OrElse Not _TargetControl.Visible Then Return False
        Return TypeOf _TargetControl Is Form OrElse _TargetControl.Parent IsNot Nothing
    End Function
    Private Sub ShowOverlayView()
        If _OverlaySurfaceShown Then
            UpdateOverlayBounds()
            _View.BringToFront()
            Return
        End If
        EnsureOverlayView()
        AttachOverlayViewToTarget()
        If _View.Parent Is Nothing Then Return
        _RestoreFocusControl = GetDeepestFocusedControl()
        CaptureTargetSnapshot()
        _View.ApplySettings()
        _View.Visible = True
        _View.BringToFront()
        _OverlaySurfaceShown = True
        _OverlayShownAt = DateTimeOffset.UtcNow
        If _BlockKeyboardInput Then FocusOverlay()
        OnOverlayShown()
    End Sub
    Private Sub HideOverlayView()
        If _View Is Nothing OrElse _View.IsDisposed Then Return
        Dim WasShown As Boolean = _OverlaySurfaceShown
        _View.Visible = False
        _OverlaySurfaceShown = False
        _View.ClearSnapshot()
        _OverlayShownAt = Nothing
        If Not WasShown Then Return
        RestorePreviousFocus()
        OnOverlayHidden()
    End Sub
    Private Sub FocusOverlay()
        If _View Is Nothing OrElse _View.IsDisposed OrElse Not _View.Visible Then Return
        _View.Select()
        _View.Focus()
    End Sub
    Private Function GetDeepestFocusedControl() As Control
        If _TargetControl Is Nothing OrElse Not _TargetControl.ContainsFocus Then Return Nothing
        Dim CurrentControl As Control = _TargetControl.FindForm()
        Dim CurrentContainer As ContainerControl = TryCast(CurrentControl, ContainerControl)
        While CurrentContainer IsNot Nothing AndAlso CurrentContainer.ActiveControl IsNot Nothing
            CurrentControl = CurrentContainer.ActiveControl
            CurrentContainer = TryCast(CurrentControl, ContainerControl)
        End While
        Return CurrentControl
    End Function
    Private Sub RestorePreviousFocus()
        Dim ControlToFocus As Control = _RestoreFocusControl
        _RestoreFocusControl = Nothing
        If ControlToFocus Is Nothing OrElse ControlToFocus.IsDisposed OrElse Not ControlToFocus.Visible OrElse Not ControlToFocus.Enabled OrElse Not ControlToFocus.CanFocus Then Return
        ControlToFocus.Focus()
    End Sub
    Private Sub CaptureTargetSnapshot()
        If _View Is Nothing OrElse _View.IsDisposed Then Return
        If Not _CaptureTarget OrElse _TargetControl Is Nothing OrElse _TargetControl.IsDisposed OrElse _View.Width <= 0 OrElse _View.Height <= 0 Then
            _View.ClearSnapshot()
            Return
        End If
        Dim Snapshot As Bitmap = Nothing
        Dim ViewWasVisible As Boolean = _View.Visible
        Try
            _View.Visible = False
            Snapshot = CreateTargetSnapshot()
            _View.SetSnapshot(Snapshot)
            Snapshot = Nothing
        Catch ex As ArgumentException
            _View.ClearSnapshot()
        Catch ex As InvalidOperationException
            _View.ClearSnapshot()
        Catch ex As ExternalException
            _View.ClearSnapshot()
        Finally
            If Snapshot IsNot Nothing Then Snapshot.Dispose()
            _View.Visible = ViewWasVisible
        End Try
    End Sub
    Private Function CreateTargetSnapshot() As Bitmap
        Dim TargetForm As Form = TryCast(_TargetControl, Form)
        If TargetForm IsNot Nothing Then Return CreateFormClientSnapshot(TargetForm)
        Dim Snapshot As New Bitmap(_View.Width, _View.Height)
        Try
            _TargetControl.DrawToBitmap(Snapshot, New Rectangle(Point.Empty, Snapshot.Size))
            Return Snapshot
        Catch
            Snapshot.Dispose()
            Throw
        End Try
    End Function
    Private Function CreateFormClientSnapshot(TargetForm As Form) As Bitmap
        Using FormSnapshot As New Bitmap(TargetForm.Width, TargetForm.Height)
            TargetForm.DrawToBitmap(FormSnapshot, New Rectangle(Point.Empty, TargetForm.Size))
            Dim ClientOrigin As Point = TargetForm.PointToScreen(Point.Empty)
            Dim ClientOffset As New Point(ClientOrigin.X - TargetForm.Bounds.Left, ClientOrigin.Y - TargetForm.Bounds.Top)
            Dim ClientSnapshot As New Bitmap(_View.Width, _View.Height)
            Try
                Using SnapshotGraphics As Graphics = Graphics.FromImage(ClientSnapshot)
                    SnapshotGraphics.DrawImageUnscaled(FormSnapshot, -ClientOffset.X, -ClientOffset.Y)
                End Using
                Return ClientSnapshot
            Catch
                ClientSnapshot.Dispose()
                Throw
            End Try
        End Using
    End Function
    Private Sub RefreshOverlayAppearance()
        If _View Is Nothing OrElse _View.IsDisposed Then Return
        _View.ApplySettings()
    End Sub
    Private Sub DisposeOverlayView()
        If _View Is Nothing Then Return
        RemoveHandler _View.CancellationClick, AddressOf OverlayView_CancellationClick
        If _View.Parent IsNot Nothing Then _View.Parent.Controls.Remove(_View)
        _View.Dispose()
        _View = Nothing
        _OverlaySurfaceShown = False
    End Sub
End Class
