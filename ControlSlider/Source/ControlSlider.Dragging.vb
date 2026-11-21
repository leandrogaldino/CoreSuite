Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

Public Partial Class ControlSlider
    Private ReadOnly _HookedControls As New List(Of Control)()
    Private _MouseSource As Control
    Private _MouseIsDown As Boolean
    Private _IsDragging As Boolean
    Private _MouseDownScreen As Point
    Private _StartChildLocation As Point

    Private Sub RefreshInputHandlers()
        CancelCurrentOperation()
        DetachInputHandlers()

        If _Child Is Nothing OrElse _Child.IsDisposed Then Return
        Dim Target As Control = If(_DragHandle, _Child)
        If Target Is Nothing OrElse Target.IsDisposed OrElse Not IsDescendantOf(Target, _Child) Then Return

        AttachHandleTree(Target, True)
    End Sub

    Private Sub AttachHandleTree(Target As Control, IsRoot As Boolean)
        If Target.IsDisposed Then Return
        If Not IsRoot AndAlso IsInteractiveControl(Target) Then Return

        _HookedControls.Add(Target)
        AddHandler Target.MouseDown, AddressOf InputMouseDown
        AddHandler Target.MouseMove, AddressOf InputMouseMove
        AddHandler Target.MouseUp, AddressOf InputMouseUp
        AddHandler Target.MouseCaptureChanged, AddressOf InputMouseCaptureChanged

        If _DragHandle IsNot Nothing AndAlso _IncludeHandleChildren Then
            AddHandler Target.ControlAdded, AddressOf HandleControlsChanged
            AddHandler Target.ControlRemoved, AddressOf HandleControlsChanged
            For Each nested As Control In Target.Controls
                AttachHandleTree(nested, False)
            Next
        End If
    End Sub

    Private Sub DetachInputHandlers()
        For Each Target As Control In _HookedControls
            RemoveHandler Target.MouseDown, AddressOf InputMouseDown
            RemoveHandler Target.MouseMove, AddressOf InputMouseMove
            RemoveHandler Target.MouseUp, AddressOf InputMouseUp
            RemoveHandler Target.MouseCaptureChanged, AddressOf InputMouseCaptureChanged
            RemoveHandler Target.ControlAdded, AddressOf HandleControlsChanged
            RemoveHandler Target.ControlRemoved, AddressOf HandleControlsChanged
        Next
        _HookedControls.Clear()
    End Sub

    Private Sub HandleControlsChanged(sender As Object, e As ControlEventArgs)
        RefreshInputHandlers()
    End Sub

    Private Function CanBeginDragging() As Boolean
        Return _Enabled AndAlso
            _Parent IsNot Nothing AndAlso Not _Parent.IsDisposed AndAlso
            _Child IsNot Nothing AndAlso Not _Child.IsDisposed AndAlso
            _Child.Parent IsNot Nothing AndAlso IsDescendantOf(_Child.Parent, _Parent) AndAlso
            (Site Is Nothing OrElse Not Site.DesignMode) AndAlso
            LicenseManager.UsageMode <> LicenseUsageMode.Designtime
    End Function

    Private Shared Function IsDescendantOf(Target As Control, Ancestor As Control) As Boolean
        Dim Current As Control = Target
        While Current IsNot Nothing
            If ReferenceEquals(Current, Ancestor) Then Return True
            Current = Current.Parent
        End While
        Return False
    End Function

    Private Shared Function IsInteractiveControl(Target As Control) As Boolean
        Return TypeOf Target Is ButtonBase OrElse
            TypeOf Target Is TextBoxBase OrElse
            TypeOf Target Is ComboBox OrElse
            TypeOf Target Is ListControl OrElse
            TypeOf Target Is UpDownBase OrElse
            TypeOf Target Is DateTimePicker OrElse
            TypeOf Target Is TrackBar OrElse
            TypeOf Target Is ScrollBar OrElse
            TypeOf Target Is LinkLabel
    End Function

    Private Sub InputMouseDown(sender As Object, e As MouseEventArgs)
        If e.Button <> MouseButtons.Left OrElse Not CanBeginDragging() Then Return

        If IsSnapAnimating Then FinishSnapAnimation()
        If Not CanBeginDragging() Then Return
        CancelGesture()

        _MouseSource = DirectCast(sender, Control)
        _MouseDownScreen = _MouseSource.PointToScreen(e.Location)
        _StartChildLocation = _Child.Location
        _MouseIsDown = True
        _MouseSource.Capture = True
    End Sub

    Private Sub InputMouseMove(sender As Object, e As MouseEventArgs)
        If Not _MouseIsDown OrElse _MouseSource Is Nothing Then Return
        If Not CanBeginDragging() Then
            CancelGesture()
            Return
        End If

        If (e.Button And MouseButtons.Left) <> MouseButtons.Left Then
            EndGesture()
            Return
        End If

        Dim CursorScreen As Point = DirectCast(sender, Control).PointToScreen(e.Location)
        Dim DeltaX As Long = CLng(CursorScreen.X) - _MouseDownScreen.X
        Dim DeltaY As Long = CLng(CursorScreen.Y) - _MouseDownScreen.Y
        If _MovementAxis = SliderMovementAxis.Vertical Then DeltaX = 0
        If _MovementAxis = SliderMovementAxis.Horizontal Then DeltaY = 0

        If Not _IsDragging Then
            If Math.Abs(DeltaX) < _DragThreshold AndAlso Math.Abs(DeltaY) < _DragThreshold Then Return
            _IsDragging = True
            If _BringToFrontOnDrag Then _Child.BringToFront()
            OnDragStarted()
            If Not _MouseIsDown OrElse Not CanBeginDragging() Then Return
        End If

        Dim Target As New Point(SaturatingAdd(_StartChildLocation.X, DeltaX),
                                SaturatingAdd(_StartChildLocation.Y, DeltaY))
        If _ConstrainToParent Then Target = ClampChildLocation(Target)

        If _Child.Location <> Target Then
            _Child.Location = Target
            OnDragging()
        End If
    End Sub

    Private Sub InputMouseUp(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then EndGesture()
    End Sub

    Private Sub InputMouseCaptureChanged(sender As Object, e As EventArgs)
        If _MouseIsDown AndAlso ReferenceEquals(sender, _MouseSource) AndAlso Not _MouseSource.Capture Then
            EndGesture()
        End If
    End Sub

    Private Sub EndGesture()
        If Not _MouseIsDown Then Return
        Dim CompletedDrag As Boolean = _IsDragging
        Dim StartLocation As Point = _StartChildLocation
        CancelGesture()
        If Not CompletedDrag OrElse _Child Is Nothing OrElse _Child.IsDisposed Then Return

        If _SnapToGrid Then
            Dim SnappedLocation As Point = CalculateSnapLocation(_Child.Location)
            If _SnapAnimation AndAlso _AnimationDuration > 0 AndAlso SnappedLocation <> _Child.Location Then
                StartSnapAnimation(SnappedLocation, StartLocation)
                Return
            End If
            _Child.Location = SnappedLocation
        End If
        OnDragCompleted(StartLocation)
    End Sub

    Private Sub CancelGesture()
        Dim Source As Control = _MouseSource
        _MouseSource = Nothing
        _MouseIsDown = False
        _IsDragging = False
        If Source IsNot Nothing AndAlso Not Source.IsDisposed AndAlso Source.Capture Then
            Source.Capture = False
        End If
    End Sub

    Private Sub CancelCurrentOperation()
        CancelGesture()
        CancelSnapAnimation()
    End Sub

    Private Shared Function SaturatingAdd(value As Integer, delta As Long) As Integer
        Dim Result As Long = CLng(value) + delta
        Return CInt(Math.Max(CLng(Integer.MinValue), Math.Min(CLng(Integer.MaxValue), Result)))
    End Function
End Class
