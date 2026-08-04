Partial Public Class TextBoxActionPanel
    Private Sub SetTargetControl(Value As TextBoxBase)
        If ReferenceEquals(_TargetControl, Value) Then Return
        If Value IsNot Nothing AndAlso Value.IsDisposed Then Throw New ObjectDisposedException(Value.Name, "A disposed TextBoxBase cannot be assigned as the target control.")
        DetachTargetControl()
        _TargetControl = Value
        AttachTargetControl()
        OnTargetControlChanged()
    End Sub
    Private Sub AttachTargetControl()
        If _TargetControl Is Nothing Then Return
        AddHandler _TargetControl.Enter, AddressOf TargetControl_Enter
        AddHandler _TargetControl.Leave, AddressOf TargetControl_Leave
        AddHandler _TargetControl.ParentChanged, AddressOf TargetControl_ParentChanged
        AddHandler _TargetControl.LocationChanged, AddressOf TargetControl_BoundsChanged
        AddHandler _TargetControl.SizeChanged, AddressOf TargetControl_BoundsChanged
        AddHandler _TargetControl.VisibleChanged, AddressOf TargetControl_StateChanged
        AddHandler _TargetControl.EnabledChanged, AddressOf TargetControl_StateChanged
        AddHandler _TargetControl.Disposed, AddressOf TargetControl_Disposed
        ObserveParentHierarchy()
    End Sub
    Private Sub DetachTargetControl()
        HidePanel()
        ClearParentHierarchyObservation()
        If _TargetControl Is Nothing Then Return
        RemoveHandler _TargetControl.Enter, AddressOf TargetControl_Enter
        RemoveHandler _TargetControl.Leave, AddressOf TargetControl_Leave
        RemoveHandler _TargetControl.ParentChanged, AddressOf TargetControl_ParentChanged
        RemoveHandler _TargetControl.LocationChanged, AddressOf TargetControl_BoundsChanged
        RemoveHandler _TargetControl.SizeChanged, AddressOf TargetControl_BoundsChanged
        RemoveHandler _TargetControl.VisibleChanged, AddressOf TargetControl_StateChanged
        RemoveHandler _TargetControl.EnabledChanged, AddressOf TargetControl_StateChanged
        RemoveHandler _TargetControl.Disposed, AddressOf TargetControl_Disposed
        _TargetControl = Nothing
    End Sub
    Private Sub ObserveParentHierarchy()
        ClearParentHierarchyObservation()
        If _TargetControl Is Nothing Then Return
        Dim CurrentControl As Control = _TargetControl.Parent
        While CurrentControl IsNot Nothing
            _ObservedAncestors.Add(CurrentControl)
            AddHandler CurrentControl.LocationChanged, AddressOf Ancestor_BoundsChanged
            AddHandler CurrentControl.SizeChanged, AddressOf Ancestor_BoundsChanged
            AddHandler CurrentControl.VisibleChanged, AddressOf Ancestor_StateChanged
            AddHandler CurrentControl.ParentChanged, AddressOf Ancestor_ParentChanged
            Dim ScrollableParent As ScrollableControl = TryCast(CurrentControl, ScrollableControl)
            If ScrollableParent IsNot Nothing Then AddHandler ScrollableParent.Scroll, AddressOf Ancestor_Scroll
            CurrentControl = CurrentControl.Parent
        End While
        _OwnerForm = _TargetControl.FindForm()
        If _OwnerForm IsNot Nothing Then
            AddHandler _OwnerForm.Activated, AddressOf OwnerForm_Activated
            AddHandler _OwnerForm.Deactivate, AddressOf OwnerForm_Deactivate
            AddHandler _OwnerForm.FormClosed, AddressOf OwnerForm_FormClosed
            AddHandler _OwnerForm.Shown, AddressOf OwnerForm_Shown
        End If
    End Sub
    Private Sub ClearParentHierarchyObservation()
        If _OwnerForm IsNot Nothing Then
            RemoveHandler _OwnerForm.Activated, AddressOf OwnerForm_Activated
            RemoveHandler _OwnerForm.Deactivate, AddressOf OwnerForm_Deactivate
            RemoveHandler _OwnerForm.FormClosed, AddressOf OwnerForm_FormClosed
            RemoveHandler _OwnerForm.Shown, AddressOf OwnerForm_Shown
            _OwnerForm = Nothing
        End If
        For Each CurrentControl As Control In _ObservedAncestors
            RemoveHandler CurrentControl.LocationChanged, AddressOf Ancestor_BoundsChanged
            RemoveHandler CurrentControl.SizeChanged, AddressOf Ancestor_BoundsChanged
            RemoveHandler CurrentControl.VisibleChanged, AddressOf Ancestor_StateChanged
            RemoveHandler CurrentControl.ParentChanged, AddressOf Ancestor_ParentChanged
            Dim ScrollableParent As ScrollableControl = TryCast(CurrentControl, ScrollableControl)
            If ScrollableParent IsNot Nothing Then RemoveHandler ScrollableParent.Scroll, AddressOf Ancestor_Scroll
        Next
        _ObservedAncestors.Clear()
    End Sub
    Private Sub TargetControl_Enter(Sender As Object, E As EventArgs)
        If _ShowOnFocus Then ShowPanel()
    End Sub
    Private Sub TargetControl_Leave(Sender As Object, E As EventArgs)
        If _HideOnLeave Then HidePanel()
    End Sub
    Private Sub TargetControl_ParentChanged(Sender As Object, E As EventArgs)
        ObserveParentHierarchy()
        RepositionPanel()
    End Sub
    Private Sub TargetControl_BoundsChanged(Sender As Object, E As EventArgs)
        RepositionPanel()
    End Sub
    Private Sub TargetControl_StateChanged(Sender As Object, E As EventArgs)
        If _TargetControl Is Nothing OrElse Not _TargetControl.Visible OrElse Not _TargetControl.Enabled Then
            HidePanel()
        ElseIf _ShowOnFocus AndAlso _TargetControl.Focused Then
            ShowPanel()
        End If
    End Sub
    Private Sub TargetControl_Disposed(Sender As Object, E As EventArgs)
        DetachTargetControl()
        OnTargetControlChanged()
    End Sub
    Private Sub Ancestor_BoundsChanged(Sender As Object, E As EventArgs)
        RepositionPanel()
    End Sub
    Private Sub Ancestor_StateChanged(Sender As Object, E As EventArgs)
        Dim Ancestor As Control = TryCast(Sender, Control)
        If Ancestor IsNot Nothing AndAlso Not Ancestor.Visible Then
            HidePanel()
        Else
            RepositionPanel()
        End If
    End Sub
    Private Sub Ancestor_ParentChanged(Sender As Object, E As EventArgs)
        ObserveParentHierarchy()
        RepositionPanel()
    End Sub
    Private Sub Ancestor_Scroll(Sender As Object, E As ScrollEventArgs)
        RepositionPanel()
    End Sub
    Private Sub OwnerForm_Activated(Sender As Object, E As EventArgs)
        If Not _ShowOnFocus OrElse _TargetControl Is Nothing OrElse _TargetControl.IsDisposed Then Return
        Dim Owner As Form = TryCast(Sender, Form)
        If Owner Is Nothing OrElse Owner.IsDisposed OrElse Not Owner.IsHandleCreated Then Return
        Owner.BeginInvoke(New MethodInvoker(AddressOf RestorePanelAfterOwnerActivation))
    End Sub
    Private Sub RestorePanelAfterOwnerActivation()
        If Not _ShowOnFocus OrElse _TargetControl Is Nothing OrElse _TargetControl.IsDisposed OrElse Not _TargetControl.IsHandleCreated OrElse Not _TargetControl.Focused Then Return
        ShowPanel()
    End Sub
    Private Sub OwnerForm_Deactivate(Sender As Object, E As EventArgs)
        HidePanel()
    End Sub
    Private Sub OwnerForm_FormClosed(Sender As Object, E As FormClosedEventArgs)
        HidePanel()
    End Sub
    Private Sub OwnerForm_Shown(Sender As Object, E As EventArgs)
        If _ShowOnFocus AndAlso _TargetControl IsNot Nothing AndAlso _TargetControl.Focused Then ShowPanel()
    End Sub
    Private Sub Actions_Changed(Sender As Object, E As EventArgs)
        RefreshPanel()
    End Sub
End Class
