Partial Public Class TextBoxActionPanel
    ''' <summary>
    ''' Displays the panel when the component, target, owner form, and at least one visible action are available.
    ''' </summary>
    ''' <remarks>The method does not move keyboard focus away from the target control.</remarks>
    Public Sub ShowPanel()
        If Not CanShowPanel() Then
            HidePanel()
            Return
        End If
        EnsurePopup()
        RebuildPopup()
        If Not _Popup.HasActions Then
            HidePanel()
            Return
        End If
        RepositionPanel()
        If _Popup.Visible Then Return
        _OwnerForm = _TargetControl.FindForm()
        If _OwnerForm Is Nothing OrElse _OwnerForm.IsDisposed Then Return
        _Popup.Show(_OwnerForm)
        OnPanelShown()
    End Sub
    ''' <summary>
    ''' Hides the floating action panel without changing the target control or configured actions.
    ''' </summary>
    Public Sub HidePanel()
        If _Popup Is Nothing OrElse Not _Popup.Visible Then Return
        _Popup.Hide()
        OnPanelHidden()
    End Sub
    ''' <summary>
    ''' Rebuilds the buttons and updates the popup position when it is currently visible.
    ''' </summary>
    Public Sub RefreshPanel()
        If IsInDesignMode Then Return
        If _Popup IsNot Nothing AndAlso _Popup.Visible Then
            If Not CanShowPanel() Then
                HidePanel()
                Return
            End If
            RebuildPopup()
            If Not _Popup.HasActions Then
                HidePanel()
                Return
            End If
            RepositionPanel()
        ElseIf _ShowOnFocus AndAlso _TargetControl IsNot Nothing AndAlso _TargetControl.Focused Then
            ShowPanel()
        End If
    End Sub
    ''' <summary>
    ''' Executes the first enabled action whose key matches the specified value without regard to case.
    ''' </summary>
    ''' <param name="Key">The action key to execute.</param>
    ''' <returns><see langword="True"/> when an enabled matching action was executed; otherwise, <see langword="False"/>.</returns>
    ''' <remarks>An action can be executed by key even when its button is not visible.</remarks>
    Public Function PerformAction(Key As String) As Boolean
        Dim Action As TextBoxAction = _Actions.FindByKey(Key)
        If Action Is Nothing OrElse Not _Enabled OrElse Not Action.Enabled OrElse _TargetControl Is Nothing OrElse _TargetControl.IsDisposed Then Return False
        ExecuteAction(Action)
        Return True
    End Function
    Private Function CanShowPanel() As Boolean
        If IsInDesignMode OrElse Not _Enabled OrElse _TargetControl Is Nothing OrElse _TargetControl.IsDisposed OrElse Not _TargetControl.IsHandleCreated OrElse Not _TargetControl.Visible OrElse Not _TargetControl.Enabled Then Return False
        Dim Owner As Form = _TargetControl.FindForm()
        If Owner Is Nothing OrElse Owner.IsDisposed OrElse Not Owner.Visible Then Return False
        Return True
    End Function
    Private Sub EnsurePopup()
        If _Popup IsNot Nothing AndAlso Not _Popup.IsDisposed Then Return
        _Popup = New TextBoxActionPopup()
        AddHandler _Popup.ActionClick, AddressOf Popup_ActionClick
    End Sub
    Private Sub RebuildPopup()
        EnsurePopup()
        _Popup.Rebuild(_Actions, _ButtonSize, _ButtonSpacing, _PanelPadding, _TransparentBackground, _ShowBorder, _PanelBackColor, _BorderColor, _ButtonBackColor, _ButtonHoverBackColor, _ButtonPressedBackColor)
    End Sub
    Private Sub RepositionPanel()
        If _Popup Is Nothing OrElse _Popup.IsDisposed OrElse Not _Popup.HasActions OrElse _TargetControl Is Nothing OrElse _TargetControl.IsDisposed OrElse Not _TargetControl.IsHandleCreated Then Return
        Dim TargetBounds As Rectangle = GetTargetScreenBounds()
        Dim WorkingArea As Rectangle = Screen.FromRectangle(TargetBounds).WorkingArea
        Dim PreferredX As Integer = TargetBounds.Right - _Popup.Width
        Dim AboveY As Integer = TargetBounds.Top - _Popup.Height - _PanelOffset
        Dim BelowY As Integer = TargetBounds.Bottom + _PanelOffset
        Dim PreferredY As Integer
        Select Case _Placement
            Case TextBoxActionPanelPlacement.Above
                PreferredY = AboveY
            Case TextBoxActionPanelPlacement.Below
                PreferredY = BelowY
            Case Else
                PreferredY = If(AboveY >= WorkingArea.Top, AboveY, BelowY)
        End Select
        Dim MaximumX As Integer = Math.Max(WorkingArea.Left, WorkingArea.Right - _Popup.Width)
        Dim MaximumY As Integer = Math.Max(WorkingArea.Top, WorkingArea.Bottom - _Popup.Height)
        Dim X As Integer = Math.Max(WorkingArea.Left, Math.Min(PreferredX, MaximumX))
        Dim Y As Integer = Math.Max(WorkingArea.Top, Math.Min(PreferredY, MaximumY))
        _Popup.Location = New Point(X, Y)
    End Sub
    Private Function GetTargetScreenBounds() As Rectangle
        If _TargetControl.Parent IsNot Nothing Then Return _TargetControl.Parent.RectangleToScreen(_TargetControl.Bounds)
        Return _TargetControl.RectangleToScreen(_TargetControl.ClientRectangle)
    End Function
    Private Sub Popup_ActionClick(Sender As Object, E As TextBoxActionPopupEventArgs)
        ExecuteAction(E.Action)
    End Sub
End Class
