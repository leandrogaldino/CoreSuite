Imports System.Windows.Forms

Partial Public Class LookupBox
    ''' <summary>Invoked whenever a recyclable UserControl is assigned a search-result object.</summary>
    Friend Sub BindTemplateControl(Template As UserControl, Item As Object, Index As Integer)
        Dim Binder As ILookupItemTemplate = TryCast(Template, ILookupItemTemplate)
        If Binder IsNot Nothing Then Binder.BindItem(Item)
        RaiseEvent ItemTemplateBinding(Me, New LookupItemTemplateBindingEventArgs(Item, Template, Index))
    End Sub

    Private Sub EnsureDropDown()
        If _DropDown IsNot Nothing AndAlso Not _DropDown.IsDisposed Then Return
        _DropDown = New LookupDropDown(Me)
        AddHandler _DropDown.ItemActivated, AddressOf DropDownItemActivated
        AddHandler _DropDown.PopupOpened, AddressOf DropDownOpenedInternal
        AddHandler _DropDown.PopupClosed, AddressOf DropDownClosedInternal
    End Sub
    Private Sub ShowDropDownStatus(Message As String)
        If _IsDisposing OrElse IsInDesignMode() OrElse Not Focused OrElse Parent Is Nothing OrElse Not Visible Then Return
        EnsureDropDown()
        _DropDown.ApplyOwnerSettings()
        _DropDown.ShowStatus(Message)
        ShowDropDownInternal()
    End Sub
    Private Sub ShowDropDownResults(Results As IReadOnlyList(Of Object))
        If _IsDisposing OrElse IsInDesignMode() OrElse Not Focused OrElse Parent Is Nothing OrElse Not Visible Then Return
        EnsureDropDown()
        _DropDown.ApplyOwnerSettings()
        _DropDown.ShowResults(Results)
        ShowDropDownInternal()
    End Sub
    Private Sub ShowDropDownInternal()
        _DropDown.ShowPopup()
    End Sub
    Private Sub RefreshDropDownAppearance()
        If _DropDown Is Nothing OrElse _DropDown.IsDisposed Then Return
        _DropDown.ApplyOwnerSettings()
    End Sub
    Private Sub RefreshVisibleResults()
        If _DropDown Is Nothing OrElse Not _DropDown.Visible OrElse _Results.Count = 0 Then Return
        _DropDown.ApplyOwnerSettings()
        _DropDown.ShowResults(_Results)
    End Sub
    Private Sub ColumnsChanged(Sender As Object, E As EventArgs)
        If HasSelection Then
            _SelectedTextSnapshot = GetDisplayText(_SelectedItem)
            SetTextWithoutSearching(_SelectedTextSnapshot)
        End If
        RefreshVisibleResults()
    End Sub
    Private Sub DropDownItemActivated(Sender As Object, E As LookupItemActivatedEventArgs)
        SelectItemCore(E.Item)
        Focus()
    End Sub
    Private Sub DropDownOpenedInternal(Sender As Object, E As EventArgs)
        RaiseEvent DropDownOpened(Me, EventArgs.Empty)
    End Sub
    Private Sub DropDownClosedInternal(Sender As Object, E As EventArgs)
        RaiseEvent DropDownClosed(Me, EventArgs.Empty)
    End Sub
    Private Sub DisposeDropDown()
        If _DropDown Is Nothing Then Return
        RemoveHandler _DropDown.ItemActivated, AddressOf DropDownItemActivated
        RemoveHandler _DropDown.PopupOpened, AddressOf DropDownOpenedInternal
        RemoveHandler _DropDown.PopupClosed, AddressOf DropDownClosedInternal
        _DropDown.Dispose()
        _DropDown = Nothing
    End Sub
End Class
