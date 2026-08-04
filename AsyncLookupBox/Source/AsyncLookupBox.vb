Imports System.ComponentModel
Imports System.Threading
''' <summary>
''' Represents a Windows Forms text box that requests results asynchronously and displays them in a configurable lookup drop-down.
''' </summary>
''' <remarks>
''' <para>The control is independent of SQL, HTTP, or any other data technology. Handle <see cref="SearchRequested"/> and call <see cref="AsyncLookupSearchRequestedEventArgs.SetSearchTask(Of TResult)"/> with the task returned by the application.</para>
''' <para>Every newer request cancels the token supplied to the preceding request. The selected business object remains available through <see cref="SelectedItem"/>, while <see cref="SelectedValue"/> is resolved from <see cref="ValueMember"/>.</para>
''' </remarks>
<DefaultEvent("SearchRequested")>
<DefaultProperty("DisplayMember")>
<DefaultBindingProperty(NameOf(Text))>
<Description("Performs application-defined asynchronous lookups and displays selectable results without depending on a specific data source.")>
<Designer(GetType(AsyncLookupBoxControlDesigner))>
<ToolboxItem(True)>
Partial Public Class AsyncLookupBox
    Inherits TextBox
    Private Const DefaultSearchInterval As Integer = 300
    Private Const DefaultMinimumCharacters As Integer = 2
    Private Const DefaultMaximumResults As Integer = 100
    Private Const DefaultDropDownHeight As Integer = 220
    Private ReadOnly _SearchTimer As System.Windows.Forms.Timer
    Private ReadOnly _Columns As AsyncLookupColumnCollection
    Private ReadOnly _ActionButton As AsyncLookupActionButton
    Private _DropDown As AsyncLookupDropDown
    Private _SearchCancellation As CancellationTokenSource
    Private _SearchVersion As Long
    Private _Results As IReadOnlyList(Of Object) = Array.Empty(Of Object)()
    Private _SelectedItem As Object
    Private _SelectedValue As Object
    Private _DisplayMember As String = String.Empty
    Private _ValueMember As String = String.Empty
    Private _SearchInterval As Integer = DefaultSearchInterval
    Private _MinimumCharacters As Integer = DefaultMinimumCharacters
    Private _MaximumResults As Integer = DefaultMaximumResults
    Private _DropDownHeight As Integer = DefaultDropDownHeight
    Private _DropDownWidth As Integer
    Private _SearchEnabled As Boolean = True
    Private _AutoSelectSingleResult As Boolean
    Private _ShowColumnHeaders As Boolean = True
    Private _ShowClearButton As Boolean = True
    Private _ClearButtonImage As Image
    Private _SelectedButtonImage As Image
    Private _HighlightSelectedItem As Boolean = True
    Private _SelectedItemBackColor As Color = Color.AliceBlue
    Private _SelectedItemForeColor As Color = Color.RoyalBlue
    Private _UnselectedBackColor As Color
    Private _UnselectedForeColor As Color
    Private _UpdatingSelectionAppearance As Boolean
    Private _LoadingText As String = "Searching..."
    Private _NoResultsText As String = "No results found."
    Private _SearchErrorText As String = "Unable to load results."
    Private _SearchNotConfiguredText As String = "Search provider is not configured."
    Private _CharactersRemainingText As String = "Enter {0} more character(s)."
    Private _ResultColumnHeaderText As String = "Result"
    Private _DropDownBackColor As Color = Color.White
    Private _DropDownForeColor As Color = SystemColors.ControlText
    Private _DropDownBorderColor As Color = SystemColors.ControlDark
    Private _SelectionBackColor As Color = SystemColors.Highlight
    Private _SelectionForeColor As Color = SystemColors.HighlightText
    Private _SuppressTextChanged As Boolean
    Private _IsSearching As Boolean
    Private _IsDisposing As Boolean
    ''' <summary>
    ''' Occurs when the control requires the application to supply a lookup operation.
    ''' </summary>
    ''' <remarks>The handler should call <see cref="AsyncLookupSearchRequestedEventArgs.SetSearchTask(Of TResult)"/> or <see cref="AsyncLookupSearchRequestedEventArgs.SetResults"/> before it returns.</remarks>
    <Category("AsyncLookupBox")>
    <Description("Occurs when the application must supply an asynchronous lookup operation.")>
    Public Event SearchRequested As EventHandler(Of AsyncLookupSearchRequestedEventArgs)
    ''' <summary>
    ''' Occurs after a current, non-canceled search completes successfully.
    ''' </summary>
    <Category("AsyncLookupBox")>
    <Description("Occurs after a current asynchronous lookup completes successfully.")>
    Public Event SearchCompleted As EventHandler(Of AsyncLookupSearchCompletedEventArgs)
    ''' <summary>
    ''' Occurs when preparing or awaiting a current search raises an exception.
    ''' </summary>
    <Category("AsyncLookupBox")>
    <Description("Occurs when an asynchronous lookup cannot be completed.")>
    Public Event SearchFailed As EventHandler(Of AsyncLookupSearchFailedEventArgs)
    ''' <summary>
    ''' Occurs when <see cref="IsSearching"/> changes.
    ''' </summary>
    <Category("AsyncLookupBox")>
    <Description("Occurs when the lookup searching state changes.")>
    Public Event IsSearchingChanged As EventHandler
    ''' <summary>
    ''' Occurs when a result is selected or the current selection is cleared.
    ''' </summary>
    <Category("AsyncLookupBox")>
    <Description("Occurs when a result is selected or the current selection is cleared.")>
    Public Event SelectionChanged As EventHandler(Of AsyncLookupSelectionChangedEventArgs)
    ''' <summary>
    ''' Occurs when the result drop-down is opened.
    ''' </summary>
    <Category("AsyncLookupBox")>
    <Description("Occurs when the result drop-down is opened.")>
    Public Event DropDownOpened As EventHandler
    ''' <summary>
    ''' Occurs when the result drop-down is closed.
    ''' </summary>
    <Category("AsyncLookupBox")>
    <Description("Occurs when the result drop-down is closed.")>
    Public Event DropDownClosed As EventHandler
    ''' <summary>
    ''' Initializes a new instance of the <see cref="AsyncLookupBox"/> class.
    ''' </summary>
    Public Sub New()
        _Columns = New AsyncLookupColumnCollection()
        AddHandler _Columns.Changed, AddressOf ColumnsChanged
        _SearchTimer = New System.Windows.Forms.Timer With {.Interval = _SearchInterval}
        AddHandler _SearchTimer.Tick, AddressOf SearchTimerTick
        _ActionButton = New AsyncLookupActionButton With {.BackColor = BackColor, .ForeColor = ForeColor, .Visible = False}
        AddHandler _ActionButton.Click, AddressOf ActionButtonClick
        Controls.Add(_ActionButton)
        InitializeSelectionAppearance()
        AccessibleRole = AccessibleRole.ComboBox
        UpdateActionButtonLayout()
    End Sub
    ''' <summary>
    ''' Immediately processes the current text and returns the retained results.
    ''' </summary>
    ''' <returns>A task containing the current result objects. An empty list is returned when the text is too short, the request is canceled, or the search fails.</returns>
    Public Function RefreshResultsAsync() As Task(Of IReadOnlyList(Of Object))
        _SearchTimer.Stop()
        Return PerformSearchAsync()
    End Function
    ''' <summary>
    ''' Cancels the current asynchronous request and closes the result drop-down.
    ''' </summary>
    Public Sub CancelSearch()
        CancelActiveSearch()
        CloseDropDown()
    End Sub
    ''' <summary>
    ''' Selects a result object and resolves its display text and selected value.
    ''' </summary>
    ''' <param name="Item">The result object to select, or <see langword="Nothing"/> to clear the selection.</param>
    Public Sub SelectItem(Item As Object)
        If Item Is Nothing Then
            ClearSelection()
            Return
        End If
        SelectItemCore(Item)
    End Sub
    ''' <summary>
    ''' Clears the selected object, selected value, text, results, and pending search.
    ''' </summary>
    Public Sub ClearSelection()
        ClearSelectionCore(True)
    End Sub
    ''' <summary>
    ''' Closes the result drop-down without changing the current selection or text.
    ''' </summary>
    Public Sub CloseDropDown()
        If _DropDown Is Nothing OrElse Not _DropDown.Visible Then Return
        _DropDown.HidePopup()
    End Sub
    ''' <summary>
    ''' Releases timers, cancellation resources, the result drop-down, and collection subscriptions.
    ''' </summary>
    ''' <param name="Disposing"><see langword="True"/> to release managed resources; otherwise, <see langword="False"/>.</param>
    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing AndAlso Not _IsDisposing Then
            _IsDisposing = True
            _SearchTimer.Stop()
            CancelActiveSearch()
            RemoveHandler _Columns.Changed, AddressOf ColumnsChanged
            RemoveHandler _SearchTimer.Tick, AddressOf SearchTimerTick
            RemoveHandler _ActionButton.Click, AddressOf ActionButtonClick
            _SearchTimer.Dispose()
            DisposeDropDown()
        End If
        MyBase.Dispose(Disposing)
    End Sub
    ''' <summary>
    ''' Clears a stale selection and schedules a new lookup after the configured debounce interval.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnTextChanged(e As EventArgs)
        MyBase.OnTextChanged(e)
        UpdateActionButtonLayout()
        If _SuppressTextChanged Then Return
        If _SelectedItem IsNot Nothing AndAlso Not String.Equals(Text, GetDisplayText(_SelectedItem), StringComparison.CurrentCulture) Then ClearSelectionCore(False)
        ScheduleSearch()
    End Sub
    ''' <summary>
    ''' Processes result navigation, selection, cancellation, and immediate-search keyboard commands.
    ''' </summary>
    ''' <param name="e">The keyboard event data.</param>
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        MyBase.OnKeyDown(e)
        If e.KeyCode = Keys.Down AndAlso _DropDown IsNot Nothing AndAlso _DropDown.Visible Then
            _DropDown.MoveSelection(1)
            e.Handled = True
            e.SuppressKeyPress = True
        ElseIf e.KeyCode = Keys.Up AndAlso _DropDown IsNot Nothing AndAlso _DropDown.Visible Then
            _DropDown.MoveSelection(-1)
            e.Handled = True
            e.SuppressKeyPress = True
        ElseIf e.KeyCode = Keys.Enter Then
            Dim Item As Object = _DropDown?.GetSelectedItem()
            If Item IsNot Nothing AndAlso _DropDown.Visible Then
                SelectItemCore(Item)
            ElseIf CanSearchCurrentText() Then
                StartSearch()
            End If
            e.Handled = True
            e.SuppressKeyPress = True
        ElseIf e.KeyCode = Keys.Escape Then
            If _DropDown IsNot Nothing AndAlso _DropDown.Visible Then
                CloseDropDown()
            ElseIf TextLength > 0 OrElse HasSelection Then
                ClearSelection()
            End If
            e.Handled = True
            e.SuppressKeyPress = True
        ElseIf e.KeyCode = Keys.F4 OrElse (e.Alt AndAlso e.KeyCode = Keys.Down) Then
            If _Results.Count > 0 Then ShowDropDownResults(_Results)
            e.Handled = True
            e.SuppressKeyPress = True
        End If
    End Sub
    ''' <summary>
    ''' Schedules a lookup when the control receives focus with searchable text.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnEnter(e As EventArgs)
        MyBase.OnEnter(e)
        If _Results.Count > 0 AndAlso CanSearchCurrentText() Then
            ShowDropDownResults(_Results)
        Else
            ScheduleSearch()
        End If
    End Sub
    ''' <summary>
    ''' Closes the result popup when keyboard focus moves away from the lookup box.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnLeave(e As EventArgs)
        MyBase.OnLeave(e)
        CloseDropDown()
    End Sub
    ''' <summary>
    ''' Updates the embedded action button after the control is resized.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        UpdateActionButtonLayout()
        If _DropDown IsNot Nothing AndAlso _DropDown.Visible Then _DropDown.ApplyOwnerSettings()
    End Sub
    ''' <summary>
    ''' Reapplies native text margins after the control handle is created.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        UpdateActionButtonLayout()
    End Sub
    ''' <summary>
    ''' Updates action-button placement when right-to-left layout changes.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnRightToLeftChanged(e As EventArgs)
        MyBase.OnRightToLeftChanged(e)
        UpdateActionButtonLayout()
    End Sub
    ''' <summary>
    ''' Synchronizes internal surfaces with the control background color.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnBackColorChanged(e As EventArgs)
        MyBase.OnBackColorChanged(e)
        If Not _UpdatingSelectionAppearance Then
            _UnselectedBackColor = MyBase.BackColor
            If HasSelection AndAlso HighlightSelectedItem Then ApplySelectionAppearance()
        End If
        If _ActionButton IsNot Nothing Then _ActionButton.BackColor = BackColor
    End Sub
    ''' <summary>
    ''' Synchronizes the action glyph with the control foreground color.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnForeColorChanged(e As EventArgs)
        MyBase.OnForeColorChanged(e)
        If Not _UpdatingSelectionAppearance Then
            _UnselectedForeColor = MyBase.ForeColor
            If HasSelection AndAlso HighlightSelectedItem Then ApplySelectionAppearance()
        End If
        If _ActionButton IsNot Nothing Then _ActionButton.ForeColor = ForeColor
    End Sub
    ''' <summary>
    ''' Synchronizes the action button with the control enabled state.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnEnabledChanged(e As EventArgs)
        MyBase.OnEnabledChanged(e)
        UpdateActionButtonLayout()
        If Not Enabled Then CancelSearch()
    End Sub
    ''' <summary>
    ''' Synchronizes the action button with the control read-only state.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnReadOnlyChanged(e As EventArgs)
        MyBase.OnReadOnlyChanged(e)
        UpdateActionButtonLayout()
    End Sub
    ''' <summary>
    ''' Closes the drop-down when the control moves.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnLocationChanged(e As EventArgs)
        MyBase.OnLocationChanged(e)
        CloseDropDown()
    End Sub
    ''' <summary>
    ''' Closes the drop-down when the control becomes invisible.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnVisibleChanged(e As EventArgs)
        MyBase.OnVisibleChanged(e)
        If Not Visible Then CloseDropDown()
    End Sub
    Private Sub ActionButtonClick(Sender As Object, E As EventArgs)
        If IsSearching Then
            CancelSearch()
        Else
            ClearSelection()
            Focus()
        End If
    End Sub
    Private Sub UpdateActionButtonLayout()
        If _ActionButton Is Nothing Then Return
        Dim ButtonSize As Integer = Math.Max(16, Math.Min(22, ClientSize.Height - 2))
        _ActionButton.Size = New Size(ButtonSize, ButtonSize)
        Dim ButtonY As Integer = Math.Max(0, (ClientSize.Height - ButtonSize) \ 2)
        Dim IsRightToLeft As Boolean = RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Dim ButtonX As Integer = If(IsRightToLeft, 1, Math.Max(0, ClientSize.Width - ButtonSize - 1))
        _ActionButton.Location = New Point(ButtonX, ButtonY)
        _ActionButton.IsSelected = False
        _ActionButton.Image = ClearButtonImage
        _ActionButton.Visible = ShowClearButton AndAlso (TextLength > 0 OrElse HasSelection)
        _ActionButton.Enabled = Enabled AndAlso (IsSearching OrElse Not Me.ReadOnly)
        _ActionButton.BringToFront()
        Dim ReservedMargin As Integer = If(ShowClearButton, ButtonSize + 3, 0)
        If IsRightToLeft Then
            AsyncLookupBoxInterop.SetTextMargins(Me, ReservedMargin, 0)
        Else
            AsyncLookupBoxInterop.SetTextMargins(Me, 0, ReservedMargin)
        End If
    End Sub
    Private Function IsInDesignMode() As Boolean
        Return LicenseManager.UsageMode = LicenseUsageMode.Designtime OrElse DesignMode
    End Function
End Class
