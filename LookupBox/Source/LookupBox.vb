Imports System.ComponentModel
Imports System.Drawing
Imports System.Threading
Imports System.Windows.Forms
''' <summary>
''' Represents a Windows Forms text box that requests results asynchronously and displays them in a configurable lookup drop-down.
''' </summary>
''' <remarks>
''' <para>The control is independent of SQL, HTTP, or any other data technology. Handle <see cref="SearchRequested"/> and call <see cref="LookupSearchRequestedEventArgs.SetSearchTask(Of TResult)"/> with the task returned by the application.</para>
''' <para>Every newer request cancels the token supplied to the preceding request. The selected business object remains available through <see cref="SelectedItem"/>, while <see cref="SelectedValue"/> is resolved from <see cref="ValueMember"/>.</para>
''' </remarks>
<DefaultEvent("SearchRequested")>
<DefaultProperty("DisplayMember")>
<DefaultBindingProperty(NameOf(Text))>
<Description("Performs application-defined asynchronous lookups and displays selectable results without depending on a specific data source.")>
<Designer(GetType(LookupBoxControlDesigner))>
<ToolboxItem(True)>
Partial Public Class LookupBox
    Inherits CTextBox

    Private _ButtonMouseOverBackColor As Color

    <Category("Appearance"), DefaultValue(GetType(Color), ""), Description("The button background while the pointer is over it. Empty uses the default color.")>
    Public Property ButtonMouseOverBackColor As Color
        Get
            Return _ButtonMouseOverBackColor
        End Get
        Set(value As Color)
            If _ButtonMouseOverBackColor = value Then Return
            _ButtonMouseOverBackColor = value
            _ButtonLayout?.SetButtonColors(ButtonMouseOverBackColor, ButtonMouseDownBackColor)
        End Set
    End Property

    Private _ButtonMouseDownBackColor As Color

    <Category("Appearance"), DefaultValue(GetType(Color), ""), Description("The button background while it is pressed. Empty uses the default color.")>
    Public Property ButtonMouseDownBackColor As Color
        Get
            Return _ButtonMouseDownBackColor
        End Get
        Set(value As Color)
            If _ButtonMouseDownBackColor = value Then Return
            _ButtonMouseDownBackColor = value
            _ButtonLayout?.SetButtonColors(ButtonMouseOverBackColor, ButtonMouseDownBackColor)
        End Set
    End Property

    Private Const DefaultSearchInterval As Integer = 300
    Private Const DefaultMinimumCharacters As Integer = 2
    Private Const DefaultMaximumResults As Integer = 100
    Private Const DefaultDropDownHeight As Integer = 220
    Private ReadOnly _SearchTimer As System.Windows.Forms.Timer
    Private ReadOnly _Columns As LookupColumnCollection
    Private ReadOnly _ActionButton As LookupActionButton
    Private ReadOnly _ButtonLayout As EmbeddedTextButtonLayout
    Private _DropDown As LookupDropDown
    Private _ResultsViewMode As LookupResultsViewMode = LookupResultsViewMode.Grid
    Private _ItemTemplateType As Type
    Private _TemplateItemHeight As Integer = 72
    Private _TemplateItemSpacing As Integer = 2
    Private _SearchCancellation As CancellationTokenSource
    Private _SearchVersion As Long
    Private _TextVersion As Long
    Private _SelectionResolveVersion As Long
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
    Private _SelectedTextSnapshot As String = String.Empty
    Private _ClearOnEdit As Boolean
    Private _AutoSelectMatchedResultOnLeave As Boolean = True
    Private _SelectFromStartOnLeave As Boolean = True
    Private _ShowStartOnSelect As Boolean
    Private _AllowSelectionLink As Boolean
    Private _DropDownExtendLeft As Integer
    Private _DropDownExtendRight As Integer
    Private _AutoStretchDropDown As Boolean
    Private _GridHeaderBackColor As Color = SystemColors.Control
    Private _GridHeaderForeColor As Color = SystemColors.ControlText
    Private _GridHeadersBold As Boolean
    Private _GridShowVerticalLines As Boolean = True
    Private _MessageBackColor As Color = Color.White
    Private _MessageForeColor As Color = SystemColors.ControlText
    Private _LinkFont As Font
    Private _NormalLinkFont As Font
    ''' <summary>
    ''' Occurs when the control requires the application to supply a lookup operation.
    ''' </summary>
    ''' <remarks>The handler should call <see cref="LookupSearchRequestedEventArgs.SetSearchTask(Of TResult)"/> or <see cref="LookupSearchRequestedEventArgs.SetResults"/> before it returns.</remarks>
    <Category("LookupBox")>
    <Description("Occurs when the application must supply an asynchronous lookup operation.")>
    Public Event SearchRequested As EventHandler(Of LookupSearchRequestedEventArgs)
    ''' <summary>
    ''' Occurs after a current, non-canceled search completes successfully.
    ''' </summary>
    <Category("LookupBox")>
    <Description("Occurs after a current asynchronous lookup completes successfully.")>
    Public Event SearchCompleted As EventHandler(Of LookupSearchCompletedEventArgs)
    ''' <summary>
    ''' Occurs when preparing or awaiting a current search raises an exception.
    ''' </summary>
    <Category("LookupBox")>
    <Description("Occurs when an asynchronous lookup cannot be completed.")>
    Public Event SearchFailed As EventHandler(Of LookupSearchFailedEventArgs)
    ''' <summary>
    ''' Occurs when <see cref="IsSearching"/> changes.
    ''' </summary>
    <Category("LookupBox")>
    <Description("Occurs when the lookup searching state changes.")>
    Public Event IsSearchingChanged As EventHandler
    ''' <summary>
    ''' Occurs when a result is selected or the current selection is cleared.
    ''' </summary>
    <Category("LookupBox")>
    <Description("Occurs when a result is selected or the current selection is cleared.")>
    Public Event SelectionChanged As EventHandler(Of LookupSelectionChangedEventArgs)
    ''' <summary>Occurs before SelectedValue changes. This event is informational and does not cancel the change.</summary>
    <Category("LookupBox")>
    Public Event SelectedValueChanging As EventHandler(Of LookupSelectedValueChangedEventArgs)
    ''' <summary>Occurs after SelectedValue changes.</summary>
    <Category("LookupBox")>
    Public Event SelectedValueChanged As EventHandler(Of LookupSelectedValueChangedEventArgs)
    ''' <summary>Occurs when a selected record is Control-clicked and AllowSelectionLink is enabled.</summary>
    <Category("LookupBox")>
    Public Event HyperlinkClicked As EventHandler(Of LookupSelectionLinkClickedEventArgs)
    ''' <summary>
    ''' Occurs when the result drop-down is opened.
    ''' </summary>
    <Category("LookupBox")>
    <Description("Occurs when the result drop-down is opened.")>
    Public Event DropDownOpened As EventHandler
    ''' <summary>
    ''' Occurs when the result drop-down is closed.
    ''' </summary>
    <Category("LookupBox")>
    <Description("Occurs when the result drop-down is closed.")>
    Public Event DropDownClosed As EventHandler
    ''' <summary>
    ''' Raised whenever a recyclable result UserControl is assigned a result object.
    ''' Use this event when the custom UserControl does not implement ILookupItemTemplate.
    ''' </summary>
    <Category("LookupBox")>
    Public Event ItemTemplateBinding As EventHandler(Of LookupItemTemplateBindingEventArgs)
    ''' <summary>
    ''' Initializes a new instance of the <see cref="LookupBox"/> class.
    ''' </summary>
    Public Sub New()
        _Columns = New LookupColumnCollection()
        AddHandler _Columns.Changed, AddressOf ColumnsChanged
        _SearchTimer = New System.Windows.Forms.Timer With {.Interval = _SearchInterval}
        AddHandler _SearchTimer.Tick, AddressOf SearchTimerTick
        _ActionButton = New LookupActionButton With {.BackColor = BackColor, .ForeColor = ForeColor, .Visible = False}
        AddHandler _ActionButton.Click, AddressOf ActionButtonClick
        Controls.Add(_ActionButton)
        _ButtonLayout = New EmbeddedTextButtonLayout(Me, _ActionButton, AddressOf _ActionButton.RenderContent, Sub() ActionButtonClick(_ActionButton, EventArgs.Empty))
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
            If _ButtonLayout IsNot Nothing Then _ButtonLayout.Dispose()
            RemoveHandler _ActionButton.Click, AddressOf ActionButtonClick
            _SearchTimer.Dispose()
            DisposeDropDown()
            If _LinkFont IsNot Nothing Then _LinkFont.Dispose()
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
        _TextVersion += 1
        If _SelectedItem IsNot Nothing AndAlso Not String.Equals(Text, _SelectedTextSnapshot, StringComparison.CurrentCulture) Then
            ClearSelectionCore(False)
            If ClearOnEdit Then
                SetTextWithoutSearching(String.Empty)
                SelectionStart = 0
            End If
        End If
        ScheduleSearch()
    End Sub
    ''' <summary>
    ''' Processes result navigation, selection, cancellation, and immediate-search keyboard commands.
    ''' </summary>
    ''' <param name="e">The keyboard event data.</param>
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        MyBase.OnKeyDown(e)
        If AllowSelectionLink AndAlso HasSelection AndAlso e.Control Then ApplySelectionLinkAppearance()
        If e.KeyCode = Keys.Down AndAlso _DropDown IsNot Nothing AndAlso _DropDown.Visible Then
            _DropDown.MoveSelection(1)
            e.Handled = True
            e.SuppressKeyPress = True
        ElseIf e.KeyCode = Keys.Up AndAlso _DropDown IsNot Nothing AndAlso _DropDown.Visible Then
            _DropDown.MoveSelection(-1)
            e.Handled = True
            e.SuppressKeyPress = True
        ElseIf _DropDown IsNot Nothing AndAlso _DropDown.Visible AndAlso e.KeyCode = Keys.PageDown Then
            _DropDown.PageSelection(True)
            e.Handled = True
            e.SuppressKeyPress = True
        ElseIf _DropDown IsNot Nothing AndAlso _DropDown.Visible AndAlso e.KeyCode = Keys.PageUp Then
            _DropDown.PageSelection(False)
            e.Handled = True
            e.SuppressKeyPress = True
        ElseIf _DropDown IsNot Nothing AndAlso _DropDown.Visible AndAlso e.KeyCode = Keys.Home Then
            _DropDown.JumpSelection(False)
            e.Handled = True
            e.SuppressKeyPress = True
        ElseIf _DropDown IsNot Nothing AndAlso _DropDown.Visible AndAlso e.KeyCode = Keys.End Then
            _DropDown.JumpSelection(True)
            e.Handled = True
            e.SuppressKeyPress = True
        ElseIf e.KeyCode = Keys.Enter Then
            Dim Item As Object = _DropDown?.GetSelectedItem()
            If Item IsNot Nothing AndAlso _DropDown.Visible Then
                SelectItemCore(Item)
            ElseIf Not HasSelection AndAlso CanSearchCurrentText() Then
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
        If HasSelection Then Return
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
        ResetSelectionLinkAppearance()
        If AutoSelectMatchedResultOnLeave AndAlso Not HasSelection Then SelectUniqueMatchingResult()
        CloseDropDown()
        If SelectFromStartOnLeave Then SelectionStart = 0
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
    Protected Overrides Sub OnClientSizeChanged(E As EventArgs)
        MyBase.OnClientSizeChanged(E)
        UpdateActionButtonLayout()
    End Sub

    Protected Overrides Sub OnFontChanged(E As EventArgs)
        MyBase.OnFontChanged(E)
        UpdateActionButtonLayout()
    End Sub

    Protected Overrides Sub OnBorderStyleChanged(E As EventArgs)
        MyBase.OnBorderStyleChanged(E)
        UpdateActionButtonLayout()
    End Sub

    Protected Overrides Sub OnDpiChangedAfterParent(E As EventArgs)
        MyBase.OnDpiChangedAfterParent(E)
        UpdateActionButtonLayout()
    End Sub

    Protected Overrides Sub OnMultilineChanged(E As EventArgs)
        MyBase.OnMultilineChanged(E)
        UpdateActionButtonLayout()
    End Sub

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
    ''' <summary>Routes wheel input from the focused textbox to the template list.</summary>
    Protected Overrides Sub OnMouseWheel(e As MouseEventArgs)
        If _DropDown IsNot Nothing AndAlso _DropDown.Visible AndAlso ResultsViewMode = LookupResultsViewMode.Template Then
            _DropDown.ScrollResults(e.Delta)
            Dim Handled As HandledMouseEventArgs = TryCast(e, HandledMouseEventArgs)
            If Handled IsNot Nothing Then Handled.Handled = True
            Return
        End If
        MyBase.OnMouseWheel(e)
    End Sub

    Protected Overrides Sub OnMouseClick(e As MouseEventArgs)
        If e.Button = MouseButtons.Left AndAlso AllowSelectionLink AndAlso HasSelection AndAlso (Control.ModifierKeys And Keys.Control) = Keys.Control Then
            RaiseEvent HyperlinkClicked(Me, New LookupSelectionLinkClickedEventArgs(SelectedItem, SelectedValue))
            ResetSelectionLinkAppearance()
            Return
        End If
        MyBase.OnMouseClick(e)
    End Sub
    Protected Overrides Sub OnKeyUp(e As KeyEventArgs)
        MyBase.OnKeyUp(e)
        If (Control.ModifierKeys And Keys.Control) <> Keys.Control Then ResetSelectionLinkAppearance()
    End Sub
    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        ResetSelectionLinkAppearance()
    End Sub
    Private Sub ActionButtonClick(Sender As Object, E As EventArgs)
        If IsSearching Then
            CancelSearch()
        Else
            ClearSelection()
            Focus()
        End If
    End Sub
    Protected Overrides Sub WndProc(ByRef M As Message)
        If _ButtonLayout IsNot Nothing AndAlso _ButtonLayout.BeforeWndProc(M) Then Return
        MyBase.WndProc(M)
        If _ButtonLayout IsNot Nothing Then _ButtonLayout.AfterWndProc(M)
    End Sub

    Private Sub UpdateActionButtonLayout()
        If _ButtonLayout Is Nothing Then Return
        Dim IsRightToLeft As Boolean = Me.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        _ActionButton.IsSelected = HasSelection
        _ActionButton.Image = If(HasSelection AndAlso SelectedButtonImage IsNot Nothing, SelectedButtonImage, ClearButtonImage)
        _ActionButton.Enabled = Enabled AndAlso (IsSearching OrElse Not Me.ReadOnly)
        Dim ButtonWidth As Integer = _ButtonLayout.UpdateLayout(22, IsRightToLeft, ShowClearButton AndAlso (TextLength > 0 OrElse HasSelection OrElse IsSearching))
        ' Retain the reserved width while the clear-button feature is enabled;
        ' only HorizontalPadding adds spacing between the text and the button.
        Dim ReservedMargin As Integer = If(ShowClearButton, ButtonWidth, 0)
        SetTextMarginReservation(If(IsRightToLeft, ReservedMargin, 0), If(IsRightToLeft, 0, ReservedMargin))
    End Sub
    Private Function IsInDesignMode() As Boolean
        Return LicenseManager.UsageMode = LicenseUsageMode.Designtime OrElse DesignMode
    End Function
End Class
