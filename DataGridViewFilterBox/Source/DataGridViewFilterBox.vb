Imports System.ComponentModel
Imports System.Drawing.Design
Imports System.Text
Imports System.Threading
''' <summary>
''' Represents a Windows Forms text box that filters an associated <see cref="DataGridView"/> after a configurable debounce interval.
''' </summary>
''' <remarks>
''' <para>Local filtering is performed through a <see cref="DataView.RowFilter"/> expression when the configured source is a <see cref="DataTable"/>, <see cref="DataView"/>, or a <see cref="BindingSource"/> backed by one of those types.</para>
''' <para>When local filtering is unavailable, <see cref="DataGridViewFilterMode.Automatic"/> falls back to <see cref="FilterRequested"/> so the application can perform a remote or custom search.</para>
''' <para>Any filter that already existed before this control applied its expression is preserved and combined with the generated filter through <c>AND</c>.</para>
''' </remarks>
<DefaultEvent("FilterApplied")>
<DefaultProperty(NameOf(DataGridView))>
<Description("Filters an associated DataGridView locally or requests application-defined remote filtering after a configurable debounce interval.")>
<Designer(GetType(DataGridViewFilterBoxControlDesigner))>
Public Class DataGridViewFilterBox
    Inherits TextBox
    Private Const DefaultFilterInterval As Integer = 300
    Private Const DefaultMinimumCharacters As Integer = 1
    Private ReadOnly _FilterTimer As System.Windows.Forms.Timer
    Private ReadOnly _ClearButton As FilterClearButton
    Private ReadOnly _FilterColumns As DataGridViewFilterColumnCollection
    Private ReadOnly _IgnoredColumns As DataGridViewFilterColumnCollection
    Private _DataGridView As DataGridView
    Private _BindingSource As BindingSource
    Private _ObservedGridBindingSource As BindingSource
    Private _FilterMode As DataGridViewFilterMode = DataGridViewFilterMode.Automatic
    Private _SearchMode As DataGridViewFilterSearchMode = DataGridViewFilterSearchMode.Contains
    Private _FilterInterval As Integer = DefaultFilterInterval
    Private _MinimumCharacters As Integer = DefaultMinimumCharacters
    Private _FilterEnabled As Boolean = True
    Private _CaseSensitive As Boolean
    Private _IncludeHiddenColumns As Boolean
    Private _ShowClearButton As Boolean = True
    Private _ClearButtonImage As Image
    Private _AppliedDataView As DataView
    Private _AppliedDataTable As DataTable
    Private _OriginalFilter As String = String.Empty
    Private _LastAppliedFilter As String = String.Empty
    Private _OriginalCaseSensitive As Boolean
    Private _AppliedCaseSensitive As Boolean
    Private _IsFilterApplied As Boolean
    Private _FilterWasRequested As Boolean
    Private _SuppressTextChanged As Boolean
    Private _IsDisposing As Boolean
    Private _RequestCancellation As CancellationTokenSource
    ''' <summary>
    ''' Occurs when custom filtering is requested or automatic local filtering is unavailable.
    ''' </summary>
    ''' <remarks>Use the supplied cancellation token to stop asynchronous work when a newer request replaces the current one.</remarks>
    <Category("DataGridViewFilterBox")>
    <Description("Occurs when custom filtering is requested or automatic local filtering is unavailable.")>
    Public Event FilterRequested As EventHandler(Of FilterRequestedEventArgs)
    ''' <summary>
    ''' Occurs after a local filter expression has been applied successfully.
    ''' </summary>
    <Category("DataGridViewFilterBox")>
    <Description("Occurs after a local filter expression has been applied successfully.")>
    Public Event FilterApplied As EventHandler(Of FilterAppliedEventArgs)
    ''' <summary>
    ''' Occurs after an active local or custom filter is cleared.
    ''' </summary>
    <Category("DataGridViewFilterBox")>
    <Description("Occurs after an active local or custom filter is cleared.")>
    Public Event FilterCleared As EventHandler
    ''' <summary>
    ''' Occurs when a local filter cannot be resolved or applied.
    ''' </summary>
    <Category("DataGridViewFilterBox")>
    <Description("Occurs when a local filter cannot be resolved or applied.")>
    Public Event FilterFailed As EventHandler(Of FilterFailedEventArgs)
    ''' <summary>
    ''' Initializes a new instance of the <see cref="DataGridViewFilterBox"/> class.
    ''' </summary>
    Public Sub New()
        _FilterColumns = New DataGridViewFilterColumnCollection()
        _IgnoredColumns = New DataGridViewFilterColumnCollection()
        AddHandler _FilterColumns.Changed, AddressOf FilterConfigurationChanged
        AddHandler _IgnoredColumns.Changed, AddressOf FilterConfigurationChanged
        _FilterTimer = New System.Windows.Forms.Timer With {.Interval = _FilterInterval}
        AddHandler _FilterTimer.Tick, AddressOf FilterTimerTick
        _ClearButton = New FilterClearButton With {.BackColor = BackColor, .ForeColor = ForeColor, .Visible = False}
        AddHandler _ClearButton.Click, AddressOf ClearButtonClick
        Controls.Add(_ClearButton)
        UpdateClearButtonLayout()
        PlaceholderText = "Filter..."
    End Sub
    ''' <summary>
    ''' Gets or sets the <see cref="DataGridView"/> whose data source is filtered by this control.
    ''' </summary>
    ''' <value>The target grid, or <see langword="Nothing"/> when no grid is associated.</value>
    <Category("DataGridViewFilterBox")>
    <Description("Defines the DataGridView whose data source is filtered by this control.")>
    Public Property DataGridView As DataGridView
        Get
            Return _DataGridView
        End Get
        Set(value As DataGridView)
            If ReferenceEquals(_DataGridView, value) Then Return
            DeactivateFilter(False)
            DetachDataGridView()
            _DataGridView = value
            AttachDataGridView()
            ScheduleFilter()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets an optional <see cref="System.Windows.Forms.BindingSource"/> that takes precedence over the grid's current data source.
    ''' </summary>
    ''' <value>
    ''' The explicit binding source to filter, or <see langword="Nothing"/> to resolve the source from <see cref="DataGridView"/>.
    ''' </value>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property BindingSource As BindingSource
        Get
            Return _BindingSource
        End Get
        Set(value As System.Windows.Forms.BindingSource)
            If ReferenceEquals(_BindingSource, value) Then Return
            DeactivateFilter(False)
            DetachBindingSource()
            _BindingSource = value
            AttachBindingSource()
            RefreshGridBindingSourceObservation()
            ScheduleFilter()
        End Set
    End Property
    ''' <summary>
    ''' Gets the columns included in filtering.
    ''' </summary>
    ''' <value>A collection of data column, grid column, or data property names. When empty, compatible visible columns are selected automatically.</value>
    <Category("DataGridViewFilterBox")>
    <Description("Defines the columns included in filtering. When empty, compatible visible columns are selected automatically.")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Editor(GetType(DataGridViewFilterColumnCollectionEditor), GetType(UITypeEditor))>
    Public ReadOnly Property FilterColumns As DataGridViewFilterColumnCollection
        Get
            Return _FilterColumns
        End Get
    End Property
    ''' <summary>
    ''' Gets the columns excluded from automatic or explicit filtering.
    ''' </summary>
    ''' <value>A collection of data column, grid column, or data property names that must not participate in filtering.</value>
    <Category("DataGridViewFilterBox")>
    <Description("Defines the columns excluded from automatic or explicit filtering.")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Editor(GetType(DataGridViewFilterColumnCollectionEditor), GetType(UITypeEditor))>
    Public ReadOnly Property IgnoredColumns As DataGridViewFilterColumnCollection
        Get
            Return _IgnoredColumns
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets how filter requests are processed.
    ''' </summary>
    ''' <value>One of the <see cref="DataGridViewFilterMode"/> values. The default is <see cref="DataGridViewFilterMode.Automatic"/>.</value>
    <Category("DataGridViewFilterBox")>
    <DefaultValue(DataGridViewFilterMode.Automatic)>
    <Description("Defines whether filtering is automatic, local-only, or delegated to the application.")>
    Public Property FilterMode As DataGridViewFilterMode
        Get
            Return _FilterMode
        End Get
        Set(value As DataGridViewFilterMode)
            If Not [Enum].IsDefined(GetType(DataGridViewFilterMode), value) Then Throw New InvalidEnumArgumentException(NameOf(value), CInt(value), GetType(DataGridViewFilterMode))
            If _FilterMode = value Then Return
            _FilterMode = value
            DeactivateFilter(False)
            ScheduleFilter()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets how entered text is matched during local filtering.
    ''' </summary>
    ''' <value>One of the <see cref="DataGridViewFilterSearchMode"/> values. The default is <see cref="DataGridViewFilterSearchMode.Contains"/>.</value>
    <Category("DataGridViewFilterBox")>
    <DefaultValue(DataGridViewFilterSearchMode.Contains)>
    <Description("Defines whether local filtering contains, starts with, ends with, or exactly matches the entered text.")>
    Public Property SearchMode As DataGridViewFilterSearchMode
        Get
            Return _SearchMode
        End Get
        Set(value As DataGridViewFilterSearchMode)
            If Not [Enum].IsDefined(GetType(DataGridViewFilterSearchMode), value) Then Throw New InvalidEnumArgumentException(NameOf(value), CInt(value), GetType(DataGridViewFilterSearchMode))
            If _SearchMode = value Then Return
            _SearchMode = value
            ScheduleFilter()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the debounce interval, in milliseconds, used before processing entered text.
    ''' </summary>
    ''' <value>A positive number of milliseconds. The default is <c>300</c>.</value>
    <Category("DataGridViewFilterBox")>
    <DefaultValue(DefaultFilterInterval)>
    <Description("Defines the debounce interval, in milliseconds, used before processing entered text.")>
    Public Property FilterInterval As Integer
        Get
            Return _FilterInterval
        End Get
        Set(value As Integer)
            If value < 1 Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "FilterInterval must be greater than zero.")
            If _FilterInterval = value Then Return
            _FilterInterval = value
            _FilterTimer.Interval = value
            ScheduleFilter()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the minimum number of characters required before filtering begins.
    ''' </summary>
    ''' <value>A positive number of characters. The default is <c>1</c>.</value>
    <Category("DataGridViewFilterBox")>
    <DefaultValue(DefaultMinimumCharacters)>
    <Description("Defines the minimum number of characters required before filtering begins.")>
    Public Property MinimumCharacters As Integer
        Get
            Return _MinimumCharacters
        End Get
        Set(value As Integer)
            If value < 1 Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "MinimumCharacters must be greater than zero.")
            If _MinimumCharacters = value Then Return
            _MinimumCharacters = value
            ScheduleFilter()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether filtering is enabled.
    ''' </summary>
    ''' <value><see langword="True"/> to process entered text; otherwise, <see langword="False"/>. The default is <see langword="True"/>.</value>
    <Category("DataGridViewFilterBox")>
    <DefaultValue(True)>
    <Description("Determines whether entered text is processed as a filter.")>
    Public Property FilterEnabled As Boolean
        Get
            Return _FilterEnabled
        End Get
        Set(value As Boolean)
            If _FilterEnabled = value Then Return
            _FilterEnabled = value
            If value Then
                ScheduleFilter()
            Else
                _FilterTimer.Stop()
                DeactivateFilter(True)
            End If
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether local text comparisons are case-sensitive.
    ''' </summary>
    ''' <value><see langword="True"/> for case-sensitive matching; otherwise, <see langword="False"/>. The default is <see langword="False"/>.</value>
    ''' <remarks>Local filtering temporarily applies this value to the target <see cref="DataTable.CaseSensitive"/> property and restores its original value when the filter is cleared.</remarks>
    <Category("DataGridViewFilterBox")>
    <DefaultValue(False)>
    <Description("Determines whether local text comparisons are case-sensitive.")>
    Public Property CaseSensitive As Boolean
        Get
            Return _CaseSensitive
        End Get
        Set(value As Boolean)
            If _CaseSensitive = value Then Return
            _CaseSensitive = value
            ScheduleFilter()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether hidden grid columns participate in automatic filtering.
    ''' </summary>
    ''' <value><see langword="True"/> to include hidden columns; otherwise, <see langword="False"/>. The default is <see langword="False"/>.</value>
    <Category("DataGridViewFilterBox")>
    <DefaultValue(False)>
    <Description("Determines whether hidden DataGridView columns participate in automatic filtering.")>
    Public Property IncludeHiddenColumns As Boolean
        Get
            Return _IncludeHiddenColumns
        End Get
        Set(value As Boolean)
            If _IncludeHiddenColumns = value Then Return
            _IncludeHiddenColumns = value
            ScheduleFilter()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether a clear button is displayed inside the text box while it contains text.
    ''' </summary>
    ''' <value><see langword="True"/> to display the clear button; otherwise, <see langword="False"/>. The default is <see langword="True"/>.</value>
    <Category("DataGridViewFilterBox")>
    <DefaultValue(True)>
    <Description("Determines whether a clear button is displayed inside the text box while it contains text.")>
    Public Property ShowClearButton As Boolean
        Get
            Return _ShowClearButton
        End Get
        Set(value As Boolean)
            If _ShowClearButton = value Then Return
            _ShowClearButton = value
            UpdateClearButtonLayout()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the optional image displayed by the clear button.
    ''' </summary>
    ''' <value>An image used by the clear button, or <see langword="Nothing"/> to draw the built-in close glyph.</value>
    <Category("DataGridViewFilterBox")>
    <Description("Defines the optional image displayed by the clear button. The built-in close glyph is used when no image is assigned.")>
    Public Property ClearButtonImage As Image
        Get
            Return _ClearButtonImage
        End Get
        Set(value As Image)
            If ReferenceEquals(_ClearButtonImage, value) Then Return
            _ClearButtonImage = value
            _ClearButton.Image = value
            _ClearButton.Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets a value indicating whether this control currently owns an active local filter.
    ''' </summary>
    ''' <value><see langword="True"/> when a local expression is active; otherwise, <see langword="False"/>.</value>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property IsFilterApplied As Boolean
        Get
            Return _IsFilterApplied
        End Get
    End Property
    ''' <summary>
    ''' Gets the complete local filter expression most recently assigned by this control.
    ''' </summary>
    ''' <value>The combined local filter expression, or an empty string when no local filter is active.</value>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property LastFilterExpression As String
        Get
            Return _LastAppliedFilter
        End Get
    End Property
    ''' <summary>
    ''' Applies the current text immediately without waiting for the debounce interval.
    ''' </summary>
    ''' <returns><see langword="True"/> when a local filter was applied or a custom filter was requested; otherwise, <see langword="False"/>.</returns>
    Public Function ApplyFilter() As Boolean
        _FilterTimer.Stop()
        If Not FilterEnabled OrElse _IsDisposing OrElse IsInDesignMode() Then Return False
        Dim FilterText As String = Text
        If String.IsNullOrEmpty(FilterText) OrElse FilterText.Length < MinimumCharacters Then
            DeactivateFilter(True)
            Return False
        End If
        Try
            Dim View As DataView = ResolveDataView()
            If FilterMode = DataGridViewFilterMode.Custom Then Return RequestCustomFilter(FilterText, ResolveRequestColumnNames(View))
            If View Is Nothing Then
                If FilterMode = DataGridViewFilterMode.Automatic Then Return RequestCustomFilter(FilterText, ResolveRequestColumnNames(Nothing))
                RaiseFilterFailed(FilterText, New NotSupportedException("Local filtering requires a DataTable, DataView, or a BindingSource backed by one of those types."))
                Return False
            End If
            Dim ColumnNames As List(Of String) = ResolveLocalColumnNames(View)
            If ColumnNames.Count = 0 Then
                RaiseFilterFailed(FilterText, New InvalidOperationException("No compatible columns are available for local filtering."))
                Return False
            End If
            CancelCustomRequest()
            _FilterWasRequested = False
            Return ApplyLocalFilter(View, FilterText, ColumnNames)
        Catch Failure As Exception
            DeactivateFilter(False)
            RaiseFilterFailed(FilterText, Failure)
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Clears the entered text, restores the source filter that existed before this control applied its filter, and cancels any pending custom request.
    ''' </summary>
    Public Sub ClearFilter()
        _FilterTimer.Stop()
        Dim HadFilter As Boolean = _IsFilterApplied OrElse _FilterWasRequested OrElse TextLength > 0
        _SuppressTextChanged = True
        Try
            Text = String.Empty
        Finally
            _SuppressTextChanged = False
        End Try
        DeactivateFilter(False)
        UpdateClearButtonLayout()
        If HadFilter Then RaiseEvent FilterCleared(Me, EventArgs.Empty)
    End Sub
    ''' <summary>
    ''' Reprocesses the current text immediately using the latest source and column configuration.
    ''' </summary>
    ''' <returns><see langword="True"/> when a local filter was applied or a custom filter was requested; otherwise, <see langword="False"/>.</returns>
    Public Function RefreshFilter() As Boolean
        Return ApplyFilter()
    End Function
    ''' <summary>
    ''' Releases managed resources, cancels pending requests, and restores any filter state owned by this control.
    ''' </summary>
    ''' <param name="Disposing"><see langword="True"/> to release managed and unmanaged resources; otherwise, <see langword="False"/>.</param>
    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing AndAlso Not _IsDisposing Then
            _IsDisposing = True
            _FilterTimer.Stop()
            DeactivateFilter(False)
            DetachDataGridView()
            DetachBindingSource()
            RemoveHandler _FilterColumns.Changed, AddressOf FilterConfigurationChanged
            RemoveHandler _IgnoredColumns.Changed, AddressOf FilterConfigurationChanged
            RemoveHandler _FilterTimer.Tick, AddressOf FilterTimerTick
            RemoveHandler _ClearButton.Click, AddressOf ClearButtonClick
            _FilterTimer.Dispose()
        End If
        MyBase.Dispose(Disposing)
    End Sub
    ''' <summary>
    ''' Processes text changes and schedules filtering after the configured debounce interval.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnTextChanged(e As EventArgs)
        MyBase.OnTextChanged(e)
        UpdateClearButtonLayout()
        If Not _SuppressTextChanged Then ScheduleFilter()
    End Sub
    ''' <summary>
    ''' Processes keyboard shortcuts for immediate filtering and clearing.
    ''' </summary>
    ''' <param name="e">The keyboard event data.</param>
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        If e.KeyCode = Keys.Escape AndAlso TextLength > 0 Then
            ClearFilter()
            e.Handled = True
            e.SuppressKeyPress = True
            Return
        End If
        If e.KeyCode = Keys.Enter AndAlso Not Multiline Then
            ApplyFilter()
            e.Handled = True
            e.SuppressKeyPress = True
            Return
        End If
        MyBase.OnKeyDown(e)
    End Sub
    ''' <summary>
    ''' Updates the embedded clear button after the control is resized.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        UpdateClearButtonLayout()
    End Sub
    ''' <summary>
    ''' Reapplies native text margins after the control handle is created.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        UpdateClearButtonLayout()
    End Sub
    ''' <summary>
    ''' Updates the embedded clear button when the text direction changes.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnRightToLeftChanged(e As EventArgs)
        MyBase.OnRightToLeftChanged(e)
        UpdateClearButtonLayout()
    End Sub
    ''' <summary>
    ''' Synchronizes the embedded clear button with the control background color.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnBackColorChanged(e As EventArgs)
        MyBase.OnBackColorChanged(e)
        If _ClearButton IsNot Nothing Then _ClearButton.BackColor = BackColor
    End Sub
    ''' <summary>
    ''' Synchronizes the embedded clear button with the control foreground color.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnForeColorChanged(e As EventArgs)
        MyBase.OnForeColorChanged(e)
        If _ClearButton IsNot Nothing Then _ClearButton.ForeColor = ForeColor
    End Sub
    ''' <summary>
    ''' Synchronizes the embedded clear button with the enabled state of the control.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnEnabledChanged(e As EventArgs)
        MyBase.OnEnabledChanged(e)
        If _ClearButton IsNot Nothing Then _ClearButton.Enabled = Enabled AndAlso Not Me.ReadOnly
    End Sub
    ''' <summary>
    ''' Synchronizes the embedded clear button with the read-only state of the control.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnReadOnlyChanged(e As EventArgs)
        MyBase.OnReadOnlyChanged(e)
        If _ClearButton IsNot Nothing Then _ClearButton.Enabled = Enabled AndAlso Not Me.ReadOnly
    End Sub
    Private Sub FilterTimerTick(sender As Object, e As EventArgs)
        _FilterTimer.Stop()
        ApplyFilter()
    End Sub
    Private Sub ClearButtonClick(sender As Object, e As EventArgs)
        ClearFilter()
        Focus()
    End Sub
    Private Sub FilterConfigurationChanged(sender As Object, e As EventArgs)
        ScheduleFilter()
    End Sub
    Private Sub DataGridViewSourceChanged(sender As Object, e As EventArgs)
        RefreshGridBindingSourceObservation()
        If _FilterWasRequested Then Return
        DeactivateFilter(False)
        ScheduleFilter()
    End Sub
    Private Sub DataGridViewColumnChanged(sender As Object, e As DataGridViewColumnEventArgs)
        If _FilterWasRequested Then Return
        ScheduleFilter()
    End Sub
    Private Sub BindingSourceDataSourceChanged(sender As Object, e As EventArgs)
        If _FilterWasRequested Then Return
        DeactivateFilter(False)
        ScheduleFilter()
    End Sub
    Private Sub AttachDataGridView()
        If _DataGridView Is Nothing Then Return
        AddHandler _DataGridView.DataSourceChanged, AddressOf DataGridViewSourceChanged
        AddHandler _DataGridView.ColumnAdded, AddressOf DataGridViewColumnChanged
        AddHandler _DataGridView.ColumnRemoved, AddressOf DataGridViewColumnChanged
        RefreshGridBindingSourceObservation()
    End Sub
    Private Sub DetachDataGridView()
        If _DataGridView Is Nothing Then Return
        DetachGridBindingSource()
        RemoveHandler _DataGridView.DataSourceChanged, AddressOf DataGridViewSourceChanged
        RemoveHandler _DataGridView.ColumnAdded, AddressOf DataGridViewColumnChanged
        RemoveHandler _DataGridView.ColumnRemoved, AddressOf DataGridViewColumnChanged
    End Sub
    Private Sub AttachBindingSource()
        If _BindingSource Is Nothing Then Return
        AddHandler _BindingSource.DataSourceChanged, AddressOf BindingSourceDataSourceChanged
    End Sub
    Private Sub DetachBindingSource()
        If _BindingSource Is Nothing Then Return
        RemoveHandler _BindingSource.DataSourceChanged, AddressOf BindingSourceDataSourceChanged
    End Sub
    Private Sub RefreshGridBindingSourceObservation()
        DetachGridBindingSource()
        If _DataGridView Is Nothing Then Return
        Dim GridBindingSource As BindingSource = TryCast(_DataGridView.DataSource, BindingSource)
        If GridBindingSource Is Nothing OrElse ReferenceEquals(GridBindingSource, _BindingSource) Then Return
        _ObservedGridBindingSource = GridBindingSource
        AddHandler _ObservedGridBindingSource.DataSourceChanged, AddressOf BindingSourceDataSourceChanged
    End Sub
    Private Sub DetachGridBindingSource()
        If _ObservedGridBindingSource Is Nothing Then Return
        RemoveHandler _ObservedGridBindingSource.DataSourceChanged, AddressOf BindingSourceDataSourceChanged
        _ObservedGridBindingSource = Nothing
    End Sub
    Private Sub ScheduleFilter()
        If _FilterTimer Is Nothing OrElse _IsDisposing OrElse IsInDesignMode() Then Return
        _FilterTimer.Stop()
        If Not FilterEnabled Then Return
        If String.IsNullOrEmpty(Text) OrElse TextLength < MinimumCharacters Then
            DeactivateFilter(True)
            Return
        End If
        _FilterTimer.Start()
    End Sub
    Private Function ApplyLocalFilter(View As DataView, FilterText As String, ColumnNames As List(Of String)) As Boolean
        If _AppliedDataView IsNot Nothing AndAlso Not ReferenceEquals(_AppliedDataView, View) Then RemoveLocalFilter()
        If _AppliedDataView Is Nothing Then
            _AppliedDataView = View
            _OriginalFilter = View.RowFilter
            _AppliedDataTable = View.Table
            _OriginalCaseSensitive = View.Table.CaseSensitive
        ElseIf Not String.Equals(View.RowFilter, _LastAppliedFilter, StringComparison.Ordinal) Then
            _OriginalFilter = View.RowFilter
        End If
        If _AppliedDataTable IsNot Nothing AndAlso _AppliedDataTable.CaseSensitive <> _AppliedCaseSensitive AndAlso _IsFilterApplied Then _OriginalCaseSensitive = _AppliedDataTable.CaseSensitive
        _AppliedCaseSensitive = CaseSensitive
        View.Table.CaseSensitive = CaseSensitive
        Dim GeneratedExpression As String = BuildFilterExpression(FilterText, ColumnNames)
        Dim CompleteExpression As String = If(String.IsNullOrWhiteSpace(_OriginalFilter), GeneratedExpression, $"({_OriginalFilter}) AND ({GeneratedExpression})")
        View.RowFilter = CompleteExpression
        _LastAppliedFilter = CompleteExpression
        _IsFilterApplied = True
        RaiseEvent FilterApplied(Me, New FilterAppliedEventArgs(FilterText, CompleteExpression, View.Count, ColumnNames))
        Return True
    End Function
    Private Function RequestCustomFilter(FilterText As String, ColumnNames As List(Of String)) As Boolean
        RemoveLocalFilter()
        CancelCustomRequest()
        _RequestCancellation = New CancellationTokenSource()
        _FilterWasRequested = True
        RaiseEvent FilterRequested(Me, New FilterRequestedEventArgs(FilterText, ColumnNames, _RequestCancellation.Token))
        Return True
    End Function
    Private Sub DeactivateFilter(RaiseClearedEvent As Boolean)
        Dim HadFilter As Boolean = _IsFilterApplied OrElse _FilterWasRequested
        CancelCustomRequest()
        RemoveLocalFilter()
        _FilterWasRequested = False
        If HadFilter AndAlso RaiseClearedEvent Then RaiseEvent FilterCleared(Me, EventArgs.Empty)
    End Sub
    Private Sub RemoveLocalFilter()
        Dim View As DataView = _AppliedDataView
        Dim Table As DataTable = _AppliedDataTable
        Dim OriginalFilter As String = _OriginalFilter
        Dim LastAppliedFilter As String = _LastAppliedFilter
        Dim OriginalCaseSensitive As Boolean = _OriginalCaseSensitive
        Dim AppliedCaseSensitive As Boolean = _AppliedCaseSensitive
        _AppliedDataView = Nothing
        _AppliedDataTable = Nothing
        _OriginalFilter = String.Empty
        _LastAppliedFilter = String.Empty
        _IsFilterApplied = False
        If View IsNot Nothing AndAlso String.Equals(View.RowFilter, LastAppliedFilter, StringComparison.Ordinal) Then View.RowFilter = OriginalFilter
        If Table IsNot Nothing AndAlso Table.CaseSensitive = AppliedCaseSensitive Then Table.CaseSensitive = OriginalCaseSensitive
    End Sub
    Private Sub CancelCustomRequest()
        If _RequestCancellation Is Nothing Then Return
        _RequestCancellation.Cancel()
        _RequestCancellation.Dispose()
        _RequestCancellation = Nothing
    End Sub
    Private Function ResolveDataView() As DataView
        If _BindingSource IsNot Nothing Then
            Dim ExplicitView As DataView = ResolveDataViewFromSource(_BindingSource)
            If ExplicitView IsNot Nothing Then Return ExplicitView
        End If
        If _DataGridView Is Nothing Then Return Nothing
        Return ResolveDataViewFromSource(_DataGridView.DataSource)
    End Function
    Private Shared Function ResolveDataViewFromSource(Source As Object) As DataView
        Dim View As DataView = TryCast(Source, DataView)
        If View IsNot Nothing Then Return View
        Dim Table As DataTable = TryCast(Source, DataTable)
        If Table IsNot Nothing Then Return Table.DefaultView
        Dim SourceBinding As BindingSource = TryCast(Source, BindingSource)
        If SourceBinding Is Nothing Then Return Nothing
        View = TryCast(SourceBinding.List, DataView)
        If View IsNot Nothing Then Return View
        Table = TryCast(SourceBinding.DataSource, DataTable)
        If Table IsNot Nothing Then Return Table.DefaultView
        View = TryCast(SourceBinding.DataSource, DataView)
        If View IsNot Nothing Then Return View
        Dim NestedBinding As BindingSource = TryCast(SourceBinding.DataSource, BindingSource)
        If NestedBinding IsNot Nothing AndAlso Not ReferenceEquals(NestedBinding, SourceBinding) Then Return ResolveDataViewFromSource(NestedBinding)
        Return Nothing
    End Function
    Private Function ResolveLocalColumnNames(View As DataView) As List(Of String)
        Dim ColumnNames As New List(Of String)()
        If FilterColumns.Count > 0 Then
            For Each Reference As DataGridViewFilterColumn In FilterColumns
                Dim ColumnName As String = ResolveDataColumnName(Reference.ColumnName, View.Table)
                If Not String.IsNullOrEmpty(ColumnName) AndAlso Not IsIgnoredColumn(ColumnName, View.Table) Then AddUniqueColumnName(ColumnNames, ColumnName)
            Next
            Return ColumnNames
        End If
        For Each Column As DataColumn In View.Table.Columns
            If IsSearchableDataType(Column.DataType) AndAlso Not IsIgnoredColumn(Column.ColumnName, View.Table) AndAlso IsColumnVisible(Column.ColumnName) Then AddUniqueColumnName(ColumnNames, Column.ColumnName)
        Next
        Return ColumnNames
    End Function
    Private Function ResolveRequestColumnNames(View As DataView) As List(Of String)
        If View IsNot Nothing Then Return ResolveLocalColumnNames(View)
        Dim ColumnNames As New List(Of String)()
        If FilterColumns.Count > 0 Then
            For Each Reference As DataGridViewFilterColumn In FilterColumns
                If Not IgnoredColumns.Contains(Reference.ColumnName) Then AddUniqueColumnName(ColumnNames, ResolveGridPropertyName(Reference.ColumnName))
            Next
            Return ColumnNames
        End If
        If _DataGridView Is Nothing Then Return ColumnNames
        For Each Column As DataGridViewColumn In _DataGridView.Columns
            If (IncludeHiddenColumns OrElse Column.Visible) AndAlso Not IsIgnoredGridColumn(Column) Then AddUniqueColumnName(ColumnNames, If(String.IsNullOrWhiteSpace(Column.DataPropertyName), Column.Name, Column.DataPropertyName))
        Next
        Return ColumnNames
    End Function
    Private Function ResolveDataColumnName(ReferenceName As String, Table As DataTable) As String
        For Each Column As DataColumn In Table.Columns
            If String.Equals(Column.ColumnName, ReferenceName, StringComparison.OrdinalIgnoreCase) Then Return Column.ColumnName
        Next
        If _DataGridView Is Nothing Then Return String.Empty
        For Each GridColumn As DataGridViewColumn In _DataGridView.Columns
            If String.Equals(GridColumn.Name, ReferenceName, StringComparison.OrdinalIgnoreCase) OrElse String.Equals(GridColumn.DataPropertyName, ReferenceName, StringComparison.OrdinalIgnoreCase) Then
                Dim PropertyName As String = If(String.IsNullOrWhiteSpace(GridColumn.DataPropertyName), GridColumn.Name, GridColumn.DataPropertyName)
                For Each Column As DataColumn In Table.Columns
                    If String.Equals(Column.ColumnName, PropertyName, StringComparison.OrdinalIgnoreCase) Then Return Column.ColumnName
                Next
            End If
        Next
        Return String.Empty
    End Function
    Private Function ResolveGridPropertyName(ReferenceName As String) As String
        If _DataGridView Is Nothing Then Return ReferenceName
        For Each Column As DataGridViewColumn In _DataGridView.Columns
            If String.Equals(Column.Name, ReferenceName, StringComparison.OrdinalIgnoreCase) OrElse String.Equals(Column.DataPropertyName, ReferenceName, StringComparison.OrdinalIgnoreCase) Then Return If(String.IsNullOrWhiteSpace(Column.DataPropertyName), Column.Name, Column.DataPropertyName)
        Next
        Return ReferenceName
    End Function
    Private Function IsIgnoredColumn(ColumnName As String, Table As DataTable) As Boolean
        For Each Reference As DataGridViewFilterColumn In IgnoredColumns
            Dim IgnoredName As String = ResolveDataColumnName(Reference.ColumnName, Table)
            If String.Equals(IgnoredName, ColumnName, StringComparison.OrdinalIgnoreCase) Then Return True
        Next
        Return False
    End Function
    Private Function IsIgnoredGridColumn(Column As DataGridViewColumn) As Boolean
        For Each Reference As DataGridViewFilterColumn In IgnoredColumns
            If String.Equals(Reference.ColumnName, Column.Name, StringComparison.OrdinalIgnoreCase) OrElse String.Equals(Reference.ColumnName, Column.DataPropertyName, StringComparison.OrdinalIgnoreCase) Then Return True
        Next
        Return False
    End Function
    Private Function IsColumnVisible(ColumnName As String) As Boolean
        If IncludeHiddenColumns OrElse _DataGridView Is Nothing OrElse _DataGridView.Columns.Count = 0 Then Return True
        For Each Column As DataGridViewColumn In _DataGridView.Columns
            Dim PropertyName As String = If(String.IsNullOrWhiteSpace(Column.DataPropertyName), Column.Name, Column.DataPropertyName)
            If String.Equals(PropertyName, ColumnName, StringComparison.OrdinalIgnoreCase) Then Return Column.Visible
        Next
        Return False
    End Function
    Private Shared Function IsSearchableDataType(DataType As Type) As Boolean
        If DataType Is GetType(Byte()) OrElse GetType(Image).IsAssignableFrom(DataType) Then Return False
        Return GetType(IConvertible).IsAssignableFrom(DataType) OrElse DataType Is GetType(Guid) OrElse DataType Is GetType(TimeSpan)
    End Function
    Private Shared Sub AddUniqueColumnName(ColumnNames As List(Of String), ColumnName As String)
        If String.IsNullOrWhiteSpace(ColumnName) Then Return
        If Not ColumnNames.Any(Function(CurrentName) String.Equals(CurrentName, ColumnName, StringComparison.OrdinalIgnoreCase)) Then ColumnNames.Add(ColumnName)
    End Sub
    Private Function BuildFilterExpression(FilterText As String, ColumnNames As IEnumerable(Of String)) As String
        Dim PatternText As String = EscapeLikeValue(FilterText)
        Select Case SearchMode
            Case DataGridViewFilterSearchMode.Contains
                PatternText = $"%{PatternText}%"
            Case DataGridViewFilterSearchMode.StartsWith
                PatternText = $"{PatternText}%"
            Case DataGridViewFilterSearchMode.EndsWith
                PatternText = $"%{PatternText}"
        End Select
        Dim Expressions As New List(Of String)()
        For Each ColumnName As String In ColumnNames
            Expressions.Add($"CONVERT([{EscapeColumnName(ColumnName)}], 'System.String') LIKE '{PatternText}'")
        Next
        Return $"({String.Join(" OR ", Expressions)})"
    End Function
    Private Shared Function EscapeColumnName(ColumnName As String) As String
        Return ColumnName.Replace("\", "\\").Replace("]", "\]")
    End Function
    Private Shared Function EscapeLikeValue(Value As String) As String
        Dim Builder As New StringBuilder(Value.Length)
        For Each Character As Char In Value
            Select Case Character
                Case "'"c
                    Builder.Append("''")
                Case "["c
                    Builder.Append("[[]")
                Case "]"c
                    Builder.Append("[]]")
                Case "%"c
                    Builder.Append("[%]")
                Case "*"c
                    Builder.Append("[*]")
                Case Else
                    Builder.Append(Character)
            End Select
        Next
        Return Builder.ToString()
    End Function
    Private Sub UpdateClearButtonLayout()
        If _ClearButton Is Nothing Then Return
        Dim ButtonSize As Integer = Math.Max(16, Math.Min(22, ClientSize.Height - 2))
        _ClearButton.Size = New Size(ButtonSize, ButtonSize)
        Dim ButtonY As Integer = Math.Max(0, (ClientSize.Height - ButtonSize) \ 2)
        Dim IsRightToLeft As Boolean = Me.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Dim ButtonX As Integer = If(IsRightToLeft, 1, Math.Max(0, ClientSize.Width - ButtonSize - 1))
        _ClearButton.Location = New Point(ButtonX, ButtonY)
        _ClearButton.Visible = ShowClearButton AndAlso TextLength > 0
        _ClearButton.Enabled = Enabled AndAlso Not Me.ReadOnly
        _ClearButton.BringToFront()
        Dim ReservedMargin As Integer = If(ShowClearButton, ButtonSize + 3, 0)
        If IsRightToLeft Then
            DataGridViewFilterBoxInterop.SetTextMargins(Me, ReservedMargin, 0)
        Else
            DataGridViewFilterBoxInterop.SetTextMargins(Me, 0, ReservedMargin)
        End If
    End Sub
    Private Sub RaiseFilterFailed(FilterText As String, Failure As Exception)
        RaiseEvent FilterFailed(Me, New FilterFailedEventArgs(FilterText, Failure))
    End Sub
    Private Function IsInDesignMode() As Boolean
        Return LicenseManager.UsageMode = LicenseUsageMode.Designtime OrElse DesignMode
    End Function
End Class
