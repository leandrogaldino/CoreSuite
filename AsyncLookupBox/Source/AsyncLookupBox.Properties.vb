Imports System.ComponentModel
Imports System.Drawing.Design
Partial Public Class AsyncLookupBox
    ''' <summary>
    ''' Gets or sets the property path used to create the text shown after a result is selected.
    ''' </summary>
    ''' <value>A property, data-column, dictionary-key, or nested path; an empty string uses the result object's text representation.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue("")>
    <Description("Defines the property path used as the text of a selected result.")>
    Public Property DisplayMember As String
        Get
            Return _DisplayMember
        End Get
        Set(value As String)
            Dim NormalizedValue As String = If(value, String.Empty).Trim()
            If String.Equals(_DisplayMember, NormalizedValue, StringComparison.Ordinal) Then Return
            _DisplayMember = NormalizedValue
            If _SelectedItem IsNot Nothing Then SetTextWithoutSearching(GetDisplayText(_SelectedItem))
            RefreshVisibleResults()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the property path used to resolve <see cref="SelectedValue"/>.
    ''' </summary>
    ''' <value>A property, data-column, dictionary-key, or nested path; an empty string uses the entire selected object.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue("")>
    <Description("Defines the property path used to resolve the selected value.")>
    Public Property ValueMember As String
        Get
            Return _ValueMember
        End Get
        Set(value As String)
            Dim NormalizedValue As String = If(value, String.Empty).Trim()
            If String.Equals(_ValueMember, NormalizedValue, StringComparison.Ordinal) Then Return
            _ValueMember = NormalizedValue
            If _SelectedItem IsNot Nothing Then _SelectedValue = ResolveValue(_SelectedItem)
        End Set
    End Property
    ''' <summary>
    ''' Gets the ordered result-column configuration.
    ''' </summary>
    ''' <value>The collection used to build the result grid. An empty collection displays only <see cref="DisplayMember"/>.</value>
    <Category("AsyncLookupBox")>
    <Description("Defines the columns displayed in the lookup result list.")>
    <Editor(GetType(AsyncLookupColumnCollectionEditor), GetType(UITypeEditor))>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public ReadOnly Property Columns As AsyncLookupColumnCollection
        Get
            Return _Columns
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets the debounce interval used before a search begins.
    ''' </summary>
    ''' <value>The delay in milliseconds. The default is <c>300</c>.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue(DefaultSearchInterval)>
    <Description("Defines the debounce interval, in milliseconds, used before a lookup begins.")>
    Public Property SearchInterval As Integer
        Get
            Return _SearchInterval
        End Get
        Set(value As Integer)
            If value < 1 Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "SearchInterval must be greater than zero.")
            If _SearchInterval = value Then Return
            _SearchInterval = value
            _SearchTimer.Interval = value
            ScheduleSearch()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the minimum number of characters required before a search begins.
    ''' </summary>
    ''' <value>A positive character count. The default is <c>2</c>.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue(DefaultMinimumCharacters)>
    <Description("Defines the minimum number of characters required before a lookup begins.")>
    Public Property MinimumCharacters As Integer
        Get
            Return _MinimumCharacters
        End Get
        Set(value As Integer)
            If value < 1 Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "MinimumCharacters must be greater than zero.")
            If _MinimumCharacters = value Then Return
            _MinimumCharacters = value
            ScheduleSearch()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the maximum number of results retained by the control.
    ''' </summary>
    ''' <value>The maximum count, or <c>0</c> for no control-side limit. The default is <c>100</c>.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue(DefaultMaximumResults)>
    <Description("Defines the maximum number of results retained by the control; zero removes the control-side limit.")>
    Public Property MaximumResults As Integer
        Get
            Return _MaximumResults
        End Get
        Set(value As Integer)
            If value < 0 Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "MaximumResults cannot be negative.")
            _MaximumResults = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether entered text starts lookup requests.
    ''' </summary>
    ''' <value><see langword="True"/> to enable searching; otherwise, <see langword="False"/>. The default is <see langword="True"/>.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue(True)>
    <Description("Determines whether entered text starts asynchronous lookup requests.")>
    Public Property SearchEnabled As Boolean
        Get
            Return _SearchEnabled
        End Get
        Set(value As Boolean)
            If _SearchEnabled = value Then Return
            _SearchEnabled = value
            If value Then
                ScheduleSearch()
            Else
                _SearchTimer.Stop()
                CancelSearch()
            End If
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether a single returned result is selected automatically.
    ''' </summary>
    ''' <value><see langword="True"/> to select a single result; otherwise, <see langword="False"/>. The default is <see langword="False"/>.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue(False)>
    <Description("Determines whether a single returned result is selected automatically.")>
    Public Property AutoSelectSingleResult As Boolean
        Get
            Return _AutoSelectSingleResult
        End Get
        Set(value As Boolean)
            _AutoSelectSingleResult = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the result drop-down width.
    ''' </summary>
    ''' <value>The width in pixels, or <c>0</c> to use at least the lookup-box width. The default is <c>0</c>.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue(0)>
    <Description("Defines the result drop-down width; zero uses at least the lookup-box width.")>
    Public Property DropDownWidth As Integer
        Get
            Return _DropDownWidth
        End Get
        Set(value As Integer)
            If value < 0 Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "DropDownWidth cannot be negative.")
            _DropDownWidth = value
            RefreshDropDownAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the result drop-down height.
    ''' </summary>
    ''' <value>The height in pixels. The default is <c>220</c>.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue(DefaultDropDownHeight)>
    <Description("Defines the result drop-down height in pixels.")>
    Public Property DropDownHeight As Integer
        Get
            Return _DropDownHeight
        End Get
        Set(value As Integer)
            If value < 40 Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "DropDownHeight must be at least 40 pixels.")
            _DropDownHeight = value
            RefreshDropDownAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether result-column headers are displayed.
    ''' </summary>
    <Category("AsyncLookupBox")>
    <DefaultValue(True)>
    <Description("Determines whether result-column headers are displayed.")>
    Public Property ShowColumnHeaders As Boolean
        Get
            Return _ShowColumnHeaders
        End Get
        Set(value As Boolean)
            _ShowColumnHeaders = value
            RefreshDropDownAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the embedded clear and cancel button is displayed.
    ''' </summary>
    <Category("AsyncLookupBox")>
    <DefaultValue(True)>
    <Description("Determines whether the embedded clear and search-cancel button is displayed.")>
    Public Property ShowClearButton As Boolean
        Get
            Return _ShowClearButton
        End Get
        Set(value As Boolean)
            If _ShowClearButton = value Then Return
            _ShowClearButton = value
            UpdateActionButtonLayout()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a custom image displayed by the clear button when no search is active.
    ''' </summary>
    <Category("AsyncLookupBox")>
    <DefaultValue(GetType(Image), Nothing)>
    <Description("Defines a custom clear-button image; the control draws a close glyph when no image is assigned.")>
    Public Property ClearButtonImage As Image
        Get
            Return _ClearButtonImage
        End Get
        Set(value As Image)
            If ReferenceEquals(_ClearButtonImage, value) Then Return
            _ClearButtonImage = value
            UpdateActionButtonLayout()
            _ActionButton.Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a custom image displayed by the action button while an item is selected.
    ''' </summary>
    ''' <value>An image that represents the selected state, or <see langword="Nothing"/> to use the built-in check mark.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue(GetType(Image), Nothing)>
    <Description("Defines a custom selected-state image; the control draws a check mark when no image is assigned.")>
    Public Property SelectedButtonImage As Image
        Get
            Return _SelectedButtonImage
        End Get
        Set(value As Image)
            If ReferenceEquals(_SelectedButtonImage, value) Then Return
            _SelectedButtonImage = value
            UpdateActionButtonLayout()
            _ActionButton.Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the text box uses distinct colors while an item is selected.
    ''' </summary>
    ''' <value><see langword="True"/> to highlight the selected state; otherwise, <see langword="False"/>. The default is <see langword="True"/>.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue(True)>
    <Description("Determines whether the text box uses distinct colors while a lookup item is selected.")>
    Public Property HighlightSelectedItem As Boolean
        Get
            Return _HighlightSelectedItem
        End Get
        Set(value As Boolean)
            If _HighlightSelectedItem = value Then Return
            _HighlightSelectedItem = value
            ApplySelectionAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the text-box background color used while an item is selected.
    ''' </summary>
    ''' <value>The selected-state background color. The default is <see cref="Color.AliceBlue"/>.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue(GetType(Color), "AliceBlue")>
    <Description("Defines the text-box background color used while a lookup item is selected.")>
    Public Property SelectedItemBackColor As Color
        Get
            Return _SelectedItemBackColor
        End Get
        Set(value As Color)
            If value.IsEmpty OrElse value.A < Byte.MaxValue Then Throw New ArgumentException("SelectedItemBackColor must be an opaque, non-empty color.", NameOf(value))
            If _SelectedItemBackColor = value Then Return
            _SelectedItemBackColor = value
            ApplySelectionAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the text color used while an item is selected.
    ''' </summary>
    ''' <value>The selected-state text color. The default is <see cref="Color.RoyalBlue"/>.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue(GetType(Color), "RoyalBlue")>
    <Description("Defines the text color used while a lookup item is selected.")>
    Public Property SelectedItemForeColor As Color
        Get
            Return _SelectedItemForeColor
        End Get
        Set(value As Color)
            If value.IsEmpty OrElse value.A < Byte.MaxValue Then Throw New ArgumentException("SelectedItemForeColor must be an opaque, non-empty color.", NameOf(value))
            If _SelectedItemForeColor = value Then Return
            _SelectedItemForeColor = value
            ApplySelectionAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the message displayed while a search is running.
    ''' </summary>
    <Category("AsyncLookupBox")>
    <DefaultValue("Searching...")>
    <Description("Defines the message displayed while an asynchronous lookup is running.")>
    Public Property LoadingText As String
        Get
            Return _LoadingText
        End Get
        Set(value As String)
            _LoadingText = If(value, String.Empty)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the message displayed when a search returns no results.
    ''' </summary>
    <Category("AsyncLookupBox")>
    <DefaultValue("No results found.")>
    <Description("Defines the message displayed when an asynchronous lookup returns no results.")>
    Public Property NoResultsText As String
        Get
            Return _NoResultsText
        End Get
        Set(value As String)
            _NoResultsText = If(value, String.Empty)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the user-facing message displayed when a search fails.
    ''' </summary>
    <Category("AsyncLookupBox")>
    <DefaultValue("Unable to load results.")>
    <Description("Defines the user-facing message displayed when an asynchronous lookup fails.")>
    Public Property SearchErrorText As String
        Get
            Return _SearchErrorText
        End Get
        Set(value As String)
            _SearchErrorText = If(value, String.Empty)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the message displayed when no search task is supplied by the application.
    ''' </summary>
    <Category("AsyncLookupBox")>
    <DefaultValue("Search provider is not configured.")>
    <Description("Defines the message displayed when the application does not supply a lookup task.")>
    Public Property SearchNotConfiguredText As String
        Get
            Return _SearchNotConfiguredText
        End Get
        Set(value As String)
            _SearchNotConfiguredText = If(value, String.Empty)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the formatted message displayed before <see cref="MinimumCharacters"/> is reached.
    ''' </summary>
    ''' <remarks>The format string receives the remaining character count as argument <c>{0}</c>.</remarks>
    <Category("AsyncLookupBox")>
    <DefaultValue("Enter {0} more character(s).")>
    <Description("Defines the formatted message displayed before the minimum character count is reached.")>
    Public Property CharactersRemainingText As String
        Get
            Return _CharactersRemainingText
        End Get
        Set(value As String)
            _CharactersRemainingText = If(value, String.Empty)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the fallback header used when <see cref="Columns"/> is empty.
    ''' </summary>
    <Category("AsyncLookupBox")>
    <DefaultValue("Result")>
    <Description("Defines the fallback result-column header used when no explicit columns are configured.")>
    Public Property ResultColumnHeaderText As String
        Get
            Return _ResultColumnHeaderText
        End Get
        Set(value As String)
            _ResultColumnHeaderText = If(value, String.Empty)
            RefreshVisibleResults()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the result-list background color.
    ''' </summary>
    <Category("AsyncLookupBox")>
    <DefaultValue(GetType(Color), "White")>
    <Description("Defines the result-list background color.")>
    Public Property DropDownBackColor As Color
        Get
            Return _DropDownBackColor
        End Get
        Set(value As Color)
            If value.IsEmpty OrElse value.A < Byte.MaxValue Then Throw New ArgumentException("DropDownBackColor must be an opaque, non-empty color.", NameOf(value))
            _DropDownBackColor = value
            RefreshDropDownAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the result-list text color.
    ''' </summary>
    <Category("AsyncLookupBox")>
    <DefaultValue(GetType(Color), "ControlText")>
    <Description("Defines the result-list text color.")>
    Public Property DropDownForeColor As Color
        Get
            Return _DropDownForeColor
        End Get
        Set(value As Color)
            _DropDownForeColor = value
            RefreshDropDownAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the result drop-down border color.
    ''' </summary>
    <Category("AsyncLookupBox")>
    <DefaultValue(GetType(Color), "ControlDark")>
    <Description("Defines the result drop-down border color.")>
    Public Property DropDownBorderColor As Color
        Get
            Return _DropDownBorderColor
        End Get
        Set(value As Color)
            _DropDownBorderColor = value
            RefreshDropDownAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the selected-row background color.
    ''' </summary>
    <Category("AsyncLookupBox")>
    <DefaultValue(GetType(Color), "Highlight")>
    <Description("Defines the selected-row background color.")>
    Public Property SelectionBackColor As Color
        Get
            Return _SelectionBackColor
        End Get
        Set(value As Color)
            _SelectionBackColor = value
            RefreshDropDownAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the selected-row text color.
    ''' </summary>
    <Category("AsyncLookupBox")>
    <DefaultValue(GetType(Color), "HighlightText")>
    <Description("Defines the selected-row text color.")>
    Public Property SelectionForeColor As Color
        Get
            Return _SelectionForeColor
        End Get
        Set(value As Color)
            _SelectionForeColor = value
            RefreshDropDownAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets the currently selected result object.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property SelectedItem As Object
        Get
            Return _SelectedItem
        End Get
    End Property
    ''' <summary>
    ''' Gets the value resolved from <see cref="SelectedItem"/> through <see cref="ValueMember"/>.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property SelectedValue As Object
        Get
            Return _SelectedValue
        End Get
    End Property
    ''' <summary>
    ''' Gets a value indicating whether a result object is selected.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property HasSelection As Boolean
        Get
            Return _SelectedItem IsNot Nothing
        End Get
    End Property
    ''' <summary>
    ''' Gets a value indicating whether a current search is running.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property IsSearching As Boolean
        Get
            Return _IsSearching
        End Get
    End Property
    ''' <summary>
    ''' Gets a value indicating whether the result drop-down is open.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property IsDropDownOpen As Boolean
        Get
            Return _DropDown IsNot Nothing AndAlso _DropDown.Visible
        End Get
    End Property
    ''' <summary>
    ''' Gets the objects retained by the most recent successful search.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property Results As IReadOnlyList(Of Object)
        Get
            Return _Results
        End Get
    End Property
End Class
