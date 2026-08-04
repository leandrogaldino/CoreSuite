Imports System.ComponentModel
''' <summary>
''' Defines a result column displayed by an <see cref="AsyncLookupBox"/> drop-down.
''' </summary>
''' <remarks>
''' <see cref="PropertyName"/> accepts a normal property name, a <see cref="DataColumn.ColumnName"/>, a dictionary key, or a nested path such as <c>Category.Name</c>.
''' </remarks>
<TypeConverter(GetType(ExpandableObjectConverter))>
<DefaultProperty("PropertyName")>
<Description("Defines a property displayed as a column in an AsyncLookupBox result list.")>
Public Class AsyncLookupColumn
    Private _PropertyName As String = String.Empty
    Private _HeaderText As String = String.Empty
    Private _Width As Integer = 120
    Private _MinimumWidth As Integer = 5
    Private _AutoSizeMode As DataGridViewAutoSizeColumnMode = DataGridViewAutoSizeColumnMode.None
    Private _FillWeight As Single = 100.0F
    Private _Format As String = String.Empty
    Private _NullValue As String = String.Empty
    Private _Visible As Boolean = True
    Friend Event Changed As EventHandler
    ''' <summary>
    ''' Initializes a new instance of the <see cref="AsyncLookupColumn"/> class.
    ''' </summary>
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="AsyncLookupColumn"/> class for the specified property.
    ''' </summary>
    ''' <param name="PropertyName">The property path whose value is displayed.</param>
    Public Sub New(PropertyName As String)
        Me.PropertyName = PropertyName
    End Sub
    ''' <summary>
    ''' Gets or sets the property path whose value is displayed.
    ''' </summary>
    ''' <value>A property name, data-column name, dictionary key, or nested property path.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue("")>
    <Description("Defines the property path whose value is displayed in this result column.")>
    Public Property PropertyName As String
        Get
            Return _PropertyName
        End Get
        Set(value As String)
            Dim NormalizedValue As String = If(value, String.Empty).Trim()
            If String.Equals(_PropertyName, NormalizedValue, StringComparison.Ordinal) Then Return
            _PropertyName = NormalizedValue
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the text displayed in the column header.
    ''' </summary>
    ''' <value>The header text, or an empty string to use <see cref="PropertyName"/>.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue("")>
    <Description("Defines the text displayed in the result column header.")>
    Public Property HeaderText As String
        Get
            Return _HeaderText
        End Get
        Set(value As String)
            Dim NormalizedValue As String = If(value, String.Empty)
            If String.Equals(_HeaderText, NormalizedValue, StringComparison.Ordinal) Then Return
            _HeaderText = NormalizedValue
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the width of the result column when automatic sizing is disabled.
    ''' </summary>
    ''' <value>The width in pixels. The default is <c>120</c>.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue(120)>
    <Description("Defines the result column width when automatic sizing is disabled.")>
    Public Property Width As Integer
        Get
            Return _Width
        End Get
        Set(value As Integer)
            If value < 2 Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "Width must be at least 2 pixels.")
            If _Width = value Then Return
            _Width = value
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the minimum width of the result column.
    ''' </summary>
    ''' <value>The minimum width in pixels. The default is <c>5</c>.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue(5)>
    <Description("Defines the minimum result column width.")>
    Public Property MinimumWidth As Integer
        Get
            Return _MinimumWidth
        End Get
        Set(value As Integer)
            If value < 2 Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "MinimumWidth must be at least 2 pixels.")
            If _MinimumWidth = value Then Return
            _MinimumWidth = value
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets how the result column adjusts its width.
    ''' </summary>
    ''' <value>A <see cref="DataGridViewAutoSizeColumnMode"/> value. The default is <see cref="DataGridViewAutoSizeColumnMode.None"/>.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue(GetType(DataGridViewAutoSizeColumnMode), NameOf(DataGridViewAutoSizeColumnMode.None))>
    <Description("Defines how the result column automatically adjusts its width.")>
    Public Property AutoSizeMode As DataGridViewAutoSizeColumnMode
        Get
            Return _AutoSizeMode
        End Get
        Set(value As DataGridViewAutoSizeColumnMode)
            If value = DataGridViewAutoSizeColumnMode.NotSet Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "NotSet is not valid for a result column.")
            If _AutoSizeMode = value Then Return
            _AutoSizeMode = value
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the relative width used when <see cref="AutoSizeMode"/> is <see cref="DataGridViewAutoSizeColumnMode.Fill"/>.
    ''' </summary>
    ''' <value>The relative fill weight. The default is <c>100</c>.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue(GetType(Single), "100")>
    <Description("Defines the relative width used when the result column fills available space.")>
    Public Property FillWeight As Single
        Get
            Return _FillWeight
        End Get
        Set(value As Single)
            If value <= 0.0F OrElse value > 65535.0F Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "FillWeight must be greater than zero and no greater than 65535.")
            If _FillWeight.Equals(value) Then Return
            _FillWeight = value
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the format string applied to values in the result column.
    ''' </summary>
    ''' <value>A standard or custom composite-format string.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue("")>
    <Description("Defines the format string applied to values in the result column.")>
    Public Property Format As String
        Get
            Return _Format
        End Get
        Set(value As String)
            Dim NormalizedValue As String = If(value, String.Empty)
            If String.Equals(_Format, NormalizedValue, StringComparison.Ordinal) Then Return
            _Format = NormalizedValue
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the text displayed when a result value is <see langword="Nothing"/> or <see cref="DBNull.Value"/>.
    ''' </summary>
    ''' <value>The replacement text for null values.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue("")>
    <Description("Defines the text displayed for null values in the result column.")>
    Public Property NullValue As String
        Get
            Return _NullValue
        End Get
        Set(value As String)
            Dim NormalizedValue As String = If(value, String.Empty)
            If String.Equals(_NullValue, NormalizedValue, StringComparison.Ordinal) Then Return
            _NullValue = NormalizedValue
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the result column is visible.
    ''' </summary>
    ''' <value><see langword="True"/> to display the column; otherwise, <see langword="False"/>. The default is <see langword="True"/>.</value>
    <Category("AsyncLookupBox")>
    <DefaultValue(True)>
    <Description("Determines whether the result column is visible.")>
    Public Property Visible As Boolean
        Get
            Return _Visible
        End Get
        Set(value As Boolean)
            If _Visible = value Then Return
            _Visible = value
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Returns a readable representation of this result-column configuration.
    ''' </summary>
    ''' <returns>The configured header or property name.</returns>
    Public Overrides Function ToString() As String
        If Not String.IsNullOrWhiteSpace(HeaderText) Then Return HeaderText
        If Not String.IsNullOrWhiteSpace(PropertyName) Then Return PropertyName
        Return MyBase.ToString()
    End Function
    Private Sub OnChanged()
        RaiseEvent Changed(Me, EventArgs.Empty)
    End Sub
End Class
