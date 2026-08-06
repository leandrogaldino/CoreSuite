Imports System.ComponentModel
''' <summary>
''' Describes one navigation button and the <see cref="UserControl"/> displayed when that button is selected.
''' </summary>
<DefaultProperty("Key")>
<Description("Defines one lazily created UserControl page displayed by a NavigationView.")>
<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class NavigationPage
    Private _Key As String = String.Empty
    Private _Text As String = String.Empty
    Private _Image As Image
    Private _ToolTipText As String = String.Empty
    Private _AccessibleName As String = String.Empty
    Private _ControlType As Type
    Private _CacheMode As NavigationPageCacheMode = NavigationPageCacheMode.KeepAlive
    Private _Visible As Boolean = True
    Private _Enabled As Boolean = True
    Private _Tag As Object
    Private _Factory As Func(Of UserControl)
    Private _ControlInstance As UserControl
    Private _Owner As NavigationPageCollection
    Friend Event Changed As EventHandler
    ''' <summary>
    ''' Initializes a new empty instance of the <see cref="NavigationPage"/> class.
    ''' </summary>
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Initializes a page that creates a control of the specified type.
    ''' </summary>
    ''' <param name="Key">The unique key used to locate the page.</param>
    ''' <param name="Text">The text displayed by the navigation button.</param>
    ''' <param name="ControlType">A non-abstract type derived from <see cref="UserControl"/> with a parameterless constructor.</param>
    ''' <param name="Image">The optional image displayed by the navigation button.</param>
    Public Sub New(Key As String, Text As String, ControlType As Type, Optional Image As Image = Nothing)
        Me.Key = Key
        Me.Text = Text
        Me.ControlType = ControlType
        Me.Image = Image
    End Sub
    ''' <summary>
    ''' Initializes a page that creates its control by invoking the specified factory.
    ''' </summary>
    ''' <param name="Key">The unique key used to locate the page.</param>
    ''' <param name="Text">The text displayed by the navigation button.</param>
    ''' <param name="Image">The optional image displayed by the navigation button.</param>
    ''' <param name="Factory">The factory that creates the page control.</param>
    Public Sub New(Key As String, Text As String, Image As Image, Factory As Func(Of UserControl))
        Me.Key = Key
        Me.Text = Text
        Me.Image = Image
        Me.Factory = Factory
    End Sub
    ''' <summary>
    ''' Gets or sets the unique key used to locate the page.
    ''' </summary>
    ''' <value>A case-insensitive key unique within the owning collection.</value>
    <Category("NavigationPage")>
    <Description("Specifies the unique key used to locate this page.")>
    <DefaultValue("")>
    <NotifyParentProperty(True)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property Key As String
        Get
            Return _Key
        End Get
        Set(Value As String)
            Dim NormalizedValue As String = If(Value, String.Empty).Trim()
            If String.Equals(_Key, NormalizedValue, StringComparison.Ordinal) Then Return
            If _Owner IsNot Nothing Then _Owner.ValidateProposedKey(NormalizedValue, Me)
            _Key = NormalizedValue
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the text displayed by the navigation button.
    ''' </summary>
    ''' <value>The navigation button text.</value>
    <Category("NavigationPage")>
    <Description("Specifies the text displayed by the navigation button.")>
    <DefaultValue("")>
    <NotifyParentProperty(True)>
    Public Property Text As String
        Get
            Return _Text
        End Get
        Set(Value As String)
            Dim NormalizedValue As String = If(Value, String.Empty)
            If String.Equals(_Text, NormalizedValue, StringComparison.Ordinal) Then Return
            _Text = NormalizedValue
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the image displayed by the navigation button.
    ''' </summary>
    ''' <value>The page image, or <see langword="Nothing"/>.</value>
    <Category("NavigationPage")>
    <Description("Specifies the image displayed by the navigation button.")>
    <NotifyParentProperty(True)>
    Public Property Image As Image
        Get
            Return _Image
        End Get
        Set(Value As Image)
            If ReferenceEquals(_Image, Value) Then Return
            _Image = Value
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the tooltip displayed for the navigation button.
    ''' </summary>
    ''' <value>The tooltip text, or an empty string to disable it.</value>
    <Category("NavigationPage")>
    <Description("Specifies the tooltip displayed for the navigation button.")>
    <DefaultValue("")>
    <NotifyParentProperty(True)>
    Public Property ToolTipText As String
        Get
            Return _ToolTipText
        End Get
        Set(Value As String)
            Dim NormalizedValue As String = If(Value, String.Empty)
            If String.Equals(_ToolTipText, NormalizedValue, StringComparison.Ordinal) Then Return
            _ToolTipText = NormalizedValue
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the name announced by accessibility clients for the navigation button.
    ''' </summary>
    ''' <value>An accessible name, or an empty string to derive it from the page text or key.</value>
    <Category("NavigationPage")>
    <Description("Specifies the name announced by accessibility clients for the navigation button.")>
    <DefaultValue("")>
    <NotifyParentProperty(True)>
    Public Property AccessibleName As String
        Get
            Return _AccessibleName
        End Get
        Set(Value As String)
            Dim NormalizedValue As String = If(Value, String.Empty)
            If String.Equals(_AccessibleName, NormalizedValue, StringComparison.Ordinal) Then Return
            _AccessibleName = NormalizedValue
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the type of <see cref="UserControl"/> created for the page.
    ''' </summary>
    ''' <value>A non-abstract <see cref="UserControl"/> type with a parameterless constructor, or <see langword="Nothing"/> when a run-time factory is used.</value>
    <Category("NavigationPage")>
    <Description("Specifies the UserControl type created for this page. The type must provide a parameterless constructor.")>
    <TypeConverter(GetType(UserControlTypeConverter))>
    <NotifyParentProperty(True)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property ControlType As Type
        Get
            Return _ControlType
        End Get
        Set(Value As Type)
            If ReferenceEquals(_ControlType, Value) Then Return
            EnsureControlIsNotCreated(NameOf(ControlType))
            ValidateControlType(Value)
            _ControlType = Value
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets how the created page control is retained when navigation leaves the page.
    ''' </summary>
    ''' <value>One of the <see cref="NavigationPageCacheMode"/> values. The default is <see cref="NavigationPageCacheMode.KeepAlive"/>.</value>
    <Category("NavigationPage")>
    <Description("Specifies whether the created control is kept alive or recreated after navigation leaves the page.")>
    <DefaultValue(NavigationPageCacheMode.KeepAlive)>
    <NotifyParentProperty(True)>
    Public Property CacheMode As NavigationPageCacheMode
        Get
            Return _CacheMode
        End Get
        Set(Value As NavigationPageCacheMode)
            If Not [Enum].IsDefined(GetType(NavigationPageCacheMode), Value) Then Throw New InvalidEnumArgumentException(NameOf(Value), CInt(Value), GetType(NavigationPageCacheMode))
            If _CacheMode = Value Then Return
            _CacheMode = Value
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the page appears in the navigation pane.
    ''' </summary>
    ''' <value><see langword="True"/> to display the navigation button; otherwise, <see langword="False"/>.</value>
    <Category("NavigationPage")>
    <Description("Determines whether this page appears in the navigation pane.")>
    <DefaultValue(True)>
    <NotifyParentProperty(True)>
    Public Property Visible As Boolean
        Get
            Return _Visible
        End Get
        Set(Value As Boolean)
            If _Visible = Value Then Return
            _Visible = Value
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether navigation to the page is allowed.
    ''' </summary>
    ''' <value><see langword="True"/> to allow navigation; otherwise, <see langword="False"/>.</value>
    <Category("NavigationPage")>
    <Description("Determines whether navigation to this page is allowed.")>
    <DefaultValue(True)>
    <NotifyParentProperty(True)>
    Public Property Enabled As Boolean
        Get
            Return _Enabled
        End Get
        Set(Value As Boolean)
            If _Enabled = Value Then Return
            _Enabled = Value
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets application-defined data associated with the page.
    ''' </summary>
    ''' <value>An arbitrary object, or <see langword="Nothing"/>.</value>
    <Category("Data")>
    <Description("Stores application-defined data associated with this page.")>
    <TypeConverter(GetType(StringConverter))>
    <NotifyParentProperty(True)>
    Public Property Tag As Object
        Get
            Return _Tag
        End Get
        Set(Value As Object)
            If ReferenceEquals(_Tag, Value) Then Return
            _Tag = Value
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the run-time factory used to create the page control.
    ''' </summary>
    ''' <value>A factory that returns a new <see cref="UserControl"/>, or <see langword="Nothing"/> to use <see cref="ControlType"/>.</value>
    ''' <remarks>The factory is not serialized by the Windows Forms Designer and takes precedence over <see cref="ControlType"/>.</remarks>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property Factory As Func(Of UserControl)
        Get
            Return _Factory
        End Get
        Set(Value As Func(Of UserControl))
            If ReferenceEquals(_Factory, Value) Then Return
            EnsureControlIsNotCreated(NameOf(Factory))
            _Factory = Value
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets a value indicating whether the page currently owns a created control instance.
    ''' </summary>
    ''' <value><see langword="True"/> when a non-disposed control is cached; otherwise, <see langword="False"/>.</value>
    <Browsable(False)>
    Public ReadOnly Property IsCreated As Boolean
        Get
            Return _ControlInstance IsNot Nothing AndAlso Not _ControlInstance.IsDisposed
        End Get
    End Property
    ''' <summary>
    ''' Gets the currently created page control without creating it.
    ''' </summary>
    ''' <value>The cached <see cref="UserControl"/>, or <see langword="Nothing"/> when the page has not been created.</value>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property CachedControl As UserControl
        Get
            If _ControlInstance IsNot Nothing AndAlso _ControlInstance.IsDisposed Then _ControlInstance = Nothing
            Return _ControlInstance
        End Get
    End Property
    ''' <summary>
    ''' Returns a display name for collection editors and diagnostic output.
    ''' </summary>
    ''' <returns>The page text, key, or <c>(Page)</c> when neither value is configured.</returns>
    Public Overrides Function ToString() As String
        If Not String.IsNullOrWhiteSpace(_Text) Then Return _Text
        If Not String.IsNullOrWhiteSpace(_Key) Then Return _Key
        Return "(Page)"
    End Function
    Friend Property Owner As NavigationPageCollection
        Get
            Return _Owner
        End Get
        Set(Value As NavigationPageCollection)
            _Owner = Value
        End Set
    End Property
    Friend Function GetOrCreateControl() As UserControl
        If IsCreated Then Return _ControlInstance
        _ControlInstance = CreateNewControl()
        Return _ControlInstance
    End Function
    Friend Function CreateNewControl() As UserControl
        Dim Result As UserControl
        If _Factory IsNot Nothing Then
            Result = _Factory.Invoke()
        ElseIf _ControlType IsNot Nothing Then
            Result = TryCast(Activator.CreateInstance(_ControlType), UserControl)
        Else
            Throw New InvalidOperationException($"Navigation page '{_Key}' does not define a ControlType or Factory.")
        End If
        If Result Is Nothing Then Throw New InvalidOperationException($"The factory for navigation page '{_Key}' returned Nothing.")
        If Result.IsDisposed Then Throw New InvalidOperationException($"The factory for navigation page '{_Key}' returned a disposed control.")
        Return Result
    End Function
    Friend Sub AssignControl(Control As UserControl)
        _ControlInstance = Control
    End Sub
    Friend Function ReleaseControl() As Boolean
        Dim Control As UserControl = CachedControl
        _ControlInstance = Nothing
        If Control Is Nothing Then Return False
        If Control.Parent IsNot Nothing Then Control.Parent.Controls.Remove(Control)
        Control.Dispose()
        Return True
    End Function
    Private Shared Sub ValidateControlType(Value As Type)
        If Value Is Nothing Then Return
        If Not GetType(UserControl).IsAssignableFrom(Value) Then Throw New ArgumentException("ControlType must derive from UserControl.", NameOf(Value))
        If Value.IsAbstract Then Throw New ArgumentException("ControlType cannot be abstract.", NameOf(Value))
        If Value.GetConstructor(Type.EmptyTypes) Is Nothing Then Throw New ArgumentException("ControlType must provide a public parameterless constructor. Use Factory for controls that require constructor arguments.", NameOf(Value))
    End Sub
    Private Sub EnsureControlIsNotCreated(PropertyName As String)
        If IsCreated Then Throw New InvalidOperationException($"{PropertyName} cannot be changed while the page control is created. Call ClosePage before changing it.")
    End Sub
    Private Sub OnChanged()
        RaiseEvent Changed(Me, EventArgs.Empty)
    End Sub
End Class
