Imports System.ComponentModel
Imports System.Drawing.Design
Partial Public Class NavigationView
    ''' <summary>
    ''' Gets the ordered collection of pages displayed and managed by the control.
    ''' </summary>
    ''' <value>The navigation page collection.</value>
    <Category(CategoryName)>
    <Description("Contains the pages displayed and managed by this NavigationView.")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Editor(GetType(NavigationPageCollectionEditor), GetType(UITypeEditor))>
    Public ReadOnly Property Pages As NavigationPageCollection
        Get
            Return _Pages
        End Get
    End Property
    ''' <summary>
    ''' Gets the page currently displayed in the content area.
    ''' </summary>
    ''' <value>The selected page, or <see langword="Nothing"/> when no page is displayed.</value>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property SelectedPage As NavigationPage
        Get
            Return _SelectedPage
        End Get
    End Property
    ''' <summary>
    ''' Gets the key of the page currently displayed in the content area.
    ''' </summary>
    ''' <value>The selected page key, or an empty string when no page is displayed.</value>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property SelectedPageKey As String
        Get
            Return If(_SelectedPage Is Nothing, String.Empty, _SelectedPage.Key)
        End Get
    End Property
    ''' <summary>
    ''' Gets the currently displayed <see cref="UserControl"/>.
    ''' </summary>
    ''' <value>The selected page control, or <see langword="Nothing"/> when no page is displayed.</value>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property SelectedControl As UserControl
        Get
            Return If(_SelectedPage Is Nothing, Nothing, _SelectedPage.CachedControl)
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the first visible enabled page is opened automatically after loading.
    ''' </summary>
    ''' <value><see langword="True"/> to navigate automatically; otherwise, <see langword="False"/>.</value>
    <Category(CategoryName)>
    <Description("Determines whether the first visible enabled page is opened automatically after loading.")>
    <DefaultValue(True)>
    Public Property AutoNavigateFirstPage As Boolean
        Get
            Return _AutoNavigateFirstPage
        End Get
        Set(Value As Boolean)
            If _AutoNavigateFirstPage = Value Then Return
            _AutoNavigateFirstPage = Value
            If Value AndAlso _IsLoaded AndAlso Not IsInDesignMode AndAlso _SelectedPage Is Nothing Then NavigateFirstAvailablePage()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the edge on which the navigation pane is displayed.
    ''' </summary>
    ''' <value>One of the <see cref="NavigationPanePosition"/> values.</value>
    <Category(CategoryName)>
    <Description("Specifies the edge on which the navigation pane is displayed.")>
    <DefaultValue(NavigationPanePosition.Left)>
    Public Property NavigationPosition As NavigationPanePosition
        Get
            Return _NavigationPosition
        End Get
        Set(Value As NavigationPanePosition)
            If Not [Enum].IsDefined(GetType(NavigationPanePosition), Value) Then Throw New InvalidEnumArgumentException(NameOf(Value), CInt(Value), GetType(NavigationPanePosition))
            If _NavigationPosition = Value Then Return
            _NavigationPosition = Value
            _NavigationPanel.Dock = If(Value = NavigationPanePosition.Left, DockStyle.Left, DockStyle.Right)
            ApplyButtonAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the width, in pixels, of the navigation pane.
    ''' </summary>
    ''' <value>A value from 80 through 600. The default is 220.</value>
    <Category(CategoryName)>
    <Description("Specifies the width, in pixels, of the navigation pane.")>
    <DefaultValue(220)>
    Public Property NavigationWidth As Integer
        Get
            Return _NavigationWidth
        End Get
        Set(Value As Integer)
            ValidateRange(NameOf(Value), Value, 80, 600)
            If _NavigationWidth = Value Then Return
            _NavigationWidth = Value
            _NavigationPanel.Width = Value
            UpdateButtonLayout()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the height, in pixels, of each navigation button.
    ''' </summary>
    ''' <value>A value from 24 through 128. The default is 44.</value>
    <Category(CategoryName)>
    <Description("Specifies the height, in pixels, of each navigation button.")>
    <DefaultValue(44)>
    Public Property ButtonHeight As Integer
        Get
            Return _ButtonHeight
        End Get
        Set(Value As Integer)
            ValidateRange(NameOf(Value), Value, 24, 128)
            If _ButtonHeight = Value Then Return
            _ButtonHeight = Value
            UpdateButtonLayout()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the vertical space, in pixels, between navigation buttons.
    ''' </summary>
    ''' <value>A value from 0 through 32. The default is 2.</value>
    <Category(CategoryName)>
    <Description("Specifies the vertical space, in pixels, between navigation buttons.")>
    <DefaultValue(2)>
    Public Property ButtonSpacing As Integer
        Get
            Return _ButtonSpacing
        End Get
        Set(Value As Integer)
            ValidateRange(NameOf(Value), Value, 0, 32)
            If _ButtonSpacing = Value Then Return
            _ButtonSpacing = Value
            UpdateButtonLayout()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the internal spacing around all navigation buttons.
    ''' </summary>
    ''' <value>The navigation pane padding. The default is 8 pixels on every side.</value>
    <Category(CategoryName)>
    <Description("Specifies the internal spacing around all navigation buttons.")>
    <DefaultValue(GetType(Padding), "8, 8, 8, 8")>
    Public Property NavigationPadding As Padding
        Get
            Return _NavigationPadding
        End Get
        Set(Value As Padding)
            If _NavigationPadding = Value Then Return
            _NavigationPadding = Value
            _NavigationFlow.Padding = Value
            UpdateButtonLayout()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the internal spacing used to lay out image and text inside each navigation button.
    ''' </summary>
    ''' <value>The button content padding. The default is 12 pixels horizontally.</value>
    <Category(CategoryName)>
    <Description("Specifies the internal spacing used to lay out image and text inside each navigation button.")>
    <DefaultValue(GetType(Padding), "12, 0, 12, 0")>
    Public Property ButtonPadding As Padding
        Get
            Return _ButtonPadding
        End Get
        Set(Value As Padding)
            If _ButtonPadding = Value Then Return
            _ButtonPadding = Value
            ApplyButtonAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the internal spacing between the content-area edges and the displayed page.
    ''' </summary>
    ''' <value>The content padding. The default is zero.</value>
    <Category(CategoryName)>
    <Description("Specifies the internal spacing between the content-area edges and the displayed page.")>
    <DefaultValue(GetType(Padding), "0, 0, 0, 0")>
    Public Property ContentPadding As Padding
        Get
            Return _ContentPadding
        End Get
        Set(Value As Padding)
            If _ContentPadding = Value Then Return
            _ContentPadding = Value
            _ContentPanel.Padding = Value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the size used to draw each page image.
    ''' </summary>
    ''' <value>The image size. The default is 20 by 20 pixels.</value>
    <Category(CategoryName)>
    <Description("Specifies the size used to draw each page image.")>
    <DefaultValue(GetType(Size), "20, 20")>
    Public Property ImageSize As Size
        Get
            Return _ImageSize
        End Get
        Set(Value As Size)
            If Value.Width < 0 OrElse Value.Width > 64 Then Throw New ArgumentOutOfRangeException(NameOf(Value), Value, "Image width must be between 0 and 64 pixels.")
            If Value.Height < 0 OrElse Value.Height > 64 Then Throw New ArgumentOutOfRangeException(NameOf(Value), Value, "Image height must be between 0 and 64 pixels.")
            If _ImageSize = Value Then Return
            _ImageSize = Value
            ApplyButtonAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the width, in pixels, of the selected-page indicator.
    ''' </summary>
    ''' <value>A value from 0 through 16. The default is 4.</value>
    <Category(CategoryName)>
    <Description("Specifies the width, in pixels, of the selected-page indicator.")>
    <DefaultValue(4)>
    Public Property SelectedIndicatorWidth As Integer
        Get
            Return _SelectedIndicatorWidth
        End Get
        Set(Value As Integer)
            ValidateRange(NameOf(Value), Value, 0, 16)
            If _SelectedIndicatorWidth = Value Then Return
            _SelectedIndicatorWidth = Value
            ApplyButtonAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether page images are displayed.
    ''' </summary>
    ''' <value><see langword="True"/> to display images; otherwise, <see langword="False"/>.</value>
    <Category(CategoryName)>
    <Description("Determines whether page images are displayed.")>
    <DefaultValue(True)>
    Public Property ShowImages As Boolean
        Get
            Return _ShowImages
        End Get
        Set(Value As Boolean)
            If _ShowImages = Value Then Return
            _ShowImages = Value
            ApplyButtonAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether page tooltips are displayed.
    ''' </summary>
    ''' <value><see langword="True"/> to display tooltips; otherwise, <see langword="False"/>.</value>
    <Category(CategoryName)>
    <Description("Determines whether page tooltips are displayed.")>
    <DefaultValue(True)>
    Public Property ShowToolTips As Boolean
        Get
            Return _ShowToolTips
        End Get
        Set(Value As Boolean)
            If _ShowToolTips = Value Then Return
            _ShowToolTips = Value
            ApplyButtonToolTips()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the background color of the navigation pane.
    ''' </summary>
    ''' <value>The navigation background color.</value>
    <Category(CategoryName)>
    <Description("Specifies the background color of the navigation pane.")>
    <DefaultValue(GetType(Color), "Control")>
    Public Property NavigationBackColor As Color
        Get
            Return _NavigationBackColor
        End Get
        Set(Value As Color)
            If _NavigationBackColor = Value Then Return
            _NavigationBackColor = Value
            _NavigationPanel.BackColor = Value
            _NavigationFlow.BackColor = Value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the background color of the content area.
    ''' </summary>
    ''' <value>The content background color.</value>
    <Category(CategoryName)>
    <Description("Specifies the background color of the content area.")>
    <DefaultValue(GetType(Color), "Window")>
    Public Property ContentBackColor As Color
        Get
            Return _ContentBackColor
        End Get
        Set(Value As Color)
            If _ContentBackColor = Value Then Return
            _ContentBackColor = Value
            _ContentPanel.BackColor = Value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the normal navigation-button background color.
    ''' </summary>
    ''' <value>The normal button background color.</value>
    <Category(CategoryName)>
    <Description("Specifies the normal navigation-button background color.")>
    <DefaultValue(GetType(Color), "Control")>
    Public Property ButtonBackColor As Color
        Get
            Return _ButtonBackColor
        End Get
        Set(Value As Color)
            If _ButtonBackColor = Value Then Return
            _ButtonBackColor = Value
            ApplyButtonAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the navigation-button background color used while the pointer is over it.
    ''' </summary>
    ''' <value>The hover button background color.</value>
    <Category(CategoryName)>
    <Description("Specifies the navigation-button background color used while the pointer is over it.")>
    <DefaultValue(GetType(Color), "ControlLight")>
    Public Property ButtonHoverBackColor As Color
        Get
            Return _ButtonHoverBackColor
        End Get
        Set(Value As Color)
            If _ButtonHoverBackColor = Value Then Return
            _ButtonHoverBackColor = Value
            ApplyButtonAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the normal navigation-button text color.
    ''' </summary>
    ''' <value>The normal button text color.</value>
    <Category(CategoryName)>
    <Description("Specifies the normal navigation-button text color.")>
    <DefaultValue(GetType(Color), "ControlText")>
    Public Property ButtonForeColor As Color
        Get
            Return _ButtonForeColor
        End Get
        Set(Value As Color)
            If _ButtonForeColor = Value Then Return
            _ButtonForeColor = Value
            ApplyButtonAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the selected navigation-button background color.
    ''' </summary>
    ''' <value>The selected button background color.</value>
    <Category(CategoryName)>
    <Description("Specifies the selected navigation-button background color.")>
    <DefaultValue(GetType(Color), "Highlight")>
    Public Property SelectedButtonBackColor As Color
        Get
            Return _SelectedButtonBackColor
        End Get
        Set(Value As Color)
            If _SelectedButtonBackColor = Value Then Return
            _SelectedButtonBackColor = Value
            ApplyButtonAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the selected navigation-button text color.
    ''' </summary>
    ''' <value>The selected button text color.</value>
    <Category(CategoryName)>
    <Description("Specifies the selected navigation-button text color.")>
    <DefaultValue(GetType(Color), "HighlightText")>
    Public Property SelectedButtonForeColor As Color
        Get
            Return _SelectedButtonForeColor
        End Get
        Set(Value As Color)
            If _SelectedButtonForeColor = Value Then Return
            _SelectedButtonForeColor = Value
            ApplyButtonAppearance()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the color of the selected-page indicator.
    ''' </summary>
    ''' <value>The selected indicator color.</value>
    <Category(CategoryName)>
    <Description("Specifies the color of the selected-page indicator.")>
    <DefaultValue(GetType(Color), "Highlight")>
    Public Property SelectedIndicatorColor As Color
        Get
            Return _SelectedIndicatorColor
        End Get
        Set(Value As Color)
            If _SelectedIndicatorColor = Value Then Return
            _SelectedIndicatorColor = Value
            ApplyButtonAppearance()
        End Set
    End Property
    Private Shared Sub ValidateRange(PropertyName As String, Value As Integer, Minimum As Integer, Maximum As Integer)
        If Value < Minimum OrElse Value > Maximum Then Throw New ArgumentOutOfRangeException(PropertyName, Value, $"The value must be between {Minimum} and {Maximum}.")
    End Sub
End Class
