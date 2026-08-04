Imports System.ComponentModel
Imports System.Drawing.Design
Imports Microsoft.DotNet.DesignTools.Editors
Partial Public Class TextBoxActionPanel
    ''' <summary>
    ''' Gets or sets the existing text box enhanced by this component.
    ''' </summary>
    ''' <value>A <see cref="TextBoxBase"/> such as <see cref="TextBox"/> or any compatible derived CoreSuite control, or <see langword="Nothing"/>.</value>
    <Category(CategoryName)>
    <Description("Specifies the existing TextBoxBase control enhanced by this action panel.")>
    <RefreshProperties(RefreshProperties.All)>
    Public Property TargetControl As TextBoxBase
        Get
            Return _TargetControl
        End Get
        Set(Value As TextBoxBase)
            SetTargetControl(Value)
        End Set
    End Property
    ''' <summary>
    ''' Gets the ordered collection of image actions displayed for the target control.
    ''' </summary>
    ''' <value>The action collection. Its first visible item appears at the right edge and following items extend to the left.</value>
    <Category(CategoryName)>
    <Description("Contains the image actions displayed from the right edge of the target toward the left.")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Editor(GetType(CollectionEditor), GetType(UITypeEditor))>
    Public ReadOnly Property Actions As TextBoxActionCollection
        Get
            Return _Actions
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the component can display and execute actions.
    ''' </summary>
    ''' <value><see langword="True"/> to enable the component; otherwise, <see langword="False"/>.</value>
    <Category(CategoryName)>
    <Description("Determines whether the component can display and execute actions.")>
    <DefaultValue(True)>
    Public Property Enabled As Boolean
        Get
            Return _Enabled
        End Get
        Set(Value As Boolean)
            If _Enabled = Value Then Return
            _Enabled = Value
            If Not Value Then
                HidePanel()
            ElseIf _ShowOnFocus AndAlso _TargetControl IsNot Nothing AndAlso _TargetControl.Focused Then
                ShowPanel()
            End If
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the panel is displayed automatically when the target receives focus.
    ''' </summary>
    ''' <value><see langword="True"/> to display the panel on focus; otherwise, <see langword="False"/>.</value>
    <Category(CategoryName)>
    <Description("Determines whether the panel is displayed automatically when the target receives focus.")>
    <DefaultValue(True)>
    Public Property ShowOnFocus As Boolean
        Get
            Return _ShowOnFocus
        End Get
        Set(Value As Boolean)
            If _ShowOnFocus = Value Then Return
            _ShowOnFocus = Value
            If Value AndAlso _TargetControl IsNot Nothing AndAlso _TargetControl.Focused Then ShowPanel()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the panel is hidden automatically when the target loses focus.
    ''' </summary>
    ''' <value><see langword="True"/> to hide the panel on leave; otherwise, <see langword="False"/>.</value>
    <Category(CategoryName)>
    <Description("Determines whether the panel is hidden automatically when the target loses focus.")>
    <DefaultValue(True)>
    Public Property HideOnLeave As Boolean
        Get
            Return _HideOnLeave
        End Get
        Set(Value As Boolean)
            _HideOnLeave = Value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the preferred position of the panel relative to the target.
    ''' </summary>
    ''' <value>One of the <see cref="TextBoxActionPanelPlacement"/> values.</value>
    <Category(CategoryName)>
    <Description("Specifies whether the panel appears automatically, above, or below the target.")>
    <DefaultValue(TextBoxActionPanelPlacement.Auto)>
    Public Property Placement As TextBoxActionPanelPlacement
        Get
            Return _Placement
        End Get
        Set(Value As TextBoxActionPanelPlacement)
            If Not [Enum].IsDefined(GetType(TextBoxActionPanelPlacement), Value) Then Throw New InvalidEnumArgumentException(NameOf(Value), CInt(Value), GetType(TextBoxActionPanelPlacement))
            If _Placement = Value Then Return
            _Placement = Value
            RepositionPanel()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the square size, in pixels, of each action button.
    ''' </summary>
    ''' <value>A value from 16 through 64. The default is 24.</value>
    <Category(CategoryName)>
    <Description("Specifies the square size, in pixels, of each action button.")>
    <DefaultValue(24)>
    Public Property ButtonSize As Integer
        Get
            Return _ButtonSize
        End Get
        Set(Value As Integer)
            ValidateRange(NameOf(Value), Value, 16, 64)
            If _ButtonSize = Value Then Return
            _ButtonSize = Value
            RefreshPanel()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the space, in pixels, between adjacent action buttons.
    ''' </summary>
    ''' <value>A value from 0 through 16. The default is 0, which places the buttons directly beside one another.</value>
    <Category(CategoryName)>
    <Description("Specifies the space, in pixels, between adjacent action buttons.")>
    <DefaultValue(0)>
    Public Property ButtonSpacing As Integer
        Get
            Return _ButtonSpacing
        End Get
        Set(Value As Integer)
            ValidateRange(NameOf(Value), Value, 0, 16)
            If _ButtonSpacing = Value Then Return
            _ButtonSpacing = Value
            RefreshPanel()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the internal space, in pixels, between the popup border and its buttons.
    ''' </summary>
    ''' <value>A value from 0 through 16. The default is 0.</value>
    <Category(CategoryName)>
    <Description("Specifies the internal space, in pixels, between the popup border and its buttons.")>
    <DefaultValue(0)>
    Public Property PanelPadding As Integer
        Get
            Return _PanelPadding
        End Get
        Set(Value As Integer)
            ValidateRange(NameOf(Value), Value, 0, 16)
            If _PanelPadding = Value Then Return
            _PanelPadding = Value
            RefreshPanel()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the distance, in pixels, between the popup and the target control.
    ''' </summary>
    ''' <value>A value from 0 through 32. The default is 0, which attaches the popup edge directly to the target edge.</value>
    <Category(CategoryName)>
    <Description("Specifies the distance, in pixels, between the popup and the target control.")>
    <DefaultValue(0)>
    Public Property PanelOffset As Integer
        Get
            Return _PanelOffset
        End Get
        Set(Value As Integer)
            ValidateRange(NameOf(Value), Value, 0, 32)
            If _PanelOffset = Value Then Return
            _PanelOffset = Value
            RepositionPanel()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the area surrounding the action buttons is transparent.
    ''' </summary>
    ''' <value><see langword="True"/> to make the popup background transparent; otherwise, <see langword="False"/>. The default is <see langword="True"/>.</value>
    <Category(CategoryName)>
    <Description("Determines whether the area surrounding the action buttons is transparent.")>
    <DefaultValue(True)>
    Public Property TransparentBackground As Boolean
        Get
            Return _TransparentBackground
        End Get
        Set(Value As Boolean)
            If _TransparentBackground = Value Then Return
            _TransparentBackground = Value
            RefreshPanel()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether a one-pixel border is drawn around the popup.
    ''' </summary>
    ''' <value><see langword="True"/> to draw the border; otherwise, <see langword="False"/>. The default is <see langword="False"/>.</value>
    <Category(CategoryName)>
    <Description("Determines whether a one-pixel border is drawn around the popup.")>
    <DefaultValue(False)>
    Public Property ShowBorder As Boolean
        Get
            Return _ShowBorder
        End Get
        Set(Value As Boolean)
            If _ShowBorder = Value Then Return
            _ShowBorder = Value
            RefreshPanel()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the popup and table-layout background color.
    ''' </summary>
    ''' <value>The background color used when <see cref="TransparentBackground"/> is <see langword="False"/>. The default is <see cref="SystemColors.Window"/>.</value>
    <Category(CategoryName)>
    <Description("Specifies the popup background color used when TransparentBackground is False.")>
    <DefaultValue(GetType(Color), "Window")>
    Public Property PanelBackColor As Color
        Get
            Return _PanelBackColor
        End Get
        Set(Value As Color)
            If _PanelBackColor = Value Then Return
            _PanelBackColor = Value
            RefreshPanel()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the color of the one-pixel popup border.
    ''' </summary>
    ''' <value>The border color used when <see cref="ShowBorder"/> is <see langword="True"/>. The default is <see cref="SystemColors.ControlDark"/>.</value>
    <Category(CategoryName)>
    <Description("Specifies the popup border color used when ShowBorder is True.")>
    <DefaultValue(GetType(Color), "ControlDark")>
    Public Property BorderColor As Color
        Get
            Return _BorderColor
        End Get
        Set(Value As Color)
            If _BorderColor = Value Then Return
            _BorderColor = Value
            RefreshPanel()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the normal background color of each action button.
    ''' </summary>
    ''' <value>The normal button color. The default is <see cref="SystemColors.Window"/>.</value>
    <Category(CategoryName)>
    <Description("Specifies the normal background color of each action button.")>
    <DefaultValue(GetType(Color), "Window")>
    Public Property ButtonBackColor As Color
        Get
            Return _ButtonBackColor
        End Get
        Set(Value As Color)
            If _ButtonBackColor = Value Then Return
            _ButtonBackColor = Value
            RefreshPanel()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the action-button background color used while the pointer is over a button.
    ''' </summary>
    ''' <value>The hover color. The default is <see cref="SystemColors.ControlLight"/>.</value>
    <Category(CategoryName)>
    <Description("Specifies the action-button background color used while the pointer is over a button.")>
    <DefaultValue(GetType(Color), "ControlLight")>
    Public Property ButtonHoverBackColor As Color
        Get
            Return _ButtonHoverBackColor
        End Get
        Set(Value As Color)
            If _ButtonHoverBackColor = Value Then Return
            _ButtonHoverBackColor = Value
            RefreshPanel()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the action-button background color used while a button is pressed.
    ''' </summary>
    ''' <value>The pressed color. The default is <see cref="SystemColors.ControlDark"/>.</value>
    <Category(CategoryName)>
    <Description("Specifies the action-button background color used while a button is pressed.")>
    <DefaultValue(GetType(Color), "ControlDark")>
    Public Property ButtonPressedBackColor As Color
        Get
            Return _ButtonPressedBackColor
        End Get
        Set(Value As Color)
            If _ButtonPressedBackColor = Value Then Return
            _ButtonPressedBackColor = Value
            RefreshPanel()
        End Set
    End Property
    Private Shared Sub ValidateRange(PropertyName As String, Value As Integer, Minimum As Integer, Maximum As Integer)
        If Value < Minimum OrElse Value > Maximum Then Throw New ArgumentOutOfRangeException(PropertyName, Value, $"The value must be between {Minimum} and {Maximum}.")
    End Sub
End Class
