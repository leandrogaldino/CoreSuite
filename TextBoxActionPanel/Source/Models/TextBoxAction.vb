Imports System.ComponentModel
''' <summary>
''' Describes one image button displayed by a <see cref="TextBoxActionPanel"/>.
''' </summary>
''' <remarks>
''' An action can be handled through <see cref="TextBoxActionPanel.ActionClicked"/> or by assigning a delegate to <see cref="ClickHandler"/>.
''' </remarks>
<DefaultProperty("Key")>
<Description("Defines one configurable image action displayed by a TextBoxActionPanel.")>
<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class TextBoxAction
    Private _Key As String = String.Empty
    Private _Image As Image
    Private _ToolTipText As String = String.Empty
    Private _AccessibleName As String = String.Empty
    Private _Visible As Boolean = True
    Private _Enabled As Boolean = True
    Private _ClickHandler As System.Action(Of TextBoxActionClickEventArgs)
    Friend Event Changed As EventHandler
    ''' <summary>
    ''' Initializes a new instance of the <see cref="TextBoxAction"/> class.
    ''' </summary>
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="TextBoxAction"/> class with its primary presentation values.
    ''' </summary>
    ''' <param name="Key">The identifier used to distinguish the action.</param>
    ''' <param name="Image">The image displayed by the action button.</param>
    ''' <param name="ToolTipText">The tooltip displayed when the pointer rests over the button.</param>
    Public Sub New(Key As String, Image As Image, ToolTipText As String)
        Me.Key = Key
        Me.Image = Image
        Me.ToolTipText = ToolTipText
    End Sub
    ''' <summary>
    ''' Gets or sets the identifier used to locate and distinguish the action.
    ''' </summary>
    ''' <value>A developer-defined identifier such as <c>View</c>, <c>Search</c>, or <c>Create</c>.</value>
    <Category("TextBoxAction")>
    <Description("Specifies the identifier used to locate and distinguish this action.")>
    <DefaultValue("")>
    Public Property Key As String
        Get
            Return _Key
        End Get
        Set(Value As String)
            Dim NormalizedValue As String = If(Value, String.Empty)
            If String.Equals(_Key, NormalizedValue, StringComparison.Ordinal) Then Return
            _Key = NormalizedValue
            OnChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the image displayed by the action button.
    ''' </summary>
    ''' <value>The action image. A 16-by-16 pixel image is recommended.</value>
    <Category("TextBoxAction")>
    <Description("Specifies the image displayed by the action button. A 16-by-16 pixel image is recommended.")>
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
    ''' Gets or sets the text displayed when the pointer rests over the action button.
    ''' </summary>
    ''' <value>The tooltip text, or an empty string to disable the tooltip.</value>
    <Category("TextBoxAction")>
    <Description("Specifies the text displayed when the pointer rests over the action button.")>
    <DefaultValue("")>
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
    ''' Gets or sets the name announced by accessibility clients for the action button.
    ''' </summary>
    ''' <value>An accessible name, or an empty string to derive it from the tooltip or key.</value>
    <Category("TextBoxAction")>
    <Description("Specifies the name announced by accessibility clients for the action button.")>
    <DefaultValue("")>
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
    ''' Gets or sets a value indicating whether the action button is displayed.
    ''' </summary>
    ''' <value><see langword="True"/> to display the button; otherwise, <see langword="False"/>.</value>
    <Category("TextBoxAction")>
    <Description("Determines whether the action button is displayed.")>
    <DefaultValue(True)>
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
    ''' Gets or sets a value indicating whether the action can be executed.
    ''' </summary>
    ''' <value><see langword="True"/> to enable the action; otherwise, <see langword="False"/>.</value>
    <Category("TextBoxAction")>
    <Description("Determines whether the action button is enabled and the action can be executed.")>
    <DefaultValue(True)>
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
    ''' Gets or sets the delegate invoked after the component raises its <see cref="TextBoxActionPanel.ActionClicked"/> event.
    ''' </summary>
    ''' <value>A delegate that handles this action, or <see langword="Nothing"/> to use only the component event.</value>
    ''' <remarks>Delegates are intended for run-time assignment and are not serialized by the Windows Forms Designer.</remarks>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property ClickHandler As System.Action(Of TextBoxActionClickEventArgs)
        Get
            Return _ClickHandler
        End Get
        Set(Value As System.Action(Of TextBoxActionClickEventArgs))
            _ClickHandler = Value
        End Set
    End Property
    ''' <summary>
    ''' Returns a display name for the action in collection editors and diagnostic output.
    ''' </summary>
    ''' <returns>The configured <see cref="Key"/>, or <c>(Action)</c> when no key is assigned.</returns>
    Public Overrides Function ToString() As String
        Return If(String.IsNullOrWhiteSpace(_Key), "(Action)", _Key)
    End Function
    Friend Sub Invoke(E As TextBoxActionClickEventArgs)
        If _ClickHandler IsNot Nothing Then _ClickHandler.Invoke(E)
    End Sub
    Private Sub OnChanged()
        RaiseEvent Changed(Me, EventArgs.Empty)
    End Sub
End Class
