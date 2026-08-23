Imports System.ComponentModel
Imports System.Drawing

''' <summary>
''' Represents an item displayed by a <see cref="StatusSelector"/>.
''' </summary>
Public Class StatusSelectorItem
    Implements INotifyPropertyChanged
    Private _Text As String
    Private _Value As Object
    Private _ForeColor As Color = SystemColors.ControlText
    Private _BackColor As Color = Color.Empty
    Private _HoverForeColor As Color = Color.Empty
    Private _HoverBackColor As Color = Color.Empty
    Private _Enabled As Boolean = True
    Private _Visible As Boolean = True
    Private _Image As Image
    Private _ToolTipText As String = String.Empty
    Private _Font As Font
    ''' <summary>
    ''' Occurs when one of the item properties changes.
    ''' </summary>
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    ''' <summary>
    ''' Initializes a new instance of the <see cref="StatusSelectorItem"/> class.
    ''' </summary>
    Public Sub New()
        Me.New(String.Empty, Nothing, SystemColors.ControlText)
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="StatusSelectorItem"/> class.
    ''' </summary>
    ''' <param name="text">The text displayed for the item.</param>
    ''' <param name="value">The value represented by the item.</param>
    Public Sub New(text As String, value As Object)
        Me.New(text, value, SystemColors.ControlText)
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="StatusSelectorItem"/> class.
    ''' </summary>
    ''' <param name="text">The text displayed for the item.</param>
    ''' <param name="value">The value represented by the item.</param>
    ''' <param name="foreColor">The foreground color used by the item.</param>
    Public Sub New(text As String, value As Object, foreColor As Color)
        _Text = If(text, String.Empty)
        _Value = value
        _ForeColor = foreColor
    End Sub
    ''' <summary>
    ''' Gets or sets the text displayed for the item.
    ''' </summary>
    <Category("StatusSelector"), Description("Specifies the text displayed for the status item.")>
    Public Property Text As String
        Get
            Return _Text
        End Get
        Set(value As String)
            Dim NewValue = If(value, String.Empty)
            If String.Equals(_Text, NewValue, StringComparison.Ordinal) Then Return
            _Text = NewValue
            OnPropertyChanged(NameOf(Text))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the value represented by the item.
    ''' </summary>
    <Category("StatusSelector"), Description("Specifies the application value represented by the status item.")>
    Public Property Value As Object
        Get
            Return _Value
        End Get
        Set(value As Object)
            If Object.Equals(_Value, value) Then Return
            _Value = value
            OnPropertyChanged(NameOf(value))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the foreground color used by the item.
    ''' </summary>
    <Category("Appearance"), Description("Specifies the foreground color used by the status item.")>
    Public Property ForeColor As Color
        Get
            Return _ForeColor
        End Get
        Set(value As Color)
            If _ForeColor = value Then Return
            _ForeColor = value
            OnPropertyChanged(NameOf(ForeColor))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the background color used by the item. <see cref="Color.Empty"/> inherits the selector menu background color.
    ''' </summary>
    <Category("Appearance"), Description("Specifies the item background color. Color.Empty inherits the selector menu background color.")>
    Public Property BackColor As Color
        Get
            Return _BackColor
        End Get
        Set(value As Color)
            If _BackColor = value Then Return
            _BackColor = value
            OnPropertyChanged(NameOf(BackColor))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the foreground color used while the pointer is over the item. <see cref="Color.Empty"/> inherits the selector hover foreground color.
    ''' </summary>
    <Category("Appearance"), Description("Specifies the item hover foreground color. Color.Empty inherits the selector hover setting.")>
    Public Property HoverForeColor As Color
        Get
            Return _HoverForeColor
        End Get
        Set(value As Color)
            If _HoverForeColor = value Then Return
            _HoverForeColor = value
            OnPropertyChanged(NameOf(HoverForeColor))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the background color used while the pointer is over the item. <see cref="Color.Empty"/> inherits the selector hover background color.
    ''' </summary>
    <Category("Appearance"), Description("Specifies the item hover background color. Color.Empty inherits the selector hover setting.")>
    Public Property HoverBackColor As Color
        Get
            Return _HoverBackColor
        End Get
        Set(value As Color)
            If _HoverBackColor = value Then Return
            _HoverBackColor = value
            OnPropertyChanged(NameOf(HoverBackColor))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the item can be selected.
    ''' </summary>
    <Category("Behavior"), Description("Specifies whether the status item can be selected."), DefaultValue(True)>
    Public Property Enabled As Boolean
        Get
            Return _Enabled
        End Get
        Set(value As Boolean)
            If _Enabled = value Then Return
            _Enabled = value
            OnPropertyChanged(NameOf(Enabled))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the item is visible in the dropdown menu.
    ''' </summary>
    <Category("Behavior"), Description("Specifies whether the status item is visible in the dropdown menu."), DefaultValue(True)>
    Public Property Visible As Boolean
        Get
            Return _Visible
        End Get
        Set(value As Boolean)
            If _Visible = value Then Return
            _Visible = value
            OnPropertyChanged(NameOf(Visible))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the image displayed beside the item.
    ''' </summary>
    <Category("Appearance"), Description("Specifies the image displayed beside the status item."), DefaultValue(GetType(Image), Nothing)>
    Public Property Image As Image
        Get
            Return _Image
        End Get
        Set(value As Image)
            If ReferenceEquals(_Image, value) Then Return
            _Image = value
            OnPropertyChanged(NameOf(Image))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the tooltip text displayed for the item.
    ''' </summary>
    <Category("Appearance"), Description("Specifies the tooltip text displayed for the status item."), DefaultValue("")>
    Public Property ToolTipText As String
        Get
            Return _ToolTipText
        End Get
        Set(value As String)
            Dim NewValue = If(value, String.Empty)
            If String.Equals(_ToolTipText, NewValue, StringComparison.Ordinal) Then Return
            _ToolTipText = NewValue
            OnPropertyChanged(NameOf(ToolTipText))
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the font used by the item. <see langword="Nothing"/> inherits the selector font.
    ''' </summary>
    <Category("Appearance"), Description("Specifies the item font. Nothing inherits the selector font."), DefaultValue(GetType(Font), Nothing)>
    Public Property Font As Font
        Get
            Return _Font
        End Get
        Set(value As Font)
            If ReferenceEquals(_Font, value) Then Return
            _Font = value
            OnPropertyChanged(NameOf(Font))
        End Set
    End Property
    ''' <summary>
    ''' Returns the display text of the item.
    ''' </summary>
    Public Overrides Function ToString() As String
        Return Text
    End Function
    Private Sub OnPropertyChanged(propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub
End Class
