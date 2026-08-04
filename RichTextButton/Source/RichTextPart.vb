Imports System.ComponentModel
Imports System.Runtime.CompilerServices
''' <summary>
''' Represents an individually formatted text segment rendered by a <see cref="RichTextButton"/>.
''' </summary>
<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class RichTextPart
    Implements INotifyPropertyChanged
    Private _Text As String = String.Empty
    Private _Font As Font = SystemFonts.DefaultFont
    Private _Color As Color = SystemColors.ControlText
    ''' <summary>
    ''' Occurs when the value of one of the text-part properties changes.
    ''' </summary>
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    ''' <summary>
    ''' Gets or sets the text displayed by this text part.
    ''' </summary>
    ''' <value>
    ''' The text rendered by the button for this part.
    ''' </value>
    <DefaultValue("")>
    Public Property Text As String
        Get
            Return _Text
        End Get
        Set(value As String)
            value = If(value, String.Empty)
            If _Text = value Then Return
            _Text = value
            OnPropertyChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the font used to render this text part.
    ''' </summary>
    ''' <value>
    ''' The font applied when rendering the text.
    ''' </value>
    Public Property Font As Font
        Get
            Return _Font
        End Get
        Set(value As Font)
            value = If(value, SystemFonts.DefaultFont)
            If Equals(_Font, value) Then Return
            _Font = value
            OnPropertyChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the color used to render this text part.
    ''' </summary>
    ''' <value>
    ''' The color applied when rendering the text.
    ''' </value>
    Public Property Color As Color
        Get
            Return _Color
        End Get
        Set(value As Color)
            If _Color = value Then Return
            _Color = value
            OnPropertyChanged()
        End Set
    End Property
    ''' <summary>
    ''' Raises the <see cref="PropertyChanged"/> event.
    ''' </summary>
    ''' <param name="PropertyName">The name of the property whose value changed.</param>
    Protected Overridable Sub OnPropertyChanged(<CallerMemberName> Optional PropertyName As String = Nothing)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(PropertyName))
    End Sub
    ''' <summary>
    ''' Returns a textual representation of this text part.
    ''' </summary>
    ''' <returns>The configured text, or <c>(Empty)</c> when no text has been defined.</returns>
    Public Overrides Function ToString() As String
        If String.IsNullOrEmpty(Text) Then Return "(Empty)"
        Return Text
    End Function
End Class