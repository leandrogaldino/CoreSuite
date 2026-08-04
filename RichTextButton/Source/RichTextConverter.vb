Imports System.ComponentModel

Imports System.Globalization

''' <summary>
''' Provides design-time conversion support for a collection of <see cref="RichTextPart"/> objects.
''' </summary>
Public Class RichTextConverter
    Inherits CollectionConverter
    ''' <summary>
    ''' Determines whether the collection can be converted to the specified destination type.
    ''' </summary>
    ''' <param name="Context">An object that provides contextual information about the component.</param>
    ''' <param name="DestinationType">The type to which the collection will be converted.</param>
    ''' <returns><see langword="True"/> when conversion is supported; otherwise, <see langword="False"/>.</returns>
    Public Overrides Function CanConvertTo(Context As ITypeDescriptorContext, DestinationType As Type) As Boolean
        Return DestinationType Is GetType(String) OrElse MyBase.CanConvertTo(Context, DestinationType)
    End Function
    ''' <summary>
    ''' Converts the collection to the specified destination type.
    ''' </summary>
    ''' <param name="Context">An object that provides contextual information about the component.</param>
    ''' <param name="Culture">The culture used during the conversion.</param>
    ''' <param name="Value">The collection value to convert.</param>
    ''' <param name="DestinationType">The type to which the value will be converted.</param>
    ''' <returns>The converted value.</returns>
    Public Overrides Function ConvertTo(Context As ITypeDescriptorContext, Culture As CultureInfo, Value As Object, DestinationType As Type) As Object
        If DestinationType Is GetType(String) AndAlso TypeOf Value Is BindingList(Of RichTextPart) Then
            Dim Parts As BindingList(Of RichTextPart) = DirectCast(Value, BindingList(Of RichTextPart))
            Return $"({Parts.Count} part{If(Parts.Count = 1, "", "s")})"
        End If
        Return MyBase.ConvertTo(Context, Culture, Value, DestinationType)
    End Function
End Class