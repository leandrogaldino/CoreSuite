Imports System.ComponentModel
Imports System.Globalization
Imports System.Reflection
''' <summary>
''' Provides selectable <see cref="UserControl"/> types for <see cref="NavigationPage.ControlType"/> in the Windows Forms Property Grid.
''' </summary>
Public Class UserControlTypeConverter
    Inherits TypeConverter
    ''' <summary>
    ''' Indicates that the converter provides a list of available values.
    ''' </summary>
    ''' <param name="Context">The format context.</param>
    ''' <returns><see langword="True"/>.</returns>
    Public Overrides Function GetStandardValuesSupported(Context As ITypeDescriptorContext) As Boolean
        Return True
    End Function
    ''' <summary>
    ''' Indicates that values outside the available list can also be assigned.
    ''' </summary>
    ''' <param name="Context">The format context.</param>
    ''' <returns><see langword="False"/>.</returns>
    Public Overrides Function GetStandardValuesExclusive(Context As ITypeDescriptorContext) As Boolean
        Return False
    End Function
    ''' <summary>
    ''' Gets the loaded, public <see cref="UserControl"/> types that can be created through a parameterless constructor.
    ''' </summary>
    ''' <param name="Context">The format context.</param>
    ''' <returns>The available control types.</returns>
    Public Overrides Function GetStandardValues(Context As ITypeDescriptorContext) As StandardValuesCollection
        Dim Values As New List(Of Type)
        For Each LoadedAssembly As Assembly In AppDomain.CurrentDomain.GetAssemblies()
            For Each Candidate As Type In GetLoadableTypes(LoadedAssembly)
                If Candidate.IsPublic AndAlso Not Candidate.IsAbstract AndAlso GetType(UserControl).IsAssignableFrom(Candidate) AndAlso Candidate.GetConstructor(Type.EmptyTypes) IsNot Nothing Then Values.Add(Candidate)
            Next
        Next
        Values.Sort(Function(Left As Type, Right As Type) StringComparer.OrdinalIgnoreCase.Compare(Left.FullName, Right.FullName))
        Return New StandardValuesCollection(Values)
    End Function
    ''' <summary>
    ''' Indicates whether a source value can be converted to a control type.
    ''' </summary>
    ''' <param name="Context">The format context.</param>
    ''' <param name="SourceType">The source type.</param>
    ''' <returns><see langword="True"/> for strings; otherwise, the base implementation result.</returns>
    Public Overrides Function CanConvertFrom(Context As ITypeDescriptorContext, SourceType As Type) As Boolean
        Return SourceType Is GetType(String) OrElse MyBase.CanConvertFrom(Context, SourceType)
    End Function
    ''' <summary>
    ''' Converts a type name into a loaded control type.
    ''' </summary>
    ''' <param name="Context">The format context.</param>
    ''' <param name="Culture">The conversion culture.</param>
    ''' <param name="Value">The value to convert.</param>
    ''' <returns>The resolved type, <see langword="Nothing"/> for an empty value, or the base conversion result.</returns>
    Public Overrides Function ConvertFrom(Context As ITypeDescriptorContext, Culture As CultureInfo, Value As Object) As Object
        Dim TypeName As String = TryCast(Value, String)
        If TypeName Is Nothing Then Return MyBase.ConvertFrom(Context, Culture, Value)
        If String.IsNullOrWhiteSpace(TypeName) Then Return Nothing
        Dim ResolvedType As Type = Type.GetType(TypeName, False, True)
        If ResolvedType IsNot Nothing Then Return ResolvedType
        For Each LoadedAssembly As Assembly In AppDomain.CurrentDomain.GetAssemblies()
            ResolvedType = LoadedAssembly.GetType(TypeName, False, True)
            If ResolvedType IsNot Nothing Then Return ResolvedType
            For Each Candidate As Type In GetLoadableTypes(LoadedAssembly)
                If String.Equals(Candidate.Name, TypeName, StringComparison.OrdinalIgnoreCase) Then Return Candidate
            Next
        Next
        Throw New FormatException($"The UserControl type '{TypeName}' could not be resolved.")
    End Function
    ''' <summary>
    ''' Indicates whether a control type can be converted to the requested destination type.
    ''' </summary>
    ''' <param name="Context">The format context.</param>
    ''' <param name="DestinationType">The destination type.</param>
    ''' <returns><see langword="True"/> for strings; otherwise, the base implementation result.</returns>
    Public Overrides Function CanConvertTo(Context As ITypeDescriptorContext, DestinationType As Type) As Boolean
        Return DestinationType Is GetType(String) OrElse MyBase.CanConvertTo(Context, DestinationType)
    End Function
    ''' <summary>
    ''' Converts a control type to a display name.
    ''' </summary>
    ''' <param name="Context">The format context.</param>
    ''' <param name="Culture">The conversion culture.</param>
    ''' <param name="Value">The value to convert.</param>
    ''' <param name="DestinationType">The destination type.</param>
    ''' <returns>The full type name for string destinations, or the base conversion result.</returns>
    Public Overrides Function ConvertTo(Context As ITypeDescriptorContext, Culture As CultureInfo, Value As Object, DestinationType As Type) As Object
        If DestinationType Is GetType(String) Then
            Dim ControlType As Type = TryCast(Value, Type)
            Return If(ControlType Is Nothing, String.Empty, ControlType.FullName)
        End If
        Return MyBase.ConvertTo(Context, Culture, Value, DestinationType)
    End Function
    Private Shared Function GetLoadableTypes(Assembly As Assembly) As IEnumerable(Of Type)
        Try
            Return Assembly.GetTypes()
        Catch Ex As ReflectionTypeLoadException
            Dim Values As New List(Of Type)
            For Each Item As Type In Ex.Types
                If Item IsNot Nothing Then Values.Add(Item)
            Next
            Return Values
        Catch Ex As NotSupportedException
            Return Array.Empty(Of Type)()
        End Try
    End Function
End Class
