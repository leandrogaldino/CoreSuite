Imports System.ComponentModel
Imports System.Reflection

''' <summary>
''' Provides helper methods for working with Enum types and their metadata,
''' such as <see cref="DescriptionAttribute"/>.
''' </summary>
Public Class EnumHelper

    ''' <summary>
    ''' Gets the enum value whose <see cref="DescriptionAttribute"/> matches the specified description.
    ''' </summary>
    ''' <typeparam name="T">
    ''' The enum type to search.
    ''' </typeparam>
    ''' <param name="Description">
    ''' The description associated with the enum value.
    ''' </param>
    ''' <returns>
    ''' The enum value whose <see cref="DescriptionAttribute"/> matches
    ''' <paramref name="Description"/>.
    ''' </returns>
    ''' <exception cref="ArgumentException">
    ''' Thrown when <typeparamref name="T"/> is not an enum type.
    ''' </exception>
    ''' <exception cref="InvalidOperationException">
    ''' Thrown when no enum value has the specified description.
    ''' </exception>
    Public Shared Function GetValueFromDescription(Of T As Structure)(Description As String) As T
        If Not GetType(T).IsEnum Then
            Throw New ArgumentException($"Type '{GetType(T).FullName}' must be an enum.")
        End If
        ArgumentException.ThrowIfNullOrWhiteSpace(Description)
        For Each Field As FieldInfo In GetType(T).GetFields(BindingFlags.Public Or BindingFlags.Static)
            Dim Attribute As DescriptionAttribute = Field.GetCustomAttribute(Of DescriptionAttribute)()
            If Attribute IsNot Nothing AndAlso String.Equals(Attribute.Description, Description, StringComparison.Ordinal) Then
                Return DirectCast(Field.GetValue(Nothing), T)
            End If
        Next
        Throw New InvalidOperationException($"No value of enum '{GetType(T).FullName}' has the description '{Description}'.")
    End Function

    ''' <summary>
    ''' Returns a collection of enum values filtered by a predicate.
    ''' If no predicate is provided, all enum values are returned.
    ''' </summary>
    ''' <typeparam name="T">
    ''' The enum type.
    ''' </typeparam>
    ''' <param name="Predicate">
    ''' An optional filter function applied to each enum field.
    ''' </param>
    ''' <returns>
    ''' An <see cref="IEnumerable(Of T)"/> containing the enum values
    ''' that satisfy the specified predicate.
    ''' </returns>
    ''' <exception cref="ArgumentException">
    ''' Thrown when <typeparamref name="T"/> is not an enum type.
    ''' </exception>
    Public Shared Function GetEnumItems(Of T As Structure)(Optional Predicate As Func(Of FieldInfo, Boolean) = Nothing) As IEnumerable(Of T)
        Dim EnumItems As New List(Of T)()
        If Not GetType(T).IsEnum Then
            Throw New ArgumentException("Type **T** must be an Enum.")
        End If
        If Predicate Is Nothing Then
            Predicate = Function(f) True
        End If
        Dim EnumType As Type = GetType(T)
        Dim Fields = EnumType.GetFields(BindingFlags.Public Or BindingFlags.Static)
        For Each Field In Fields
            If Predicate(Field) Then
                Dim EnumValue As T = CType([Enum].Parse(GetType(T), Field.Name), T)
                EnumItems.Add(EnumValue)
            End If
        Next
        Return EnumItems
    End Function

    ''' <summary>
    ''' Returns a collection of description strings from the <see cref="DescriptionAttribute"/>
    ''' of each enum value, optionally filtered by a predicate.
    ''' </summary>
    ''' <typeparam name="T">
    ''' The enum type.
    ''' </typeparam>
    ''' <param name="Predicate">
    ''' An optional filter function applied to each enum field.
    ''' </param>
    ''' <returns>
    ''' An <see cref="IEnumerable(Of String)"/> containing the description values
    ''' of the enum items.
    ''' </returns>
    Public Shared Function GetEnumDescriptions(Of T)(Optional Predicate As Func(Of FieldInfo, Boolean) = Nothing) As IEnumerable(Of String)
        Dim Descriptions As New List(Of String)
        If Predicate Is Nothing Then
            Predicate = Function(f) True
        End If
        Dim EnumType As Type = GetType(T)
        Dim Fields = EnumType.GetFields(BindingFlags.Public Or BindingFlags.Static)
        For Each Field In Fields
            If Predicate(Field) Then
                Dim Attribute = TryCast(Field.GetCustomAttributes(GetType(DescriptionAttribute), True).FirstOrDefault(), DescriptionAttribute)
                If Attribute IsNot Nothing Then
                    Descriptions.Add(Attribute.Description)
                Else
                    Descriptions.Add(String.Empty)
                End If
            End If
        Next

        Return Descriptions
    End Function

    ''' <summary>
    ''' Gets the value of the <see cref="DescriptionAttribute"/> applied
    ''' to the specified enum value.
    ''' </summary>
    ''' <param name="Value">
    ''' The enum value.
    ''' </param>
    ''' <returns>
    ''' The description text defined by <see cref="DescriptionAttribute"/>,
    ''' or <c>Nothing</c> if the attribute is not present.
    ''' </returns>
    Public Shared Function GetEnumDescription(Value As [Enum]) As String
        Dim FieldInfo As FieldInfo = Value.GetType().GetField(Value.ToString())
        If FieldInfo Is Nothing Then Return Nothing
        Dim Attribute = CType(FieldInfo.GetCustomAttribute(GetType(DescriptionAttribute)), DescriptionAttribute)
        Return Attribute?.Description
    End Function

    ''' <summary>
    ''' Attempts to convert a textual value to a defined member of the specified enum type.
    ''' </summary>
    ''' <typeparam name="T">
    ''' The enum type to which the supplied value should be converted.
    ''' </typeparam>
    ''' <param name="Value">
    ''' The enum member name or numeric value to convert.
    ''' </param>
    ''' <param name="Result">
    ''' When this method returns <see langword="True"/>, contains the converted enum value.
    ''' When this method returns <see langword="False"/>, contains the default value of
    ''' <typeparamref name="T"/>.
    ''' </param>
    ''' <returns>
    ''' <see langword="True"/> when <paramref name="Value"/> represents a defined enum member;
    ''' otherwise, <see langword="False"/>.
    ''' </returns>
    ''' <exception cref="ArgumentException">
    ''' Thrown when <typeparamref name="T"/> is not an enum type.
    ''' </exception>
    Public Shared Function TryParseEnum(Of T As Structure)(Value As String, ByRef Result As T) As Boolean
        If Not GetType(T).IsEnum Then
            Throw New ArgumentException($"Type '{GetType(T).FullName}' must be an enum.")
        End If
        Result = Nothing
        If String.IsNullOrWhiteSpace(Value) Then Return False
        If Not [Enum].TryParse(Value, True, Result) Then
            Result = Nothing
            Return False
        End If
        If Not [Enum].IsDefined(GetType(T), Result) Then
            Result = Nothing
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' Converts a textual value to a defined member of the specified enum type or returns a
    ''' fallback value when conversion is not possible.
    ''' </summary>
    ''' <typeparam name="T">
    ''' The enum type to which the supplied value should be converted.
    ''' </typeparam>
    ''' <param name="Value">
    ''' The enum member name or numeric value to convert.
    ''' </param>
    ''' <param name="DefaultValue">
    ''' The enum value returned when <paramref name="Value"/> is null, empty, invalid or does
    ''' not represent a defined enum member.
    ''' </param>
    ''' <returns>
    ''' The converted enum value, or <paramref name="DefaultValue"/> when conversion fails.
    ''' </returns>
    ''' <exception cref="ArgumentException">
    ''' Thrown when <typeparamref name="T"/> is not an enum type.
    ''' </exception>
    Public Shared Function ParseEnumOrDefault(Of T As Structure)(Value As String, DefaultValue As T) As T
        Dim Result As T = Nothing
        If TryParseEnum(Value, Result) Then Return Result
        Return DefaultValue
    End Function
End Class