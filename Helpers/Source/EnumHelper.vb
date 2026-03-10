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
    ''' The enum type.
    ''' </typeparam>
    ''' <param name="Description">
    ''' The description text associated with the enum value.
    ''' </param>
    ''' <returns>
    ''' The enum value whose <see cref="DescriptionAttribute"/> matches the specified description.
    ''' </returns>
    ''' <exception cref="Exception">
    ''' Thrown when no enum value with the specified description is found.
    ''' </exception>
    Public Shared Function GetEnumValue(Of T)(Description As String) As T
        For Each Field In GetType(T).GetFields()
            Dim Attr As DescriptionAttribute = TryCast(Attribute.GetCustomAttribute(Field, GetType(DescriptionAttribute)), DescriptionAttribute)
            If Attr Is Nothing Then Continue For
            If Attr.Description = Description Then
                Return Field.GetValue(Nothing)
            End If
        Next
        Throw New Exception("Nenhum elemento encontrado com essa descrição.")
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
            Throw New ArgumentException("O tipo T deve ser um Enum.")
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

End Class