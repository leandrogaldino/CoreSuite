Imports System.Reflection

''' <summary>
''' Provides utility methods to retrieve custom attributes
''' from types, members, and enumerations using reflection.
''' </summary>
Public Class AttributeHelper

    ''' <summary>
    ''' Retrieves an attribute from a member of a class, structure, or enumeration
    ''' based on the member name.
    ''' </summary>
    ''' <typeparam name="TAttribute">
    ''' The type of attribute to retrieve.
    ''' </typeparam>
    ''' <param name="Type">
    ''' The <see cref="System.Type"/> that contains the specified member.
    ''' </param>
    ''' <param name="MemberName">
    ''' The name of the member from which the attribute will be retrieved.
    ''' </param>
    ''' <returns>
    ''' The attribute of type <typeparamref name="TAttribute"/> associated with the specified member,
    ''' or <c>Nothing</c> if the attribute is not present.
    ''' </returns>
    ''' <exception cref="ArgumentNullException">
    ''' Thrown when <paramref name="Type"/> or <paramref name="MemberName"/> is <c>Nothing</c> or empty.
    ''' </exception>
    ''' <exception cref="ArgumentException">
    ''' Thrown when the member with the specified name is not found in the provided type.
    ''' </exception>
    Public Shared Function GetAttribute(Of TAttribute As Attribute)(Type As Type, MemberName As String) As TAttribute
        If String.IsNullOrEmpty(MemberName) Then
            Throw New ArgumentNullException(NameOf(MemberName))
        End If
        If Type Is Nothing Then
            Throw New ArgumentNullException(NameOf(Type))
        End If
        Dim Member = Type.GetMember(MemberName, BindingFlags.Public Or BindingFlags.Static).FirstOrDefault()
        If Member Is Nothing Then
            Throw New ArgumentException($"The member '{MemberName}' was not found in type '{Type.Name}'.")
        End If
        Return CType(Member.GetCustomAttribute(GetType(TAttribute)), TAttribute)
    End Function

End Class