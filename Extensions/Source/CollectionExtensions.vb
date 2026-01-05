Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports CoreSuite.Helpers
Imports CoreSuite.Attributes

Namespace Extensions
    Public Module CollectionExtensions
        ''' <summary>
        ''' Attempts to return the single element of a sequence. 
        ''' If no element exists, a new instance of the element's type is created and returned.
        ''' </summary>
        ''' <typeparam name="T">The type of element in the sequence.</typeparam>
        ''' <param name="Source">The sequence to search.</param>
        ''' <param name="Predicate">An optional function to filter the elements.</param>
        ''' <returns>
        ''' The single element of the sequence matching the predicate, 
        ''' or a new instance of <typeparamref name="T"/> if none exists.
        ''' </returns>
        ''' <remarks>
        ''' <para>
        ''' Use this extension when you expect at most one item from a collection, 
        ''' but want to ensure a valid object is returned even if the sequence is empty.
        ''' </para>
        ''' <para>
        ''' If the type <typeparamref name="T"/> does not have a parameterless constructor, 
        ''' <c>Nothing</c> is returned.
        ''' </para>
        ''' </remarks>
        Public Function FirstOrNew(Of T)(Source As IEnumerable(Of T), Optional Predicate As Func(Of T, Boolean) = Nothing) As T
            Dim Element As T = Source.SingleOrDefault(Predicate)
            If Element Is Nothing Then
                Try
                    Return Activator.CreateInstance(GetType(T))
                Catch ex As Exception
                    Return Nothing
                End Try
            Else
                Return Element
            End If
        End Function

        ''' <summary>
        ''' Converts a collection of objects to a <see cref="DataTable"/>.
        ''' </summary>
        ''' <param name="Source">The collection to convert.</param>
        ''' <returns>A <see cref="DataTable"/> representing the collection's data.</returns>
        ''' <exception cref="InvalidOperationException">Thrown if the source is a string.</exception>
        ''' <remarks>
        ''' <para>
        ''' Each public property of the collection's item type becomes a column in the <see cref="DataTable"/>.
        ''' Properties marked with <see cref="IgnoreInToTableAttribute"/> are ignored.
        ''' Lazy-loaded properties (<c>Lazy(Of T)</c>) are represented as strings showing the type 
        ''' without evaluating the value.
        ''' </para>
        ''' <para>
        ''' This extension is useful for:
        ''' </para>
        ''' <list type="bullet">
        ''' <item><description>Binding collections to UI elements like grids.</description></item>
        ''' <item><description>Exporting collections to Excel, CSV, or other tabular formats.</description></item>
        ''' <item><description>Debugging or inspecting object data in a table format.</description></item>
        ''' </list>
        ''' <para>
        ''' Use <see cref="IgnoreInToTableAttribute"/> on properties you do not want included in the table.
        ''' </para>
        ''' </remarks>
        <Extension>
        Public Function ToTable(Source As IEnumerable) As DataTable
            If TypeOf Source Is String Then
                Throw New InvalidOperationException("A string cannot be converted to a DataTable.")
            End If
            Dim CollectionType As Type = ReflectionHelper.GetCollectionPropertyType(Source.GetType)
            Dim Table As New DataTable
            For Each p As PropertyInfo In CollectionType.GetProperties()
                If p.GetCustomAttribute(GetType(IgnoreInToTableAttribute)) IsNot Nothing Then
                    Continue For
                End If
                Dim propType As Type = p.PropertyType
                If propType.IsGenericType AndAlso propType.GetGenericTypeDefinition() = GetType(Lazy(Of )) Then
                    Table.Columns.Add(p.Name, GetType(String))
                Else
                    Table.Columns.Add(p.Name, propType)
                End If
            Next
            For Each item In Source
                Dim RowObj As New List(Of Object)
                For Each p As PropertyInfo In item.GetType.GetProperties()
                    If p.GetCustomAttribute(GetType(IgnoreInToTableAttribute)) IsNot Nothing Then
                        Continue For
                    End If
                    If p.PropertyType.IsGenericType AndAlso p.PropertyType.GetGenericTypeDefinition() = GetType(Lazy(Of )) Then
                        RowObj.Add($"Lazy<{p.PropertyType.GetGenericArguments()(0).Name}>")
                    Else
                        Dim value As Object = p.GetValue(item)
                        RowObj.Add(value)
                    End If
                Next
                Table.Rows.Add(RowObj.ToArray())
            Next

            Return Table
        End Function
    End Module
End Namespace
