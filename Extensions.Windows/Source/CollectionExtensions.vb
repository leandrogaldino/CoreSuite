Imports System.Reflection
Imports System.Runtime.CompilerServices

Namespace Extensions
    Public Module CollectionExtensions
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



        ''' <summary>
        ''' Converts a collection of dictionaries to a <see cref="DataTable"/>.
        ''' </summary>
        ''' <param name="Source">The collection of dictionaries to convert.</param>
        ''' <returns>A <see cref="DataTable"/> representing the dictionary data.</returns>
        ''' <remarks>
        ''' Each dictionary key becomes a column and each dictionary instance becomes a row.
        ''' Missing keys are represented as <see cref="DBNull.Value"/>.
        ''' </remarks>
        <Extension>
        Public Function ToTable(Source As IEnumerable(Of Dictionary(Of String, Object))) As DataTable
            If Source Is Nothing Then
                Return Nothing
            End If
            Dim Table As New DataTable()
            Dim Columns As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each Item In Source
                If Item Is Nothing Then Continue For
                For Each Key In Item.Keys
                    Columns.Add(Key)
                Next
            Next Item
            For Each ColumnName In Columns
                Table.Columns.Add(ColumnName, GetType(Object))
            Next ColumnName
            For Each Item In Source
                If Item Is Nothing Then
                    Continue For
                End If
                Dim Row As DataRow = Table.NewRow()
                For Each ColumnName In Columns
                    Dim Value As Object = Nothing
                    If Item.TryGetValue(ColumnName, Value) AndAlso Value IsNot Nothing Then
                        Row(ColumnName) = Value
                    Else
                        Row(ColumnName) = DBNull.Value
                    End If
                Next ColumnName
                Table.Rows.Add(Row)
            Next Item
            Return Table
        End Function
    End Module
End Namespace