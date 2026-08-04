Imports System.Collections.ObjectModel
Imports System.Linq
''' <summary>
''' Represents one result set returned by a query or stored procedure.
''' </summary>
Public NotInheritable Class MySqlResultSet
    Private ReadOnly _Columns As IReadOnlyList(Of String)
    Private ReadOnly _Rows As IReadOnlyList(Of IReadOnlyDictionary(Of String, Object))
    Friend Sub New(Columns As IEnumerable(Of String), Rows As IEnumerable(Of IReadOnlyDictionary(Of String, Object)))
        ArgumentNullException.ThrowIfNull(Columns)
        ArgumentNullException.ThrowIfNull(Rows)
        _Columns = New ReadOnlyCollection(Of String)(Columns.ToList())
        _Rows = New ReadOnlyCollection(Of IReadOnlyDictionary(Of String, Object))(Rows.ToList())
    End Sub
    ''' <summary>
    ''' Gets the unique column names used as row dictionary keys.
    ''' </summary>
    Public ReadOnly Property Columns As IReadOnlyList(Of String)
        Get
            Return _Columns
        End Get
    End Property
    ''' <summary>
    ''' Gets the returned rows. Database null values are represented by <see langword="Nothing"/>.
    ''' </summary>
    Public ReadOnly Property Rows As IReadOnlyList(Of IReadOnlyDictionary(Of String, Object))
        Get
            Return _Rows
        End Get
    End Property
End Class
