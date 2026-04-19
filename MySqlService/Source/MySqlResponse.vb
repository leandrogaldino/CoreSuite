''' <summary>
''' Represents the response of an SQL or CRUD operation executed
''' by the <see cref="MySqlRequest"/> class.
''' Contains returned data, number of affected rows, and the last inserted ID.
''' </summary>
Public Class MySqlResponse
    Private ReadOnly _Data As List(Of Dictionary(Of String, Object))
    Private ReadOnly _AffectedRows As Long
    Private ReadOnly _LastInsertedID As Long

    Friend Sub New(Optional Data As List(Of Dictionary(Of String, Object)) = Nothing, Optional AffectedRows As Long = 0, Optional LastInsertedID As Long = 0)
        _Data = Data
        _AffectedRows = AffectedRows
        _LastInsertedID = LastInsertedID
    End Sub

    ''' <summary>
    ''' Gets the ID of the last row inserted by an INSERT operation.
    ''' </summary>
    Public ReadOnly Property LastInsertedID As Long
        Get
            Return _LastInsertedID
        End Get
    End Property

    ''' <summary>
    ''' Gets the number of rows affected by the operation
    ''' (INSERT, UPDATE, or DELETE).
    ''' </summary>
    Public ReadOnly Property AffectedRows As Long
        Get
            Return _AffectedRows
        End Get
    End Property

    ''' <summary>
    ''' Indicates whether the operation returned any data.
    ''' </summary>
    Public ReadOnly Property HasData As Boolean
        Get
            If Data IsNot Nothing AndAlso Data.Count > 0 Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    ''' <summary>
    ''' Gets the returned data as a list of dictionaries (column/value).
    ''' </summary>
    Public ReadOnly Property Data As List(Of Dictionary(Of String, Object))
        Get
            Return _Data
        End Get
    End Property

    ''' <summary>
    ''' Gets the returned data as a <see cref="DataTable"/>.
    ''' Each dictionary key becomes a column, and each item becomes a row.
    ''' </summary>
    Public ReadOnly Property Table As DataTable
        Get
            If _Data Is Nothing Then
                Return Nothing
            Else
                Return GetDataTable(_Data)
            End If
        End Get
    End Property

    Private Function GetDataTable(dictList As List(Of Dictionary(Of String, Object))) As DataTable
        Dim Table As New DataTable()
        Dim AllKeys As New HashSet(Of String)()
        For Each Dict In dictList
            For Each Key In Dict.Keys
                AllKeys.Add(Key)
            Next Key
        Next Dict
        For Each Key In AllKeys
            Table.Columns.Add(Key, GetType(Object))
        Next Key
        For Each Dict In dictList
            Dim RowValues As Object() = Table.Columns.Cast(Of DataColumn)().
                                    Select(Function(Col)
                                               Dim Val As Object = Nothing
                                               Dict.TryGetValue(Col.ColumnName, Val)
                                               Return If(Val, DBNull.Value)
                                           End Function).ToArray()
            Table.Rows.Add(RowValues)
        Next Dict
        Return Table
    End Function
End Class