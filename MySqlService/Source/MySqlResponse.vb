''' <summary>
''' Represents the response of an SQL or CRUD operation executed
''' by the <see cref="MySqlRequest"/> class.
''' Contains returned data, number of affected rows, and the last inserted ID.
''' </summary>
Public Class MySqlResponse
    Private ReadOnly _Data As List(Of Dictionary(Of String, Object))
    Private ReadOnly _AffectedRows As Long
    Private ReadOnly _LastInsertedID As Long
    ''' <summary>
    ''' Initializes a new instance of the <see cref="MySqlResponse"/> class with the
    ''' specified data, affected row count, and last inserted ID.
    ''' </summary>
    ''' <param name="Data">The data returned by the operation, if any.</param>
    ''' <param name="AffectedRows">The number of rows affected by the operation.</param>
    ''' <param name="LastInsertedID">The ID of the last row inserted by an INSERT operation.</param>
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
End Class