Imports System.Data.Common
Imports MySql.Data.MySqlClient

''' <summary>A database error exposed without a reference to the underlying driver.</summary>
Public NotInheritable Class MySqlServiceException
    Inherits DbException

    Friend Sub New(providerError As MySqlException)
        MyBase.New(providerError.Message, providerError)
        Number = providerError.Number
        SqlState = providerError.SqlState
        HResult = providerError.ErrorCode
    End Sub

    Public ReadOnly Property Number As Integer
    Public ReadOnly Property SqlState As String
End Class
