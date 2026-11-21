Imports MySql.Data.Types

''' <summary>Preserves MySQL dates, including zero dates, without exposing provider types.</summary>
Public NotInheritable Class MySqlDateValue
    Private ReadOnly _Value As MySqlDateTime
    Friend Sub New(value As MySqlDateTime)
        _Value = value
    End Sub
    Public Sub New(value As DateTime)
        _Value = New MySqlDateTime(value)
    End Sub
    Public Sub New(year As Integer, month As Integer, day As Integer, Optional hour As Integer = 0, Optional minute As Integer = 0, Optional second As Integer = 0, Optional microsecond As Integer = 0)
        _Value = New MySqlDateTime(year, month, day, hour, minute, second, microsecond)
    End Sub
    Friend Function ToProviderValue() As MySqlDateTime
        Return _Value
    End Function
    Public Function GetDateTime() As DateTime
        Return _Value.GetDateTime()
    End Function
    Public Overrides Function ToString() As String
        Return _Value.ToString()
    End Function
    Public ReadOnly Property Year As Integer
        Get
            Return _Value.Year
        End Get
    End Property
    Public ReadOnly Property Month As Integer
        Get
            Return _Value.Month
        End Get
    End Property
    Public ReadOnly Property Day As Integer
        Get
            Return _Value.Day
        End Get
    End Property
    Public ReadOnly Property Hour As Integer
        Get
            Return _Value.Hour
        End Get
    End Property
    Public ReadOnly Property Minute As Integer
        Get
            Return _Value.Minute
        End Get
    End Property
    Public ReadOnly Property Second As Integer
        Get
            Return _Value.Second
        End Get
    End Property
    Public ReadOnly Property Millisecond As Integer
        Get
            Return _Value.Millisecond
        End Get
    End Property
    Public ReadOnly Property Microsecond As Integer
        Get
            Return _Value.Microsecond
        End Get
    End Property
    Public ReadOnly Property TimezoneOffset As Integer
        Get
            Return _Value.TimezoneOffset
        End Get
    End Property
    Public ReadOnly Property IsNull As Boolean
        Get
            Return _Value.IsNull
        End Get
    End Property
    Public ReadOnly Property IsValidDateTime As Boolean
        Get
            Return _Value.IsValidDateTime
        End Get
    End Property
End Class
