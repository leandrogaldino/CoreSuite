Imports System.Data
Imports MySql.Data.MySqlClient
''' <summary>
''' Describes a stored procedure parameter, including output and return-value metadata.
''' </summary>
Public NotInheritable Class MySqlProcedureParameter
    ''' <summary>
    ''' Initializes a new input parameter.
    ''' </summary>
    ''' <param name="name">The parameter name.</param>
    ''' <param name="value">The parameter value.</param>
    Public Sub New(name As String, Optional value As Object = Nothing)
        Me.Name = NormalizeParameterName(name)
        Me.Value = value
    End Sub
    ''' <summary>
    ''' Gets the normalized parameter name.
    ''' </summary>
    Public ReadOnly Property Name As String
    ''' <summary>
    ''' Gets or sets the parameter value.
    ''' </summary>
    Public Property Value As Object
    ''' <summary>
    ''' Gets or sets the parameter direction. The default is <see cref="ParameterDirection.Input"/>.
    ''' </summary>
    Public Property Direction As ParameterDirection = ParameterDirection.Input
    ''' <summary>
    ''' Gets or sets the optional provider-specific MySQL data type.
    ''' </summary>
    Public Property MySqlDbType As MySqlDbType?
    ''' <summary>
    ''' Gets or sets the optional parameter size.
    ''' </summary>
    Public Property Size As Integer?
    ''' <summary>
    ''' Gets or sets the optional numeric precision.
    ''' </summary>
    Public Property Precision As Byte?
    ''' <summary>
    ''' Gets or sets the optional numeric scale.
    ''' </summary>
    Public Property Scale As Byte?
End Class
