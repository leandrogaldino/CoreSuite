
Imports System.ComponentModel

''' <summary>
''' Represents a reference to a database column used in query definitions.
''' </summary>
<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class QueryColumnReference
    ''' <summary>
    ''' Gets or sets the name of the referenced database column.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the name of the referenced database column.")>
    Public Property ColumnName As String
    ''' <summary>
    ''' Initializes a new instance of the <see cref="QueryColumnReference"/> class.
    ''' </summary>
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="QueryColumnReference"/> class
    ''' with the specified column name.
    ''' </summary>
    ''' <param name="Name">
    ''' The name of the database column.
    ''' </param>
    Public Sub New(Name As String)
        Me.ColumnName = Name
    End Sub
    ''' <summary>
    ''' Returns a string representation of the column reference.
    ''' </summary>
    ''' <returns>
    ''' The name of the referenced column.
    ''' </returns>
    Public Overrides Function ToString() As String
        Return ColumnName
    End Function
End Class
