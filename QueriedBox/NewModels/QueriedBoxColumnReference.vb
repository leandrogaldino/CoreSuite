
Imports System.ComponentModel

<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class QueriedBoxColumnReference
    Public Property ColumnName As String
    Public Sub New()
    End Sub
    Public Sub New(Name As String)
        Me.ColumnName = Name
    End Sub
    Public Overrides Function ToString() As String
        Return ColumnName
    End Function
End Class
