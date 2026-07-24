Imports System.ComponentModel

<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class QueriedBoxColumn
    Public Property ColumnName As String
    Public Property ColumnAlias As String
    Public Property Freeze As Boolean = True
    Public Property Display As Boolean = True
    Public Property Prefix As String
    Public Property Suffix As String
    Public Sub New()
    End Sub
    Public Sub New(Name As String, [Alias] As String, Freeze As Boolean, Display As Boolean, Prefix As String, Suffix As String)
        Me.ColumnName = Name
        Me.ColumnAlias = [Alias]
        Me.Freeze = Freeze
        Me.Display = Display
        Me.Prefix = Prefix
        Me.Suffix = Suffix
    End Sub
    Public Overrides Function ToString() As String
        Dim HasAlias As Boolean = Not String.IsNullOrEmpty(ColumnAlias)

        If HasAlias Then
            Return $"{ColumnName} AS {ColumnAlias}"
        Else
            Return $"{ColumnName}"
        End If

    End Function
End Class
