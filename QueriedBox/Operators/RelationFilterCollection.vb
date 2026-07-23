Imports System.ComponentModel
Imports System.ComponentModel.TypeConverter

Friend Class RelationFilterCollection
    Inherits StringConverter
    Private Shared ReadOnly _JoinList As New List(Of String) From {
        "LEFT",
        "RIGHT",
        "INNER",
        "CROSS"
    }
    Public Overrides Function GetStandardValuesSupported(context As ITypeDescriptorContext) As Boolean
        Return True
    End Function
    Public Overrides Function GetStandardValuesExclusive(context As ITypeDescriptorContext) As Boolean
        Return True
    End Function
    Public Overrides Function GetStandardValues(context As ITypeDescriptorContext) As StandardValuesCollection
        Return New StandardValuesCollection(_JoinList)
    End Function
End Class
