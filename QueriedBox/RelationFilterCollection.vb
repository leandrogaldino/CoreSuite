Imports System.ComponentModel

Friend Class RelationFilterCollection
    Inherits StringConverter
    Private Shared _JoinList As New List(Of String) From {
        "LEFT",
        "RIGHT",
        "INNER",
        "CROSS"
    }
    Public Overrides Function GetStandardValuesSupported(ByVal context As ITypeDescriptorContext) As Boolean
        Return True
    End Function
    Public Overrides Function GetStandardValuesExclusive(ByVal context As ITypeDescriptorContext) As Boolean
        Return True
    End Function
    Public Overrides Function GetStandardValues(ByVal context As ITypeDescriptorContext) As StandardValuesCollection
        Return New StandardValuesCollection(_JoinList)
    End Function
End Class
