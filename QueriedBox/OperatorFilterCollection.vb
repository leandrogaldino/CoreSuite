Imports System.ComponentModel
Friend Class OperatorFilterCollection
    Inherits StringConverter
    Private Shared _OperatorList As New List(Of String) From {
            "=",
            "<>",
            ">",
            ">=",
            "<",
            "<=",
            "BETWEEN",
            "LIKE"
        }
    Public Overrides Function GetStandardValuesSupported(ByVal context As ITypeDescriptorContext) As Boolean
        Return True
    End Function
    Public Overrides Function GetStandardValuesExclusive(ByVal context As ITypeDescriptorContext) As Boolean
        Return True
    End Function
    Public Overrides Function GetStandardValues(ByVal context As ITypeDescriptorContext) As StandardValuesCollection
        Return New StandardValuesCollection(_OperatorList)
    End Function
End Class
