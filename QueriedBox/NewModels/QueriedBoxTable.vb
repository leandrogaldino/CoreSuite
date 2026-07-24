Public Class QueriedBoxTable
    Public Property TableName As String
    Public Property [Alias] As String
    Public Sub New()
    End Sub
    Public Sub New(Name As String, [Alias] As String)
        Me.TableName = Name
        Me.Alias = [Alias]
    End Sub
    Overrides Function ToString() As String
        Dim HasAlias As Boolean = Not String.IsNullOrEmpty([Alias])
        If HasAlias Then
            Return $"{TableName} AS {[Alias]}"
        Else
            Return $"{TableName}"
        End If
    End Function
End Class
