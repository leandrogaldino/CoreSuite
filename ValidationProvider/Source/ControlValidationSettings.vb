Friend Class ControlValidationSettings
    Public Property Required As Boolean
    Public Property CustomValidationEnabled As Boolean
    Public Property ValidationGroup As String = String.Empty
    Public Property ValidationMessage As String = String.Empty
    Public Property ValidationDisplayName As String = String.Empty
    Public Property MinimumLength As Integer
    Public Property MaximumLength As Integer
    Public Property RegularExpression As String = String.Empty
    Public Property CompiledRegularExpression As Text.RegularExpressions.Regex
    Public Property CompareWith As Control
    Public Property ValuePropertyName As String = String.Empty
    Public Property IsRegistered As Boolean
    Public ReadOnly Property IsDefault As Boolean
        Get
            Return Not Required AndAlso Not CustomValidationEnabled AndAlso String.IsNullOrEmpty(ValidationGroup) AndAlso String.IsNullOrEmpty(ValidationMessage) AndAlso String.IsNullOrEmpty(ValidationDisplayName) AndAlso MinimumLength = 0 AndAlso MaximumLength = 0 AndAlso String.IsNullOrEmpty(RegularExpression) AndAlso CompareWith Is Nothing AndAlso String.IsNullOrEmpty(ValuePropertyName)
        End Get
    End Property
End Class
