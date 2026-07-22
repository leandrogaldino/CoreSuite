Public Class CMessageBoxOptions
    Private _ErrorImage As Image = My.Resources.ImageResources._Error
    Private _SuccessImage As Image = My.Resources.ImageResources.Success
    Private _InformationImage As Image = My.Resources.ImageResources.Information
    Private _WarningImage As Image = My.Resources.ImageResources.Warning
    Private _QuestionImage As Image = My.Resources.ImageResources.Question
    Public Property ShowExceptionDetails As Boolean = False
    Public Property ExceptionEmail As CMessageBoxExceptionEmail
    Public Property TitleFont As Font = New Font("Segoe UI", 11.25, FontStyle.Bold)
    Public Property TitleForeColor As Color = SystemColors.ControlText
    Public Property MessageFont As Font = New Font("Segoe UI", 9.75)
    Public Property MessageForeColor As Color = SystemColors.ControlText
    Public Property AdditionalInformations As Dictionary(Of String, Object)
    Public Property ErrorImage As Image
        Get
            Return _ErrorImage
        End Get
        Set(value As Image)
            If value IsNot Nothing AndAlso (value.Width <> 64 OrElse value.Height <> 64) Then
                Throw New ArgumentException("A imagem deve ter exatamente 64x64 pixels.", NameOf(value))
            End If
            _ErrorImage = value
        End Set
    End Property
    Public Property SuccessImage As Image
        Get
            Return _SuccessImage
        End Get
        Set(value As Image)
            If value IsNot Nothing AndAlso (value.Width <> 64 OrElse value.Height <> 64) Then
                Throw New ArgumentException("A imagem deve ter exatamente 64x64 pixels.", NameOf(value))
            End If
            _SuccessImage = value
        End Set
    End Property
    Public Property InformationImage As Image
        Get
            Return _InformationImage
        End Get
        Set(value As Image)
            If value IsNot Nothing AndAlso (value.Width <> 64 OrElse value.Height <> 64) Then
                Throw New ArgumentException("A imagem deve ter exatamente 64x64 pixels.", NameOf(value))
            End If
            _InformationImage = value
        End Set
    End Property
    Public Property WarningImage As Image
        Get
            Return _WarningImage
        End Get
        Set(value As Image)
            If value IsNot Nothing AndAlso (value.Width <> 64 OrElse value.Height <> 64) Then
                Throw New ArgumentException("A imagem deve ter exatamente 64x64 pixels.", NameOf(value))
            End If
            _WarningImage = value
        End Set
    End Property
    Public Property QuestionImage As Image
        Get
            Return _QuestionImage
        End Get
        Set(value As Image)
            If value IsNot Nothing AndAlso (value.Width <> 64 OrElse value.Height <> 64) Then
                Throw New ArgumentException("A imagem deve ter exatamente 64x64 pixels.", NameOf(value))
            End If
            _QuestionImage = value
        End Set
    End Property
End Class
