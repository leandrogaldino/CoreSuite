Imports System.ComponentModel

Partial Public Class LookupBox
    Private _PlaceholderText As String = String.Empty

    ''' <summary>Placeholder displayed by the native edit control when the lookup is empty and not focused.</summary>
    <DefaultValue(""), Localizable(True), Category("Appearance"), Description("Text displayed when the control is empty and does not have focus.")>
    Public Property PlaceholderText As String
        Get
            Return _PlaceholderText
        End Get
        Set(value As String)
            _PlaceholderText = If(value, String.Empty)
            If IsHandleCreated Then LookupBoxInterop.SetCueBanner(Me, _PlaceholderText)
        End Set
    End Property
End Class
