Imports CoreSuite.Helpers
Imports System.ComponentModel
Partial Public Class ActionTextBox
    Private _Placeholder As TextBoxPlaceholder
    <DefaultValue(""), Localizable(True), Category("Appearance"), Description("Text shown when the control is empty and does not have focus.")>
    Public Overridable Property PlaceholderText As String
        Get
            Return If(_Placeholder Is Nothing, String.Empty, _Placeholder.Text)
        End Get
        Set(Value As String)
            If _Placeholder Is Nothing Then _Placeholder = New TextBoxPlaceholder(Me)
            _Placeholder.Text = If(Value, String.Empty)
            Invalidate()
        End Set
    End Property
End Class
