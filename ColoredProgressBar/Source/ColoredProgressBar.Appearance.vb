Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Drawing2D

Public Partial Class ColoredProgressBar

    ' --------------------------------------------------
    ' Appearance
    ' --------------------------------------------------

    <Category("Appearance")>
    <Description("Gets or sets the initial gradient color.")>
    Public Property ProgressStartColor As Color
        Get
            Return _ProgressStartColor
        End Get
        Set(value As Color)
            If _ProgressStartColor = value Then Return

            _ProgressStartColor = value
            Invalidate()
        End Set
    End Property

    <Category("Appearance")>
    <Description("Gets or sets the final gradient color.")>
    Public Property ProgressEndColor As Color
        Get
            Return _ProgressEndColor
        End Get
        Set(value As Color)
            If _ProgressEndColor = value Then Return

            _ProgressEndColor = value
            Invalidate()
        End Set
    End Property

    <Category("Appearance")>
    <Description("Gets or sets the progress track color.")>
    Public Property TrackColor As Color
        Get
            Return _TrackColor
        End Get
        Set(value As Color)
            If _TrackColor = value Then Return

            _TrackColor = value
            Invalidate()
        End Set
    End Property

    <Category("Appearance")>
    <Description("Gets or sets the border color.")>
    Public Property BorderColor As Color
        Get
            Return _BorderColor
        End Get
        Set(value As Color)
            If _BorderColor = value Then Return

            _BorderColor = value
            Invalidate()
        End Set
    End Property

    <Category("Appearance")>
    <Description("Gets or sets the border thickness.")>
    <DefaultValue(0)>
    Public Property BorderThickness As Integer
        Get
            Return _BorderThickness
        End Get
        Set(value As Integer)
            Dim NewValue As Integer = Math.Max(0, value)

            If _BorderThickness = NewValue Then Return

            _BorderThickness = NewValue
            Invalidate()
        End Set
    End Property

    <Category("Appearance")>
    <Description("Gets or sets the corner radius.")>
    <DefaultValue(9)>
    Public Property CornerRadius As Integer
        Get
            Return _CornerRadius
        End Get
        Set(value As Integer)
            Dim NewValue As Integer = Math.Max(0, value)

            If _CornerRadius = NewValue Then Return

            _CornerRadius = NewValue
            Invalidate()
        End Set
    End Property

    <Category("Appearance")>
    <Description("Gets or sets the gradient direction.")>
    <DefaultValue(LinearGradientMode.Vertical)>
    Public Property GradientDirection As LinearGradientMode
        Get
            Return _GradientDirection
        End Get
        Set(value As LinearGradientMode)
            If Not [Enum].IsDefined(GetType(LinearGradientMode), value) Then
                Throw New InvalidEnumArgumentException(
                    NameOf(value),
                    CInt(value),
                    GetType(LinearGradientMode))
            End If

            If _GradientDirection = value Then Return

            _GradientDirection = value
            Invalidate()
        End Set
    End Property

    <Category("Appearance")>
    <Description("Gets or sets whether the percentage is displayed.")>
    <DefaultValue(False)>
    Public Property ShowPercentage As Boolean
        Get
            Return _ShowPercentage
        End Get
        Set(value As Boolean)
            If _ShowPercentage = value Then Return

            _ShowPercentage = value
            Invalidate()
        End Set
    End Property

    <Category("Appearance")>
    <Description("Gets or sets the text color over the filled area.")>
    Public Property ProgressTextColor As Color
        Get
            Return _ProgressTextColor
        End Get
        Set(value As Color)
            If _ProgressTextColor = value Then Return

            _ProgressTextColor = value
            Invalidate()
        End Set
    End Property

End Class
