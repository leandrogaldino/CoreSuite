Imports System.ComponentModel

Public Partial Class ColoredProgressBar

    ' --------------------------------------------------
    ' Behavior
    ' --------------------------------------------------

    <Category("Behavior")>
    <Description("Gets or sets the minimum progress value.")>
    <DefaultValue(0)>
    Public Property Minimum As Integer
        Get
            Return _Minimum
        End Get
        Set(value As Integer)
            Dim NewMinimum As Integer = Math.Max(0, value)

            If _Minimum = NewMinimum Then Return

            _Minimum = NewMinimum

            If _Maximum < _Minimum Then
                _Maximum = _Minimum
            End If

            NormalizeValue()
        End Set
    End Property

    <Category("Behavior")>
    <Description("Gets or sets the maximum progress value.")>
    <DefaultValue(100)>
    Public Property Maximum As Integer
        Get
            Return _Maximum
        End Get
        Set(value As Integer)
            Dim NewMaximum As Integer = Math.Max(0, value)

            If _Maximum = NewMaximum Then Return

            _Maximum = NewMaximum

            If _Minimum > _Maximum Then
                _Minimum = _Maximum
            End If

            NormalizeValue()
        End Set
    End Property

    <Category("Behavior")>
    <Description("Gets or sets the current progress value.")>
    <DefaultValue(0)>
    Public Property Value As Integer
        Get
            Return _Value
        End Get
        Set(value As Integer)
            Dim NewValue As Integer =
                Math.Max(_Minimum, Math.Min(_Maximum, value))

            If _Value = NewValue Then Return

            _Value = NewValue

            UpdateVisualProgress()
            OnValueChanged(EventArgs.Empty)
        End Set
    End Property

    <Category("Behavior")>
    <Description("Gets or sets the increment used by PerformStep.")>
    <DefaultValue(10)>
    Public Property [Step] As Integer
        Get
            Return _Step
        End Get
        Set(value As Integer)
            _Step = value
        End Set
    End Property

    ' --------------------------------------------------
    ' Read-only properties
    ' --------------------------------------------------

    <Browsable(False)>
    Public ReadOnly Property Percentage As Double
        Get
            If _Maximum <= _Minimum Then Return 0.0R

            Dim CurrentRange As Double =
                CDbl(_Maximum) - CDbl(_Minimum)

            Dim CurrentValue As Double =
                CDbl(_Value) - CDbl(_Minimum)

            Return Math.Max(
                0.0R,
                Math.Min(100.0R, CurrentValue / CurrentRange * 100.0R))
        End Get
    End Property

    <Browsable(False)>
    Public ReadOnly Property DisplayedPercentage As Double
        Get
            Return _DisplayedProgress * 100.0R
        End Get
    End Property

    ' --------------------------------------------------
    ' Events
    ' --------------------------------------------------

    Public Event ValueChanged As EventHandler

    Protected Overridable Sub OnValueChanged(e As EventArgs)
        RaiseEvent ValueChanged(Me, e)
    End Sub

    ' --------------------------------------------------
    ' Public methods
    ' --------------------------------------------------

    Public Sub PerformStep()
        Increment(_Step)
    End Sub

    Public Sub Increment(amount As Integer)
        Dim NewValue As Long = CLng(_Value) + CLng(amount)

        NewValue = Math.Max(
            CLng(_Minimum),
            Math.Min(CLng(_Maximum), NewValue))

        Value = CInt(NewValue)
    End Sub

    ' --------------------------------------------------
    ' Progress normalization
    ' --------------------------------------------------

    Private Sub NormalizeValue()
        Dim OldValue As Integer = _Value

        _Value = Math.Max(
            _Minimum,
            Math.Min(_Maximum, _Value))

        UpdateVisualProgress()

        If OldValue <> _Value Then
            OnValueChanged(EventArgs.Empty)
        End If
    End Sub

End Class
