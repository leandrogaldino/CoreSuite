
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Represents a button that suppresses focus cues and provides
''' built-in tooltip support.
''' </summary>
Public Class NoFocusCueButton
    Inherits Button

    Private _ShowDefaultBorder As Boolean
    Private _IsDefaultButton As Boolean
    Private _FocusedBackColor As Color = Color.FromArgb(233, 235, 238)
    Private _FocusedBorderColor As Color = Color.FromArgb(133, 139, 147)
    Private _FocusAppearanceApplied As Boolean
    Private _NormalBackColor As Color
    Private _NormalBorderColor As Color
    Private _NormalUseVisualStyleBackColor As Boolean

    ''' <summary>Controls the additional default-button border.</summary>
    <Category("Appearance"), DefaultValue(False)>
    Public Property ShowDefaultBorder As Boolean
        Get
            Return _ShowDefaultBorder
        End Get
        Set(Value As Boolean)
            If _ShowDefaultBorder = Value Then Return
            _ShowDefaultBorder = Value
            MyBase.NotifyDefault(_IsDefaultButton AndAlso Value)
            Invalidate()
        End Set
    End Property

    <Category("Appearance"), DefaultValue(GetType(Color), "233, 235, 238")>
    <Description("Background color of a focused flat button. Empty preserves the normal color.")>
    Public Property FocusedBackColor As Color
        Get
            Return _FocusedBackColor
        End Get
        Set(Value As Color)
            If _FocusedBackColor = Value Then Return
            RestoreFocusAppearance()
            _FocusedBackColor = Value
            ApplyFocusAppearance()
        End Set
    End Property

    <Category("Appearance"), DefaultValue(GetType(Color), "133, 139, 147")>
    <Description("Border color of a focused flat button. Empty preserves the normal color.")>
    Public Property FocusedBorderColor As Color
        Get
            Return _FocusedBorderColor
        End Get
        Set(Value As Color)
            If _FocusedBorderColor = Value Then Return
            RestoreFocusAppearance()
            _FocusedBorderColor = Value
            ApplyFocusAppearance()
        End Set
    End Property

    Public Overrides Sub NotifyDefault(Value As Boolean)
        _IsDefaultButton = Value
        MyBase.NotifyDefault(Value AndAlso _ShowDefaultBorder)
    End Sub

    Protected Overrides Sub OnGotFocus(e As EventArgs)
        ApplyFocusAppearance()
        MyBase.OnGotFocus(e)
    End Sub

    Protected Overrides Sub OnLostFocus(e As EventArgs)
        RestoreFocusAppearance()
        MyBase.OnLostFocus(e)
    End Sub

    Protected Overrides Sub OnEnabledChanged(e As EventArgs)
        RestoreFocusAppearance()
        ApplyFocusAppearance()
        MyBase.OnEnabledChanged(e)
    End Sub

    Private Sub ApplyFocusAppearance()
        If _FocusAppearanceApplied OrElse Not Focused OrElse Not Enabled OrElse DesignMode Then Return
        If FlatStyle <> FlatStyle.Flat OrElse FlatAppearance.BorderSize = 0 Then Return
        If _FocusedBackColor.IsEmpty AndAlso _FocusedBorderColor.IsEmpty Then Return

        _NormalBackColor = BackColor
        _NormalBorderColor = FlatAppearance.BorderColor
        _NormalUseVisualStyleBackColor = UseVisualStyleBackColor
        _FocusAppearanceApplied = True

        If Not _FocusedBackColor.IsEmpty Then
            BackColor = _FocusedBackColor
            UseVisualStyleBackColor = False
        End If
        If Not _FocusedBorderColor.IsEmpty Then FlatAppearance.BorderColor = _FocusedBorderColor
        Invalidate()
    End Sub

    Private Sub RestoreFocusAppearance()
        If Not _FocusAppearanceApplied Then Return
        _FocusAppearanceApplied = False
        If Not _FocusedBackColor.IsEmpty Then
            If BackColor = _FocusedBackColor Then
                BackColor = _NormalBackColor
                UseVisualStyleBackColor = _NormalUseVisualStyleBackColor
            End If
        End If
        If Not _FocusedBorderColor.IsEmpty AndAlso FlatAppearance.BorderColor = _FocusedBorderColor Then
            FlatAppearance.BorderColor = _NormalBorderColor
        End If
        Invalidate()
    End Sub

    Private ReadOnly _ToolTipComponent As New ToolTip With {
        .ShowAlways = True
    }

    Private _ToolTipTextValue As String = String.Empty

    ''' <summary>
    ''' Gets a value indicating that focus cues should never
    ''' be displayed for this button.
    ''' </summary>
    Protected Overrides ReadOnly Property ShowFocusCues As Boolean
        Get
            Return False
        End Get
    End Property

    ''' <summary>
    ''' Gets or sets the tooltip text displayed when the
    ''' mouse pointer hovers over this button.
    ''' </summary>
    <Category("Behavior")>
    <Description("Specifies the tooltip text displayed when hovering over the button.")>
    <DefaultValue("")>
    <Localizable(True)>
    Public Property TooltipText As String
        Get
            Return _ToolTipTextValue
        End Get
        Set(Value As String)
            Dim NewValue As String = If(Value, String.Empty)

            If String.Equals(_ToolTipTextValue, NewValue, StringComparison.Ordinal) Then
                Return
            End If

            _ToolTipTextValue = NewValue
            _ToolTipComponent.SetToolTip(Me, NewValue)
        End Set
    End Property

    ''' <summary>
    ''' Releases the resources used by the button.
    ''' </summary>
    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing Then
            _ToolTipComponent.Dispose()
        End If

        MyBase.Dispose(Disposing)
    End Sub

End Class
