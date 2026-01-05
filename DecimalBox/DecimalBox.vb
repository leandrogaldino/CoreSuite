Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Windows.Forms.Design

<Designer(GetType(DecimalBox.DecimalBoxDesigner))>
Public Class DecimalBox
    Inherits TextBox

#Region "ENUMS"

    Public Enum DecimalTextBoxBorderStyles
        Custom
        FixedSingle
        Fixed3D
        None
    End Enum

#End Region

#Region "FIELDS"

    Private _DecimalOnly As Boolean = True
    Private _DecimalPlaces As Integer = 2
    Private _DecimalValue As Decimal = 0

    Private ReadOnly _DecimalList As New List(Of String) From {
        "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ",", ".", "+", "-"
    }

    Private _SuspendValueChange As Boolean

#End Region

#Region "PROPERTIES"

#Region "NEW PROPERTIES"

    Private _IncludeThousandSeparator As TriState = TriState.True

    Public Property IncludeThousandSeparator As TriState
        Get
            Return _IncludeThousandSeparator
        End Get
        Set(value As TriState)
            _IncludeThousandSeparator = value
            If IsNumeric(Text) Then
                Dim s As String = FormatNumber(Text, DecimalPlaces, TriState.False, TriState.False, TriState.False)
                Text = FormatNumber(s, DecimalPlaces, TriState.True, TriState.False, value)
            End If
        End Set
    End Property

    ''' <summary>
    ''' Especifica se o controle deve aceitar somente números.
    ''' </summary>
    <Category("Comportamento")>
    <Description("Especifica se o controle deve aceitar somente números.")>
    Public Property DecimalOnly As Boolean
        Get
            Return _DecimalOnly
        End Get
        Set(value As Boolean)
            _SuspendValueChange = True
            _DecimalOnly = value
            If value Then
                TextAlign = HorizontalAlignment.Right
                If IsNumeric(Text) Then
                    _DecimalValue = CDec(Text)
                    MyBase.Text = FormatNumber(Text, DecimalPlaces, TriState.True, TriState.False, IncludeThousandSeparator)
                Else
                    _DecimalValue = 0
                    If Not MyBase.Text = Nothing Then
                        MyBase.Text = FormatNumber(0, DecimalPlaces, TriState.True, TriState.False, IncludeThousandSeparator)
                    End If

                End If
            Else
                TextAlign = HorizontalAlignment.Left
                _DecimalValue = 0
                'MyBase.Text = Nothing
            End If
            _SuspendValueChange = False
        End Set
    End Property

    ''' <summary>
    ''' Especifica a quantidade de casas decimais a serem mostradas no controle caso a propriedade DecimalOnly seja verdadeira (A propriedade DecimalValue guarda todas as casas decimais).
    ''' </summary>
    <Category("Comportamento")>
    <Description("Especifica a quantidade de casas decimais a serem mostradas no controle caso a propriedade DecimalOnly seja verdadeira (A propriedade DecimalValue guarda todas as casas decimais).")>
    Public Property DecimalPlaces As Integer
        Get
            Return _DecimalPlaces
        End Get
        Set(value As Integer)
            _SuspendValueChange = True
            If value > 99 Then value = 99
            _DecimalPlaces = value
            If DecimalOnly Then
                If Not MyBase.Text = Nothing Then
                    MyBase.Text = FormatNumber(DecimalValue, value, TriState.True, TriState.False, IncludeThousandSeparator)
                End If

            End If
            _SuspendValueChange = False
        End Set
    End Property

    ''' <summary>
    ''' Retorna o valor armazenado no controle com todas as casas decimais se a propriedade DecimalOnly for verdadeira.
    ''' </summary>
    <Category("Aparência")>
    <Description("Retorna o valor armazenado no controle com todas as casas decimais se a propriedade DecimalOnly for verdadeira.")>
    Public ReadOnly Property DecimalValue As Decimal
        Get
            Return _DecimalValue
        End Get
    End Property

#End Region

#Region "OVERRIDED PROPERTIES"

    Public Overrides Property Multiline As Boolean
        Get
            Return MyBase.Multiline
        End Get
        Set(value As Boolean)
            If value Then
                DecimalOnly = False
            End If
            MyBase.Multiline = value
        End Set
    End Property

    <RefreshProperties(RefreshProperties.All)>
    Public Overrides Property Text As String
        Get
            Return MyBase.Text
        End Get
        Set(value As String)
            If DesignMode Then
                If DecimalOnly Then
                    If IsNumeric(value) Then
                        If Not _SuspendValueChange Then _DecimalValue = CDec(value)
                        MyBase.Text = FormatNumber(value, DecimalPlaces, TriState.True, TriState.False, IncludeThousandSeparator)
                    Else
                        _DecimalValue = 0
                        MyBase.Text = Nothing
                    End If
                Else
                    _DecimalValue = 0
                    MyBase.Text = value
                End If
            Else
                If DecimalOnly Then
                    If IsNumeric(value) Then
                        _DecimalValue = CDec(value)
                        MyBase.Text = FormatNumber(value, DecimalPlaces, TriState.True, TriState.False, IncludeThousandSeparator)
                    End If
                Else
                    If value = Nothing Then
                        _DecimalValue = 0
                        MyBase.Text = Nothing
                    Else
                        _DecimalValue = 0
                        MyBase.Text = value
                    End If
                End If
            End If
        End Set
    End Property

#End Region

#End Region

#Region "PUBLIC SUBS"

    Public Sub New()
        TextAlign = HorizontalAlignment.Right
    End Sub

#End Region

#Region "OVERRIDED SUBS"

    Protected Overrides Sub OnTextChanged(e As EventArgs)
        If Not _SuspendValueChange And Not DesignMode Then
            If DecimalOnly And IsNumeric(Text) Then
                _DecimalValue = CDec(Text)
            Else
                _DecimalValue = FormatNumber(0, DecimalPlaces, TriState.True, TriState.False, IncludeThousandSeparator)
            End If
            MyBase.OnTextChanged(e)
        End If
    End Sub

    Protected Overrides Sub OnLostFocus(e As EventArgs)
        MyBase.OnLostFocus(e)
        _SuspendValueChange = True
        If DecimalOnly And IsNumeric(Text) Then
            If Text <> FormatNumber(_DecimalValue, DecimalPlaces, TriState.True, TriState.False, IncludeThousandSeparator) Then
                _DecimalValue = CDec(Text)
            End If
            MyBase.Text = FormatNumber(Text, DecimalPlaces, TriState.True, TriState.False, IncludeThousandSeparator)
        Else
            _DecimalValue = 0
            MyBase.Text = FormatNumber(0, DecimalPlaces, TriState.True, TriState.False, IncludeThousandSeparator)
            'MyBase.OnTextChanged(e)
        End If
        _SuspendValueChange = False
    End Sub

    Protected Overrides Sub OnEnter(e As EventArgs)
        MyBase.OnEnter(e)
        _SuspendValueChange = True
        If DecimalOnly Then
            If _DecimalValue = 0 Then
                MyBase.Text = Nothing
            End If
        End If
        _SuspendValueChange = False
    End Sub

    Protected Overrides Sub OnKeyPress(e As KeyPressEventArgs)
        MyBase.OnKeyPress(e)
        If DecimalOnly And Not Char.IsControl(e.KeyChar) Then
            If Not _DecimalList.Contains(e.KeyChar) Then e.Handled = True
            If Text.Length = 28 Then e.Handled = True
        End If
    End Sub

#End Region

#Region "INTERNAL CLASSES"

    Class DecimalBoxDesigner
        Inherits ControlDesigner

        Public Sub New()
        End Sub

        Public Overrides Sub InitializeNewComponent(ByVal defaultValues As IDictionary)
            MyBase.InitializeNewComponent(defaultValues)
            Dim myTextBox As DecimalBox = TryCast(Component, DecimalBox)
            If myTextBox IsNot Nothing Then
                myTextBox.Text = "0"
            End If
        End Sub

    End Class

#End Region

End Class