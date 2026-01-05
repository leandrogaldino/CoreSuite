Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Windows.Forms.Design
Public Class PercentageBox
    Inherits TextBox
    Public Event PercentageValueChanged(sender As Object, e As EventArgs)
#Region "ENUMS"
    Public Enum PercentageBoxBorderStyles
        Custom
        FixedSingle
        Fixed3D
        None
    End Enum
    Public Enum ReturnFormat
        [Integer]
        Percent
    End Enum
#End Region
#Region "FIELDS"
    Private _DesignerHost As IDesignerHost
    Private _BorderStyle As PercentageBoxBorderStyles = PercentageBoxBorderStyles.Custom
    Private _BorderColorDefault As Color = SystemColors.WindowFrame
    Private _BorderColorFocused As Color = SystemColors.HotTrack
    Private _ShowPercentSymbol As Boolean = True
    Private _PercentageValue As Integer = 0
    Private ReadOnly _DecimalList As New List(Of String) From {
        "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ",", ".", "+", "-"
    }
    Private _SuspendValueChange As Boolean
#End Region
#Region "PROPERTIES"
#Region "OVERLOADED PROPERTIES"
    ''' <summary>
    ''' Indica se o controle de edição deve ter uma borda.
    ''' </summary>
    <DefaultValue(GetType(PercentageBoxBorderStyles), "Custom")>
    <Category("Aparência")>
    <Description("Indica se o controle de edição deve ter uma borda.")>
    Overloads Property BorderStyle As PercentageBoxBorderStyles
        Get
            Return _BorderStyle
        End Get
        Set(value As PercentageBoxBorderStyles)
            _BorderStyle = value
            MyBase.BorderStyle = Windows.Forms.BorderStyle.FixedSingle
            Select Case value
                Case Is = PercentageBoxBorderStyles.Fixed3D
                    MyBase.BorderStyle = Windows.Forms.BorderStyle.Fixed3D
                Case Is = PercentageBoxBorderStyles.None
                    MyBase.BorderStyle = Windows.Forms.BorderStyle.None
            End Select
        End Set
    End Property
#End Region
#Region "NEW PROPERTIES"
    ''' <summary>
    ''' Especifica se a propriedade PercentageValue vai retornar um valor inteiro ou o valor real da porcentagem.
    ''' </summary>
    <Category("Comportamento")>
    <Description("Especifica se a propriedade PercentageValue vai retornar um valor inteiro ou o valor real da porcentagem.")>
    Public Property PercentageValueFormat As ReturnFormat = ReturnFormat.Percent

    ''' <summary>
    ''' Especifica se o simbolo de porcentagem (%) será mostrado.
    ''' </summary>
    <Category("Comportamento")>
    <Description("Especifica se o simbolo de porcentagem (%) será mostrado.")>
    Public Property ShowPercentSymbol As Boolean
        Get
            Return _ShowPercentSymbol
        End Get
        Set(value As Boolean)
            _SuspendValueChange = True
            _ShowPercentSymbol = value
            If value Then
                If IsNumeric(Text) Then
                    _PercentageValue = CInt(Text)
                    RaiseEvent PercentageValueChanged(Me, EventArgs.Empty)
                    MyBase.Text = Text & "%"
                Else
                    _PercentageValue = 0
                    RaiseEvent PercentageValueChanged(Me, EventArgs.Empty)
                    MyBase.Text = 0
                End If
            Else
                _PercentageValue = Text.Replace("%", Nothing)
                RaiseEvent PercentageValueChanged(Me, EventArgs.Empty)
                MyBase.Text = Text.Replace("%", Nothing)
            End If
            _SuspendValueChange = False
        End Set
    End Property
    ''' <summary>
    ''' Retorna o valor armazenado no controle com todas as casas decimais se a propriedade DecimalOnly for verdadeira.
    ''' </summary>
    <Category("Aparência")>
    <Description("Retorna o valor armazenado no controle com todas as casas decimais se a propriedade DecimalOnly for verdadeira.")>
    Public ReadOnly Property PercentageValue As Decimal
        Get
            If PercentageValueFormat = ReturnFormat.Integer Then
                Return _PercentageValue
            Else
                Return _PercentageValue / 100
            End If


        End Get
    End Property
    ''' <summary>
    ''' Define a cor da borda quando o controle está com foco.
    ''' </summary>
    <DefaultValue(GetType(Color), "HotTrack")>
    <Category("Aparência")>
    <Description("Define a cor da borda quando o controle está com foco.")>
    Public Property BorderColorFocused As Color
        Get
            Return _BorderColorFocused
        End Get
        Set(value As Color)
            _BorderColorFocused = value
            If BorderStyle = PercentageBoxBorderStyles.Custom Then
                WndProc(New Message With {.Msg = 133})
            End If
        End Set
    End Property
    ''' <summary>
    '''"Define a cor da borda quando o controle está sem foco.
    ''' </summary>
    <DefaultValue(GetType(Color), "WindowFrame")>
    <Category("Aparência")>
    <Description("Define a cor da borda quando o controle está sem foco.")>
    Public Property BorderColorDefault As Color
        Get
            Return _BorderColorDefault
        End Get
        Set(value As Color)
            _BorderColorDefault = value
            If BorderStyle = PercentageBoxBorderStyles.Custom Then
                WndProc(New Message With {.Msg = 133})
            End If
        End Set
    End Property
#End Region
#Region "OVERRIDED PROPERTIES"
    Public Overrides Property Multiline As Boolean
        Get
            Return MyBase.Multiline
        End Get
        Set(value As Boolean)
            If value Then
                ShowPercentSymbol = False
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
                If ShowPercentSymbol Then
                    If IsNumeric(value.Replace("%", Nothing)) Then
                        If Not _SuspendValueChange Then
                            _PercentageValue = CInt(value.Replace("%", Nothing))
                            RaiseEvent PercentageValueChanged(Me, EventArgs.Empty)
                        End If
                        MyBase.Text = CInt(value.Replace("%", Nothing)) & "%"
                    Else
                        _PercentageValue = 0
                        RaiseEvent PercentageValueChanged(Me, EventArgs.Empty)
                        MyBase.Text = "0%"
                    End If
                Else
                    If IsNumeric(value.Replace("%", Nothing)) Then
                        _PercentageValue = value.Replace("%", Nothing)
                        RaiseEvent PercentageValueChanged(Me, EventArgs.Empty)
                        MyBase.Text = value.Replace("%", Nothing)
                    Else
                        _PercentageValue = 0
                        RaiseEvent PercentageValueChanged(Me, EventArgs.Empty)
                        MyBase.Text = "0"
                    End If
                End If
            Else
                If ShowPercentSymbol Then
                    If IsNumeric(value.Replace("%", Nothing)) Then
                        _PercentageValue = CInt(value.Replace("%", Nothing))
                        RaiseEvent PercentageValueChanged(Me, EventArgs.Empty)
                        _SuspendValueChange = True
                        MyBase.Text = CInt(value.Replace("%", Nothing)) & "%"
                        _SuspendValueChange = False
                    End If
                Else
                    _PercentageValue = 0
                    RaiseEvent PercentageValueChanged(Me, EventArgs.Empty)
                    MyBase.Text = CInt(value.Replace("%", Nothing))
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
            If ShowPercentSymbol Then
                If IsNumeric(Text.Replace("%", Nothing)) Then
                    _PercentageValue = Text.Replace("%", Nothing)
                    RaiseEvent PercentageValueChanged(Me, EventArgs.Empty)
                    MyBase.Text = Text.Replace("%", Nothing) '& "%"
                Else
                    _PercentageValue = 0
                    RaiseEvent PercentageValueChanged(Me, EventArgs.Empty)
                    MyBase.Text = Nothing
                End If
            Else
                If IsNumeric(Text.Replace("%", Nothing)) Then
                    _PercentageValue = Text.Replace("%", Nothing)
                    RaiseEvent PercentageValueChanged(Me, EventArgs.Empty)
                    MyBase.Text = Text.Replace("%", Nothing)
                Else
                    _PercentageValue = 0
                    RaiseEvent PercentageValueChanged(Me, EventArgs.Empty)
                    MyBase.Text = Nothing
                End If
            End If
        End If
        MyBase.OnTextChanged(e)
    End Sub
    Protected Overrides Sub OnLostFocus(e As EventArgs)
        MyBase.OnLostFocus(e)
        _SuspendValueChange = True
        If ShowPercentSymbol Then
            If Not IsNumeric(Text.Replace("%", Nothing)) Then
                _PercentageValue = "0"
                RaiseEvent PercentageValueChanged(Me, EventArgs.Empty)
                MyBase.Text = 0 & "%"
            Else
                If Text.Replace("%", Nothing) <> _PercentageValue Then
                    _PercentageValue = CInt(Text.Replace("%", Nothing))
                    RaiseEvent PercentageValueChanged(Me, EventArgs.Empty)
                End If
                MyBase.Text = CInt(Text.Replace("%", Nothing)) & "%"
            End If
        Else
            If Not IsNumeric(Text.Replace("%", Nothing)) Then
                _PercentageValue = "0"
                RaiseEvent PercentageValueChanged(Me, EventArgs.Empty)
                MyBase.Text = "0"
            Else
                If Text.Replace("%", Nothing) <> _PercentageValue Then
                    _PercentageValue = CInt(Text.Replace("%", Nothing))
                    RaiseEvent PercentageValueChanged(Me, EventArgs.Empty)
                End If
                MyBase.Text = CInt(Text.Replace("%", Nothing))
            End If
        End If
        _SuspendValueChange = False
    End Sub

    Protected Overrides Sub OnGotFocus(e As EventArgs)
        MyBase.OnGotFocus(e)
        _SuspendValueChange = True
        If ShowPercentSymbol Then
            MyBase.Text = MyBase.Text.Replace("%", Nothing)
            If MyBase.Text = "0" Then MyBase.Text = Nothing
        End If
        Me.Select(MyBase.Text.Length, 0)
        _SuspendValueChange = False
    End Sub
    Protected Overrides Sub WndProc(ByRef m As Message)
        MyBase.WndProc(m)
        Dim dc As IntPtr
        If m.Msg = 133 AndAlso Focused AndAlso BorderStyle = PercentageBoxBorderStyles.Custom Then
            dc = GetWindowDC(Handle)
            Using g As Graphics = Graphics.FromHdc(dc)
                Using p As Pen = New Pen(BorderColorFocused)
                    g.DrawRectangle(p, 0, 0, Width - 1, Height - 1)
                End Using
            End Using
        ElseIf m.Msg = 133 AndAlso Not Focused AndAlso BorderStyle = PercentageBoxBorderStyles.Custom Then
            dc = GetWindowDC(Handle)
            Using g As Graphics = Graphics.FromHdc(dc)
                Using p As Pen = New Pen(BorderColorDefault)
                    g.DrawRectangle(p, 0, 0, Width - 1, Height - 1)
                End Using
            End Using
        End If
    End Sub
    Protected Overrides Sub OnHandleCreated(ByVal e As EventArgs)
        MyBase.OnHandleCreated(e)

        If DesignMode AndAlso Site IsNot Nothing Then
            _DesignerHost = TryCast(Site.GetService(GetType(IDesignerHost)), IDesignerHost)
            If _DesignerHost IsNot Nothing Then
                Dim designer = CType(_DesignerHost.GetDesigner(Me), ControlDesigner)
                If designer IsNot Nothing Then
                    Dim actions = designer.ActionLists(0)
                    designer.ActionLists.Clear()
                    designer.ActionLists.Add(New PercentageBoxControlDesignerActionList(designer, actions))
                End If
            End If
        End If
    End Sub
    Protected Overrides Sub OnKeyPress(e As KeyPressEventArgs)
        MyBase.OnKeyPress(e)
        If ShowPercentSymbol And Not Char.IsControl(e.KeyChar) Then
            If Not _DecimalList.Contains(e.KeyChar) Then e.Handled = True
        End If
    End Sub
#End Region
#Region "PRIVATE FUNCTIONS"
    Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As IntPtr) As IntPtr
#End Region
#Region "INTERNAL CLASSES"
    Private Class PercentageBoxControlDesignerActionList
        Inherits DesignerActionList
        Private Control As PercentageBox
        Private Designer As ControlDesigner
        Private ActionList As DesignerActionList
        Public Sub New(ByVal Designer As ControlDesigner, ByVal ActionList As DesignerActionList)
            MyBase.New(Designer.Component)
            Me.Designer = Designer
            Me.ActionList = ActionList
            Control = CType(Designer.Control, PercentageBox)
        End Sub
        Public Overrides Function GetSortedActionItems() As DesignerActionItemCollection
            Dim Items = New DesignerActionItemCollection From {
                New DesignerActionPropertyItem("PercentageValueFormat", "Percentage Value Format", "Comportamento", "Especifica se a propriedade PercentageValue vai retornar um valor inteiro ou o valor real da porcentagem."),
                New DesignerActionPropertyItem("ShowPercentSymbol", "Show Percent Symbol", "Comportamento", "Especifica se o simbolo de porcentagem (%) será mostrado.")
            }
            Return Items
        End Function

        Public Property PercentageValueFormat As ReturnFormat
            Get
                Return Control.PercentageValueFormat
            End Get
            Set(ByVal value As ReturnFormat)
                TypeDescriptor.GetProperties(Component)("PercentageValueFormat").SetValue(Component, value)
            End Set
        End Property

        Public Property ShowPercentSymbol As Boolean
            Get
                Return Control.ShowPercentSymbol
            End Get
            Set(ByVal value As Boolean)
                TypeDescriptor.GetProperties(Component)("ShowPercentSymbol").SetValue(Component, value)
            End Set
        End Property
    End Class
#End Region
End Class
