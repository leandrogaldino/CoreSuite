Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Public Class ToggleButton
    Inherits Control
#Region "EVENTS"
    <Category("Propriedade Alterada")>
    Public Event StateChanged(sender As Control, e As EventArgs)
#End Region
#Region "ENUMS"
    Public Enum ToggleButtonStates
        [OFF]
        [ON]
    End Enum
#End Region
#Region "FIELDS"
    Private _BorderColor As Color = SystemColors.WindowFrame
    Private _State As ToggleButtonStates = ToggleButtonStates.OFF
#End Region
#Region "CONSTRUCTOR"
    Public Sub New()
        Size = New Size(75, 20)
        Cursor = Cursors.Hand
        SetStyle(ControlStyles.Selectable, True)
    End Sub
#End Region
#Region "PROPERTIES"
#Region "HIDDEN PROPRTIES"
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Overrides Property BackColor As Color
        Get
            Return MyBase.BackColor
        End Get
        Set(value As Color)
            MyBase.BackColor = value
        End Set
    End Property
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Overrides Property BackgroundImage As Image
        Get
            Return MyBase.BackgroundImage
        End Get
        Set(value As Image)
            MyBase.BackgroundImage = value
        End Set
    End Property
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Overrides Property BackgroundImageLayout As ImageLayout
        Get
            Return MyBase.BackgroundImageLayout
        End Get
        Set(value As ImageLayout)
            MyBase.BackgroundImageLayout = value
        End Set
    End Property
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Overrides Property Font As Font
        Get
            Return MyBase.Font
        End Get
        Set(value As Font)
            MyBase.Font = value
        End Set
    End Property
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Overrides Property ForeColor As Color
        Get
            Return MyBase.ForeColor
        End Get
        Set(value As Color)
            MyBase.ForeColor = value
        End Set
    End Property
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Overrides Property RightToLeft As RightToLeft
        Get
            Return MyBase.RightToLeft
        End Get
        Set(value As RightToLeft)
            MyBase.RightToLeft = value
        End Set
    End Property
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Overloads ReadOnly Property Text As String
        Get
            Return If(State = ToggleButtonStates.OFF, OffStyle.Text, OnStyle.Text)
        End Get
    End Property
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Overrides Property AllowDrop As Boolean
        Get
            Return MyBase.AllowDrop
        End Get
        Set(value As Boolean)
            MyBase.AllowDrop = value
        End Set
    End Property
#End Region
#Region "OVERRIDED PROPERTIES"
    <DefaultValue(GetType(Cursor), "Hand")>
    Public Overrides Property Cursor As Cursor
        Get
            Return MyBase.Cursor
        End Get
        Set(value As Cursor)
            MyBase.Cursor = value
        End Set
    End Property
#End Region
#Region "NEW PROPERTIES"
    <TypeConverter(GetType(ExpandableObjectConverter))>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Category("Aparência")>
    Public Property OffStyle() As Style = New Style(Me) With {.Text = "OFF"}
    <TypeConverter(GetType(ExpandableObjectConverter))>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Category("Aparência")>
    Public Property OnStyle() As Style = New Style(Me) With {.Text = "ON"}
    ''' <summary>
    ''' A cor da borda do controle.
    ''' </summary>
    <DefaultValue(GetType(Color), "WindowFrame")>
    <Category("Aparência")>
    <Description("A cor da borda do controle.")>
    Public Property BorderColor As Color
        Get
            Return _BorderColor
        End Get
        Set(value As Color)
            _BorderColor = value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' O estado do controle.
    ''' </summary>
    <Category("Comportamento")>
    <DefaultValue(GetType(ToggleButtonStates), "OFF")>
    <Description("O estado do controle.")>
    Public Property State As ToggleButtonStates
        Get
            Return _State
        End Get
        Set(value As ToggleButtonStates)
            Dim PrevState = _State
            _State = value
            If value <> PrevState Then
                RaiseEvent StateChanged(Me, New EventArgs)
            End If
            Invalidate()
        End Set
    End Property
#End Region
#End Region
#Region "OVERRIDED SUBS"
    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        OffStyle.Font = Font
        OnStyle.Font = Font
    End Sub
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        Dim ControlArea As Rectangle
        Dim SwitchArea As Rectangle
        Dim TextArea As Rectangle
        Dim SwitchAreaPoint As Point
        Dim SwitchAreaSize As Size
        Dim TextAreaPoint As Point
        Dim TextAreaSize As Size
        Dim SwitchBackColor As Color
        Dim TextForeColor As Color
        Dim Text As String = Nothing
        Dim Font As Font = Nothing
        Dim SwitchBorderColor As Color
        Dim SwitchBorderSize As Integer
        Dim HorizontalTextAlign As StringAlignment
        Dim VerticalTextAlign As StringAlignment
        Dim TextBackColor As Color
        Dim Image As Bitmap = Nothing
        Dim ScratchVisible As Boolean
        Dim ScratchColor As Color
        ControlArea = ClientRectangle
        If State = ToggleButtonStates.OFF Then
            If Enabled Then
                SwitchAreaSize = New Size(ControlArea.Width - (ControlArea.Width * (OffStyle.SwitchRatio / 100)), ControlArea.Height - 2)
                SwitchAreaPoint = New Point(ControlArea.X + 1, ControlArea.Y + 1)
                TextAreaSize = New Size(ControlArea.Width - SwitchAreaSize.Width, ControlArea.Height)
                TextAreaPoint = New Point(ControlArea.X + SwitchAreaSize.Width, ControlArea.Y)
                SwitchBackColor = OffStyle.SwitchBackColor
                TextForeColor = OffStyle.TextForeColor
                Text = OffStyle.Text
                Font = OffStyle.Font
                SwitchBorderColor = OffStyle.SwitchBorderColor
                SwitchBorderSize = OffStyle.SwitchBorderSize
                HorizontalTextAlign = GetStringAlign(OffStyle.TextAlign).ElementAt(1)
                VerticalTextAlign = GetStringAlign(OffStyle.TextAlign).ElementAt(0)
                TextBackColor = OffStyle.TextBackColor
                Image = OffStyle.SwitchImage
                ScratchVisible = OffStyle.ScratchVisible
                ScratchColor = OffStyle.ScratchColor
            ElseIf Not Enabled And Not DesignMode Then
                SwitchAreaSize = New Size(ControlArea.Width - (ControlArea.Width * (OffStyle.SwitchRatio / 100)), ControlArea.Height - 2)
                SwitchAreaPoint = New Point(ControlArea.X + 1, ControlArea.Y + 1)
                TextAreaSize = New Size(ControlArea.Width - SwitchAreaSize.Width, ControlArea.Height)
                TextAreaPoint = New Point(ControlArea.X + SwitchAreaSize.Width, ControlArea.Y)
                SwitchBackColor = SystemColors.ButtonFace
                TextForeColor = SystemColors.ControlDark
                Text = OffStyle.Text
                Font = OffStyle.Font
                SwitchBorderColor = SystemColors.ControlDark
                SwitchBorderSize = OffStyle.SwitchBorderSize
                HorizontalTextAlign = GetStringAlign(OffStyle.TextAlign).ElementAt(1)
                VerticalTextAlign = GetStringAlign(OffStyle.TextAlign).ElementAt(0)
                TextBackColor = OffStyle.TextBackColor
                Image = OffStyle.SwitchImage
                ScratchVisible = OffStyle.ScratchVisible
                ScratchColor = OffStyle.ScratchColor
            End If
        Else
            If Enabled Then
                TextAreaSize = New Size(ControlArea.Width - (ControlArea.Width * (OnStyle.SwitchRatio / 100)), ControlArea.Height - 2)
                TextAreaPoint = New Point(ControlArea.X + 1, ControlArea.Y + 1)
                SwitchAreaSize = New Size(ControlArea.Width - TextAreaSize.Width - 2, ControlArea.Height - 2)
                SwitchAreaPoint = New Point(ControlArea.X + TextAreaSize.Width + 1, ControlArea.Y + 1)
                SwitchBackColor = OnStyle.SwitchBackColor
                TextForeColor = OnStyle.TextForeColor
                Text = OnStyle.Text
                Font = OnStyle.Font
                SwitchBorderColor = OnStyle.SwitchBorderColor
                SwitchBorderSize = OnStyle.SwitchBorderSize
                HorizontalTextAlign = GetStringAlign(OnStyle.TextAlign).ElementAt(1)
                VerticalTextAlign = GetStringAlign(OnStyle.TextAlign).ElementAt(0)
                TextBackColor = OnStyle.TextBackColor
                Image = OnStyle.SwitchImage
                ScratchVisible = OnStyle.ScratchVisible
                ScratchColor = OnStyle.ScratchColor
            ElseIf Not Enabled And Not DesignMode Then
                TextAreaSize = New Size(ControlArea.Width - (ControlArea.Width * (OnStyle.SwitchRatio / 100)), ControlArea.Height - 2)
                TextAreaPoint = New Point(ControlArea.X + 1, ControlArea.Y + 1)
                SwitchAreaSize = New Size(ControlArea.Width - TextAreaSize.Width - 2, ControlArea.Height - 2)
                SwitchAreaPoint = New Point(ControlArea.X + TextAreaSize.Width + 1, ControlArea.Y + 1)
                SwitchBackColor = SystemColors.ButtonFace
                TextForeColor = SystemColors.ControlDark
                Text = OnStyle.Text
                Font = OnStyle.Font
                SwitchBorderColor = SystemColors.ControlDark
                SwitchBorderSize = OnStyle.SwitchBorderSize
                HorizontalTextAlign = GetStringAlign(OnStyle.TextAlign).ElementAt(1)
                VerticalTextAlign = GetStringAlign(OnStyle.TextAlign).ElementAt(0)
                TextBackColor = OnStyle.TextBackColor
                Image = OnStyle.SwitchImage
                ScratchVisible = OnStyle.ScratchVisible
                ScratchColor = OnStyle.ScratchColor
            End If
        End If
        SwitchArea = New Rectangle(SwitchAreaPoint, SwitchAreaSize)
        TextArea = New Rectangle(TextAreaPoint, TextAreaSize)
        Using Brush As New SolidBrush(SwitchBackColor)
            e.Graphics.FillRectangle(Brush, SwitchArea)
        End Using

        If ScratchVisible Then
            Using pen As New Pen(OffStyle.ScratchColor)
                e.Graphics.DrawLine(pen, New Point(SwitchArea.X + (SwitchArea.Width / 2) - 3, SwitchArea.Top + 3),
                                    New Point(SwitchArea.X + (SwitchArea.Width / 2) - 3, SwitchArea.Y + SwitchArea.Bottom - 5))
                e.Graphics.DrawLine(pen, New Point(SwitchArea.X + (SwitchArea.Width / 2) + 3, SwitchArea.Top + 3),
                                    New Point(SwitchArea.X + (SwitchArea.Width / 2) + 3, SwitchArea.Y + SwitchArea.Bottom - 5))
            End Using
        End If
        Using Brush As New SolidBrush(TextBackColor)
            e.Graphics.FillRectangle(Brush, TextArea)
        End Using
        If Image Is Nothing Then
            Using Brush As New SolidBrush(TextForeColor)
                Using StringFormat As New StringFormat()
                    StringFormat.Alignment = HorizontalTextAlign
                    StringFormat.LineAlignment = VerticalTextAlign
                    e.Graphics.DrawString(Text, Font, Brush, TextArea, StringFormat)
                End Using
            End Using
        Else
            e.Graphics.DrawImage(Image, New Point(TextAreaPoint.X + ((TextAreaSize.Width - 13) / 2), TextAreaPoint.Y + ((TextAreaSize.Height - 13) / 2)))
        End If
        If SwitchBorderSize > SwitchArea.Width Then SwitchBorderSize = SwitchArea.Width
        If SwitchBorderSize > SwitchArea.Height Then SwitchBorderSize = SwitchArea.Height
        ControlPaint.DrawBorder(CreateGraphics, SwitchArea,
                                SwitchBorderColor, SwitchBorderSize, ButtonBorderStyle.Solid,
                                SwitchBorderColor, SwitchBorderSize, ButtonBorderStyle.Solid,
                                SwitchBorderColor, SwitchBorderSize, ButtonBorderStyle.Solid,
                                SwitchBorderColor, SwitchBorderSize, ButtonBorderStyle.Solid)
        ControlPaint.DrawBorder(CreateGraphics, ControlArea, If(Focused, SystemColors.Highlight, _BorderColor), ButtonBorderStyle.Solid)

    End Sub
    Protected Overrides Sub OnClick(e As EventArgs)
        MyBase.OnClick(e)
        If State = ToggleButtonStates.OFF Then
            State = ToggleButtonStates.ON
        Else
            State = ToggleButtonStates.OFF
        End If
    End Sub
    Protected Overrides Sub OnGotFocus(ByVal e As EventArgs)
        Invalidate()
        MyBase.OnGotFocus(e)
    End Sub
    Protected Overrides Sub OnLostFocus(ByVal e As EventArgs)
        MyBase.OnLostFocus(e)
        Invalidate()
    End Sub
    Protected Overrides Sub OnMouseDown(ByVal e As MouseEventArgs)
        If Not Focused Then Focus()
        MyBase.OnMouseDown(e)
    End Sub
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        MyBase.OnKeyDown(e)
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Space Then
            State = If(State = ToggleButtonStates.OFF, ToggleButtonStates.ON, ToggleButtonStates.OFF)
        End If
    End Sub
#End Region
#Region "PRIVATE FUNCTIONS"
    Private Function GetStringAlign(ByVal Align As ContentAlignment) As StringAlignment()
        Select Case Align
            Case Is = ContentAlignment.TopLeft
                Return {StringAlignment.Near, StringAlignment.Near}
            Case Is = ContentAlignment.TopCenter
                Return {StringAlignment.Near, StringAlignment.Center}
            Case Is = ContentAlignment.TopRight
                Return {StringAlignment.Near, StringAlignment.Far}
            Case Is = ContentAlignment.MiddleLeft
                Return {StringAlignment.Center, StringAlignment.Near}
            Case Is = ContentAlignment.TopCenter
                Return {StringAlignment.Center, StringAlignment.Center}
            Case Is = ContentAlignment.MiddleRight
                Return {StringAlignment.Center, StringAlignment.Far}
            Case Is = ContentAlignment.BottomLeft
                Return {StringAlignment.Far, StringAlignment.Near}
            Case Is = ContentAlignment.BottomCenter
                Return {StringAlignment.Far, StringAlignment.Center}
            Case Is = ContentAlignment.BottomRight
                Return {StringAlignment.Far, StringAlignment.Far}
            Case Else
                Return {StringAlignment.Center, StringAlignment.Center}
        End Select
    End Function
#End Region
#Region "INTERNAL CLASSES"
    Public Class Style
        Private _Control As ToggleButton
        Private _ScratchVisible As Boolean = True
        Private _ScratchColor As Color = SystemColors.ControlDarkDark
        Private _SwitchRatio As Integer = 50
        Private _SwitchBorderSize As Integer = 1
        Private _SwitchBorderColor As Color = SystemColors.ButtonShadow
        Private _SwitchBackColor As Color = SystemColors.ButtonFace
        Private _SwitchImage As Bitmap = Nothing
        Private _Text As String
        Private _Font As Font = New Font("Microsoft Sans Serif", 8.25)
        Private _TextForeColor As Color = SystemColors.ControlText
        Private _TextBackColor As Color = SystemColors.Window
        Private _TextAlign As ContentAlignment = ContentAlignment.MiddleCenter
        Public Sub New(Optional ByVal Control As ToggleButton = Nothing)
            _Control = Control
        End Sub
        ''' <summary>
        ''' Define se o risco do gatilho está visível.
        ''' </summary>
        <Description("Define se o risco do gatilho está visível.")>
        <Category("Aparência")>
        <DefaultValue(GetType(Boolean), "True")>
        Public Property ScratchVisible As Boolean
            Get
                Return _ScratchVisible
            End Get
            Set(value As Boolean)
                _ScratchVisible = value
                If _Control IsNot Nothing Then _Control.Refresh()
            End Set
        End Property
        ''' <summary>
        ''' Define a cor do risco do gatilho.
        ''' </summary>
        <Description("Define a cor do risco do gatilho.")>
        <DefaultValue(GetType(Color), "ControlDarkDark")>
        Public Property ScratchColor As Color
            Get
                Return _ScratchColor
            End Get
            Set(value As Color)
                _ScratchColor = value
                If _Control IsNot Nothing Then _Control.Refresh()
            End Set
        End Property
        ''' <summary>
        ''' Define a cor da borda do gatilho.
        ''' </summary>
        <Description("Define a cor da borda do gatilho.")>
        <DefaultValue(GetType(Color), "ButtonShadow")>
        Public Property SwitchBorderColor As Color
            Get
                Return _SwitchBorderColor
            End Get
            Set(value As Color)
                _SwitchBorderColor = value
                If _Control IsNot Nothing Then _Control.Refresh()
            End Set
        End Property
        ''' <summary>
        ''' Define a cor de fundo do gatilho.
        ''' </summary>
        <Description("Define a cor de fundo do gatilho.")>
        <DefaultValue(GetType(Color), "ButtonFace")>
        Public Property SwitchBackColor As Color
            Get
                Return _SwitchBackColor
            End Get
            Set(value As Color)
                _SwitchBackColor = value
                If _Control IsNot Nothing Then _Control.Refresh()
            End Set
        End Property
        ''' <summary>
        ''' Define a imagem de fundo do gatilho. Essa propriedade sobrepõem a propriedade Text.
        ''' </summary>
        <DefaultValue(GetType(Bitmap), "Nenhum")>
        <Description("Define a imagem de fundo do gatilho. Essa propriedade sobrepõem a propriedade Text.")>
        Public Property SwitchImage As Bitmap
            Get
                Return _SwitchImage
            End Get
            Set(value As Bitmap)
                _SwitchImage = value
                If _Control IsNot Nothing Then _Control.Refresh()
            End Set
        End Property
        ''' <summary>
        ''' Define a porcentagem ocupada pelo gatilho em relação a area total do controle.
        ''' </summary>
        <Description("Define a porcentagem ocupada pelo gatilho em relação a area total do controle.")>
        <DefaultValue(GetType(Integer), "50")>
        Public Property SwitchRatio As Integer
            Get
                Return _SwitchRatio
            End Get
            Set(value As Integer)
                If value < 0 Then value = 0
                If value > 100 Then value = 100
                _SwitchRatio = value
                If _Control IsNot Nothing Then _Control.Refresh()
            End Set
        End Property
        ''' <summary>
        ''' Define a espessura da borda do gatilho.
        ''' </summary>
        <Description("Define a espessura da borda do gatilho.")>
        <DefaultValue(GetType(Integer), "1")>
        Public Property SwitchBorderSize As Integer
            Get
                Return _SwitchBorderSize
            End Get
            Set(value As Integer)
                If value < 0 Then value = 0
                If value > 20 Then value = 20
                _SwitchBorderSize = value
                If _Control IsNot Nothing Then _Control.Refresh()
            End Set
        End Property
        ''' <summary>
        ''' Define o texto do controle.
        ''' </summary>
        <Description("Define o texto do controle.")>
        Public Property Text As String
            Get
                Return _Text
            End Get
            Set(value As String)
                _Text = value
                If _Control IsNot Nothing Then _Control.Refresh()
            End Set
        End Property
        ''' <summary>
        ''' Define a fonte do texto do controle.
        ''' </summary>
        <Description("Define a fonte do texto do controle.")>
        Public Property Font As Font
            Get
                Return _Control.Font
            End Get
            Set(value As Font)
                _Control.Font = value
                If _Control IsNot Nothing Then _Control.Refresh()
            End Set
        End Property
        ''' <summary>
        ''' Define a cor do texto do controle.
        ''' </summary>
        <DefaultValue(GetType(Color), "ControlText")>
        <Description("Define a cor do texto do controle.")>
        Public Property TextForeColor As Color
            Get
                Return _TextForeColor
            End Get
            Set(value As Color)
                _TextForeColor = value
                If _Control IsNot Nothing Then _Control.Refresh()
            End Set
        End Property
        ''' <summary>
        ''' Define a cor de fundo do texto do controle.
        ''' </summary>
        <Description("Define a cor de fundo do texto do controle.")>
        <DefaultValue(GetType(Color), "Window")>
        Public Property TextBackColor As Color
            Get
                Return _TextBackColor
            End Get
            Set(value As Color)
                _TextBackColor = value
                If _Control IsNot Nothing Then _Control.Refresh()
            End Set
        End Property
        ''' <summary>
        ''' Define o alinhamento do texto do controle.
        ''' </summary>
        <Description("Define o alinhamento do texto do controle.")>
        <DefaultValue(GetType(ContentAlignment), "MiddleCenter")>
        Public Property TextAlign As ContentAlignment
            Get
                Return _TextAlign
            End Get
            Set(value As ContentAlignment)
                _TextAlign = value
                If _Control IsNot Nothing Then _Control.Refresh()
            End Set
        End Property
        Public Overrides Function ToString() As String
            Return Nothing
        End Function
    End Class
#End Region
End Class
