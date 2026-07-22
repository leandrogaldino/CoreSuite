Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Reflection
Imports Microsoft.DotNet.DesignTools.Designers
Imports Microsoft.DotNet.DesignTools.Designers.Actions

Partial Public Class QueriedBox
    Public Class Condition
        ''' <summary>
        ''' Define o nome ou apelido da tabela do banco de dados onde será aplicada a condição.
        ''' </summary>
        <Description("Define o nome ou apelido da tabela do banco de dados onde será aplicada a condição.")>
        <Category("Query")>
        Public Property TableNameOrAlias As String
        ''' <summary>
        ''' Define o nome do campo do banco de dados onde será aplicada a condição.
        ''' </summary>
        <Description("Define o nome do campo do banco de dados onde será aplicada a condição.")>
        <Category("Query")>
        Public Property FieldName As String
        ''' <summary>
        ''' Define o operador da condição. Para o operador BETWEEN, separar os dois valores por ponto e vírgula (;).
        ''' </summary>
        <Description("Define o operador da condição. Para o operador BETWEEN, separar os dois valores por ponto e vírgula (;).")>
        <Category("Query")>
        <TypeConverter(GetType(OperatorFilterCollection))>
        Public Property [Operator] As String
        ''' <summary>
        ''' Define o valor a ser testado na condição.
        ''' </summary>
        <Description("Define o valor a ser testado na condição.")>
        <Category("Query")>
        Public Property Value As String
        Public Overrides Function ToString() As String
            If TableNameOrAlias <> Nothing And FieldName <> Nothing And [Operator] <> Nothing And Value <> Nothing Then
                If [Operator] = "BETWEEN" Then
                    If Value.Split(";").Length = 2 Then
                        Return String.Format("{0}.{1} {2} {3} AND {4}", TableNameOrAlias, FieldName, [Operator], Value.Split(";").ElementAt(0), Value.Split(";").ElementAt(1))
                    Else
                        Return "New Undefined" & MyBase.GetType.Name
                    End If
                Else
                    Return String.Format("{0}.{1} {2} {3}", TableNameOrAlias, FieldName, [Operator], Value)
                End If
            ElseIf TableNameOrAlias <> Nothing And FieldName <> Nothing And [Operator] <> Nothing And Value = Nothing Then
                If [Operator] = "BETWEEN" Then
                    Return String.Format("{0}.{1} {2} {3} AND {4}", TableNameOrAlias, FieldName, [Operator], "Nothing", "Nothing")
                Else
                    Return String.Format("{0}.{1} {2} {3}", TableNameOrAlias, FieldName, [Operator], "Nothing")
                End If
            Else
                Return "New Undefined" & MyBase.GetType.Name
            End If
        End Function
    End Class
    Class OperatorFilterCollection
        Inherits StringConverter
        Private Shared ReadOnly _OperatorList As New List(Of String) From {
                "=",
                "<>",
                ">",
                ">=",
                "<",
                "<=",
                "BETWEEN",
                "LIKE"
            }
        Public Overrides Function GetStandardValuesSupported(context As ITypeDescriptorContext) As Boolean
            Return True
        End Function
        Public Overrides Function GetStandardValuesExclusive(context As ITypeDescriptorContext) As Boolean
            Return True
        End Function
        Public Overrides Function GetStandardValues(context As ITypeDescriptorContext) As StandardValuesCollection
            Return New StandardValuesCollection(_OperatorList)
        End Function
    End Class
    Public Class OtherField
        ''' <summary>
        ''' Define o nome da tabela a qual está relacionada aos resultados que serão exibidos.
        ''' </summary>
        <Description("Define o nome da tabela a qual está relacionada aos resultados que serão exibidos.")>
        <Category("Query")>
        Public Property DisplayTableName As String
        ''' <summary>
        ''' Define o apelido da tabela a qual está relacionada aos resultados que serão exibidos.
        ''' </summary>
        <Description("Define o apelido da tabela a qual está relacionada aos resultados que serão exibidos.")>
        <Category("Query")>
        Public Property DisplayTableAlias As String
        ''' <summary>
        ''' Define o nome do campo que será que será exibido em conjunto nos resultados da pesquisa.
        ''' </summary>
        <Description("Define o nome do campo que será que será exibido em conjunto nos resultados da pesquisa.")>
        <Category("Query")>
        Public Property DisplayFieldName As String
        ''' <summary>
        ''' Define o apelido do campo que será que será exibido em conjunto nos resultados da pesquisa.
        ''' </summary>
        <Description("Define o apelido do campo que será que será exibido em conjunto nos resultados da pesquisa.")>
        <Category("Query")>
        Public Property DisplayFieldAlias As String
        ''' <summary>
        ''' Define o nome do campo da tabela que está atribuído como Primary Key.
        ''' </summary>
        <Description("Define o nome do campo da tabela que está atribuído como Primary Key.")>
        <Category("Query")>
        Public Property DisplayMainFieldName As String
        ''' <summary>
        ''' Define se o valor desse campo será congelado no controle.
        ''' </summary>
        <Description("Define se o valor desse campo será congelado no controle.")>
        <Category("Comportamento")>
        Public Property Freeze As Boolean
        ''' <summary>
        ''' Define um prefixo que será congelado com o resultado.
        ''' </summary>
        <Category("Comportamento")>
        <Description("Define um prefixo que será congelado com o resultado.")>
        Public Property Prefix As String
        ''' <summary>
        ''' Define um sufixo que será congelado com o resultado.
        ''' </summary>
        <Category("Comportamento")>
        <Description("Define um sufixo que será congelado com o resultado.")>
        Public Property Suffix As String
        ''' <summary>
        ''' Define se o campo será exibido no resultado.
        ''' </summary>
        <Category("Comportamento")>
        <Description("Define se o campo será exibido no resultado.")>
        Public Property Display As Boolean = True
        ''' <summary>
        ''' Define um valor a ser retornado caso a consulta retorne um valor nulo para esse campo.
        ''' </summary>
        <Description("Define um valor a ser retornado caso a consulta retorne um valor nulo para esse campo.")>
        <Category("Query")>
        Public Property IfNull As String
        ''' <summary>
        ''' Define como a largura da coluna será ajustada nos resultados.
        ''' </summary>
        <Category("Aparência")>
        <DefaultValue(GetType(Color), "Blue")>
        <Description("Define como a largura da coluna será ajustada nos resultados.")>
        Public Property DisplayFieldAutoSizeColumnMode As DataGridViewAutoSizeColumnMode
        Public Overrides Function ToString() As String
            If DisplayTableName <> Nothing Or DisplayFieldName <> Nothing Then
                Return If(DisplayTableAlias = Nothing, DisplayTableName, DisplayTableAlias) & "." & If(DisplayFieldAlias = Nothing, DisplayFieldName, DisplayFieldAlias)
            Else
                Return "Undefined " & MyBase.GetType.Name
            End If
        End Function
    End Class
    Public Class Parameter
        ''' <summary>
        ''' Define o nome do parâmetro utilizado na Query.
        ''' </summary>
        <Description("Define o nome do parâmetro utilizado nas condições da Query.")>
        <Category("Query")>
        Public Property ParameterName As String
        ''' <summary>
        ''' Define o valor do parâmetro utilizado na Query.
        ''' </summary>
        <Description("Define o valor do parâmetro utilizado nas condições da Query.")>
        <Category("Query")>
        Public Property ParameterValue As String
        Public Overrides Function ToString() As String
            If ParameterName <> Nothing And ParameterValue <> Nothing Then
                Return ParameterName & " = " & ParameterValue
            ElseIf ParameterName <> Nothing And ParameterValue = Nothing Then
                Return ParameterName & " = Nothing"
            Else
                Return "New Undefined" & MyBase.GetType.Name
            End If
        End Function
    End Class
    Public Class Relation
        ''' <summary>
        ''' Define condições adicionais para o relacionamento.
        ''' </summary>
        <Description("Define condições adicionais para o relacionamento.")>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
        <Category("Query")>
        Public Property Conditions As New Collection(Of Condition)

        ''' <summary>
        ''' Define o tipo de relacionamento.
        ''' </summary>
        <Description("Define o tipo de relacionamento.")>
        <TypeConverter(GetType(RelationFilterCollection))>
        <Category("Query")>
        Public Property RelationType As String
        ''' <summary>
        ''' Define o operador a ser utilizado na relação.
        ''' </summary>
        <Description("Define o operador a ser utilizado na relação.")>
        <Category("Query")>
        <TypeConverter(GetType(OperatorFilterCollection))>
        Public Property [Operator] As String
        ''' <summary>
        ''' Define o nome da tabela que será relacionada.
        ''' </summary>
        <Description("Define o nome da tabela que será relacionada.")>
        <Category("Query")>
        Public Property RelateTableName As String
        ''' <summary>
        ''' Define o apelido da tabela que será relacionada.
        ''' </summary>
        <Description("Define o apelido da tabela que será relacionada.")>
        <Category("Query")>
        Public Property RelateTableAlias As String
        ''' <summary>
        ''' Define o nome do campo que se será relacionado.
        ''' </summary>
        <Description("Define o nome do campo que se será relacionado.")>
        <Category("Query")>
        Public Property RelateFieldName As String
        ''' <summary>
        ''' Define o nome da tabela que será relacionada com a tabela principal.
        ''' </summary>
        <Description("Define o nome da tabela que será relacionada com a tabela principal.")>
        <Category("Query")>
        Public Property WithTableName As String
        ''' <summary>
        ''' Define o apelido da tabela que se será relacionada com a tabela principal.
        ''' </summary>
        <Description("Define o apelido da tabela que se será relacionada com a tabela principal.")>
        <Category("Query")>
        Public Property WithTableAlias As String
        ''' <summary>
        ''' Define o nome do campo que será relacionado com o campo da tabela principal.
        ''' </summary>
        <Description("Define o nome do campo que será relacionado com o campo da tabela principal.")>
        <Category("Query")>
        Public Property WithFieldName As String
        Public Sub New()
        End Sub
        Public Overrides Function ToString() As String
            If RelateTableName = Nothing Or WithTableName = Nothing Or RelateFieldName = Nothing Or WithFieldName = Nothing Or [Operator] = Nothing Or RelationType = Nothing Then
                Return "New Undefined " & MyBase.GetType.Name
            Else
                Return String.Format("{0} JOIN {1} AS {2} ON {2}.{3} {4} {5}.{6}",
                                     RelationType, RelateTableName, If(RelateTableAlias <> Nothing, RelateTableAlias, RelateTableName),
                                     RelateFieldName, [Operator], If(WithTableAlias <> Nothing, WithTableAlias, WithTableName), WithFieldName)
            End If
        End Function


    End Class
    Class RelationFilterCollection
        Inherits StringConverter
        Private Shared ReadOnly _JoinList As New List(Of String) From {
            "LEFT",
            "RIGHT",
            "INNER",
            "CROSS"
        }
        Public Overrides Function GetStandardValuesSupported(context As ITypeDescriptorContext) As Boolean
            Return True
        End Function
        Public Overrides Function GetStandardValuesExclusive(context As ITypeDescriptorContext) As Boolean
            Return True
        End Function
        Public Overrides Function GetStandardValues(context As ITypeDescriptorContext) As StandardValuesCollection
            Return New StandardValuesCollection(_JoinList)
        End Function
    End Class
    Private Class QueriedTextBoxControlDesignerActionList
        Inherits DesignerActionList
        Private ReadOnly _Control As QueriedBox
        Public Sub New(Designer As ControlDesigner)
            MyBase.New(Designer.Component)
            _Control = CType(Designer.Control, QueriedBox)
        End Sub
        Public Overrides Function GetSortedActionItems() As DesignerActionItemCollection
            Dim Items = New DesignerActionItemCollection From {
                New DesignerActionPropertyItem("QueryEnabled", "QueryEnabled", "Category1", "Define se as pesquisas estão habilitadas."),
                New DesignerActionPropertyItem("CharactersToQuery", "Characters To Query", "Category1", "Define a quantidade de caracteres necessários para iniciar a pesquisa."),
                New DesignerActionPropertyItem("MainTableName", "MainTableName", "Category2", "Define o nome da tabela principal."),
                New DesignerActionPropertyItem("MainTableAlias", "MainTableAlias", "Category2", "Define o apelido da tabela principal."),
                New DesignerActionPropertyItem("MainReturnFieldName", "MainReturnFieldName", "Category2", "Define o nome do campo atribuido como Primary Key na tabela principal."),
                New DesignerActionPropertyItem("DisplayTableName", "DisplayTableName", "Category3", "Define o nome da tabela a qual está relacionada aos resultados que serão exibidos."),
                New DesignerActionPropertyItem("DisplayTableAlias", "DisplayTableAlias", "Category3", "Define o apelido da tabela a qual está relacionada aos resultados que serão exibidos."),
                New DesignerActionPropertyItem("DisplayFieldName", "DisplayFieldName", "Category3", "Define o nome do campo que será congelado no controle quando escolhido pelo usuário."),
                New DesignerActionPropertyItem("DisplayFieldAlias", "DisplayFieldAlias", "Category3", "Define o apelido do campo que será congelado no controle quando escolhido pelo usuário."),
                New DesignerActionPropertyItem("DisplayMainFieldName", "DisplayMainFieldName", "Category3", "Define o nome do campo da tabela que está atribuído como Primary Key."),
                New DesignerActionPropertyItem("Limit", "Limit", "Category4", "Define o máximo de resultados que podem ser retornados pela pesquisa.")
            }
            Return Items
        End Function

        Public Property QueryEnabled As Boolean
            Get
                Return _Control.QueryEnabled
            End Get
            Set(value As Boolean)
                TypeDescriptor.GetProperties(Component)("QueryEnabled").SetValue(Component, value)
            End Set
        End Property
        Public Property QueryInterval As Integer
            Get
                Return _Control.QueryInterval
            End Get
            Set(value As Integer)
                TypeDescriptor.GetProperties(Component)("QueryInterval").SetValue(Component, value)
            End Set
        End Property
        Public Property CharactersToQuery As Integer
            Get
                Return _Control.CharactersToQuery
            End Get
            Set(value As Integer)
                TypeDescriptor.GetProperties(Component)("CharactersToQuery").SetValue(Component, value)
            End Set
        End Property
        Public Property MainTableName As String
            Get
                Return _Control.MainTableName
            End Get
            Set(value As String)
                TypeDescriptor.GetProperties(Component)("MainTableName").SetValue(Component, value)
            End Set
        End Property
        Public Property MainTableAlias As String
            Get
                Return _Control.MainTableAlias
            End Get
            Set(value As String)
                TypeDescriptor.GetProperties(Component)("MainTableAlias").SetValue(Component, value)
            End Set
        End Property
        Public Property MainReturnFieldName As String
            Get
                Return _Control.MainReturnFieldName
            End Get
            Set(value As String)
                TypeDescriptor.GetProperties(Component)("MainReturnFieldName").SetValue(Component, value)
            End Set
        End Property
        Public Property DisplayTableName As String
            Get
                Return _Control.DisplayTableName
            End Get
            Set(value As String)
                TypeDescriptor.GetProperties(Component)("DisplayTableName").SetValue(Component, value)
            End Set
        End Property
        Public Property DisplayTableAlias As String
            Get
                Return _Control.DisplayTableAlias
            End Get
            Set(value As String)
                TypeDescriptor.GetProperties(Component)("DisplayTableAlias").SetValue(Component, value)
            End Set
        End Property
        Public Property DisplayFieldName As String
            Get
                Return _Control.DisplayFieldName
            End Get
            Set(value As String)
                TypeDescriptor.GetProperties(Component)("DisplayFieldName").SetValue(Component, value)
            End Set
        End Property
        Public Property DisplayFieldAlias As String
            Get
                Return _Control.DisplayFieldAlias
            End Get
            Set(value As String)
                TypeDescriptor.GetProperties(Component)("DisplayFieldAlias").SetValue(Component, value)
            End Set
        End Property
        Public Property DisplayMainFieldName As String
            Get
                Return _Control.DisplayMainFieldName
            End Get
            Set(value As String)
                TypeDescriptor.GetProperties(Component)("DisplayMainFieldName").SetValue(Component, value)
            End Set
        End Property
        Public Property Limit As Integer
            Get
                Return _Control.Limit
            End Get
            Set(value As Integer)
                TypeDescriptor.GetProperties(Component)("Limit").SetValue(Component, value)
            End Set
        End Property
    End Class
    Private Class FormDropDownResults
        Inherits Form
        Public Textbox As Control
        Private ReadOnly _QueriedBox As QueriedBox
        Public Sub New(Control As QueriedBox)
            SuspendLayout()
            _QueriedBox = Control
            InitializeComponent()
            BackColor = _QueriedBox.DropDownBorderColor
            Font = Control.Font
            FormBorderStyle = FormBorderStyle.None
            Padding = New Padding(1)
            Size = New Size(300, 120)
            DoubleBuffered = True
            TopMost = True
            KeyPreview = True
            ResumeLayout(True)
        End Sub
        Private Sub InitializeComponent()
            Dim ShowVGridLines As Boolean = _QueriedBox.ShowVerticalGridLines
            PanelContainer = New Panel With {
                .Dock = DockStyle.Fill
            }
            DgvResults = New DataGridView With {
                .AllowUserToAddRows = False,
                .AllowUserToDeleteRows = False,
                .AllowUserToResizeColumns = True,
                .AllowUserToResizeRows = False,
                .AllowUserToOrderColumns = True,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                .BackgroundColor = Color.White,
                .BorderStyle = BorderStyle.None,
                .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                .CellBorderStyle = If(_QueriedBox.ShowVerticalGridLines, DataGridViewCellBorderStyle.Single, DataGridViewCellBorderStyle.SingleHorizontal),
                .ColumnHeadersBorderStyle = If(_QueriedBox.ShowVerticalGridLines, DataGridViewHeaderBorderStyle.Raised, DataGridViewCellBorderStyle.None),
                .ColumnHeadersVisible = _QueriedBox.GridHeaderVisible,
                .DefaultCellStyle = New DataGridViewCellStyle With {
                    .SelectionBackColor = _QueriedBox.GridSelectionBackColor,
                    .SelectionForeColor = _QueriedBox.GridSelectionForeColor,
                    .BackColor = _QueriedBox.GridBackColor,
                    .ForeColor = _QueriedBox.GridForeColor
                },
                .ColumnHeadersDefaultCellStyle = New DataGridViewCellStyle With {
                    .BackColor = _QueriedBox.GridHeaderBackColor,
                    .ForeColor = _QueriedBox.GridHeaderForeColor
                },
                .Dock = DockStyle.Fill,
                .MultiSelect = False,
                .[ReadOnly] = True,
                .RowHeadersVisible = False,
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                .Visible = False,
                .EnableHeadersVisualStyles = False
            }
            DgvResults.ColumnHeadersDefaultCellStyle.Font = New Font(_QueriedBox.Font, If(_QueriedBox.GridHeadersBold, FontStyle.Bold, FontStyle.Regular))
            EnableDoubleBuffered(DgvResults)
            LblCharsRemaining = New Label With {
                .AutoSize = False,
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleCenter,
                .BackColor = _QueriedBox.LabelBackColor,
                .ForeColor = _QueriedBox.LabelForeColor,
                .Visible = False
            }
            PanelContainer.Controls.AddRange({DgvResults, LblCharsRemaining})
            Controls.Add(PanelContainer)
        End Sub
        Private Sub DgvResults_DataSourceChanged(sender As Object, e As EventArgs) Handles DgvResults.DataSourceChanged
            If DgvResults.Rows.Count = 0 Then
                DgvResults.Visible = False
                LblCharsRemaining.Text = "Nenhum Registro Encontrado."
                LblCharsRemaining.Visible = True
            End If
        End Sub
        Private Sub DropDownResults_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            Application.AddMessageFilter(New PopupWindowHelperMessageFilter(Me, Textbox))
        End Sub
        Private Shared Sub EnableDoubleBuffered(dgv As DataGridView)
            Dim DgvType As Type = dgv.[GetType]()
            Dim Info As PropertyInfo = DgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)
            Info.SetValue(dgv, True, Nothing)
        End Sub
        Protected Overrides ReadOnly Property CreateParams As CreateParams
            Get
                Dim Rect As CreateParams = MyBase.CreateParams
                Rect.Style = CInt(Flags.WindowStyles.WS_SYSMENU) Or CInt(Flags.WindowStyles.WS_CHILD)
                Rect.ExStyle = Rect.ExStyle Or CInt(Flags.WindowStyles.WS_EX_NOACTIVATE) Or CInt(Flags.WindowStyles.WS_EX_TOOLWINDOW)
                Rect.X = Me.Location.X
                Rect.Y = Me.Location.Y
                Return Rect
            End Get
        End Property
        Private Sub DataGridView_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles DgvResults.PreviewKeyDown
            If e.KeyCode = Keys.Tab Then
                Close()
                Me.Select()
            End If
        End Sub
        Private Sub DataGridView_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles DgvResults.MouseDoubleClick
            Dim Click As DataGridView.HitTestInfo = DgvResults.HitTest(e.X, e.Y)
            If Click.Type = DataGridViewHitTestType.Cell Then
                _QueriedBox.AutoFreeze()
                _QueriedBox.Focus()
                Close()
            End If
        End Sub
        Friend WithEvents DgvResults As DataGridView
        Friend WithEvents LblCharsRemaining As Label
        Friend WithEvents PanelContainer As Panel
    End Class
    Private Class PopupWindowHelperMessageFilter
        Implements IMessageFilter
        Private Const WM_LBUTTONDOWN As Integer = &H201
        Private Const WM_RBUTTONDOWN As Integer = &H204
        Private Const WM_MBUTTONDOWN As Integer = &H207
        Private Const WM_NCLBUTTONDOWN As Integer = &HA1
        Private Const WM_NCRBUTTONDOWN As Integer = &HA4
        Private Const WM_NCMBUTTONDOWN As Integer = &HA7
        Private TextBox As Control = Nothing
        Public Property Popup As Form = Nothing
        Public Sub New(popupW As Form, textbox As Control)
            Popup = popupW
            Me.TextBox = textbox
        End Sub
        <DebuggerStepThrough>
        Private Sub OnMouseDown()
            Dim CursorPos As Point = Cursor.Position
            If TextBox.Parent Is Nothing Then Exit Sub
            If Not Popup.Bounds.Contains(CursorPos) Then
                If Not TextBox.Bounds.Contains(TextBox.Parent.PointToClient(CursorPos)) Then
                    If CType(TextBox, QueriedBox).DropDownResultsForm IsNot Nothing AndAlso CType(TextBox, QueriedBox).DropDownResultsForm.DgvResults.SelectedRows.Count = 1 Then
                        If UCase(CType(TextBox, QueriedBox).Text) = UCase(CType(TextBox, QueriedBox).DropDownResultsForm.DgvResults.SelectedRows(0).Cells(If(CType(TextBox, QueriedBox).DisplayFieldAlias = Nothing, CType(TextBox, QueriedBox).DisplayFieldName, CType(TextBox, QueriedBox).DisplayFieldAlias)).Value.ToString) Then
                            CType(TextBox, QueriedBox).AutoFreeze()
                        End If
                    End If
                    Application.RemoveMessageFilter(Me)
                    Popup.Close()
                End If
            End If
        End Sub
        <DebuggerStepThrough>
        Private Function IMessageFilter_PreFilterMessage(ByRef m As Message) As Boolean Implements IMessageFilter.PreFilterMessage
            If Popup IsNot Nothing Then
                Select Case m.Msg
                    Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN, WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN, WM_NCMBUTTONDOWN
                        OnMouseDown()
                End Select
            End If
            Return False
        End Function
    End Class
    Private Class Flags
        <Flags()>
        Public Enum WindowStyles As UInteger
            WS_OVERLAPPED = 0
            WS_POPUP = 2147483648
            WS_CHILD = 1073741824
            WS_MINIMIZE = 536870912
            WS_VISIBLE = 268435456
            WS_DISABLED = 134217728
            WS_CLIPSIBLINGS = 67108864
            WS_CLIPCHILDREN = 33554432
            WS_MAXIMIZE = 16777216
            WS_BORDER = 8388608
            WS_DLGFRAME = 4194304
            WS_VSCROLL = 2097152
            WS_HSCROLL = 1048576
            WS_SYSMENU = 524288
            WS_THICKFRAME = 262144
            WS_GROUP = 131072
            WS_TABSTOP = 65536
            WS_MINIMIZEBOX = 131072
            WS_MAXIMIZEBOX = 65536
            WS_CAPTION = WS_BORDER Or WS_DLGFRAME
            WS_TILED = WS_OVERLAPPED
            WS_ICONIC = WS_MINIMIZE
            WS_SIZEBOX = WS_THICKFRAME
            WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
            WS_OVERLAPPEDWINDOW = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
            WS_POPUPWINDOW = WS_POPUP Or WS_BORDER Or WS_SYSMENU
            WS_CHILDWINDOW = WS_CHILD
            WS_EX_DLGMODALFRAME = 1
            WS_EX_NOPARENTNOTIFY = 4
            WS_EX_TOPMOST = 8
            WS_EX_ACCEPTFILES = 16
            WS_EX_TRANSPARENT = 32
            WS_EX_MDICHILD = 64
            WS_EX_TOOLWINDOW = 128
            WS_EX_WINDOWEDGE = 256
            WS_EX_CLIENTEDGE = 512
            WS_EX_CONTEXTHELP = 1024
            WS_EX_RIGHT = 4096
            WS_EX_LEFT = 0
            WS_EX_RTLREADING = 8192
            WS_EX_LTRREADING = 0
            WS_EX_LEFTSCROLLBAR = 16384
            WS_EX_RIGHTSCROLLBAR = 0
            WS_EX_CONTROLPARENT = 65536
            WS_EX_STATICEDGE = 131072
            WS_EX_APPWINDOW = 262144
            WS_EX_OVERLAPPEDWINDOW = WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE
            WS_EX_PALETTEWINDOW = WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST
            WS_EX_LAYERED = 524288
            WS_EX_NOINHERITLAYOUT = 1048576
            WS_EX_LAYOUTRTL = 4194304
            WS_EX_COMPOSITED = 33554432
            WS_EX_NOACTIVATE = 134217728
        End Enum
    End Class
End Class
