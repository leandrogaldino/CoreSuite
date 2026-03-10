Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports System.Data.Common
Imports System.Drawing
Imports System.Reflection
Imports System.Windows.Forms
Imports System.Windows.Forms.Design
Imports CoreSuite.Helpers

Public Class QueriedBox
    Inherits TextBox

    Private WithEvents Timer As Timer

    Private _CharactersToQuery As Integer = 3

    Private _CtrlHyperLink As Boolean = False

    Private _DesignerHost As IDesignerHost

    Private _DropDownAutoStretchRight As Boolean

    Private _DropDownBorderColor As Color = SystemColors.HotTrack

    Private _DropDownStretchRight As Integer

    Private _FirstEnter As Boolean = False

    Private _FreezeColor As Color = Color.Blue

    Private _FreezedPrimaryKey As Long

    Private _FreezedValue As String

    Private _Freezing As Boolean

    Private _GridBackColor As Color = SystemColors.Window

    Private _GridForeColor As Color = SystemColors.ControlText

    Private _GridHeaderBackColor As Color = SystemColors.Window

    Private _GridHeaderForeColor As Color = SystemColors.ControlText

    Private _GridSelectionBackColor As Color = SystemColors.HotTrack

    Private _GridSelectionForeColor As Color = SystemColors.Window

    Private _IsFreezed As Boolean

    Private _IsHyperLink As Boolean = False

    Private _KeyDown As Boolean

    Private _LabelBackColor As Color = SystemColors.Window

    Private _LabelForeColor As Color = SystemColors.ControlText

    Private _QueryEnabled As Boolean = True

    Private _QueryInterval As Integer = 300

    Private _RawFreezedValues As New List(Of (String, String, Object)) From {("Table", "Field", New Object())}

    Private _UnFreezeColor As Color

    Private DropDownResultsForm As FormDropDownResults

    <Category("Propriedade Alterada")>
    Public Event FreezedPrimaryKeyChanged(sender As Object, e As EventArgs)

    <Category("Propriedade Alterada")>
    Public Event FreezedPrimaryKeyChanging(sender As Object, e As EventArgs)

    <Category("Ação")>
    Public Event HyperlinkClicked(sender As Object, e As EventArgs)
    ''' <summary>
    ''' Define se será permitido o uso de hyperlinks em registros selecionados com a tecla control.
    ''' </summary>
    ''' <returns></returns>
    <DefaultValue(GetType(Boolean), "False")>
    <Category("Comportamento")>
    <Description("Define se será permitido o uso de hyperlinks em registros selecionados com a tecla control.")>
    Public Property AllowHyperlink As Boolean = False

    ''' <summary>
    ''' Define a quantidade de caracteres necessários para iniciar a pesquisa.
    ''' </summary>
    <Category("Comportamento")>
    <DefaultValue(GetType(Integer), "3")>
    <Description("Define a quantidade de caracteres necessários para iniciar a pesquisa.")>
    Public Property CharactersToQuery As Integer
        Get
            Return _CharactersToQuery
        End Get
        Set(value As Integer)
            If value < 1 Then value = 1
            _CharactersToQuery = value
        End Set
    End Property

    ''' <summary>
    ''' Limpa todo o conteúdo cado um valor seja descongelado.
    ''' </summary>
    <Category("Comportamento")>
    <DefaultValue(GetType(Boolean), "False")>
    <Description("Limpa todo o conteúdo cado um valor seja descongelado.")>
    Public Property ClearOnUnfreeze As Boolean

    ''' <summary>
    ''' Define o nome do campo da tabela a ser pesquisado.
    ''' </summary>
    <Category("Query")>
    <Description("Define condições para a pesquisa. Deve ser definida com a sintaxe SQL.")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <MergableProperty(False)>
    Public Property Conditions As New Collection(Of Condition)

    ''' <summary>
    ''' Define a Conexão a ser utilizada em todos os controles para o acesso aos dados.
    ''' </summary>
    ''' <returns></returns>
    <Description("Define a Conexão a ser utilizada em todos os controles para o acesso aos dados.")>
    Public Shared Property Connection As DbConnection

    ''' <summary>
    ''' Depura a consulta configurada no controle quando o texto é alterado.
    ''' </summary>
    <Category("Diversos")>
    <Description("Depura a consulta configurada no controle quando o texto é alterado.")>
    Public Property DebugOnTextChanged As Boolean

    ''' <summary>
    ''' Define o apelido do campo que será congelado no controle quando escolhido pelo usuário.
    ''' </summary>
    <Category("Query")>
    <Description("Define o apelido do campo que será congelado no controle quando escolhido pelo usuário.")>
    Public Property DisplayFieldAlias As String

    ''' <summary>
    ''' Define como a largura da coluna será ajustada nos resultados.
    ''' </summary>
    <Category("Aparência")>
    <DefaultValue(GetType(Color), "Blue")>
    <Description("Define como a largura da coluna será ajustada nos resultados.")>
    Public Property DisplayFieldAutoSizeColumnMode As DataGridViewAutoSizeColumnMode

    ''' <summary>
    ''' Define o nome do campo que será congelado no controle quando escolhido pelo usuário.
    ''' </summary>
    <Category("Query")>
    <Description("Define o nome do campo que será congelado no controle quando escolhido pelo usuário.")>
    Public Property DisplayFieldName As String

    ''' <summary>
    ''' Define o nome do campo da tabela que está atribuído como PRIMARYKEY.
    ''' </summary>
    <Category("Query")>
    <Description("Define o nome do campo da tabela que está atribuído como Primary Key.")>
    Public Property DisplayMainFieldName As String

    ''' <summary>
    ''' Define o apelido da tabela a qual está relacionada aos resultados que serão exibidos.
    ''' </summary>
    <Category("Query")>
    <Description("Define o apelido da tabela a qual está relacionada aos resultados que serão exibidos.")>
    <MergableProperty(False)>
    Public Property DisplayTableAlias As String

    ''' <summary>
    ''' Define o nome da tabela a qual está relacionada aos resultados que serão exibidos.
    ''' </summary>
    <Category("Query")>
    <Description("Define o nome da tabela a qual está relacionada aos resultados que serão exibidos.")>
    <MergableProperty(False)>
    Public Property DisplayTableName As String

    ''' <summary>
    ''' Define se o resultado deve ser distinguido (Ocultar resultados idênticos provenientes da Query).
    ''' </summary>
    <Category("Query")>
    <Description("Define se o resultado deve ser distinguido (Ocultar resultados idênticos provenientes da Query).")>
    Public Property Distinct As Boolean

    ''' <summary>
    ''' Define o quanto o painel será esticado automaticamente para a direita até mostrar todos os resultados.
    ''' </summary>
    <Category("Aparência")>
    <DefaultValue(GetType(Integer), "120")>
    <Description("Define o quanto o painel será esticado automaticamente para a direita até mostrar todos os resultados.")>
    Public Property DropDownAutoStretchRight As Boolean
        Get
            Return _DropDownAutoStretchRight
        End Get
        Set(value As Boolean)
            _DropDownAutoStretchRight = value
            If value Then DropDownStretchRight = 0
        End Set
    End Property

    ''' <summary>
    ''' Define a cor da borda do DropDown.
    ''' </summary>
    <Category("Aparência")>
    <DefaultValue(GetType(Color), "HotTrack")>
    <Description("Define a cor da borda do DropDown.")>
    Public Property DropDownBorderColor As Color
        Get
            Return _DropDownBorderColor
        End Get
        Set(value As Color)
            If value <> Color.Transparent Then
                _DropDownBorderColor = value
            Else
                Throw New ArgumentException("O controle não dá suporte a cores de borda transparente.")
            End If
        End Set
    End Property

    ''' <summary>
    ''' Define o quanto o painel será esticado para baixo.
    ''' </summary>
    <Category("Aparência")>
    <DefaultValue(GetType(Integer), "120")>
    <Description("Define o quanto o painel será esticado para baixo.")>
    Public Property DropDownStretchDown As Integer = 120

    ''' <summary>
    ''' Define o quanto o painel será esticado para a esquerda.
    ''' </summary>
    <Category("Aparência")>
    <DefaultValue(GetType(Integer), "0")>
    <Description("Define o quanto o painel será esticado para a esquerda.")>
    Public Property DropDownStretchLeft As Integer

    ''' <summary>
    ''' Define o quanto o painel será esticado para a direita.
    ''' </summary>
    <Category("Aparência")>
    <DefaultValue(GetType(Integer), "0")>
    <Description("Define o quanto o painel será esticado para a direita.")>
    Public Property DropDownStretchRight As Integer
        Get
            Return _DropDownStretchRight
        End Get
        Set(value As Integer)
            If DropDownAutoStretchRight Then value = 0
            _DropDownStretchRight = value
        End Set
    End Property

    Public Overrides Property ForeColor As Color
        Get
            Return MyBase.ForeColor
        End Get
        Set(value As Color)
            MyBase.ForeColor = value
            If Not _Freezing Then
                _UnFreezeColor = value
            End If
        End Set
    End Property

    ''' <summary>
    ''' Define cor do texto para quando um resultado da pesquisa for selecionado.
    ''' </summary>
    <Category("Aparência")>
    <DefaultValue(GetType(Color), "Blue")>
    <Description("Define cor do texto para quando um resultado da pesquisa for selecionado.")>
    Public Property FreezeColor As Color
        Get
            Return _FreezeColor
        End Get
        Set(value As Color)
            If value <> Color.Transparent Then
                _FreezeColor = value
            Else
                Throw New ArgumentException("O controle não dá suporte a cores de texto transparente.")
            End If
        End Set
    End Property

    ''' <summary>
    ''' Retorna a chave primária do registro selecionado ou 0 quando não há registro selecionado.
    ''' </summary>
    ''' <returns></returns>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Always)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property FreezedPrimaryKey As Long
        Get
            Return _FreezedPrimaryKey
        End Get
    End Property

    ''' <summary>
    ''' Retorna o valor do registro selecionado.
    ''' </summary>
    ''' <returns></returns>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Always)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property FreezedValue As String
        Get
            Return _FreezedValue
        End Get
    End Property

    ''' <summary>
    ''' Define a cor de fundo do grid.
    ''' </summary>
    <Category("Aparência")>
    <DefaultValue(GetType(Color), "Window")>
    <Description("Define a cor de fundo do grid.")>
    Public Property GridBackColor As Color
        Get
            Return _GridBackColor
        End Get
        Set(value As Color)
            If value <> Color.Transparent Then
                _GridBackColor = value
            Else
                Throw New ArgumentException("O controle não dá suporte a cores de fundo transparente.")
            End If
        End Set
    End Property

    ''' <summary>
    ''' Define a cor do texto do grid.
    ''' </summary>
    <Category("Aparência")>
    <DefaultValue(GetType(Color), "ControlText")>
    <Description("Define a cor do texto do grid.")>
    Public Property GridForeColor As Color
        Get
            Return _GridForeColor
        End Get
        Set(value As Color)
            If value <> Color.Transparent Then
                _GridForeColor = value
            Else
                Throw New ArgumentException("O controle não dá suporte a cores de texto transparente.")
            End If
        End Set
    End Property

    ''' <summary>
    ''' Define a cor de fundo do cabeçalho do grid.
    ''' </summary>
    <Category("Aparência")>
    <DefaultValue(GetType(Color), "Control")>
    <Description("Define a cor de fundo do cabeçalho do grid.")>
    Public Property GridHeaderBackColor As Color
        Get
            Return _GridHeaderBackColor
        End Get
        Set(value As Color)
            If value <> Color.Transparent Then
                _GridHeaderBackColor = value
            Else
                Throw New ArgumentException("O controle não dá suporte a cores de texto transparente.")
            End If
        End Set
    End Property

    ''' <summary>
    ''' Define a cor do texto do cabeçalho do grid.
    ''' </summary>
    <Category("Aparência")>
    <DefaultValue(GetType(Color), "ControlText")>
    <Description("Define a cor do texto do cabeçalho do grid.")>
    Public Property GridHeaderForeColor As Color
        Get
            Return _GridHeaderForeColor
        End Get
        Set(value As Color)
            If value <> Color.Transparent Then
                _GridHeaderForeColor = value
            Else
                Throw New ArgumentException("O controle não dá suporte a cores de texto transparente.")
            End If
        End Set
    End Property

    ''' <summary>
    ''' Define a fonte do cabeçalho do grid como negrito.
    ''' </summary>
    <Category("Aparência")>
    <DefaultValue(GetType(Boolean), "False")>
    <Description("Define a fonte do cabeçalho do grid como negrito.")>
    Public Property GridHeadersBold As Boolean = False

    ''' <summary>
    ''' Define o cabeçalho do grid como visível.
    ''' </summary>
    <Category("Aparência")>
    <DefaultValue(GetType(Boolean), "True")>
    <Description("Define o cabeçalho do grid como visível.")>
    Public Property GridHeaderVisible As Boolean = True

    ''' <summary>
    ''' Define a cor de seleção do grid.
    ''' </summary>
    <Category("Aparência")>
    <DefaultValue(GetType(Color), "HotTrack")>
    <Description("Define a cor de seleção do grid.")>
    Public Property GridSelectionBackColor As Color
        Get
            Return _GridSelectionBackColor
        End Get
        Set(value As Color)
            If value <> Color.Transparent Then
                _GridSelectionBackColor = value
            Else
                Throw New ArgumentException("O controle não dá suporte a cores de seleção transparente.")
            End If
        End Set
    End Property

    ''' <summary>
    ''' Define a cor de seleção do texto do grid.
    ''' </summary>
    <Category("Aparência")>
    <DefaultValue(GetType(Color), "Window")>
    <Description("Define a cor de seleção do texto do grid.")>
    Public Property GridSelectionForeColor As Color
        Get
            Return _GridSelectionForeColor
        End Get
        Set(value As Color)
            If value <> Color.Transparent Then
                _GridSelectionForeColor = value
            Else
                Throw New ArgumentException("O controle não dá suporte a cores de texto transparente.")
            End If
        End Set
    End Property

    ''' <summary>
    ''' Define um valor caso o resultado seja nulo.
    ''' </summary>
    <Category("Query")>
    <Description("Define um valor caso o resultado seja nulo.")>
    Public Property IfNull As String

    ''' <summary>
    ''' Retorna se existe um registro selecionado.
    ''' </summary>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Always)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property IsFreezed As Boolean
        Get
            Return _IsFreezed
        End Get
    End Property

    ''' <summary>
    ''' Define a cor de fundo da label que informa quantos caracteres faltam.
    ''' </summary>
    <Category("Aparência")>
    <DefaultValue(GetType(Color), "Window")>
    <Description("Define a cor de fundo da label que informa quantos caracteres faltam.")>
    Public Property LabelBackColor As Color
        Get
            Return _LabelBackColor
        End Get
        Set(value As Color)
            If value <> Color.Transparent Then
                _LabelBackColor = value
            Else
                Throw New ArgumentException("O controle não dá suporte a cores de texto transparente.")
            End If
        End Set
    End Property

    ''' <summary>
    ''' Define a cor do texto da label que informa quantos caracteres faltam.
    ''' </summary>
    <Category("Aparência")>
    <DefaultValue(GetType(Color), "ControlText")>
    <Description("Define a cor do texto da label que informa quantos caracteres faltam.")>
    Public Property LabelForeColor As Color
        Get
            Return _LabelForeColor
        End Get
        Set(value As Color)
            If value <> Color.Transparent Then
                _LabelForeColor = value
            Else
                Throw New ArgumentException("O controle não dá suporte a cores de texto transparente.")
            End If
        End Set
    End Property

    ''' <summary>
    ''' Define o nome do campo da tabela a ser pesquisado.
    ''' </summary>
    <Category("Query")>
    <DefaultValue(GetType(Integer), "1000")>
    <Description("Define o máximo de resultados que podem ser retornados pela pesquisa.")>
    Public Property Limit As Integer = 1000

    ''' <summary>
    ''' Define o nome do campo atribuido como Primary Key na tabela principal.
    ''' </summary>
    <Category("Query")>
    <Description("Define o nome do campo atribuido como Primary Key na tabela principal.")>
    <MergableProperty(False)>
    Public Property MainReturnFieldName As String

    ''' <summary>
    ''' Define o apelido da tabela principal.
    ''' </summary>
    <Category("Query")>
    <Description("Define o apelido da tabela principal.")>
    <MergableProperty(False)>
    Public Property MainTableAlias As String

    ''' <summary>
    ''' Define o nome da tabela principal.
    ''' </summary>
    <Category("Query")>
    <Description("Define o nome da tabela principal.")>
    <MergableProperty(False)>
    Public Property MainTableName As String

    Public Overrides Property Multiline As Boolean
        Get
            Return MyBase.Multiline
        End Get
        Set(value As Boolean)
            If value Then
                QueryEnabled = False
            End If
            MyBase.Multiline = value
        End Set
    End Property
    ''' <summary>
    ''' Define os outros campos que serão mostrados nos resultados da pesquisa.
    ''' </summary>
    <Category("Query")>
    <Description("Define os outros campos que serão mostrados nos resultados da pesquisa.")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <MergableProperty(False)>
    Public Property OtherFields As New Collection(Of OtherField)
    ''' <summary>
    ''' Define os parâmetos utilizados nas condições da Query.
    ''' </summary>
    <Category("Query")>
    <Description("Define os parâmetos utilizados nas condições da Query.")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <MergableProperty(False)>
    Public Property Parameters As New Collection(Of Parameter)

    ''' <summary>
    ''' Define um prefixo que será congelado com o resultado.
    ''' </summary>
    <Category("Comportamento")>
    <DefaultValue(GetType(String), Nothing)>
    <Description("Define um prefixo que será congelado com o resultado.")>
    Public Property Prefix As String

    ''' <summary>
    ''' Define se as pesquisas estão habilitadas.
    ''' </summary>
    <Category("Comportamento")>
    <DefaultValue(GetType(Boolean), "True")>
    <Description("Define se as pesquisas estão habilitadas.")>
    Public Property QueryEnabled As Boolean
        Get
            Return _QueryEnabled
        End Get
        Set(value As Boolean)
            If value Then Multiline = False
            _QueryEnabled = value
        End Set
    End Property

    ''' <summary>
    ''' Define o intervalo em que será feita uma nova pesquisa entre cada caractere digitado.
    ''' </summary>
    <Category("Comportamento")>
    <DefaultValue(GetType(Integer), "300")>
    <Description("Define o intervalo em que será feita uma nova pesquisa entre cada caractere digitado.")>
    Public Property QueryInterval As Integer
        Get
            Return _QueryInterval
        End Get
        Set(value As Integer)
            If value < 0 Then value = 0
            _QueryInterval = value
        End Set
    End Property

    ''' <summary>
    ''' Define o nome do campo da tabela a ser pesquisado.
    ''' </summary>
    <Category("Query")>
    <Description("Define condições para a pesquisa. Deve ser definida com a sintaxe SQL.")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <MergableProperty(False)>
    Public Property Relations As New Collection(Of Relation)
    ''' <summary>
    ''' Define se será mostrado o inicio do conteúdo ao congelar um resultado.
    ''' </summary>
    <Category("Comportamento")>
    <DefaultValue(GetType(Boolean), "False")>
    <Description("Define se será mostrado o inicio do conteúdo ao congelar um resultado.")>
    Public Property ShowStartOnFreeze As Boolean = False

    ''' <summary>
    ''' Define se será mostrado o inicio do conteúdo ao sair do controle.
    ''' </summary>
    <Category("Comportamento")>
    <DefaultValue(GetType(Boolean), "True")>
    <Description("Define se será mostrado o inicio do conteúdo ao sair do controle.")>
    Public Property ShowStartOnLeave As Boolean = True

    ''' <summary>
    ''' Define se o Grid mostrará linhas verticais no DropDown.
    ''' </summary>
    ''' <returns></returns>
    <Category("Aparência")>
    <DefaultValue(GetType(Boolean), "True")>
    <Description("Define se o Grid mostrará linhas verticais no DropDown.")>
    Public Property ShowVerticalGridLines As Boolean = True
    ''' <summary>
    ''' Define um sufixo que será congelado com o resultado.
    ''' </summary>
    <Category("Comportamento")>
    <DefaultValue(GetType(String), Nothing)>
    <Description("Define um sufixo que será congelado com o resultado.")>
    Public Property Suffix As String
    ''' <summary>
    ''' Retorna a cor padrão de quando um registro é deselecionado (Mesma cor de fundo do controle).
    ''' </summary>
    ''' <returns></returns>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Always)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property UnFreezeColor As Color
        Get
            Return _UnFreezeColor
        End Get
    End Property

    Public Sub DebugQuery()
        Debug.Print(GetQuery())
    End Sub

    ''' <summary>
    ''' Retorna se o dropdown de resultados está visível ou não.
    ''' </summary>
    ''' <returns>Se o dropdown de resultados está visivel ou não</returns>
    Public Function DropDownVisible() As Boolean
        If DropDownResultsForm Is Nothing Then
            Return False
        Else
            If DropDownResultsForm.Visible Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    ''' <summary>
    ''' Congela manualmente um registro usando as propriedades definidas no controle.
    ''' </summary>
    ''' <param name="ID">ID do regristro a ser selecionado no banco de dados.</param>
    Public Sub Freeze(ID As Long)
        Dim OldPrimaryKey As Long = _FreezedPrimaryKey
        Dim Query As String
        Dim TableResults As DataTable
        Dim MainValue As String
        Dim OtherValue As String
        Dim FullValue As String = String.Empty
        Dim ParameterList As Dictionary(Of String, Object)
        If QueryEnabled Then
            _Freezing = True
            Query = GetQuery()
            Query += String.Format("{0}WHERE{0}{1}{2}.{3} = {4}",
                                    vbNewLine,
                                    vbTab,
                                    If(MainTableAlias = Nothing, MainTableName, MainTableAlias),
                                    MainReturnFieldName,
                                    ID
                               )
            ParameterList = New Dictionary(Of String, Object)
            For Each p As Parameter In Parameters
                ParameterList.Add(p.ParameterName, p.ParameterValue)
            Next p
            TableResults = ExecuteQuery(Query, ParameterList)
            If TableResults.Rows.Count = 1 Then
                QueryEnabled = False
                MainValue = TableResults.Rows(0).Item(If(DisplayFieldAlias = Nothing, DisplayFieldName, DisplayFieldAlias)).ToString
                _RawFreezedValues.Add((DisplayTableName, DisplayFieldName, MainValue))
                If Not String.IsNullOrEmpty(MainValue) Then
                    FullValue &= Prefix & MainValue & Suffix
                End If
                For Each o As OtherField In OtherFields
                    If o.Freeze Then
                        OtherValue = TableResults.Rows(0).Item(If(o.DisplayFieldAlias = Nothing, o.DisplayFieldName, o.DisplayFieldAlias)).ToString
                        _RawFreezedValues.Add((o.DisplayTableName, o.DisplayFieldName, OtherValue))
                        If Not String.IsNullOrEmpty(OtherValue) Then
                            FullValue &= o.Prefix & OtherValue & o.Suffix
                        End If
                    End If
                Next o
                If OldPrimaryKey <> ID Then
                    RaiseEvent FreezedPrimaryKeyChanging(Me, EventArgs.Empty)
                End If
                Text = FullValue
                ForeColor = FreezeColor
                _FreezedValue = FullValue
                _FreezedPrimaryKey = ID
                _IsFreezed = True
                If _FreezedPrimaryKey <> OldPrimaryKey Then
                    RaiseEvent FreezedPrimaryKeyChanged(Me, EventArgs.Empty)
                End If
                If ShowStartOnFreeze Then
                    Me.Select(0, 0)
                Else
                    Me.Select(Me.TextLength, 0)
                End If
                _CtrlHyperLink = True
                QueryEnabled = True
                CloseDropDown()
                _Freezing = False
            Else
                Unfreeze()
            End If
        End If
    End Sub

    ''' <summary>
    ''' Seleciona manualmente um registro espcificando ignorando as propredades definidas no controle.
    ''' </summary>
    ''' <param name="Table">Nome da tabela do banco de dados.</param>
    ''' <param name="Field">Nome do campo do banco de dados.</param>
    ''' <param name="ID">ID do regristro a ser selecionado no banco de dados.</param>
    Public Sub Freeze(Table As String, Field As String, ID As Long)
        Dim OldPrimaryKey As Long = _FreezedPrimaryKey
        Dim Query As String
        If QueryEnabled Then
            _Freezing = True
            Query = String.Format("SELECT {0} FROM {1} WHERE ID = @ID", Field, Table, ID)
            Dim TableResults As DataTable = ExecuteQuery(Query, New Dictionary(Of String, Object) From {{"@ID", ID}})
            If TableResults.Rows.Count = 1 Then
                QueryEnabled = False
                If ShowStartOnFreeze Then
                    Me.Select(0, 0)
                Else
                    Me.Select(Me.TextLength, 0)
                End If
                _CtrlHyperLink = True
                If OldPrimaryKey <> ID Then
                    RaiseEvent FreezedPrimaryKeyChanging(Me, EventArgs.Empty)
                End If
                Text = TableResults.Rows(0).Item(Field).ToString
                ForeColor = FreezeColor
                _IsFreezed = True
                _FreezedPrimaryKey = ID
                If _FreezedPrimaryKey <> OldPrimaryKey Then
                    RaiseEvent FreezedPrimaryKeyChanged(Me, EventArgs.Empty)
                End If
                _RawFreezedValues.Add((Table, Field, TableResults.Rows(0).Item(Field).ToString))
                QueryEnabled = True
                CloseDropDown()
                _Freezing = False
            Else
                Unfreeze()
            End If
        End If
    End Sub

    ''' <summary>
    ''' Retorna o valor congelado referente à tabela e campo informados.
    ''' </summary>
    ''' <param name="TableName">O nome da tabela.</param>
    ''' <param name="FieldName">O nome do campo.</param>
    ''' <returns>O valor congelado referente ao campo informado.</returns>
    Public Function GetRawFreezedValueOf(TableName As String, FieldName As String) As Object
        Dim Match = _RawFreezedValues.Find(Function(t) t.Item1 = TableName AndAlso t.Item2 = FieldName)
        If String.IsNullOrEmpty(Match.Item1) Then
            Throw New Exception($"A tabela '{TableName}' não foi configurada no controle.")
        End If
        If String.IsNullOrEmpty(Match.Item2) Then
            Throw New Exception($"O campo '{FieldName}' da tabela '{TableName}' não foi configurado no controle.")
        End If
        Return Match.Item3
    End Function

    ''' <summary>
    ''' Descongela manualmente um registro.
    ''' </summary>
    Public Sub Unfreeze()
        Dim OldPrimaryKey As Long
        If QueryEnabled Then
            QueryEnabled = False
            OldPrimaryKey = _FreezedPrimaryKey
            If OldPrimaryKey <> 0 Then
                RaiseEvent FreezedPrimaryKeyChanging(Me, EventArgs.Empty)
            End If
            Text = Nothing
            ForeColor = UnFreezeColor
            _FreezedValue = Nothing
            _FreezedPrimaryKey = 0
            _IsFreezed = False
            If _FreezedPrimaryKey <> OldPrimaryKey Then
                RaiseEvent FreezedPrimaryKeyChanged(Me, EventArgs.Empty)
            End If
            _RawFreezedValues = New List(Of (String, String, Object))
            QueryEnabled = True
            CloseDropDown()
        End If
    End Sub

    Protected Overrides Sub OnEnter(e As EventArgs)
        MyBase.OnEnter(e)
        SelectionStart = Text.Length
        If Not _FirstEnter Then
            If Not DesignMode AndAlso Parent IsNot Nothing Then
                AddHandler Parent.FindForm.Deactivate, AddressOf Form_Deactivate
                Timer = New Timer With {
                    .Interval = QueryInterval
                }
            End If
            _FirstEnter = True
        End If
    End Sub

    Protected Overrides Sub OnForeColorChanged(e As EventArgs)
        MyBase.OnForeColorChanged(e)
        If QueryEnabled Then
            _UnFreezeColor = ForeColor
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
                    designer.ActionLists.Add(New QueriedTextBoxControlDesignerActionList(designer, actions))
                End If
            End If
        End If
    End Sub
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        MyBase.OnKeyDown(e)
        _KeyDown = True
        If QueryEnabled Then
            Select Case e.KeyCode
                Case Is = Keys.Enter, Keys.Escape
                    e.SuppressKeyPress = True
                Case Is = Keys.Down
                    e.Handled = True
                Case Is = Keys.Up
                    e.Handled = True
            End Select
        End If
    End Sub

    Protected Overrides Sub OnKeyUp(e As KeyEventArgs)
        MyBase.OnKeyUp(e)
        If QueryEnabled Then
            FormatTextBox(False)
        End If
    End Sub

    Protected Overrides Sub OnLeave(e As EventArgs)
        MyBase.OnLeave(e)
        If QueryEnabled Then
            If _IsHyperLink Then FormatTextBox(False)
            If DropDownResultsForm IsNot Nothing AndAlso DropDownResultsForm.DgvResults.SelectedRows.Count = 1 Then
                If UCase(Text) = UCase(DropDownResultsForm.DgvResults.SelectedRows(0).Cells(If(DisplayFieldAlias = Nothing, DisplayFieldName, DisplayFieldAlias)).Value.ToString) Then
                    AutoFreeze()
                End If
            End If
            CloseDropDown()
        End If
        If ShowStartOnLeave Then
            SelectionStart = 0
        End If
    End Sub

    Protected Overrides Sub OnMouseClick(e As MouseEventArgs)
        MyBase.OnMouseClick(e)
        If QueryEnabled Then
            If _IsHyperLink AndAlso e.Button = MouseButtons.Left Then
                RaiseEvent HyperlinkClicked(Me, EventArgs.Empty)
                FormatTextBox(False)
            End If
        End If
    End Sub

    Protected Overrides Sub OnPreviewKeyDown(e As PreviewKeyDownEventArgs)
        MyBase.OnPreviewKeyDown(e)
        Dim Row As Long
        If QueryEnabled Then
            If AllowHyperlink Then FormatTextBox(e.Control)
            If DropDownResultsForm IsNot Nothing Then
                If DropDownResultsForm.DgvResults.SelectedRows.Count = 1 Then
                    Select Case e.KeyCode
                        Case Is = Keys.Tab
                            If DropDownResultsForm IsNot Nothing AndAlso DropDownResultsForm.DgvResults.SelectedRows.Count = 1 Then
                                If UCase(Text) = UCase(DropDownResultsForm.DgvResults.SelectedRows(0).Cells(If(DisplayFieldAlias = Nothing, DisplayFieldName, DisplayFieldAlias)).Value.ToString) Then
                                    AutoFreeze()
                                End If
                            End If
                            CloseDropDown()
                        Case Is = Keys.Enter
                            AutoFreeze()
                            Me.Select(TextLength, 0)
                            CloseDropDown()
                        Case Is = Keys.Down
                            Row = DropDownResultsForm.DgvResults.SelectedRows(0).Index
                            If DropDownResultsForm IsNot Nothing AndAlso DropDownResultsForm.DgvResults.Rows.Count > Row + 1 Then
                                DropDownResultsForm.DgvResults.Rows(Row + 1).Selected = True
                                If DropDownResultsForm.DgvResults.SelectedRows(0).Index = DropDownResultsForm.DgvResults.Rows.Count - 1 Then
                                    DropDownResultsForm.DgvResults.FirstDisplayedScrollingRowIndex = DropDownResultsForm.DgvResults.SelectedRows(0).Index
                                    Row += 1
                                Else
                                    Row += 2
                                End If
                                If DropDownResultsForm.DgvResults.Rows(Row).Displayed = False Then
                                    If Row >= 3 Then
                                        DropDownResultsForm.DgvResults.FirstDisplayedScrollingRowIndex = DropDownResultsForm.DgvResults.SelectedRows(0).Index - 2
                                    End If
                                End If
                            End If
                        Case Is = Keys.Up
                            Row = DropDownResultsForm.DgvResults.SelectedRows(0).Index
                            If DropDownResultsForm.DgvResults.Visible = True And Row > 0 Then
                                DropDownResultsForm.DgvResults.Rows(Row - 1).Selected = True
                                If DropDownResultsForm.DgvResults.Rows(Row - 1).Displayed = False Then
                                    DropDownResultsForm.DgvResults.FirstDisplayedScrollingRowIndex = DropDownResultsForm.DgvResults.SelectedRows(0).Index
                                End If
                            End If
                    End Select
                End If
                If e.KeyCode = Keys.Escape Then CloseDropDown()
            End If
        End If
    End Sub

    <DebuggerStepThrough>
    Protected Overrides Sub OnTextChanged(e As EventArgs)
        MyBase.OnTextChanged(e)
        Dim Chars As Integer
        Dim CharsDif As Integer
        If Not DesignMode AndAlso _KeyDown Then
            If QueryEnabled Then
                If FreezedValue <> Nothing And Text <> FreezedValue Then
                    AutoUnfreeze()
                End If
                If Parent IsNot Nothing Then
                    If DropDownResultsForm Is Nothing Then
                        DropDownResultsForm = New FormDropDownResults(Me)
                        CType(DropDownResultsForm, FormDropDownResults).Textbox = Me
                        DropDownResultsForm.Location = Me.Parent.PointToScreen(New Point(Me.Left - DropDownStretchLeft, Me.Bottom))
                        DropDownResultsForm.Width = DropDownStretchLeft + Me.Width + If(DropDownAutoStretchRight, 0, DropDownStretchRight)
                        DropDownResultsForm.Height = DropDownStretchDown
                        AddHandler DropDownResultsForm.FormClosed, AddressOf DropDownResultsForm_FormClosed
                        DropDownResultsForm.Show()
                        DropDownVisible()
                    End If
                    Chars = Text.Replace("%", Nothing).Count
                    CharsDif = CharactersToQuery - Chars
                    If Chars = 0 Then
                        CloseDropDown()
                    Else
                        If Chars < CharactersToQuery Then
                            DropDownResultsForm.DgvResults.DataSource = Nothing
                            DropDownResultsForm.LblCharsRemaining.Visible = True
                            DropDownResultsForm.DgvResults.Visible = False
                            If CharsDif > 1 Then
                                DropDownResultsForm.LblCharsRemaining.Text = String.Format("Digite mais {0} caracteres para consultar.", CharactersToQuery - Chars)
                            ElseIf CharsDif = 1 Then
                                DropDownResultsForm.LblCharsRemaining.Text = String.Format("Digite mais {0} caractere para consultar.", CharactersToQuery - Chars)
                            End If
                        ElseIf Chars >= CharactersToQuery Then
                            Timer.Stop() : Timer.Start()
                            DropDownResultsForm.LblCharsRemaining.Visible = False
                            DropDownResultsForm.DgvResults.Visible = True
                        End If
                    End If
                End If
            End If
        End If
        _KeyDown = False
    End Sub
    Private Sub AutoFreeze()
        Dim OldPrimaryKey As Long
        Dim MainValue As String
        Dim OtherValue As String
        Dim FullValue As String = String.Empty
        If QueryEnabled Then
            _Freezing = True
            OldPrimaryKey = _FreezedPrimaryKey
            If DropDownResultsForm IsNot Nothing AndAlso DropDownResultsForm.DgvResults.SelectedRows.Count = 1 Then
                QueryEnabled = False
                MainValue = DropDownResultsForm.DgvResults.SelectedRows(0).Cells(If(DisplayFieldAlias = Nothing, DisplayFieldName, DisplayFieldAlias)).Value.ToString
                _RawFreezedValues.Add((DisplayTableName, DisplayFieldName, MainValue))
                If Not String.IsNullOrEmpty(MainValue) Then
                    FullValue &= Prefix & MainValue & Suffix
                End If
                For Each o As OtherField In OtherFields
                    If o.Freeze Then
                        OtherValue = DropDownResultsForm.DgvResults.SelectedRows(0).Cells(If(o.DisplayFieldAlias = Nothing, o.DisplayFieldName, o.DisplayFieldAlias)).Value.ToString
                        _RawFreezedValues.Add((o.DisplayTableName, o.DisplayFieldName, OtherValue))
                        If Not String.IsNullOrEmpty(OtherValue) Then
                            FullValue &= o.Prefix & OtherValue & o.Suffix
                        End If
                    End If
                Next o
                If OldPrimaryKey <> DropDownResultsForm.DgvResults.SelectedRows(0).Cells("mainid").Value Then
                    RaiseEvent FreezedPrimaryKeyChanging(Me, EventArgs.Empty)
                End If
                Text = FullValue
                ForeColor = FreezeColor
                _FreezedValue = FullValue
                If Not IsNumeric(DropDownResultsForm.DgvResults.SelectedRows(0).Cells("mainid").Value) Then
                    Throw New Exception("Não houve retorno de chave primária, verifique as relações.")
                End If
                _FreezedPrimaryKey = DropDownResultsForm.DgvResults.SelectedRows(0).Cells("mainid").Value
                _IsFreezed = True
                If _FreezedPrimaryKey <> OldPrimaryKey Then
                    RaiseEvent FreezedPrimaryKeyChanged(Me, EventArgs.Empty)
                End If
                If ShowStartOnFreeze Then
                    Me.Select(0, 0)
                Else
                    Me.Select(Me.TextLength, 0)
                End If
                _CtrlHyperLink = True
                QueryEnabled = True
                CloseDropDown()
            End If
            _Freezing = False
        End If
    End Sub

    Private Sub AutoUnfreeze()
        Dim OldPrimaryKey As Long
        If QueryEnabled Then
            OldPrimaryKey = _FreezedPrimaryKey
            QueryEnabled = False
            If ClearOnUnfreeze Then
                Text = Nothing
            End If
            If OldPrimaryKey <> 0 Then
                RaiseEvent FreezedPrimaryKeyChanging(Me, EventArgs.Empty)
            End If
            ForeColor = UnFreezeColor
            _FreezedValue = Nothing
            _IsFreezed = False
            _FreezedPrimaryKey = 0
            If _FreezedPrimaryKey <> OldPrimaryKey Then
                RaiseEvent FreezedPrimaryKeyChanged(Me, EventArgs.Empty)
            End If
            _RawFreezedValues = New List(Of (String, String, Object))
            _CtrlHyperLink = False
            QueryEnabled = True
        End If
    End Sub

    <DebuggerStepThrough>
    Private Sub CloseDropDown()
        If DropDownResultsForm IsNot Nothing Then
            DropDownResultsForm.Close()
            DropDownResultsForm = Nothing
        End If
    End Sub

    Private Sub DropDownResultsForm_FormClosed(ByVal sender As Object, ByVal e As FormClosedEventArgs)
        DropDownResultsForm.Dispose()
        DropDownResultsForm = Nothing
    End Sub

    Private Function ExecuteQuery(Query As String, Optional Parameters As Dictionary(Of String, Object) = Nothing) As DataTable
        Dim Table As New DataTable
        Dim Par As DbParameter
        Dim Factory As DbProviderFactory = DbProviderFactories.GetFactory(Connection)
        Using Cmd As IDbCommand = Connection.CreateCommand
            Cmd.CommandText = Query
            If Parameters IsNot Nothing Then
                For Each P In Parameters
                    Par = Cmd.CreateParameter
                    Par.ParameterName = P.Key
                    Par.Value = P.Value
                    Cmd.Parameters.Add(Par)
                Next P
            End If
            If DebugOnTextChanged Then DatabaseHelper.DebugQuery(Cmd)
            Using Adp As DbDataAdapter = Factory.CreateDataAdapter()
                Adp.SelectCommand = Cmd
                Connection.Open()
                Try
                    Adp.Fill(Table)
                Catch ex As Exception
                    If DropDownResultsForm IsNot Nothing Then
                        DropDownResultsForm.Close()
                        DropDownResultsForm = Nothing
                    End If
                    Throw ex
                Finally
                    Connection.Close()
                End Try
                Return Table
            End Using
        End Using
    End Function

    <DebuggerStepThrough>
    Private Sub Form_Deactivate(sender As Object, e As EventArgs)
        If DropDownResultsForm IsNot Nothing AndAlso DropDownResultsForm.DgvResults.SelectedRows.Count = 1 Then
            If UCase(Text) = UCase(DropDownResultsForm.DgvResults.SelectedRows(0).Cells(If(DisplayFieldAlias = Nothing, DisplayFieldName, DisplayFieldAlias)).Value.ToString) Then
                AutoFreeze()
            End If
        End If
        CloseDropDown()
    End Sub

    Private Sub FormatTextBox(ByVal ShowAsLink As Boolean)
        If Not _CtrlHyperLink Then Return
        If ShowAsLink Then
            Font = New Font(Font, FontStyle.Underline)
            Cursor = Cursors.Hand
            _IsHyperLink = True
        Else
            Font = New Font(Font, FontStyle.Regular)
            Cursor = Cursors.IBeam
            _IsHyperLink = False
        End If
    End Sub
    Private Function GetQuery() As String
        Dim Query As String
        Query = String.Format("SELECT {0}{1}{2}{3}.{4} AS 'mainid',{1}{2}{8}{5}.{6}{9} AS '{7}'",
                                  If(Distinct, "DISTINCT", Nothing), '0
                                  vbNewLine, '1
                                  vbTab, '2
                                  If(MainTableAlias = Nothing, MainTableName, MainTableAlias), '3
                                  MainReturnFieldName, '4
                                  If(DisplayTableAlias = Nothing, DisplayTableName, DisplayTableAlias), '5
                                  DisplayFieldName, '6
                                  If(DisplayFieldAlias = Nothing, DisplayFieldName, DisplayFieldAlias), '7
                                  If(IfNull <> Nothing, "IFNULL(", Nothing), '8
                                  If(IfNull <> Nothing, ", " & IfNull & ") ", Nothing) '9
                              )
        For Each o As OtherField In OtherFields
            Query += String.Format(",{0}{1}{5}{2}.{3}{6} AS '{4}'",
                                       vbNewLine, '0
                                       vbTab, '1
                                       If(o.DisplayTableAlias = Nothing, o.DisplayTableName, o.DisplayTableAlias), '2
                                       o.DisplayFieldName, '3
                                       If(o.DisplayFieldAlias = Nothing, o.DisplayFieldName, o.DisplayFieldAlias), '4
                                       If(o.IfNull <> Nothing, "IFNULL(", Nothing),'5
                                       If(o.IfNull <> Nothing, ", " & o.IfNull & ") ", Nothing)
                                   )
        Next o
        Query += String.Format("{0}FROM {1} AS {2}",
                               vbNewLine, '0
                               MainTableName, '1
                               If(MainTableAlias = Nothing, MainTableName, MainTableAlias)) '2)
        For Each r As Relation In Relations
            Query += String.Format("{0}{1} JOIN {2} AS {3} ON {3}.{4} {5} {6}.{7}",
                                       vbNewLine, '0
                                       r.RelationType, '1
                                       r.RelateTableName, '2
                                       If(r.RelateTableAlias = Nothing, r.RelateTableName, r.RelateTableAlias), '3
                                       r.RelateFieldName, '4
                                       r.Operator, '5
                                       If(r.WithTableAlias = Nothing, r.WithTableName, r.WithTableAlias), '6
                                       r.WithFieldName '7
                                   )
            For Each c As Condition In r.Conditions
                Query += String.Format(" AND{0}{1}{2}.{3} {4} {5}",
                                           vbNewLine,
                                           vbTab,
                                           c.TableNameOrAlias,
                                           c.FieldName,
                                           c.Operator,
                                           c.Value
                                       )
            Next c
        Next r
        Return Query
    End Function

    Private Sub Timer_Tick(sender As Object, e As EventArgs) Handles Timer.Tick
        Dim Query As String
        Dim ParameterList As Dictionary(Of String, Object)
        Dim TableResult As DataTable
        Dim ValueParameter As String
        Timer.Interval = QueryInterval
        Timer.Stop()
        ValidateTick()
        Query = GetQuery()
        ValueParameter = "@" & Guid.NewGuid.ToString("N")
        Query += String.Format("{0}WHERE{0}({0}{1}{2}.{3} LIKE {4}",
                                    vbNewLine,
                                    vbTab,
                                    If(DisplayTableAlias = Nothing, DisplayTableName, DisplayTableAlias),
                                    DisplayFieldName,
                                    ValueParameter
                               )
        For Each o As OtherField In OtherFields
            Query += String.Format(" OR{0}{1}{2}.{3} LIKE {4}",
                                   vbNewLine,
                                   vbTab,
                                   If(o.DisplayTableAlias = Nothing, o.DisplayTableName, o.DisplayTableAlias),
                                   o.DisplayFieldName,
                                   ValueParameter
                               )
        Next o
        Query += String.Format("{0})", vbNewLine)
        If Conditions.Count > 0 Then
            Query += String.Format(" AND{0}(", vbNewLine)
            For Each c As Condition In Conditions
                Query += String.Format("{0}{1}{2}.{3} {4} {5} AND",
                                       vbNewLine,
                                       vbTab,
                                       c.TableNameOrAlias,
                                       c.FieldName,
                                       c.Operator,
                                       c.Value
                                   )
            Next c
            Query = Strings.Left(Query, Query.Length - 4)
            Query += String.Format("{0})", vbNewLine)
        End If
        If Limit > 0 Then
            Query += $"LIMIT {Limit};"
        Else
            Query += ";"
        End If
        ParameterList = New Dictionary(Of String, Object) From {
            {ValueParameter, "%" & Text & "%"}
        }
        For Each p As Parameter In Parameters
            ParameterList.Add(p.ParameterName, p.ParameterValue)
        Next p
        Try
            TableResult = New DataTable
            TableResult = ExecuteQuery(Query, ParameterList)
            If DropDownResultsForm IsNot Nothing Then
                DropDownResultsForm.DgvResults.DataSource = TableResult
                DropDownResultsForm.DgvResults.Columns("mainid").Visible = False
                DropDownResultsForm.DgvResults.Columns.Cast(Of DataGridViewColumn).First(Function(c) c.Name = If(String.IsNullOrEmpty(DisplayFieldAlias), DisplayFieldName, DisplayFieldAlias)).AutoSizeMode = DisplayFieldAutoSizeColumnMode
                For Each OtherField In OtherFields
                    DropDownResultsForm.DgvResults.Columns.Cast(Of DataGridViewColumn).First(Function(c) c.Name = If(String.IsNullOrEmpty(OtherField.DisplayFieldAlias), OtherField.DisplayFieldName, OtherField.DisplayFieldAlias)).AutoSizeMode = OtherField.DisplayFieldAutoSizeColumnMode
                Next OtherField
                If DropDownAutoStretchRight Then
                    For Each c In DropDownResultsForm.DgvResults.Controls
                        If c.GetType() Is GetType(HScrollBar) Then
                            Dim vbar As HScrollBar = DirectCast(c, HScrollBar)
                            If vbar.Visible = True AndAlso DropDownResultsForm.DgvResults.Rows.Count > 0 Then
                                Do Until vbar.Visible = False
                                    DropDownResultsForm.Width += 10
                                Loop
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            CloseDropDown()
            Throw ex
        End Try
    End Sub
    Private Sub ValidateTick()
        Dim FieldHeaders As New List(Of String)
        Dim ParametersNames As New List(Of String)
        If Connection Is Nothing Then
            DropDownResultsForm.Close() : DropDownResultsForm = Nothing
            Throw New Exception("A propriedade Connection não foi definida.")
        End If
        If MainTableName = Nothing Then
            DropDownResultsForm.Close() : DropDownResultsForm = Nothing
            Throw New Exception("A propriedade MainTableName não foi definida.")
        End If
        If MainReturnFieldName = Nothing Then
            DropDownResultsForm.Close() : DropDownResultsForm = Nothing
            Throw New Exception("A propriedade MainReturnFieldName não foi definida.")
        End If
        If DisplayTableName = Nothing Then
            DropDownResultsForm.Close() : DropDownResultsForm = Nothing
            Throw New Exception("A propriedade DisplayTableName não foi definida.")
        End If
        If DisplayFieldName = Nothing Then
            DropDownResultsForm.Close() : DropDownResultsForm = Nothing
            Throw New Exception("A propriedade DisplayFieldName não foi definida.")
        End If
        If DisplayMainFieldName = Nothing Then
            DropDownResultsForm.Close() : DropDownResultsForm = Nothing
            Throw New Exception("A propriedade DisplayMainFieldName  não foi definida.")
        End If
        If DisplayFieldAlias <> Nothing Then
            FieldHeaders.Add(DisplayFieldAlias)
        End If
        If FieldHeaders.Count <> FieldHeaders.Distinct.Count Then
            DropDownResultsForm.Close() : DropDownResultsForm = Nothing
            Throw New Exception("Existe mais de um controle com o mesmo valor para a propriedade FieldHeader.")
        End If
        For i = 0 To OtherFields.Count - 1
            If OtherFields(i).DisplayTableName = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe OtherField com a propriedade TableName não definida.")
            End If
            If OtherFields(i).DisplayFieldName = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe OtherField com a propriedade FieldName não definida.")
            End If
        Next i
        For i = 0 To Parameters.Count - 1
            If Parameters(i).ParameterName = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                ParametersNames.Clear()
                Throw New Exception("Existe Parameter com a propriedade ParameterName não definida.")
            Else
                ParametersNames.Add(Parameters(i).ParameterName)
            End If
        Next
        If ParametersNames.Count <> ParametersNames.Distinct.Count Then
            DropDownResultsForm.Close() : DropDownResultsForm = Nothing
            Throw New Exception("Existe Parameter com a propriedade ParameterName duplicada.")
            Exit Sub
        End If
        For i = 0 To Conditions.Count - 1
            If Conditions(i).TableNameOrAlias = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe Condition com a propriedade TableName não definida.")
                Exit Sub
            End If
            If Conditions(i).FieldName = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe Condition com a propriedade FieldName não definida.")
                Exit Sub
            End If
            If Conditions(i).Operator = "BETWEEN" And Conditions(i).Value <> Nothing AndAlso Conditions(i).Value.Split(";").Count <> 2 Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe Condition com a propriedade Value não definida para o operador BETWEEN.")
                Exit Sub
            End If
        Next i
        For i = 0 To Relations.Count - 1
            If Relations(i).RelationType = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe Relation com a propriedade RelationType não definida.")
                Exit Sub
            End If
            If Relations(i).Operator = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe Relation com a propriedade Operator não definida.")
                Exit Sub
            End If
            If Relations(i).RelateTableName = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe Relation com a propriedade TableNameX não definida.")
                Exit Sub
            End If
            If Relations(i).WithTableName = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe Relation com a propriedade TableNameY não definida.")
                Exit Sub
            End If
            If Relations(i).RelateFieldName = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe Relation com a propriedade FieldNameX não definida.")
                Exit Sub
            End If
            If Relations(i).WithFieldName = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe Relation com a propriedade FieldNameY não definida.")
                Exit Sub
            End If
        Next i
    End Sub
    Public Class Condition

        ''' <summary>
        ''' Define o operador da condição. Para o operador BETWEEN, separar os dois valores por ponto e vírgula (;).
        ''' </summary>
        <Description("Define o operador da condição. Para o operador BETWEEN, separar os dois valores por ponto e vírgula (;).")>
        <Category("Query")>
        <TypeConverter(GetType(OperatorFilterCollection))>
        Public Property [Operator] As String

        ''' <summary>
        ''' Define o nome do campo do banco de dados onde será aplicada a condição.
        ''' </summary>
        <Description("Define o nome do campo do banco de dados onde será aplicada a condição.")>
        <Category("Query")>
        Public Property FieldName As String

        ''' <summary>
        ''' Define o nome ou apelido da tabela do banco de dados onde será aplicada a condição.
        ''' </summary>
        <Description("Define o nome ou apelido da tabela do banco de dados onde será aplicada a condição.")>
        <Category("Query")>
        Public Property TableNameOrAlias As String
        ''' <summary>
        ''' Define o valor a ser testado na condição.
        ''' </summary>
        <Description("Define o valor a ser testado na condição.")>
        <Category("Query")>
        Public Property Value As String

        Public Overrides Function ToString() As String
            If TableNameOrAlias <> Nothing And FieldName <> Nothing And [Operator] <> Nothing And Value <> Nothing Then
                If [Operator] = "BETWEEN" Then
                    If Value.Split(";").Count = 2 Then
                        Return String.Format("{0}.{1} {2} {3} AND {3}", TableNameOrAlias, FieldName, [Operator], Value.Split(";").ElementAt(0), Value.Split(";").ElementAt(1))
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

        Public Overrides Function GetStandardValues(ByVal context As ITypeDescriptorContext) As StandardValuesCollection
            Return New StandardValuesCollection(_OperatorList)
        End Function

        Public Overrides Function GetStandardValuesExclusive(ByVal context As ITypeDescriptorContext) As Boolean
            Return True
        End Function

        Public Overrides Function GetStandardValuesSupported(ByVal context As ITypeDescriptorContext) As Boolean
            Return True
        End Function
    End Class

    Public Class OtherField

        ''' <summary>
        ''' Define o apelido do campo que será que será exibido em conjunto nos resultados da pesquisa.
        ''' </summary>
        <Description("Define o apelido do campo que será que será exibido em conjunto nos resultados da pesquisa.")>
        <Category("Query")>
        Public Property DisplayFieldAlias As String

        ''' <summary>
        ''' Define como a largura da coluna será ajustada nos resultados.
        ''' </summary>
        <Category("Aparência")>
        <DefaultValue(GetType(Color), "Blue")>
        <Description("Define como a largura da coluna será ajustada nos resultados.")>
        Public Property DisplayFieldAutoSizeColumnMode As DataGridViewAutoSizeColumnMode

        ''' <summary>
        ''' Define o nome do campo que será que será exibido em conjunto nos resultados da pesquisa.
        ''' </summary>
        <Description("Define o nome do campo que será que será exibido em conjunto nos resultados da pesquisa.")>
        <Category("Query")>
        Public Property DisplayFieldName As String

        ''' <summary>
        ''' Define o nome do campo da tabela que está atribuído como Primary Key.
        ''' </summary>
        <Description("Define o nome do campo da tabela que está atribuído como Primary Key.")>
        <Category("Query")>
        Public Property DisplayMainFieldName As String

        ''' <summary>
        ''' Define o apelido da tabela a qual está relacionada aos resultados que serão exibidos.
        ''' </summary>
        <Description("Define o apelido da tabela a qual está relacionada aos resultados que serão exibidos.")>
        <Category("Query")>
        Public Property DisplayTableAlias As String

        ''' <summary>
        ''' Define o nome da tabela a qual está relacionada aos resultados que serão exibidos.
        ''' </summary>
        <Description("Define o nome da tabela a qual está relacionada aos resultados que serão exibidos.")>
        <Category("Query")>
        Public Property DisplayTableName As String
        ''' <summary>
        ''' Define se o valor desse campo será congelado no controle.
        ''' </summary>
        <Description("Define se o valor desse campo será congelado no controle.")>
        <Category("Comportamento")>
        Public Property Freeze As Boolean

        ''' <summary>
        ''' Define um valor a ser retornado caso a consulta retorne um valor nulo para esse campo.
        ''' </summary>
        <Description("Define um valor a ser retornado caso a consulta retorne um valor nulo para esse campo.")>
        <Category("Query")>
        Public Property IfNull As String

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

        Public Sub New()
        End Sub

        ''' <summary>
        ''' Define o operador a ser utilizado na relação.
        ''' </summary>
        <Description("Define o operador a ser utilizado na relação.")>
        <Category("Query")>
        <TypeConverter(GetType(OperatorFilterCollection))>
        Public Property [Operator] As String

        ''' <summary>
        ''' Define condições adicionais para o relacionamento.
        ''' </summary>
        <Description("Define condições adicionais para o relacionamento.")>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
        <Category("Query")>
        Public Property Conditions As New Collection(Of Condition)

        ''' <summary>
        ''' Define o nome do campo que se será relacionado.
        ''' </summary>
        <Description("Define o nome do campo que se será relacionado.")>
        <Category("Query")>
        Public Property RelateFieldName As String

        ''' <summary>
        ''' Define o apelido da tabela que será relacionada.
        ''' </summary>
        <Description("Define o apelido da tabela que será relacionada.")>
        <Category("Query")>
        Public Property RelateTableAlias As String

        ''' <summary>
        ''' Define o nome da tabela que será relacionada.
        ''' </summary>
        <Description("Define o nome da tabela que será relacionada.")>
        <Category("Query")>
        Public Property RelateTableName As String

        ''' <summary>
        ''' Define o tipo de relacionamento.
        ''' </summary>
        <Description("Define o tipo de relacionamento.")>
        <TypeConverter(GetType(RelationFilterCollection))>
        <Category("Query")>
        Public Property RelationType As String
        ''' <summary>
        ''' Define o nome do campo que será relacionado com o campo da tabela principal.
        ''' </summary>
        <Description("Define o nome do campo que será relacionado com o campo da tabela principal.")>
        <Category("Query")>
        Public Property WithFieldName As String

        ''' <summary>
        ''' Define o apelido da tabela que se será relacionada com a tabela principal.
        ''' </summary>
        <Description("Define o apelido da tabela que se será relacionada com a tabela principal.")>
        <Category("Query")>
        Public Property WithTableAlias As String

        ''' <summary>
        ''' Define o nome da tabela que será relacionada com a tabela principal.
        ''' </summary>
        <Description("Define o nome da tabela que será relacionada com a tabela principal.")>
        <Category("Query")>
        Public Property WithTableName As String
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

        Public Overrides Function GetStandardValues(ByVal context As ITypeDescriptorContext) As StandardValuesCollection
            Return New StandardValuesCollection(_JoinList)
        End Function

        Public Overrides Function GetStandardValuesExclusive(ByVal context As ITypeDescriptorContext) As Boolean
            Return True
        End Function

        Public Overrides Function GetStandardValuesSupported(ByVal context As ITypeDescriptorContext) As Boolean
            Return True
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

    Private Class FormDropDownResults
        Inherits Form
        Friend WithEvents DgvResults As DataGridView
        Friend WithEvents LblCharsRemaining As Label
        Friend WithEvents PanelContainer As Panel
        Public Textbox As Control
        Private ReadOnly _QueriedBox As QueriedBox

        Public Sub New(ByVal SearchBox As QueriedBox)
            SuspendLayout()
            _QueriedBox = SearchBox
            InitializeComponent()
            BackColor = _QueriedBox.DropDownBorderColor
            Font = SearchBox.Font
            FormBorderStyle = FormBorderStyle.None
            Padding = New Padding(1)
            Size = New Size(300, 120)
            DoubleBuffered = True
            TopMost = True
            KeyPreview = True
            ResumeLayout(True)
        End Sub

        Protected Overrides ReadOnly Property CreateParams As CreateParams
            Get
                Dim ret As CreateParams = MyBase.CreateParams
                ret.Style = CInt(Flags.WindowStyles.WS_SYSMENU) Or CInt(Flags.WindowStyles.WS_CHILD)
                ret.ExStyle = ret.ExStyle Or CInt(Flags.WindowStyles.WS_EX_NOACTIVATE) Or CInt(Flags.WindowStyles.WS_EX_TOOLWINDOW)
                ret.X = Me.Location.X
                ret.Y = Me.Location.Y
                Return ret
            End Get
        End Property

        Private Sub DataGridView_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles DgvResults.MouseDoubleClick
            Dim Click As DataGridView.HitTestInfo = DgvResults.HitTest(e.X, e.Y)
            If Click.Type = DataGridViewHitTestType.Cell Then
                _QueriedBox.AutoFreeze()
                _QueriedBox.Focus()
                Close()
            End If
        End Sub

        Private Sub DataGridView_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles DgvResults.PreviewKeyDown
            If e.KeyCode = Keys.Tab Then
                Close()
                Me.Select()
            End If
        End Sub

        Private Sub DgvResults_DataSourceChanged(sender As Object, e As EventArgs) Handles DgvResults.DataSourceChanged
            If DgvResults.Rows.Count = 0 Then
                DgvResults.Visible = False
                LblCharsRemaining.Text = "Nenhum Registro Encontrado."
                LblCharsRemaining.Visible = True
            End If
        End Sub

        Private Sub DropDownResults_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
            Application.AddMessageFilter(New PopupWindowHelperMessageFilter(Me, Textbox))
        End Sub

        Private Sub EnableDoubleBuffered(ByVal dgv As DataGridView)
            Dim DgvType As Type = dgv.[GetType]()
            Dim pi As PropertyInfo = DgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)
            pi.SetValue(dgv, True, Nothing)
        End Sub

        Private Sub InitializeComponent()
            Dim sg As Boolean = _QueriedBox.ShowVerticalGridLines
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
    End Class

    Private Class PopupWindowHelperMessageFilter
        Implements IMessageFilter
        Private Const WM_LBUTTONDOWN As Integer = &H201
        Private Const WM_MBUTTONDOWN As Integer = &H207
        Private Const WM_NCLBUTTONDOWN As Integer = &HA1
        Private Const WM_NCMBUTTONDOWN As Integer = &HA7
        Private Const WM_NCRBUTTONDOWN As Integer = &HA4
        Private Const WM_RBUTTONDOWN As Integer = &H204
        Private ReadOnly TextBox As Control = Nothing
        Public Sub New(ByVal popupW As Form, ByVal textbox As Control)
            Popup = popupW
            Me.TextBox = textbox
        End Sub

        Public Property Popup As Form = Nothing
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

        <DebuggerStepThrough>
        Private Sub OnMouseDown()
            Dim cursorPos As Point = Cursor.Position
            If TextBox.Parent Is Nothing Then Exit Sub
            If Not Popup.Bounds.Contains(cursorPos) Then
                If Not TextBox.Bounds.Contains(TextBox.Parent.PointToClient(cursorPos)) Then
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
    End Class

    Private Class QueriedTextBoxControlDesignerActionList
        Inherits DesignerActionList
        Private ReadOnly ActionList As DesignerActionList
        Private ReadOnly Control As QueriedBox
        Private ReadOnly Designer As ControlDesigner
        Public Sub New(ByVal Designer As ControlDesigner, ByVal ActionList As DesignerActionList)
            MyBase.New(Designer.Component)
            Me.Designer = Designer
            Me.ActionList = ActionList
            Control = CType(Designer.Control, QueriedBox)
        End Sub

        Public Property CharactersToQuery As Integer
            Get
                Return Control.CharactersToQuery
            End Get
            Set(ByVal value As Integer)
                TypeDescriptor.GetProperties(Component)("CharactersToQuery").SetValue(Component, value)
            End Set
        End Property

        Public Property DisplayFieldAlias As String
            Get
                Return Control.DisplayFieldAlias
            End Get
            Set(ByVal value As String)
                TypeDescriptor.GetProperties(Component)("DisplayFieldAlias").SetValue(Component, value)
            End Set
        End Property

        Public Property DisplayFieldName As String
            Get
                Return Control.DisplayFieldName
            End Get
            Set(ByVal value As String)
                TypeDescriptor.GetProperties(Component)("DisplayFieldName").SetValue(Component, value)
            End Set
        End Property

        Public Property DisplayMainFieldName As String
            Get
                Return Control.DisplayMainFieldName
            End Get
            Set(ByVal value As String)
                TypeDescriptor.GetProperties(Component)("DisplayMainFieldName").SetValue(Component, value)
            End Set
        End Property

        Public Property DisplayTableAlias As String
            Get
                Return Control.DisplayTableAlias
            End Get
            Set(ByVal value As String)
                TypeDescriptor.GetProperties(Component)("DisplayTableAlias").SetValue(Component, value)
            End Set
        End Property

        Public Property DisplayTableName As String
            Get
                Return Control.DisplayTableName
            End Get
            Set(ByVal value As String)
                TypeDescriptor.GetProperties(Component)("DisplayTableName").SetValue(Component, value)
            End Set
        End Property

        Public Property Limit As Integer
            Get
                Return Control.Limit
            End Get
            Set(ByVal value As Integer)
                TypeDescriptor.GetProperties(Component)("Limit").SetValue(Component, value)
            End Set
        End Property

        Public Property MainReturnFieldName As String
            Get
                Return Control.MainReturnFieldName
            End Get
            Set(ByVal value As String)
                TypeDescriptor.GetProperties(Component)("MainReturnFieldName").SetValue(Component, value)
            End Set
        End Property

        Public Property MainTableAlias As String
            Get
                Return Control.MainTableAlias
            End Get
            Set(ByVal value As String)
                TypeDescriptor.GetProperties(Component)("MainTableAlias").SetValue(Component, value)
            End Set
        End Property

        Public Property MainTableName As String
            Get
                Return Control.MainTableName
            End Get
            Set(ByVal value As String)
                TypeDescriptor.GetProperties(Component)("MainTableName").SetValue(Component, value)
            End Set
        End Property

        Public Property QueryEnabled As Boolean
            Get
                Return Control.QueryEnabled
            End Get
            Set(ByVal value As Boolean)
                TypeDescriptor.GetProperties(Component)("QueryEnabled").SetValue(Component, value)
            End Set
        End Property

        Public Property QueryInterval As Integer
            Get
                Return Control.QueryInterval
            End Get
            Set(ByVal value As Integer)
                TypeDescriptor.GetProperties(Component)("QueryInterval").SetValue(Component, value)
            End Set
        End Property

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
    End Class
End Class