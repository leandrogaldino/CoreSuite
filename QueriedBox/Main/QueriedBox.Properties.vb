Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Data.Common

Partial Public Class QueriedBox
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
    ''' Depura a consulta configurada no controle quando o texto é alterado.
    ''' </summary>
    <Category("Diversos")>
    <Description("Depura a consulta configurada no controle quando o texto é alterado.")>
    Public Property DebugOnTextChanged As Boolean
    ''' <summary>
    ''' Define se o resultado deve ser distinguido (Ocultar resultados idênticos provenientes da Query).
    ''' </summary>
    <Category("Query")>
    <Description("Define se o resultado deve ser distinguido (Ocultar resultados idênticos provenientes da Query).")>
    Public Property Distinct As Boolean
    ''' <summary>
    ''' Define um valor caso o resultado seja nulo.
    ''' </summary>
    <Category("Query")>
    <Description("Define um valor caso o resultado seja nulo.")>
    Public Property IfNull As String
    ''' <summary>
    ''' Define o nome da tabela principal.
    ''' </summary>
    <Category("Query")>
    <Description("Define o nome da tabela principal.")>
    <MergableProperty(False)>
    Public Property MainTableName As String
    ''' <summary>
    ''' Define o apelido da tabela principal.
    ''' </summary>
    <Category("Query")>
    <Description("Define o apelido da tabela principal.")>
    <MergableProperty(False)>
    Public Property MainTableAlias As String
    ''' <summary>
    ''' Define o nome do campo atribuido como Primary Key na tabela principal.
    ''' </summary>
    <Category("Query")>
    <Description("Define o nome do campo atribuido como Primary Key na tabela principal.")>
    <MergableProperty(False)>
    Public Property MainReturnFieldName As String
    ''' <summary>
    ''' Define o nome da tabela a qual está relacionada aos resultados que serão exibidos.
    ''' </summary>
    <Category("Query")>
    <Description("Define o nome da tabela a qual está relacionada aos resultados que serão exibidos.")>
    <MergableProperty(False)>
    Public Property DisplayTableName As String
    ''' <summary>
    ''' Define o apelido da tabela a qual está relacionada aos resultados que serão exibidos.
    ''' </summary>
    <Category("Query")>
    <Description("Define o apelido da tabela a qual está relacionada aos resultados que serão exibidos.")>
    <MergableProperty(False)>
    Public Property DisplayTableAlias As String
    ''' <summary>
    ''' Define o nome do campo que será congelado no controle quando escolhido pelo usuário.
    ''' </summary>
    <Category("Query")>
    <Description("Define o nome do campo que será congelado no controle quando escolhido pelo usuário.")>
    Public Property DisplayFieldName As String
    ''' <summary>
    ''' Define o apelido do campo que será congelado no controle quando escolhido pelo usuário.
    ''' </summary>
    <Category("Query")>
    <Description("Define o apelido do campo que será congelado no controle quando escolhido pelo usuário.")>
    Public Property DisplayFieldAlias As String
    ''' <summary>
    ''' Define o nome do campo da tabela que está atribuído como PRIMARYKEY.
    ''' </summary>
    <Category("Query")>
    <Description("Define o nome do campo da tabela que está atribuído como Primary Key.")>
    Public Property DisplayMainFieldName As String
    ''' <summary>
    ''' Define o nome do campo da tabela a ser pesquisado.
    ''' </summary>
    <Category("Query")>
    <DefaultValue(GetType(Integer), "1000")>
    <Description("Define o máximo de resultados que podem ser retornados pela pesquisa.")>
    Public Property Limit As Integer = 1000
    ''' <summary>
    ''' Define os outros campos que serão mostrados nos resultados da pesquisa.
    ''' </summary>
    <Category("Query")>
    <Description("Define os outros campos que serão mostrados nos resultados da pesquisa.")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <MergableProperty(False)>
    Public Property OtherFields As New Collection(Of OtherField)
    ''' <summary>
    ''' Define o nome do campo da tabela a ser pesquisado.
    ''' </summary>
    <Category("Query")>
    <Description("Define condições para a pesquisa. Deve ser definida com a sintaxe SQL.")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <MergableProperty(False)>
    Public Property Conditions As New Collection(Of Condition)
    ''' <summary>
    ''' Define os parâmetos utilizados nas condições da Query.
    ''' </summary>
    <Category("Query")>
    <Description("Define os parâmetos utilizados nas condições da Query.")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <MergableProperty(False)>
    Public Property Parameters As New Collection(Of Parameter)
    ''' <summary>
    ''' Define o nome do campo da tabela a ser pesquisado.
    ''' </summary>
    <Category("Query")>
    <Description("Define condições para a pesquisa. Deve ser definida com a sintaxe SQL.")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <MergableProperty(False)>
    Public Property Relations As New Collection(Of Relation)
    ''' <summary>
    ''' Define como a largura da coluna será ajustada nos resultados.
    ''' </summary>
    <Category("Aparência")>
    <DefaultValue(GetType(Color), "Blue")>
    <Description("Define como a largura da coluna será ajustada nos resultados.")>
    Public Property DisplayFieldAutoSizeColumnMode As DataGridViewAutoSizeColumnMode
    ''' <summary>
    ''' Define cor do texto para quando um resultado da pesquisa for selecionado.
    ''' </summary>
    <LocalizedCategory("Appearance")>
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
    ''' Define se o Grid mostrará linhas verticais no DropDown.
    ''' </summary>
    ''' <returns></returns>
    <Category("Aparência")>
    <DefaultValue(GetType(Boolean), "True")>
    <Description("Define se o Grid mostrará linhas verticais no DropDown.")>
    Public Property ShowVerticalGridLines As Boolean = True
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
    ''' <summary>
    ''' Define um prefixo que será congelado com o resultado.
    ''' </summary>
    <Category("Comportamento")>
    <DefaultValue(GetType(String), Nothing)>
    <Description("Define um prefixo que será congelado com o resultado.")>
    Public Property Prefix As String
    ''' <summary>
    ''' Define um sufixo que será congelado com o resultado.
    ''' </summary>
    <Category("Comportamento")>
    <DefaultValue(GetType(String), Nothing)>
    <Description("Define um sufixo que será congelado com o resultado.")>
    Public Property Suffix As String
    ''' <summary>
    ''' Limpa todo o conteúdo cado um valor seja descongelado.
    ''' </summary>
    <Category("Comportamento")>
    <DefaultValue(GetType(Boolean), "False")>
    <Description("Limpa todo o conteúdo cado um valor seja descongelado.")>
    Public Property ClearOnUnfreeze As Boolean
    ''' <summary>
    ''' Define se será mostrado o inicio do conteúdo ao sair do controle.
    ''' </summary>
    <Category("Comportamento")>
    <DefaultValue(GetType(Boolean), "True")>
    <Description("Define se será mostrado o inicio do conteúdo ao sair do controle.")>
    Public Property ShowStartOnLeave As Boolean = True

    ''' <summary>
    ''' Define se será mostrado o inicio do conteúdo ao congelar um resultado.
    ''' </summary>
    <Category("Comportamento")>
    <DefaultValue(GetType(Boolean), "False")>
    <Description("Define se será mostrado o inicio do conteúdo ao congelar um resultado.")>
    Public Property ShowStartOnFreeze As Boolean = False
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
    ''' Define se será permitido o uso de hyperlinks em registros selecionados com a tecla control.
    ''' </summary>
    ''' <returns></returns>
    <DefaultValue(GetType(Boolean), "False")>
    <Category("Comportamento")>
    <Description("Define se será permitido o uso de hyperlinks em registros selecionados com a tecla control.")>
    Public Property AllowHyperlink As Boolean = False
    ''' <summary>
    ''' Define a Conexão a ser utilizada em todos os controles para o acesso aos dados.
    ''' </summary>
    ''' <returns></returns>
    <Description("Define a Conexão a ser utilizada em todos os controles para o acesso aos dados.")>
    Public Shared Property Connection As DbConnection
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
End Class
