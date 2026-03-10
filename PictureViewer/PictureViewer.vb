Imports System.ComponentModel
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports CoreSuite.Helpers

''' <summary>
''' Controle personalizado para exibição e manipulação de imagens.
''' Permite navegar, adicionar, remover e salvar imagens em um visualizador.
''' </summary>
Public Class PictureViewer
    Inherits Panel
    Public Event PictureAdded(Path As String)
    Public Event PictureRemoved(Path As String)
    Friend WithEvents TlpControls As TableLayoutPanel
    Friend WithEvents PbxPicture As PictureBox
    Friend WithEvents LblCounter As Label
    Friend WithEvents BtnFirst As NoFocusCueButton
    Friend WithEvents BtnPrevious As NoFocusCueButton
    Friend WithEvents BtnNext As NoFocusCueButton
    Friend WithEvents BtnLast As NoFocusCueButton
    Friend WithEvents BtnSave As NoFocusCueButton
    Friend WithEvents BtnRemove As NoFocusCueButton
    Friend WithEvents BtnInclude As NoFocusCueButton
    Private _CounterMask As String
    Private _SelectedPicture As String = String.Empty
    Private _SelectedIndex As Integer = -1
    Private _MaximumPictures As Integer? = Nothing
    Private _ShowCounterBar As Boolean
    Private ReadOnly _Pictures As List(Of String)

    ''' <summary>
    ''' Obtém ou define o número máximo de imagens que podem ser carregadas no controle.
    ''' </summary>
    <Description("Obtém ou define o número máximo de imagens que podem ser carregadas no controle.")>
    <Category("Comportamento")>
    Public Property MaximumPictures As Integer?
        Get
            Return _MaximumPictures
        End Get
        Set(value As Integer?)
            _MaximumPictures = value
            LblCounter.Text = _CounterMask.Replace("{0}", 0).Replace("{1}", 0).Replace("{2}", If(MaximumPictures.HasValue, MaximumPictures, "∞"))
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Obtém a coleção somente leitura das imagens carregadas.
    ''' </summary>
    <Category("Dados")>
    <Browsable(False)>
    Public ReadOnly Property Pictures As IReadOnlyCollection(Of String)
        Get
            Return _Pictures
        End Get
    End Property

    ''' <summary>
    ''' Obtém ou define a imagem atualmente selecionada no controle.
    ''' </summary>
    <Category("Dados")>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property SelectedPicture As String
        Get
            Return _SelectedPicture
        End Get
        Set(value As String)
            _SelectedPicture = value
            _SelectedIndex = _Pictures.IndexOf(_SelectedPicture)
        End Set
    End Property

    ''' <summary>
    ''' Obtém ou define a imagem exibida no botão "Primeiro".
    ''' </summary>
    <Description("Obtém ou define a imagem exibida no botão 'Primeiro'")>
    <Category("Aparência")>
    Public Property FirstButtonImage As Image
        Get
            Return BtnFirst.Image
        End Get
        Set(value As Image)
            BtnFirst.Image = value
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Obtém ou define a imagem exibida no botão "Anterior".
    ''' </summary>
    <Description("Obtém ou define a imagem exibida no botão 'Anterior'")>
    <Category("Aparência")>
    Public Property PreviousButtonImage As Image
        Get
            Return BtnPrevious.Image
        End Get
        Set(value As Image)
            BtnPrevious.Image = value
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Obtém ou define a imagem exibida no botão "Próximo".
    ''' </summary>
    <Description("Obtém ou define a imagem exibida no botão 'Próximo'")>
    <Category("Aparência")>
    Public Property NextButtonImage As Image
        Get
            Return BtnNext.Image
        End Get
        Set(value As Image)
            BtnNext.Image = value
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Obtém ou define a imagem exibida no botão "Último".
    ''' </summary>
    <Description("Obtém ou define a imagem exibida no botão 'Último'")>
    <Category("Aparência")>
    Public Property LastButtonImage As Image
        Get
            Return BtnLast.Image
        End Get
        Set(value As Image)
            BtnLast.Image = value
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Obtém ou define a imagem exibida no botão "Salvar".
    ''' </summary>
    <Description("Obtém ou define a imagem exibida no botão 'Salvar'")>
    <Category("Aparência")>
    Public Property SaveButtonImage As Image
        Get
            Return BtnSave.Image
        End Get
        Set(value As Image)
            BtnSave.Image = value
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Obtém ou define a imagem exibida no botão "Remover".
    ''' </summary>
    <Description("Obtém ou define a imagem exibida no botão 'Remover'")>
    <Category("Aparência")>
    Public Property RemoveButtonImage As Image
        Get
            Return BtnRemove.Image
        End Get
        Set(value As Image)
            BtnRemove.Image = value
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Obtém ou define a imagem exibida no botão "Incluir".
    ''' </summary>
    <Description("Obtém ou define a imagem exibida no botão 'Incluir'")>
    <Category("Aparência")>
    Public Property IncludeButtonImage As Image
        Get
            Return BtnInclude.Image
        End Get
        Set(value As Image)
            BtnInclude.Image = value
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Obtém ou define a máscara de formatação exibida no contador.
    ''' Os placeholders {0}, {1} e {2} representam:
    ''' índice atual, total de imagens e limite máximo permitido, respectivamente.
    ''' </summary>
    <Description("Obtém ou define a máscara de formatação exibida no contador. Os placeholders {0}, {1} e {2} representam: índice atual, total de imagens e limite máximo permitido, respectivamente.")>
    <Category("Aparência")>
    Public Property CounterMask As String
        Get
            Return _CounterMask
        End Get
        Set(value As String)
            _CounterMask = value
            LblCounter.Text = _CounterMask.Replace("{0}", 0).Replace("{1}", 0).Replace("{2}", If(MaximumPictures.HasValue, MaximumPictures, "∞"))
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Obtém ou define se a barra de controle será exibida.
    ''' </summary>
    <Description("Obtém ou define se a barra de controle será exibida.")>
    <Category("Aparência")>
    Public Property ShowControlBar As Boolean
        Get
            Return TlpControls.Visible
        End Get
        Set(value As Boolean)
            TlpControls.Visible = value
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Obtém ou define se a barra de contador será exibida.
    ''' </summary>
    <Description("Obtém ou define se a barra de contador será exibida.")>
    <Category("Aparência")>
    Public Property ShowCounterBar As Boolean
        Get
            Return _ShowCounterBar
        End Get
        Set(value As Boolean)
            _ShowCounterBar = value
            LblCounter.Visible = _ShowCounterBar
            RefreshControls()
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Obtém ou define a cor de fundo do controle e da área da imagem.
    ''' </summary>
    <Description("Obtém ou define a cor de fundo do controle e da área da imagem.")>
    <Category("Aparência")>
    Overloads Property BackColor As Color
        Get
            Return MyBase.BackColor
        End Get
        Set(value As Color)
            MyBase.BackColor = value
            PbxPicture.BackColor = value
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Obtém ou define a cor de fundo da barra de contador.
    ''' </summary>
    <Description("Obtém ou define a cor de fundo da barra de contador.")>
    <Category("Aparência")>
    Public Property CounterBarBackColor As Color
        Get
            Return LblCounter.BackColor
        End Get
        Set(value As Color)
            LblCounter.BackColor = value
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Obtém ou define a cor de fundo da barra de controle.
    ''' </summary>
    <Description("Obtém ou define a cor de fundo da barra de controle.")>
    <Category("Aparência")>
    Public Property ControlBarBackColor As Color
        Get
            Return TlpControls.BackColor
        End Get
        Set(value As Color)
            TlpControls.BackColor = value
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Inicializa uma nova instância do <see cref="PictureViewer"/>.
    ''' </summary>
    Public Sub New()
        InitializeComponents()
        _Pictures = New List(Of String)
        RefreshControls()
    End Sub

    Private Sub InitializeComponents()
        BtnFirst = New NoFocusCueButton With {
            .Anchor = AnchorStyles.None,
            .BackgroundImageLayout = ImageLayout.Center,
            .BackColor = Color.Transparent,
            .FlatStyle = FlatStyle.Flat
        }
        BtnFirst.FlatAppearance.BorderSize = 0
        BtnFirst.TooltipText = "Primeiro"
        BtnFirst.UseVisualStyleBackColor = False
        BtnFirst.Image = My.Resources.NavFirst
        BtnPrevious = New NoFocusCueButton With {
            .Anchor = AnchorStyles.None,
            .BackgroundImageLayout = ImageLayout.Center,
            .BackColor = Color.Transparent,
            .FlatStyle = FlatStyle.Flat
        }
        BtnPrevious.FlatAppearance.BorderSize = 0
        BtnPrevious.TooltipText = "Anterior"
        BtnPrevious.UseVisualStyleBackColor = False
        BtnPrevious.Image = My.Resources.NavPrevious
        BtnNext = New NoFocusCueButton With {
            .Anchor = AnchorStyles.None,
            .BackgroundImageLayout = ImageLayout.Center,
            .BackColor = Color.Transparent,
            .FlatStyle = FlatStyle.Flat
        }
        BtnNext.FlatAppearance.BorderSize = 0
        BtnNext.TooltipText = "Próximo"
        BtnNext.UseVisualStyleBackColor = False
        BtnNext.Image = My.Resources.NavNext
        BtnLast = New NoFocusCueButton With {
            .Anchor = AnchorStyles.None,
            .BackgroundImageLayout = ImageLayout.Center,
            .BackColor = Color.Transparent,
            .FlatStyle = FlatStyle.Flat
        }
        BtnLast.FlatAppearance.BorderSize = 0
        BtnLast.TooltipText = "Último"
        BtnLast.UseVisualStyleBackColor = False
        BtnLast.Image = My.Resources.NavLast
        BtnSave = New NoFocusCueButton With {
            .Anchor = AnchorStyles.None,
            .BackgroundImageLayout = ImageLayout.Center,
            .BackColor = Color.Transparent,
            .FlatStyle = FlatStyle.Flat
        }
        BtnSave.FlatAppearance.BorderSize = 0
        BtnSave.TooltipText = "Salvar"
        BtnSave.UseVisualStyleBackColor = False
        BtnSave.Image = My.Resources.ImageSave
        BtnRemove = New NoFocusCueButton With {
            .Anchor = AnchorStyles.None,
            .BackgroundImageLayout = ImageLayout.Center,
            .BackColor = Color.Transparent,
            .FlatStyle = FlatStyle.Flat
        }
        BtnRemove.FlatAppearance.BorderSize = 0
        BtnRemove.TooltipText = "Remover"
        BtnRemove.UseVisualStyleBackColor = False
        BtnRemove.Image = My.Resources.ImageDelete
        BtnInclude = New NoFocusCueButton With {
            .Anchor = AnchorStyles.None,
            .BackgroundImageLayout = ImageLayout.Center,
            .BackColor = Color.Transparent,
            .FlatStyle = FlatStyle.Flat
        }
        BtnInclude.FlatAppearance.BorderSize = 0
        BtnInclude.TooltipText = "Incluir"
        BtnInclude.UseVisualStyleBackColor = False
        BtnInclude.Image = My.Resources.ImageInclude
        TlpControls = New TableLayoutPanel With {
            .BackColor = Color.White,
            .ColumnCount = 9
        }
        TlpControls.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        TlpControls.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 30.0!))
        TlpControls.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 30.0!))
        TlpControls.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 30.0!))
        TlpControls.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 30.0!))
        TlpControls.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 30.0!))
        TlpControls.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 30.0!))
        TlpControls.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 30.0!))
        TlpControls.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        TlpControls.Controls.Add(BtnLast, 7, 0)
        TlpControls.Controls.Add(BtnNext, 6, 0)
        TlpControls.Controls.Add(BtnSave, 5, 0)
        TlpControls.Controls.Add(BtnRemove, 4, 0)
        TlpControls.Controls.Add(BtnInclude, 3, 0)
        TlpControls.Controls.Add(BtnPrevious, 2, 0)
        TlpControls.Controls.Add(BtnFirst, 1, 0)
        TlpControls.Dock = DockStyle.Top
        TlpControls.RowCount = 1
        TlpControls.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        TlpControls.Height = 28
        PbxPicture = New PictureBox With {
            .BackColor = Color.White,
            .Dock = DockStyle.Fill,
            .SizeMode = PictureBoxSizeMode.Zoom
        }
        LblCounter = New Label With {
            .BackColor = Color.White,
            .Dock = DockStyle.Bottom,
            .TextAlign = ContentAlignment.MiddleCenter
        }
        Padding = New Padding(1)
        Controls.Add(TlpControls)
        Controls.Add(LblCounter)
        Controls.Add(PbxPicture)
        Size = New Size(240, 150)
        CounterMask = "{0}/{1}"
    End Sub

    Private Sub ShowSelectedPicture()
        If _SelectedPicture IsNot Nothing Then
            If PbxPicture.Image IsNot Nothing Then PbxPicture.Image.Dispose()
            Using FileStream As New FileStream(_SelectedPicture, FileMode.Open, FileAccess.Read)
                Dim img As Image = Image.FromStream(FileStream)
                PbxPicture.Image = CType(img.Clone(), Image)
            End Using
        Else
            PbxPicture.Image = Nothing
        End If
    End Sub

    ''' <summary>
    ''' Adiciona uma única imagem ao controle.
    ''' Caso a imagem já exista, ela não será duplicada.
    ''' </summary>
    ''' <param name="Path">Caminho do arquivo de imagem.</param>
    Public Sub AddPicture(Path As String)
        If File.Exists(Path) Then
            If Not (MaximumPictures.HasValue AndAlso _Pictures.Count = MaximumPictures) Then
                _Pictures.Add(Path)
                RaiseEvent PictureAdded(Path)
                SelectedPicture = Path
            End If
            ShowSelectedPicture()
            RefreshControls()
        End If
    End Sub

    ''' <summary>
    ''' Adiciona múltiplas imagens ao controle.
    ''' Caso alguma imagem já exista, ela não será duplicada.
    ''' </summary>
    ''' <param name="Paths">Coleção de caminhos de arquivos de imagem.</param>
    Public Sub AddPictures(Paths As IEnumerable(Of String), Optional SelectedIndex As Integer? = Nothing)
        For Each Path In Paths.Reverse()
            If Not (MaximumPictures.HasValue AndAlso _Pictures.Count = MaximumPictures) Then
                If File.Exists(Path) Then
                    _Pictures.Insert(0, Path)
                    RaiseEvent PictureAdded(Path)
                End If
            End If
        Next Path
        If _Pictures.Count > 0 Then
            If SelectedIndex.HasValue AndAlso SelectedIndex >= 0 AndAlso SelectedIndex < _Pictures.Count Then
                SelectedPicture = _Pictures(SelectedIndex.Value)
            Else
                SelectedPicture = _Pictures.Last()
            End If
            ShowSelectedPicture()
            RefreshControls()
        End If
    End Sub

    ''' <summary>
    ''' Remove uma imagem previamente adicionada ao controle.
    ''' </summary>
    ''' <param name="Path">Caminho do arquivo de imagem a ser removido.</param>
    Public Sub RemovePicture(Path As String)
        If String.IsNullOrEmpty(Path) OrElse Not _Pictures.Contains(Path) Then Return
        Dim index As Integer = _Pictures.IndexOf(Path)
        _Pictures.RemoveAt(index)
        RaiseEvent PictureRemoved(Path)
        If _Pictures.Count = 0 Then
            SelectedPicture = Nothing
        ElseIf index < _Pictures.Count Then
            SelectedPicture = _Pictures(index)
        Else
            SelectedPicture = _Pictures(_Pictures.Count - 1)
        End If
        ShowSelectedPicture()
        RefreshControls()
    End Sub

    Public Sub Clear()
        For Each Picture In _Pictures.ToList()
            RemovePicture(Picture)
        Next Picture
    End Sub

    Private Sub BtnInclude_Click(sender As Object, e As EventArgs) Handles BtnInclude.Click
        Using OpenFileDialog As New OpenFileDialog()
            OpenFileDialog.Filter = "Imagens|*.jpg;*.jpeg;*.png;*.bmp|Todos os arquivos|*.*"
            OpenFileDialog.Title = "Selecionar Imagem"
            OpenFileDialog.Multiselect = True
            If OpenFileDialog.ShowDialog() = DialogResult.OK Then
                Cursor = Cursors.WaitCursor
                If OpenFileDialog.FileNames().Count = 1 Then
                    AddPicture(OpenFileDialog.FileName)
                Else
                    AddPictures(OpenFileDialog.FileNames())
                End If

                Cursor = Cursors.Default
            End If
        End Using
    End Sub

    Private Sub BtnSavePicture_Click(sender As Object, e As EventArgs) Handles BtnSave.Click
        Using SaveFileDialog As New SaveFileDialog()
            SaveFileDialog.Filter = "JPEG Image|*.jpg|PNG Image|*.png|BMP Image|*.bmp"
            SaveFileDialog.Title = "Salvar Imagem"
            SaveFileDialog.FileName = "Foto"
            If SaveFileDialog.ShowDialog() = DialogResult.OK Then
                Dim FileExtension As String = IO.Path.GetExtension(SaveFileDialog.FileName).ToLower()
                Dim ImageFormat As Imaging.ImageFormat
                Select Case FileExtension
                    Case ".jpg"
                        ImageFormat = Imaging.ImageFormat.Jpeg
                    Case ".png"
                        ImageFormat = Imaging.ImageFormat.Png
                    Case ".bmp"
                        ImageFormat = Imaging.ImageFormat.Bmp
                    Case Else
                        MessageBox.Show("Formato de arquivo não suportado.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                End Select
                Using Bitmap As New Bitmap(PbxPicture.Image.Width, PbxPicture.Image.Height, PbxPicture.Image.PixelFormat)
                    Using Graphic As Graphics = Graphics.FromImage(Bitmap)
                        Graphic.DrawImage(PbxPicture.Image, 0, 0, PbxPicture.Image.Width, PbxPicture.Image.Height)
                    End Using
                    Bitmap.Save(SaveFileDialog.FileName, ImageFormat)
                End Using
            End If
        End Using
    End Sub

    Private Sub BtnRemovePicture_Click(sender As Object, e As EventArgs) Handles BtnRemove.Click
        RemovePicture(_SelectedPicture)
    End Sub

    Private Sub BtnPrevious_Click(sender As Object, e As EventArgs) Handles BtnPrevious.Click
        SelectedPicture = Pictures(_Pictures.IndexOf(SelectedPicture) - 1)
        ShowSelectedPicture()
        RefreshControls()
    End Sub

    Private Sub BtnNext_Click(sender As Object, e As EventArgs) Handles BtnNext.Click
        SelectedPicture = Pictures(_Pictures.IndexOf(SelectedPicture) + 1)
        ShowSelectedPicture()
        RefreshControls()
    End Sub

    Private Sub BtnFirst_Click(sender As Object, e As EventArgs) Handles BtnFirst.Click
        SelectedPicture = Pictures(0)
        ShowSelectedPicture()
        RefreshControls()
    End Sub

    Private Sub BtnLast_Click(sender As Object, e As EventArgs) Handles BtnLast.Click
        SelectedPicture = Pictures(Pictures.Count - 1)
        ShowSelectedPicture()
        RefreshControls()
    End Sub

    Private Sub RefreshControls()
        Dim PictureCount As Integer = Pictures.Count
        Dim PictureIndex As Integer = _Pictures.IndexOf(SelectedPicture)
        If PictureCount < 1 Then
            If Not IsDesignTime() Then
                LblCounter.Visible = False
            End If
            BtnSave.Enabled = False
            BtnRemove.Enabled = False
            BtnFirst.Enabled = False
            BtnPrevious.Enabled = False
            BtnNext.Enabled = False
            BtnLast.Enabled = False
            PbxPicture.Image = Nothing
        Else
            If ShowCounterBar Then
                LblCounter.Visible = True
            End If
            LblCounter.Text = _CounterMask.Replace("{0}", PictureIndex + 1).Replace("{1}", PictureCount).Replace("{2}", If(MaximumPictures.HasValue, MaximumPictures, "∞"))
            BtnRemove.Enabled = True
            BtnSave.Enabled = True
            BtnFirst.Enabled = (PictureIndex > 0)
            BtnPrevious.Enabled = (PictureIndex > 0)
            BtnNext.Enabled = (PictureIndex < PictureCount - 1)
            BtnLast.Enabled = (PictureIndex < PictureCount - 1)
        End If
        If _Pictures.Count = MaximumPictures Then
            BtnInclude.Enabled = False
        Else
            BtnInclude.Enabled = True
        End If
    End Sub

    Private Function IsDesignTime() As Boolean
        Dim c As Control = Me
        While c IsNot Nothing
            If c.Site IsNot Nothing AndAlso c.Site.DesignMode Then
                Return True
            End If
            c = c.Parent
        End While
        Return System.ComponentModel.LicenseManager.UsageMode = System.ComponentModel.LicenseUsageMode.Designtime
    End Function

    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing Then
                If PbxPicture IsNot Nothing AndAlso PbxPicture.Image IsNot Nothing Then
                    PbxPicture.Image.Dispose()
                    PbxPicture.Image = Nothing
                End If
                _Pictures.Clear()
                For Each Button As NoFocusCueButton In {BtnFirst, BtnPrevious, BtnNext, BtnLast, BtnSave, BtnRemove, BtnInclude}
                    If Button IsNot Nothing AndAlso Button.Image IsNot Nothing Then
                        Button.Image.Dispose()
                        Button.Image = Nothing
                    End If
                Next
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

End Class