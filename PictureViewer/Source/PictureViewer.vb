Imports System.ComponentModel
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Runtime.InteropServices
Imports CoreSuite.Controls.My.Resources
''' <summary>
''' Provides an image viewer that supports adding, removing, navigating, and saving image files.
''' </summary>
Public Class PictureViewer
    Inherits Panel
    ''' <summary>
    ''' Occurs after an image is added to the viewer.
    ''' </summary>
    <Category("PictureViewer")>
    Public Event PictureAdded(Path As String)
    ''' <summary>
    ''' Occurs after an image is removed from the viewer.
    ''' </summary>
    <Category("PictureViewer")>
    Public Event PictureRemoved(Path As String)
    ''' <summary>
    ''' Occurs when the selected image changes.
    ''' </summary>
    <Category("PictureViewer")>
    Public Event SelectedPictureChanged(Path As String)
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
    Private _CounterMask As String = "{0}/{1}"
    Private _SelectedPicture As String
    Private _SelectedIndex As Integer = -1
    Private _MaximumPictures As Integer?
    Private _ShowCounterBar As Boolean
    Private ReadOnly _ToolTips As New PictureViewerToolTips()
    Private ReadOnly _Pictures As New List(Of String)
    ''' <summary>
    ''' Gets or sets the maximum number of images that can be added to the viewer.
    ''' </summary>
    <Description("Gets or sets the maximum number of images that can be added to the viewer.")>
    <Category("PictureViewer")>
    <DefaultValue(GetType(Integer), Nothing)>
    Public Property MaximumPictures As Integer?
        Get
            Return _MaximumPictures
        End Get
        Set(value As Integer?)
            If value.HasValue AndAlso value.Value < 1 Then Throw New ArgumentOutOfRangeException(NameOf(value), "MaximumPictures must be greater than zero.")
            If value.HasValue AndAlso value.Value < _Pictures.Count Then Throw New InvalidOperationException("MaximumPictures cannot be less than the current number of images.")
            If _MaximumPictures = value Then Return
            _MaximumPictures = value
            UpdateCounterText()
            RefreshControls()
        End Set
    End Property
    ''' <summary>
    ''' Gets a read-only list containing the paths of the images added to the viewer.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property Pictures As IReadOnlyList(Of String)
        Get
            Return _Pictures.AsReadOnly()
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets the path of the currently selected image.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property SelectedPicture As String
        Get
            Return _SelectedPicture
        End Get
        Set(value As String)
            Dim NewSelectedPicture As String = Nothing
            Dim NewSelectedIndex As Integer = -1
            If Not String.IsNullOrWhiteSpace(value) Then
                NewSelectedIndex = FindPictureIndex(value)
                If NewSelectedIndex < 0 Then Throw New ArgumentException($"Picture '{value}' is not contained in the viewer.", NameOf(value))
                NewSelectedPicture = _Pictures(NewSelectedIndex)
            End If
            If String.Equals(_SelectedPicture, NewSelectedPicture, StringComparison.OrdinalIgnoreCase) Then Return
            _SelectedPicture = NewSelectedPicture
            _SelectedIndex = NewSelectedIndex
            ShowSelectedPicture()
            RefreshControls()
            RaiseEvent SelectedPictureChanged(_SelectedPicture)
        End Set
    End Property
    ''' <summary>
    ''' Gets the index of the currently selected image.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property SelectedIndex As Integer
        Get
            Return _SelectedIndex
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets the image displayed by the first-image button.
    ''' </summary>
    <Description("Gets or sets the image displayed by the first-image button.")>
    <Category("PictureViewer")>
    Public Property FirstButtonImage As Image
        Get
            Return BtnFirst.Image
        End Get
        Set(value As Image)
            BtnFirst.Image = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the image displayed by the previous-image button.
    ''' </summary>
    <Description("Gets or sets the image displayed by the previous-image button.")>
    <Category("PictureViewer")>
    Public Property PreviousButtonImage As Image
        Get
            Return BtnPrevious.Image
        End Get
        Set(value As Image)
            BtnPrevious.Image = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the image displayed by the next-image button.
    ''' </summary>
    <Description("Gets or sets the image displayed by the next-image button.")>
    <Category("PictureViewer")>
    Public Property NextButtonImage As Image
        Get
            Return BtnNext.Image
        End Get
        Set(value As Image)
            BtnNext.Image = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the image displayed by the last-image button.
    ''' </summary>
    <Description("Gets or sets the image displayed by the last-image button.")>
    <Category("PictureViewer")>
    Public Property LastButtonImage As Image
        Get
            Return BtnLast.Image
        End Get
        Set(value As Image)
            BtnLast.Image = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the image displayed by the save button.
    ''' </summary>
    <Description("Gets or sets the image displayed by the save button.")>
    <Category("PictureViewer")>
    Public Property SaveButtonImage As Image
        Get
            Return BtnSave.Image
        End Get
        Set(value As Image)
            BtnSave.Image = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the image displayed by the remove button.
    ''' </summary>
    <Description("Gets or sets the image displayed by the remove button.")>
    <Category("PictureViewer")>
    Public Property RemoveButtonImage As Image
        Get
            Return BtnRemove.Image
        End Get
        Set(value As Image)
            BtnRemove.Image = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the image displayed by the include button.
    ''' </summary>
    <Description("Gets or sets the image displayed by the include button.")>
    <Category("PictureViewer")>
    Public Property IncludeButtonImage As Image
        Get
            Return BtnInclude.Image
        End Get
        Set(value As Image)
            BtnInclude.Image = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the format mask displayed by the image counter.
    ''' The {0}, {1}, and {2} placeholders represent the current position, image count, and maximum image count.
    ''' </summary>
    <Description("Gets or sets the format mask displayed by the image counter. The {0}, {1}, and {2} placeholders represent the current position, image count, and maximum image count.")>
    <Category("PictureViewer")>
    <DefaultValue("{0}/{1}")>
    Public Property CounterMask As String
        Get
            Return _CounterMask
        End Get
        Set(value As String)
            If value Is Nothing Then value = String.Empty
            If _CounterMask = value Then Return
            _CounterMask = value
            UpdateCounterText()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether the control bar is visible.
    ''' </summary>
    <Description("Gets or sets whether the control bar is visible.")>
    <Category("PictureViewer")>
    <DefaultValue(True)>
    Public Property ShowControlBar As Boolean
        Get
            Return TlpControls.Visible
        End Get
        Set(value As Boolean)
            TlpControls.Visible = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether the image counter bar is visible.
    ''' </summary>
    <Description("Gets or sets whether the image counter bar is visible.")>
    <Category("PictureViewer")>
    <DefaultValue(False)>
    Public Property ShowCounterBar As Boolean
        Get
            Return _ShowCounterBar
        End Get
        Set(value As Boolean)
            If _ShowCounterBar = value Then Return
            _ShowCounterBar = value
            RefreshControls()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the background color of the viewer and image area.
    ''' </summary>
    <Description("Gets or sets the background color of the viewer and image area.")>
    <Category("PictureViewer")>
    Public Overrides Property BackColor As Color
        Get
            Return MyBase.BackColor
        End Get
        Set(value As Color)
            MyBase.BackColor = value
            If PbxPicture IsNot Nothing Then PbxPicture.BackColor = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the background color of the image counter bar.
    ''' </summary>
    <Description("Gets or sets the background color of the image counter bar.")>
    <Category("PictureViewer")>
    Public Property CounterBarBackColor As Color
        Get
            Return LblCounter.BackColor
        End Get
        Set(value As Color)
            LblCounter.BackColor = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the background color of the control bar.
    ''' </summary>
    <Description("Gets or sets the background color of the control bar.")>
    <Category("PictureViewer")>
    Public Property ControlBarBackColor As Color
        Get
            Return TlpControls.BackColor
        End Get
        Set(value As Color)
            TlpControls.BackColor = value
        End Set
    End Property
    ''' <summary>
    ''' Gets the tooltip texts displayed by the toolbar buttons.
    ''' </summary>
    <Description("Gets the tooltip texts displayed by the toolbar buttons.")>
    <Category("PictureViewer")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public ReadOnly Property ToolTips As PictureViewerToolTips
        Get
            Return _ToolTips
        End Get
    End Property
    ''' <summary>
    ''' Initializes a new instance of the <see cref="PictureViewer"/> class.
    ''' </summary>
    Public Sub New()
        InitializeComponents()
        AddHandler _ToolTips.Changed, AddressOf UpdateToolTips
        UpdateToolTips()
        RefreshControls()
    End Sub
    ''' <summary>
    ''' Initializes the internal controls that compose the picture viewer.
    ''' </summary>
    Private Sub InitializeComponents()
        BtnFirst = CreateButton(ToolTips.First, Images.NavFirst)
        BtnPrevious = CreateButton(ToolTips.Previous, Images.NavPrevious)
        BtnNext = CreateButton(ToolTips.Next, Images.NavNext)
        BtnLast = CreateButton(ToolTips.Last, Images.NavLast)
        BtnSave = CreateButton(ToolTips.Save, Images.ImageSave)
        BtnRemove = CreateButton(ToolTips.Remove, Images.ImageDelete)
        BtnInclude = CreateButton(ToolTips.Include, Images.ImageInclude)
        TlpControls = New TableLayoutPanel With {.BackColor = Color.White, .ColumnCount = 9, .Dock = DockStyle.Top, .Height = 28, .RowCount = 1}
        TlpControls.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        TlpControls.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 30.0!))
        TlpControls.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 30.0!))
        TlpControls.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 30.0!))
        TlpControls.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 30.0!))
        TlpControls.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 30.0!))
        TlpControls.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 30.0!))
        TlpControls.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 30.0!))
        TlpControls.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        TlpControls.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        TlpControls.Controls.Add(BtnFirst, 1, 0)
        TlpControls.Controls.Add(BtnPrevious, 2, 0)
        TlpControls.Controls.Add(BtnInclude, 3, 0)
        TlpControls.Controls.Add(BtnRemove, 4, 0)
        TlpControls.Controls.Add(BtnSave, 5, 0)
        TlpControls.Controls.Add(BtnNext, 6, 0)
        TlpControls.Controls.Add(BtnLast, 7, 0)
        PbxPicture = New PictureBox With {.BackColor = Color.White, .Dock = DockStyle.Fill, .SizeMode = PictureBoxSizeMode.Zoom}
        LblCounter = New Label With {.BackColor = Color.White, .Dock = DockStyle.Bottom, .TextAlign = ContentAlignment.MiddleCenter, .Visible = False}
        Padding = New Padding(1)
        Controls.Add(PbxPicture)
        Controls.Add(LblCounter)
        Controls.Add(TlpControls)
        Size = New Size(240, 150)
    End Sub
    ''' <summary>
    ''' Creates a toolbar button with the specified tooltip text and image.
    ''' </summary>
    ''' <param name="TooltipText">The tooltip text displayed for the button.</param>
    ''' <param name="Image">The image displayed on the button.</param>
    ''' <returns>A configured <see cref="NoFocusCueButton"/> instance.</returns>
    Private Shared Function CreateButton(TooltipText As String, Image As Image) As NoFocusCueButton
        Dim Button = New NoFocusCueButton With {
            .Anchor = AnchorStyles.None,
            .BackgroundImageLayout = ImageLayout.Center,
            .BackColor = Color.Transparent,
            .FlatStyle = FlatStyle.Flat,
            .Image = Image,
            .TooltipText = TooltipText,
            .UseVisualStyleBackColor = False
        }
        Button.FlatAppearance.BorderSize = 0
        Return Button
    End Function
    ''' <summary>
    ''' Updates the tooltip texts displayed by all toolbar buttons.
    ''' </summary>
    Private Sub UpdateToolTips()
        BtnFirst.TooltipText = _ToolTips.First
        BtnPrevious.TooltipText = _ToolTips.Previous
        BtnNext.TooltipText = _ToolTips.Next
        BtnLast.TooltipText = _ToolTips.Last
        BtnInclude.TooltipText = _ToolTips.Include
        BtnRemove.TooltipText = _ToolTips.Remove
        BtnSave.TooltipText = _ToolTips.Save
    End Sub
    ''' <summary>
    ''' Finds the index of the specified image path.
    ''' </summary>
    ''' <param name="Path">The image path to locate.</param>
    ''' <returns>The zero-based index of the image if found; otherwise, -1.</returns>
    Private Function FindPictureIndex(Path As String) As Integer
        For Index As Integer = 0 To _Pictures.Count - 1
            If String.Equals(_Pictures(Index), Path, StringComparison.OrdinalIgnoreCase) Then Return Index
        Next
        Return -1
    End Function
    ''' <summary>
    ''' Determines whether another image can be added to the viewer.
    ''' </summary>
    ''' <returns><see langword="True"/> if another image can be added; otherwise, <see langword="False"/>.</returns>
    Private Function CanAddPicture() As Boolean
        Return Not MaximumPictures.HasValue OrElse _Pictures.Count < MaximumPictures.Value
    End Function
    ''' <summary>
    ''' Displays the currently selected image.
    ''' </summary>
    Private Sub ShowSelectedPicture()
        DisposeDisplayedImage()
        If String.IsNullOrWhiteSpace(_SelectedPicture) Then Return
        If Not File.Exists(_SelectedPicture) Then Throw New FileNotFoundException("The selected image file could not be found.", _SelectedPicture)
        Try
            Using Stream As New FileStream(_SelectedPicture, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Using SourceImage As Image = Image.FromStream(Stream)
                    PbxPicture.Image = New Bitmap(SourceImage)
                End Using
            End Using
        Catch Ex As Exception When TypeOf Ex Is ArgumentException OrElse TypeOf Ex Is ExternalException
            Throw New InvalidDataException($"The file '{_SelectedPicture}' does not contain a valid supported image.", Ex)
        End Try
    End Sub
    ''' <summary>
    ''' Releases the currently displayed image from memory.
    ''' </summary>
    Private Sub DisposeDisplayedImage()
        If PbxPicture Is Nothing OrElse PbxPicture.Image Is Nothing Then Return
        Dim CurrentImage As Image = PbxPicture.Image
        PbxPicture.Image = Nothing
        CurrentImage.Dispose()
    End Sub
    ''' <summary>
    ''' Updates the text displayed by the image counter.
    ''' </summary>
    Private Sub UpdateCounterText()
        If LblCounter Is Nothing Then Return
        Dim CurrentPosition As Integer = If(_SelectedIndex >= 0, _SelectedIndex + 1, 0)
        Dim MaximumValue As Object = If(MaximumPictures, "∞")
        LblCounter.Text = _CounterMask.Replace("{0}", CurrentPosition.ToString()).Replace("{1}", _Pictures.Count.ToString()).Replace("{2}", MaximumValue.ToString())
    End Sub
    ''' <summary>
    ''' Adds an image file to the viewer.
    ''' </summary>
    ''' <param name="Path">The path of the image file to add.</param>
    ''' <returns><see langword="True"/> if the image was added; otherwise, <see langword="False"/>.</returns>
    Public Function AddPicture(Path As String) As Boolean
        If String.IsNullOrWhiteSpace(Path) OrElse Not File.Exists(Path) Then Return False
        If FindPictureIndex(Path) >= 0 OrElse Not CanAddPicture() Then Return False
        _Pictures.Add(Path)
        RaiseEvent PictureAdded(Path)
        SelectedPicture = Path
        Return True
    End Function
    ''' <summary>
    ''' Adds multiple image files to the viewer.
    ''' </summary>
    ''' <param name="Paths">The paths of the image files to add.</param>
    ''' <param name="SelectedIndex">The index to select after adding the images.</param>
    ''' <returns>The number of images successfully added.</returns>
    Public Function AddPictures(Paths As IEnumerable(Of String), Optional SelectedIndex As Integer? = Nothing) As Integer
        ArgumentNullException.ThrowIfNull(Paths)
        Dim AddedPictures As New List(Of String)
        For Each Path As String In Paths
            If Not CanAddPicture() Then Exit For
            If String.IsNullOrWhiteSpace(Path) OrElse Not File.Exists(Path) OrElse FindPictureIndex(Path) >= 0 Then Continue For
            _Pictures.Add(Path)
            AddedPictures.Add(Path)
            RaiseEvent PictureAdded(Path)
        Next
        If _Pictures.Count = 0 Then
            RefreshControls()
            Return 0
        End If
        If SelectedIndex.HasValue Then
            If SelectedIndex.Value < 0 OrElse SelectedIndex.Value >= _Pictures.Count Then Throw New ArgumentOutOfRangeException(NameOf(SelectedIndex), "SelectedIndex must reference an existing image.")
            SelectedPicture = _Pictures(SelectedIndex.Value)
        ElseIf AddedPictures.Count > 0 Then
            SelectedPicture = AddedPictures(AddedPictures.Count - 1)
        ElseIf _SelectedIndex < 0 Then
            SelectedPicture = _Pictures(0)
        Else
            RefreshControls()
        End If
        Return AddedPictures.Count
    End Function
    ''' <summary>
    ''' Removes an image from the viewer.
    ''' </summary>
    ''' <param name="Path">The path of the image to remove.</param>
    ''' <returns><see langword="True"/> if the image was removed; otherwise, <see langword="False"/>.</returns>
    Public Function RemovePicture(Path As String) As Boolean
        If String.IsNullOrWhiteSpace(Path) Then Return False
        Dim RemovedIndex As Integer = FindPictureIndex(Path)
        If RemovedIndex < 0 Then Return False
        Dim RemovedPath As String = _Pictures(RemovedIndex)
        Dim WasSelected As Boolean = RemovedIndex = _SelectedIndex
        _Pictures.RemoveAt(RemovedIndex)
        RaiseEvent PictureRemoved(RemovedPath)
        If _Pictures.Count = 0 Then
            SelectedPicture = Nothing
        ElseIf WasSelected Then
            SelectedPicture = _Pictures(Math.Min(RemovedIndex, _Pictures.Count - 1))
        Else
            _SelectedIndex = FindPictureIndex(_SelectedPicture)
            RefreshControls()
        End If
        Return True
    End Function
    ''' <summary>
    ''' Removes all images from the viewer.
    ''' </summary>
    Public Sub Clear()
        If _Pictures.Count = 0 Then Return
        Dim RemovedPictures As String() = _Pictures.ToArray()
        _Pictures.Clear()
        _SelectedPicture = Nothing
        _SelectedIndex = -1
        DisposeDisplayedImage()
        For Each Path As String In RemovedPictures
            RaiseEvent PictureRemoved(Path)
        Next
        RefreshControls()
        RaiseEvent SelectedPictureChanged(Nothing)
    End Sub
    ''' <summary>
    ''' Selects the first image in the viewer.
    ''' </summary>
    Public Sub MoveFirst()
        If _Pictures.Count = 0 OrElse _SelectedIndex = 0 Then Return
        SelectedPicture = _Pictures(0)
    End Sub
    ''' <summary>
    ''' Selects the previous image in the viewer.
    ''' </summary>
    Public Sub MovePrevious()
        If _SelectedIndex <= 0 Then Return
        SelectedPicture = _Pictures(_SelectedIndex - 1)
    End Sub
    ''' <summary>
    ''' Selects the next image in the viewer.
    ''' </summary>
    Public Sub MoveNext()
        If _SelectedIndex < 0 OrElse _SelectedIndex >= _Pictures.Count - 1 Then Return
        SelectedPicture = _Pictures(_SelectedIndex + 1)
    End Sub
    ''' <summary>
    ''' Selects the last image in the viewer.
    ''' </summary>
    Public Sub MoveLast()
        If _Pictures.Count = 0 OrElse _SelectedIndex = _Pictures.Count - 1 Then Return
        SelectedPicture = _Pictures(_Pictures.Count - 1)
    End Sub
    ''' <summary>
    ''' Saves the currently displayed image to the specified file.
    ''' </summary>
    ''' <param name="Path">The destination file path.</param>
    Public Sub SaveSelectedPicture(Path As String)
        If PbxPicture.Image Is Nothing Then Throw New InvalidOperationException("There is no selected image to save.")
        If String.IsNullOrWhiteSpace(Path) Then Throw New ArgumentException("A destination file path must be provided.", NameOf(Path))
        Dim Extension As String = IO.Path.GetExtension(Path).ToLowerInvariant()
        Dim Format As ImageFormat
        Select Case Extension
            Case ".jpg", ".jpeg"
                Format = ImageFormat.Jpeg
            Case ".png"
                Format = ImageFormat.Png
            Case ".bmp"
                Format = ImageFormat.Bmp
            Case Else
                Throw New NotSupportedException($"The '{Extension}' image format is not supported.")
        End Select
        PbxPicture.Image.Save(Path, Format)
    End Sub
    ''' <summary>
    ''' Handles the Include button click.
    ''' </summary>
    Private Sub BtnInclude_Click(sender As Object, e As EventArgs) Handles BtnInclude.Click
        Using Dialog As New OpenFileDialog With {.Filter = "Image files|*.jpg;*.jpeg;*.png;*.bmp|All files|*.*", .Title = "Select images", .Multiselect = True}
            If Dialog.ShowDialog() <> DialogResult.OK Then Return
            Dim PreviousCursor As Cursor = Cursor
            Try
                Cursor = Cursors.WaitCursor
                AddPictures(Dialog.FileNames)
            Finally
                Cursor = PreviousCursor
            End Try
        End Using
    End Sub
    ''' <summary>
    ''' Handles the Save button click.
    ''' </summary>
    Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles BtnSave.Click
        If PbxPicture.Image Is Nothing Then Return
        Using Dialog As New SaveFileDialog With {.Filter = "JPEG image|*.jpg|PNG image|*.png|Bitmap image|*.bmp", .Title = "Save image", .FileName = "Picture"}
            If Dialog.ShowDialog() <> DialogResult.OK Then Return
            Try
                SaveSelectedPicture(Dialog.FileName)
            Catch Ex As Exception
                MessageBox.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub
    ''' <summary>
    ''' Handles the Remove button click.
    ''' </summary>
    Private Sub BtnRemove_Click(sender As Object, e As EventArgs) Handles BtnRemove.Click
        RemovePicture(_SelectedPicture)
    End Sub
    ''' <summary>
    ''' Handles the Previous button click.
    ''' </summary>
    Private Sub BtnPrevious_Click(sender As Object, e As EventArgs) Handles BtnPrevious.Click
        MovePrevious()
    End Sub
    ''' <summary>
    ''' Handles the Next button click.
    ''' </summary>
    Private Sub BtnNext_Click(sender As Object, e As EventArgs) Handles BtnNext.Click
        MoveNext()
    End Sub
    ''' <summary>
    ''' Handles the First button click.
    ''' </summary>
    Private Sub BtnFirst_Click(sender As Object, e As EventArgs) Handles BtnFirst.Click
        MoveFirst()
    End Sub
    ''' <summary>
    ''' Handles the Last button click.
    ''' </summary>
    Private Sub BtnLast_Click(sender As Object, e As EventArgs) Handles BtnLast.Click
        MoveLast()
    End Sub
    ''' <summary>
    ''' Updates the enabled state and visibility of the viewer controls.
    ''' </summary>
    Private Sub RefreshControls()
        Dim PictureCount As Integer = _Pictures.Count
        Dim HasPictures As Boolean = PictureCount > 0
        If HasPictures AndAlso (_SelectedIndex < 0 OrElse _SelectedIndex >= PictureCount) Then
            _SelectedIndex = FindPictureIndex(_SelectedPicture)
        End If
        UpdateCounterText()
        LblCounter.Visible = _ShowCounterBar AndAlso (HasPictures OrElse IsDesignTime())
        BtnSave.Enabled = HasPictures
        BtnRemove.Enabled = HasPictures
        BtnFirst.Enabled = HasPictures AndAlso _SelectedIndex > 0
        BtnPrevious.Enabled = HasPictures AndAlso _SelectedIndex > 0
        BtnNext.Enabled = HasPictures AndAlso _SelectedIndex >= 0 AndAlso _SelectedIndex < PictureCount - 1
        BtnLast.Enabled = HasPictures AndAlso _SelectedIndex >= 0 AndAlso _SelectedIndex < PictureCount - 1
        BtnInclude.Enabled = CanAddPicture()
        If Not HasPictures Then DisposeDisplayedImage()
    End Sub
    ''' <summary>
    ''' Determines whether the control is currently hosted in a design environment.
    ''' </summary>
    ''' <returns><see langword="True"/> if the control is in design mode; otherwise, <see langword="False"/>.</returns>
    Private Function IsDesignTime() As Boolean
        Dim CurrentControl As Control = Me
        While CurrentControl IsNot Nothing
            If CurrentControl.Site IsNot Nothing AndAlso CurrentControl.Site.DesignMode Then Return True
            CurrentControl = CurrentControl.Parent
        End While
        Return LicenseManager.UsageMode = LicenseUsageMode.Designtime
    End Function
    ''' <summary>
    ''' Releases the resources used by the <see cref="PictureViewer"/>.
    ''' </summary>
    ''' <param name="disposing"><see langword="True"/> to release managed resources; otherwise, <see langword="False"/>.</param>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing Then
                DisposeDisplayedImage()
                _Pictures.Clear()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub
End Class