Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms
Imports CoreSuite.Helpers
''' <summary>
''' Provides an asynchronous loading overlay mechanism for WinForms,
''' allowing a child loader form to be displayed while blocking or covering
''' the parent form during long-running operations.
''' </summary>
Public Class AsyncLoader
    Private _ControlsState As Dictionary(Of Control, Boolean)
    Private _ParentBackColor As Color
    Private _OverlayPanel As Panel
    Private _BackColor As Color = Color.White
    Private _BorderRadius As Integer = 0
    Private _IsRunning As Boolean
    Private _AutoClose As Boolean
    Private _MaximizeBoxVisible As Boolean
    Private _MinimizeBoxVisible As Boolean
    Private _BorderStyle As FormBorderStyle
    Private _ControlBox As Boolean
    ''' <summary>
    ''' Indicates whether the parent form should be visually covered.
    ''' </summary>
    Public Property CoverParent As Boolean
    ''' <summary>
    ''' Parent container form that will be blocked during loading.
    ''' </summary>
    Public Property Container As Form

    ''' <summary>
    ''' Loader form displayed on top of the parent container.
    ''' </summary>
    Public Property Loader As Form
    ''' <summary>
    ''' Gets or sets the background color of the overlay panel.
    ''' When <see cref="CoverParent"/> is enabled, the color cannot be transparent.
    ''' </summary>
    Public Property BackColor As Color
        Get
            Return _BackColor
        End Get
        Set(value As Color)
            If CoverParent Then
                If value = Nothing OrElse value = Color.Transparent Then
                    Throw New ArgumentException("The background color cannot be null or transparent. If you want a transparency effect, set the CoverParent property to false.")
                End If
            End If
            _BackColor = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the border radius applied to the loader form.
    ''' </summary>
    Public Property BorderRadius As Integer
        Get
            Return _BorderRadius
        End Get
        Set(value As Integer)
            _BorderRadius = value
            If _BorderRadius < 0 Then _BorderRadius = 0
        End Set
    End Property
    ''' <summary>
    ''' Gets a value indicating whether the loader is active.
    ''' </summary>
    Public ReadOnly Property IsRunning As Boolean
        Get
            Return _IsRunning
        End Get
    End Property
    ''' <summary>
    ''' Initializes a new instance of the <see cref="AsyncLoader"/> class.
    ''' </summary>
    ''' <param name="Container">Parent form to be blocked.</param>
    ''' <param name="Loader">Loader form to be displayed.</param>
    ''' <param name="BorderRadius">Rounded corner radius for the loader.</param>
    ''' <param name="CoverParent">Indicates whether the parent form is visually covered.</param>
    ''' <param name="BackColor">Overlay background color.</param>
    Public Sub New(Container As Form, Loader As Form, BorderRadius As Integer, CoverParent As Boolean, BackColor As Color)
        Me.Container = Container
        Me.Loader = Loader
        Me.CoverParent = CoverParent
        Me.BorderRadius = BorderRadius
        Me.BackColor = BackColor
    End Sub
    ''' <summary>
    ''' Starts the asynchronous loader process.
    ''' </summary>
    ''' <param name="Delay">Optional delay before completing the start process (milliseconds).</param>
    Public Async Function Start(Optional Delay As Integer = 1000) As Task
        _IsRunning = True
        Await Task.Yield()
        If CoverParent Then
            CreateOverlayPanel(Container)
        End If
        DisableParentControls()
        _ControlsState = New Dictionary(Of Control, Boolean)()
        If CoverParent Then
            For Each ctrl As Control In ControlHelper.GetAllControls(Container, False)
                _ControlsState(ctrl) = ctrl.Visible
            Next
            For Each ctrl As Control In ControlHelper.GetAllControls(Container, False)
                ctrl.Visible = False
            Next
        Else
            For Each ctrl As Control In ControlHelper.GetAllControls(Container, False)
                _ControlsState(ctrl) = ctrl.Enabled
            Next
            For Each ctrl As Control In ControlHelper.GetAllControls(Container, False)
                ctrl.Enabled = False
            Next
        End If
        AdjustChildFormSize()
        Loader.Owner = Container
        UpdateChildFormPosition()
        AddHandler Container.Resize, AddressOf OnParentResize
        AddHandler Container.Move, AddressOf OnParentMove
        AddHandler Loader.FormClosed, AddressOf OnChildFormClosed
        AddHandler Loader.Load, AddressOf OnChildFormLoad
        AddHandler Loader.FormClosing, AddressOf OnLoaderFormClosing
        Loader.ShowInTaskbar = False
        Loader.ShowIcon = False
        Loader.FormBorderStyle = FormBorderStyle.None
        ControlHelper.EnableFormDoubleBuffer(Loader, True)
        If (_OverlayPanel IsNot Nothing) Then ControlHelper.EnableControlDoubleBuffer(_OverlayPanel, True)
        Loader.Show()
        Await Task.Yield()
        If Delay > 0 Then Await Task.Delay(Delay)
    End Function
    ''' <summary>
    ''' Stops the loader and restores the parent form state.
    ''' </summary>
    ''' <param name="Delay">Optional delay before closing the loader (milliseconds).</param>
    Public Async Function [Stop](Optional Delay As Integer = 0) As Task
        _IsRunning = False
        If Loader IsNot Nothing Then
            If Delay > 0 Then Await Task.Delay(Delay)
            _AutoClose = True
            Loader.Close()
            Loader.Dispose()
            Loader = Nothing
            _AutoClose = False
        End If
    End Function
    Private Sub OnChildFormLoad(sender As Object, e As EventArgs)
        If BorderRadius > 0 Then
            Dim path As New GraphicsPath()
            path.AddArc(0, 0, BorderRadius, BorderRadius, 180, 90)
            path.AddArc(Loader.Width - BorderRadius, 0, BorderRadius, BorderRadius, 270, 90)
            path.AddArc(Loader.Width - BorderRadius, Loader.Height - BorderRadius, BorderRadius, BorderRadius, 0, 90)
            path.AddArc(0, Loader.Height - BorderRadius, BorderRadius, BorderRadius, 90, 90)
            path.CloseFigure()
            Loader.Region = New Region(path)
        End If
    End Sub

    Private Sub OnLoaderFormClosing(sender As Object, e As FormClosingEventArgs)
        If Not _AutoClose Then
            e.Cancel = True
        End If
    End Sub
    Private Sub DisableParentControls()
        If Container IsNot Nothing Then
            _MaximizeBoxVisible = Container.MaximizeBox
            _MinimizeBoxVisible = Container.MinimizeBox
            _BorderStyle = Container.FormBorderStyle
            _ControlBox = Container.ControlBox
            Container.MaximizeBox = False
            Container.MinimizeBox = False
            Container.ControlBox = False
            If Container.FormBorderStyle = FormBorderStyle.Sizable Or Container.FormBorderStyle = FormBorderStyle.SizableToolWindow Then
                Container.FormBorderStyle = FormBorderStyle.FixedSingle
            End If
        End If
    End Sub
    Private Sub RestoreParentControls()
        If Container IsNot Nothing Then
            Container.MaximizeBox = _MaximizeBoxVisible
            Container.MinimizeBox = _MinimizeBoxVisible
            Container.FormBorderStyle = _BorderStyle
            Container.ControlBox = _ControlBox
        End If
    End Sub
    Private Sub CreateOverlayPanel(Parent As Form)
        If _OverlayPanel IsNot Nothing AndAlso Not _OverlayPanel.IsDisposed Then
            Parent.Controls.Remove(_OverlayPanel)
            _OverlayPanel.Dispose()
        End If
        _OverlayPanel = New Panel With {
            .Dock = DockStyle.Fill,
            .Visible = True
        }
        If CoverParent Then
            _OverlayPanel.BackColor = BackColor
            _ParentBackColor = Parent.BackColor
            Parent.BackColor = BackColor
        Else
            _OverlayPanel.BackColor = Color.Transparent
        End If
        Parent.Controls.Add(_OverlayPanel)
        Parent.Controls.SetChildIndex(_OverlayPanel, 0)
        _OverlayPanel.BringToFront()
    End Sub
    Private Sub AdjustChildFormSize()
        If Container IsNot Nothing AndAlso Loader IsNot Nothing Then
            Dim MaxWidth As Integer = Container.ClientSize.Width - 40
            Dim MaxHeight As Integer = Container.ClientSize.Height - 40
            If Loader.Width > MaxWidth Then
                Loader.Width = MaxWidth
            End If
            If Loader.Height > MaxHeight Then
                Loader.Height = MaxHeight
            End If
            Loader.MinimumSize = New Size(Loader.Width, Loader.Height)
        End If
    End Sub
    Private Sub UpdateChildFormPosition()
        If Container IsNot Nothing AndAlso Loader IsNot Nothing Then
            Loader.StartPosition = FormStartPosition.Manual
            Loader.Location = New Point(
            Container.Left + (Container.Width - Loader.Width) \ 2, Container.Top + (Container.Height - Loader.Height) \ 2 + If(Container.FormBorderStyle <> FormBorderStyle.None, 10, 0))
            If Loader.Location.Y < Container.Top Then
                Loader.Location = New Point(Loader.Location.X, Container.Top)
            End If
        End If
    End Sub
    Private Sub OnParentResize(sender As Object, e As EventArgs)
        If Loader IsNot Nothing AndAlso Loader.Visible Then
            UpdateChildFormPosition()
        End If
    End Sub
    Private Sub OnParentMove(sender As Object, e As EventArgs)
        If Loader IsNot Nothing AndAlso Loader.Visible Then
            UpdateChildFormPosition()
        End If
    End Sub
    Private Sub OnChildFormClosed(sender As Object, e As FormClosedEventArgs)
        If CoverParent Then
            If Container IsNot Nothing AndAlso _ControlsState IsNot Nothing Then
                For Each kvp As KeyValuePair(Of Control, Boolean) In _ControlsState
                    kvp.Key.Visible = kvp.Value
                Next
            End If
            Container.BackColor = _ParentBackColor
        Else
            If Container IsNot Nothing AndAlso _ControlsState IsNot Nothing Then
                For Each kvp As KeyValuePair(Of Control, Boolean) In _ControlsState
                    kvp.Key.Enabled = kvp.Value
                Next
            End If
        End If
        If _OverlayPanel IsNot Nothing Then
            Container.Controls.Remove(_OverlayPanel)
            _OverlayPanel.Dispose()
        End If
        RestoreParentControls()
        RemoveHandler Container.Resize, AddressOf OnParentResize
        RemoveHandler Container.Move, AddressOf OnParentMove
        RemoveHandler Loader.FormClosing, AddressOf OnLoaderFormClosing
    End Sub
End Class