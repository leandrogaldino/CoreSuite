Imports System.ComponentModel
Imports System.Windows.Forms

''' <summary>Theme-aware renderer for MenuStrip, ToolStrip, ContextMenuStrip and StatusStrip.</summary>
''' <remarks>All public settings should be changed on the UI thread.</remarks>
Public Partial Class CToolStripRenderer
    Inherits ToolStripProfessionalRenderer

    Private ReadOnly _ColorTable As ToolStripColorTable
    Private ReadOnly _Targets As New List(Of WeakReference(Of ToolStrip))()
    Private _Palette As ToolStripPalette
    Private _Theme As ToolStripTheme = ToolStripTheme.Light
    Private _CheckedButtonStyle As CheckedButtonStyle = CheckedButtonStyle.Filled
    Private _CornerRadius As Integer = 6
    Private _ShowToolStripBorder As Boolean = False
    Private _ShowDropDownBorder As Boolean = True
    Private _ShowSeparators As Boolean = True
    Private _ShowImageMargin As Boolean = True
    Private _ShowGrip As Boolean = True
    Private _UseGradients As Boolean
    Private _UseSystemColors As Boolean
    Private _RespectHighContrast As Boolean = True

    Public Event ThemeChanged As EventHandler
    Public Event AppearanceChanged As EventHandler

    Public Sub New()
        Me.New(ToolStripTheme.Light)
    End Sub

    Public Sub New(Theme As ToolStripTheme)
        MyBase.New(New ToolStripColorTable())
        _ColorTable = DirectCast(MyBase.ColorTable, ToolStripColorTable)
        MyBase.RoundedEdges = False
        SetThemeCore(Theme)
    End Sub

    <Category("Appearance"), Description("Predefined theme. Assigning a custom Palette switches the theme to Custom.")>
    Public Property Theme As ToolStripTheme
        Get
            Return _Theme
        End Get
        Set(value As ToolStripTheme)
            If value = ToolStripTheme.Custom Then
                If _Theme <> ToolStripTheme.Custom Then
                    _Theme = ToolStripTheme.Custom
                    RaiseEvent ThemeChanged(Me, EventArgs.Empty)
                    RefreshAttached()
                End If
                Return
            End If
            If _Theme = value Then Return
            SetThemeCore(value)
            RaiseEvent ThemeChanged(Me, EventArgs.Empty)
            RefreshAttached()
        End Set
    End Property

    Private Sub SetThemeCore(Theme As ToolStripTheme)
        If _Palette IsNot Nothing Then RemoveHandler _Palette.Changed, AddressOf OnPaletteChanged
        _Theme = Theme
        _Palette = ToolStripPalette.Create(Theme)
        AddHandler _Palette.Changed, AddressOf OnPaletteChanged
        _ColorTable.SetPalette(_Palette)
    End Sub

    <Category("Appearance"), Description("Editable theme colors; color changes invalidate registered ToolStrips.")>
    Public Property Palette As ToolStripPalette
        Get
            Return _Palette
        End Get
        Set(value As ToolStripPalette)
            If value Is Nothing Then Throw New ArgumentNullException(NameOf(value))
            If Object.ReferenceEquals(_Palette, value) Then Return
            RemoveHandler _Palette.Changed, AddressOf OnPaletteChanged
            _Palette = value
            AddHandler _Palette.Changed, AddressOf OnPaletteChanged
            _ColorTable.SetPalette(value)
            _Theme = ToolStripTheme.Custom
            RaiseEvent ThemeChanged(Me, EventArgs.Empty)
            RefreshAttached()
        End Set
    End Property

    <Category("Appearance"), DefaultValue(GetType(CheckedButtonStyle), "Filled")>
    Public Property CheckedButtonStyle As CheckedButtonStyle
        Get
            Return _CheckedButtonStyle
        End Get
        Set(value As CheckedButtonStyle)
            If _CheckedButtonStyle = value Then Return
            _CheckedButtonStyle = value
            RefreshAttached()
        End Set
    End Property

    ''' <summary>Compatibility switch for the original checked-background behavior.</summary>
    <Category("Appearance"), Description("True draws a filled checked state; False hides it. Use CheckedButtonStyle for other styles.")>
    Public Property ShowButtonCheckedRectangle As Boolean
        Get
            Return _CheckedButtonStyle <> CheckedButtonStyle.None
        End Get
        Set(value As Boolean)
            CheckedButtonStyle = If(value, CheckedButtonStyle.Filled, CheckedButtonStyle.None)
        End Set
    End Property

    <Category("Appearance"), DefaultValue(6)>
    Public Property CornerRadius As Integer
        Get
            Return _CornerRadius
        End Get
        Set(value As Integer)
            If value < 0 OrElse value > 32 Then Throw New ArgumentOutOfRangeException(NameOf(value), "Range: 0–32.")
            If _CornerRadius = value Then Return
            _CornerRadius = value
            RefreshAttached()
        End Set
    End Property

    <Category("Borders"), DefaultValue(False)>
    Public Property ShowToolStripBorder As Boolean
        Get
            Return _ShowToolStripBorder
        End Get
        Set(value As Boolean)
            If _ShowToolStripBorder = value Then Return
            _ShowToolStripBorder = value
            RefreshAttached()
        End Set
    End Property

    <Category("Borders"), DefaultValue(True)>
    Public Property ShowDropDownBorder As Boolean
        Get
            Return _ShowDropDownBorder
        End Get
        Set(value As Boolean)
            If _ShowDropDownBorder = value Then Return
            _ShowDropDownBorder = value
            RefreshAttached()
        End Set
    End Property

    <Category("Appearance"), DefaultValue(True)>
    Public Property ShowSeparators As Boolean
        Get
            Return _ShowSeparators
        End Get
        Set(value As Boolean)
            If _ShowSeparators = value Then Return
            _ShowSeparators = value
            RefreshAttached()
        End Set
    End Property

    <Category("Appearance"), DefaultValue(True)>
    Public Property ShowImageMargin As Boolean
        Get
            Return _ShowImageMargin
        End Get
        Set(value As Boolean)
            If _ShowImageMargin = value Then Return
            _ShowImageMargin = value
            RefreshAttached()
        End Set
    End Property

    <Category("Appearance"), DefaultValue(True)>
    Public Property ShowGrip As Boolean
        Get
            Return _ShowGrip
        End Get
        Set(value As Boolean)
            If _ShowGrip = value Then Return
            _ShowGrip = value
            RefreshAttached()
        End Set
    End Property

    <Category("Appearance"), DefaultValue(False)>
    Public Property UseGradients As Boolean
        Get
            Return _UseGradients
        End Get
        Set(value As Boolean)
            If _UseGradients = value Then Return
            _UseGradients = value
            RefreshAttached()
        End Set
    End Property

    <Category("Accessibility"), DefaultValue(False)>
    Public Property UseSystemColors As Boolean
        Get
            Return _UseSystemColors
        End Get
        Set(value As Boolean)
            If _UseSystemColors = value Then Return
            _UseSystemColors = value
            RefreshAttached()
        End Set
    End Property

    <Category("Accessibility"), DefaultValue(True)>
    Public Property RespectHighContrast As Boolean
        Get
            Return _RespectHighContrast
        End Get
        Set(value As Boolean)
            If _RespectHighContrast = value Then Return
            _RespectHighContrast = value
            RefreshAttached()
        End Set
    End Property

    Private Sub OnPaletteChanged(sender As Object, e As EventArgs)
        RefreshAttached()
    End Sub

    Private Function ShouldUseGradients() As Boolean
        Return _UseGradients AndAlso Not (_UseSystemColors OrElse _Theme = ToolStripTheme.System OrElse
               (_RespectHighContrast AndAlso SystemInformation.HighContrast))
    End Function

    Private Function IsSystemPaletteActive() As Boolean
        Return _UseSystemColors OrElse _Theme = ToolStripTheme.System OrElse
               (_RespectHighContrast AndAlso SystemInformation.HighContrast)
    End Function

    Private Function EffectivePalette() As ToolStripPalette
        If IsSystemPaletteActive() Then
            Return ToolStripPalette.Create(ToolStripTheme.System)
        End If
        Return _Palette
    End Function

    ''' <summary>Initializes ToolStrips even when the renderer is assigned directly through Renderer.</summary>
    Protected Overrides Sub Initialize(Strip As ToolStrip)
        MyBase.Initialize(Strip)
        If Strip Is Nothing OrElse Strip.IsDisposed Then Return
        TrackStrip(Strip)
        ApplyDropDownSettings(Strip)
        For Each Item As ToolStripItem In Strip.Items
            HookDropDownOpening(Item)
        Next
    End Sub

    ''' <summary>Tracks dropdown items added after the owning ToolStrip was initialized.</summary>
    Protected Overrides Sub InitializeItem(Item As ToolStripItem)
        MyBase.InitializeItem(Item)
        HookDropDownOpening(Item)
    End Sub

    Private Sub HookDropDownOpening(Item As ToolStripItem)
        Dim DropDownItem As ToolStripDropDownItem = TryCast(Item, ToolStripDropDownItem)
        If DropDownItem Is Nothing Then Return
        RemoveHandler DropDownItem.DropDownOpening, AddressOf OnDropDownOpening
        AddHandler DropDownItem.DropDownOpening, AddressOf OnDropDownOpening
    End Sub

    Private Sub OnDropDownOpening(Sender As Object, e As EventArgs)
        Dim Item As ToolStripDropDownItem = TryCast(Sender, ToolStripDropDownItem)
        If Item Is Nothing OrElse Item.Owner Is Nothing Then Return
        If Not Object.ReferenceEquals(Item.Owner.Renderer, Me) Then Return

        ' The dropdown may be created lazily; apply layout options before it opens.
        RegisterStrip(Item.DropDown)
    End Sub

    Private Sub ApplyDropDownSettings(Strip As ToolStrip)
        Dim Menu As ToolStripDropDownMenu = TryCast(Strip, ToolStripDropDownMenu)
        If Menu IsNot Nothing AndAlso Menu.ShowImageMargin <> _ShowImageMargin Then
            ' Changing the control property triggers a layout pass; skipping painting does not.
            Menu.ShowImageMargin = _ShowImageMargin
        End If
    End Sub

    Private Sub TrackStrip(Strip As ToolStrip)
        For i As Integer = _Targets.Count - 1 To 0 Step -1
            Dim Candidate As ToolStrip = Nothing
            If Not _Targets(i).TryGetTarget(Candidate) OrElse Candidate Is Nothing OrElse Candidate.IsDisposed Then
                _Targets.RemoveAt(i)
            ElseIf Object.ReferenceEquals(Candidate, Strip) Then
                Return
            End If
        Next
        _Targets.Add(New WeakReference(Of ToolStrip)(Strip))
    End Sub

    ''' <summary>Assigns this renderer to controls and all currently created nested dropdowns.</summary>
    Public Sub ApplyTo(ParamArray Strips As ToolStrip())
        If Strips Is Nothing Then Throw New ArgumentNullException(NameOf(Strips))
        For Each Strip As ToolStrip In Strips
            If Strip Is Nothing Then Continue For
            RegisterStrip(Strip)
        Next
        RefreshAttached()
    End Sub

    Private Sub RegisterStrip(Strip As ToolStrip)
        If Strip Is Nothing OrElse Strip.IsDisposed Then Return
        Strip.Renderer = Me
        TrackStrip(Strip)
        ApplyDropDownSettings(Strip)

        For Each Item As ToolStripItem In Strip.Items
            HookDropDownOpening(Item)
            Dim DropDownItem As ToolStripDropDownItem = TryCast(Item, ToolStripDropDownItem)
            If DropDownItem IsNot Nothing AndAlso DropDownItem.HasDropDownItems Then
                RegisterStrip(DropDownItem.DropDown)
            End If
        Next
    End Sub

    ''' <summary>Repaints registered controls and updates menu layout options.</summary>
    Public Sub RefreshAttached()
        For i As Integer = _Targets.Count - 1 To 0 Step -1
            Dim Strip As ToolStrip = Nothing
            If Not _Targets(i).TryGetTarget(Strip) OrElse Strip Is Nothing OrElse Strip.IsDisposed Then
                _Targets.RemoveAt(i)
            ElseIf Object.ReferenceEquals(Strip.Renderer, Me) Then
                ApplyDropDownSettings(Strip)
                For Each Item As ToolStripItem In Strip.Items
                    HookDropDownOpening(Item)
                Next
                Strip.Invalidate(True)
                For Each Item As ToolStripItem In Strip.Items
                    Item.Invalidate()
                Next
            End If
        Next
        RaiseEvent AppearanceChanged(Me, EventArgs.Empty)
    End Sub
End Class
