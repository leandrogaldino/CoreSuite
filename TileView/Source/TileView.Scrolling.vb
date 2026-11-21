Imports System.Windows.Forms

Public Partial Class TileView
    Private Const WheelDelta As Integer = 120
    Private ReadOnly _VerticalBar As New VScrollBar()
    Private _ScrollOffset As Integer
    Private _SettingScrollBar As Boolean
    Private _WheelRemainder As Integer
    Private _ThumbTracking As Boolean

    Private Sub InitializeVirtualScrollBar()
        _VerticalBar.Name = "VirtualVerticalScrollBar"
        _VerticalBar.Dock = DockStyle.Right
        _VerticalBar.Visible = False
        _VerticalBar.TabStop = False
        _VerticalBar.SmallChange = 1
        AddHandler _VerticalBar.ValueChanged, AddressOf VirtualScrollBar_ValueChanged
        AddHandler _VerticalBar.Scroll, AddressOf VirtualScrollBar_Scroll
        AddHandler _VerticalBar.MouseCaptureChanged, AddressOf VirtualScrollBar_MouseCaptureChanged
        Controls.Add(_VerticalBar)
    End Sub

    ''' <summary>
    ''' Updates scrollbar range and viewport without moving child HWNDs. The effective
    ''' max Value for WinForms scrollbars is Maximum - LargeChange + 1.
    ''' </summary>
    Private Sub ConfigureVirtualScrollBar(ContentHeight As Integer)
        Dim PageHeight As Integer = Math.Max(1, ClientSize.Height)
        Dim ScrollRequired As Boolean = ContentHeight > ClientSize.Height
        Dim Maximum As Integer = Math.Max(PageHeight, ContentHeight) - 1
        Dim MaxOffset As Integer = Math.Max(0, ContentHeight - ClientSize.Height)

        _SettingScrollBar = True
        Try
            _VerticalBar.Minimum = 0
            _VerticalBar.Maximum = Maximum
            _VerticalBar.LargeChange = PageHeight
            _VerticalBar.SmallChange = Math.Max(1, _ItemSize.Height + _ItemSpacing)
            _ScrollOffset = Math.Max(0, Math.Min(_ScrollOffset, MaxOffset))
            If _VerticalBar.Value <> _ScrollOffset Then _VerticalBar.Value = _ScrollOffset
            If _VerticalBar.Visible <> ScrollRequired Then _VerticalBar.Visible = ScrollRequired
        Finally
            _SettingScrollBar = False
        End Try
    End Sub

    Private Sub VirtualScrollBar_ValueChanged(sender As Object, e As EventArgs)
        If _SettingScrollBar OrElse IsDisposed OrElse Disposing Then Return
        Dim MaxOffset As Integer = Math.Max(0, _ContentHeight - ClientSize.Height)
        Dim Target As Integer = Math.Max(0, Math.Min(_VerticalBar.Value, MaxOffset))
        If Target <> _VerticalBar.Value Then
            _SettingScrollBar = True
            Try
                _VerticalBar.Value = Target
            Finally
                _SettingScrollBar = False
            End Try
        End If
        If _ScrollOffset = Target Then Return
        _ScrollOffset = Target
        OnUserScrolled()
    End Sub

    Private Sub VirtualScrollBar_Scroll(sender As Object, e As ScrollEventArgs)
        Select Case e.Type
            Case ScrollEventType.ThumbTrack
                _ThumbTracking = True
            Case ScrollEventType.EndScroll
                _ThumbTracking = False
                FinishScrollPreview()
        End Select
    End Sub

    Private Sub VirtualScrollBar_MouseCaptureChanged(sender As Object, e As EventArgs)
        If Not _VerticalBar.Capture AndAlso _ThumbTracking Then
            _ThumbTracking = False
            FinishScrollPreview()
        End If
    End Sub

    ''' <summary>Scrolls to an absolute pixel offset; clamps to the document's bounds.</summary>
    Public Sub ScrollToOffset(VerticalOffset As Integer)
        FinishScrollPreview()
        If _layoutDirty Then UpdateViewport()
        Dim Target As Integer = Math.Max(0, Math.Min(VerticalOffset, Math.Max(0, _ContentHeight - ClientSize.Height)))
        If Target = _ScrollOffset Then Return
        _SettingScrollBar = True
        Try
            _ScrollOffset = Target
            _VerticalBar.Value = Target
        Finally
            _SettingScrollBar = False
        End Try
        UpdateViewport()
    End Sub

    Private Sub ResetVirtualScroll()
        CancelScrollPreview()
        _ScrollOffset = 0
        _WheelRemainder = 0
        If _VerticalBar Is Nothing Then Return
        _SettingScrollBar = True
        Try
            If _VerticalBar.Value <> 0 Then _VerticalBar.Value = 0
        Finally
            _SettingScrollBar = False
        End Try
    End Sub

    ''' <summary>Uses WinForms wheel settings, keeping high-resolution wheel deltas.</summary>
    Private Sub ScrollByWheelDelta(Delta As Integer)
        If _VerticalBar Is Nothing OrElse Not _VerticalBar.Visible Then Return
        Dim Lines As Integer = SystemInformation.MouseWheelScrollLines
        If Lines = 0 Then Return
        _WheelRemainder += Delta
        Dim Ticks As Integer = _WheelRemainder \ WheelDelta
        _WheelRemainder -= Ticks * WheelDelta
        If Ticks = 0 Then Return

        Dim StepPixels As Long
        If Lines < 0 Then
            StepPixels = Math.Max(1, ClientSize.Height - _ItemSize.Height)
        Else
            StepPixels = CLng(Lines) * Math.Max(1, (_ItemSize.Height + _ItemSpacing) \ 3)
        End If
        Dim NewOffset As Long = CLng(_ScrollOffset) - CLng(Ticks) * StepPixels
        Dim Maximum As Long = Math.Max(0L, CLng(_ContentHeight) - ClientSize.Height)
        UserScrollToOffset(CInt(Math.Max(0L, Math.Min(NewOffset, Maximum))))
    End Sub

    Friend Sub ScrollFromItem(Delta As Integer)
        ScrollByWheelDelta(Delta)
    End Sub

    Private Sub UserScrollToOffset(Target As Integer)
        If _layoutDirty Then UpdateViewport()
        Target = Math.Max(0, Math.Min(Target, Math.Max(0, _ContentHeight - ClientSize.Height)))
        If Target = _ScrollOffset Then Return
        _ScrollOffset = Target
        _SettingScrollBar = True
        Try
            _VerticalBar.Value = Target
        Finally
            _SettingScrollBar = False
        End Try
        OnUserScrolled()
    End Sub

    Private Sub OnUserScrolled()
        If _SmoothScrolling Then
            BeginScrollPreview()
        Else
            UpdateViewport()
        End If
    End Sub
End Class
