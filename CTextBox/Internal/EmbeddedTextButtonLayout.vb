Imports System.ComponentModel
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

''' <summary>
''' Coordinates the child button and the extra non-client strip used by the
''' centered single-line text frame. The child retains its own click event.
''' </summary>
<EditorBrowsable(EditorBrowsableState.Never)>
Public NotInheritable Class EmbeddedTextButtonLayout
    Implements IDisposable

    Private ReadOnly _Owner As System.Windows.Forms.TextBoxBase
    Private ReadOnly _Button As Control
    Private ReadOnly _Render As Action(Of Graphics, Size)
    Private ReadOnly _Click As Action
    Private _TopExtension As Integer
    Private _Shown As Boolean
    Private _Hot As Boolean
    Private _Pressed As Boolean
    Private _ExtensionPressed As Boolean
    Private _Disposed As Boolean
    Private _Updating As Boolean
    Private _Painting As Boolean

    Public Sub New(Owner As System.Windows.Forms.TextBoxBase, Button As Control, Render As Action(Of Graphics, Size), Click As Action)
        _Owner = Owner
        _Button = Button
        _Render = Render
        _Click = Click
        AddHandler _Button.Paint, AddressOf PaintChild
        AddHandler _Button.Invalidated, AddressOf ButtonInvalidated
        AddHandler _Button.MouseEnter, AddressOf ButtonEnter
        AddHandler _Button.MouseLeave, AddressOf ButtonLeave
        AddHandler _Button.MouseDown, AddressOf ButtonDown
        AddHandler _Button.MouseUp, AddressOf ButtonUp
        AddHandler _Button.MouseCaptureChanged, AddressOf ButtonCaptureChanged
        AddHandler _Owner.Disposed, AddressOf OwnerDisposed
    End Sub

    ''' <summary>Fills the client area vertically; returns the actual reserved width in physical pixels.</summary>
    Public Function UpdateLayout(LogicalWidth As Integer, LeftSide As Boolean, Visible As Boolean) As Integer
        If _Disposed OrElse _Updating OrElse _Owner.IsDisposed OrElse _Owner.Disposing Then Return 0
        _Updating = True
        Try
            Dim ClientWidth As Integer = _Owner.ClientSize.Width
            Dim ClientHeight As Integer = _Owner.ClientSize.Height
            _TopExtension = 0
            If _Owner.IsHandleCreated Then
                Dim Info As New NativeWindowInfo With {.Size = CUInt(Marshal.SizeOf(GetType(NativeWindowInfo)))}
                If GetWindowInfo(_Owner.Handle, Info) Then
                    ClientWidth = Math.Max(0, Info.ClientBounds.Right - Info.ClientBounds.Left)
                    ClientHeight = Math.Max(0, Info.ClientBounds.Bottom - Info.ClientBounds.Top)
                    If Not _Owner.Multiline AndAlso _Owner.BorderStyle = BorderStyle.FixedSingle Then
                        _TopExtension = Math.Max(0, Info.ClientBounds.Top - Info.WindowBounds.Top - 1)
                    End If
                End If
            End If
            Dim Width As Integer = Math.Max(0, Math.Min(ClientWidth, CInt(Math.Round(LogicalWidth * _Owner.DeviceDpi / 96.0R, MidpointRounding.AwayFromZero))))
            _Shown = Visible AndAlso Width > 0 AndAlso ClientHeight > 0
            _Button.Bounds = New Rectangle(If(LeftSide, 0, Math.Max(0, ClientWidth - Width)), 0, Width, ClientHeight)
            _Button.Visible = _Shown
            If Not _Shown Then
                If _ExtensionPressed AndAlso _Owner.Capture Then _Owner.Capture = False
                _Hot = False
                _Pressed = False
                _ExtensionPressed = False
            End If
            _Button.BringToFront()
            _Button.Invalidate()
            If _Owner.IsHandleCreated Then RedrawWindow(_Owner.Handle, IntPtr.Zero, IntPtr.Zero, &H401UI)
            Return Width
        Finally
            _Updating = False
            PaintExtension()
        End Try
    End Function

    Public Function BeforeWndProc(ByRef M As Message) As Boolean
        If _Disposed Then Return False
        Select Case M.Msg
            Case &H20 ' WM_SETCURSOR
                If IsOver(True) AndAlso _Button.Enabled Then
                    Cursor.Current = _Button.Cursor
                    M.Result = New IntPtr(1)
                    Return True
                End If
            Case &H201, &H203 ' WM_LBUTTONDOWN / DBLCLK
                If IsOver(True) Then
                    If _Button.Enabled Then
                        _ExtensionPressed = True
                        _Hot = True
                        _Pressed = True
                        _Owner.Capture = True
                        RefreshButton()
                    End If
                    M.Result = IntPtr.Zero
                    Return True
                End If
            Case &H202 ' WM_LBUTTONUP
                If _ExtensionPressed Then
                    Dim InvokeClick As Boolean = IsOver(False) AndAlso _Button.Enabled
                    _ExtensionPressed = False
                    _Pressed = False
                    _Owner.Capture = False
                    _Hot = IsOver(False)
                    RefreshButton()
                    If InvokeClick Then _Click()
                    M.Result = IntPtr.Zero
                    Return True
                End If
            Case &H200 ' WM_MOUSEMOVE
                If _ExtensionPressed Then
                    _Hot = IsOver(False)
                    _Pressed = _Hot
                    RefreshButton()
                    M.Result = IntPtr.Zero
                    Return True
                End If
            Case &H215 ' WM_CAPTURECHANGED
                If _ExtensionPressed Then
                    _ExtensionPressed = False
                    _Pressed = False
                    RefreshButton()
                End If
        End Select
        Return False
    End Function

    Public Sub AfterWndProc(ByRef M As Message)
        If _Disposed Then Return
        Select Case M.Msg
            Case &HF, &H85, &H86, &H7, &H8
                PaintExtension()
            Case &H200, &H2A3 ' WM_MOUSEMOVE / WM_MOUSELEAVE
                If Not _ExtensionPressed Then
                    Dim Hot As Boolean = IsOver(False) AndAlso _Button.Enabled
                    If _Hot <> Hot Then
                        _Hot = Hot
                        RefreshButton()
                    End If
                End If
        End Select
    End Sub

    Private Function HasExtension() As Boolean
        Return Not _Disposed AndAlso _TopExtension > 0 AndAlso _Shown AndAlso _Button.Visible AndAlso
            _Owner.IsHandleCreated AndAlso Not _Owner.IsDisposed AndAlso Not _Owner.Disposing
    End Function

    Private Function IsOver(ExtensionOnly As Boolean) As Boolean
        If Not HasExtension() Then Return False
        Dim P As Point = _Owner.PointToClient(Control.MousePosition)
        Return New Rectangle(_Button.Left, -_TopExtension, _Button.Width, If(ExtensionOnly, _TopExtension, _Button.Height + _TopExtension)).Contains(P)
    End Function

    Private Sub PaintChild(Sender As Object, E As PaintEventArgs)
        If _Disposed Then Return
        RenderSurface(E.Graphics, New Rectangle(0, -_TopExtension, _Button.Width, _Button.Height + _TopExtension))
    End Sub

    Private Sub RenderSurface(G As Graphics, Bounds As Rectangle)
        If Bounds.Width <= 0 OrElse Bounds.Height <= 0 Then Return
        Dim State = G.Save()
        Try
            G.TranslateTransform(Bounds.X, Bounds.Y)
            Dim Background As Color = _Button.BackColor
            If Not _Button.Enabled Then
                Background = SystemColors.Control
            ElseIf _Pressed Then
                Background = SystemColors.ControlDark
            ElseIf _Hot Then
                Background = SystemColors.ControlLight
            End If
            Using Brush As New SolidBrush(Background)
                G.FillRectangle(Brush, 0, 0, Bounds.Width, Bounds.Height)
            End Using
            _Render(G, Bounds.Size)
        Finally
            G.Restore(State)
        End Try
    End Sub

    Public Shared Sub DrawImage(G As Graphics, Image As Image, Bounds As Rectangle, Enabled As Boolean, Background As Color)
        If Image Is Nothing OrElse Bounds.Width <= 0 OrElse Bounds.Height <= 0 Then Return
        If Enabled Then
            G.DrawImage(Image, Bounds)
        Else
            Using Scaled As New Bitmap(Image, Bounds.Size)
                ControlPaint.DrawImageDisabled(G, Scaled, Bounds.X, Bounds.Y, Background)
            End Using
        End If
    End Sub

    Public Shared Sub DrawCenteredImage(G As Graphics, Image As Image, SurfaceSize As Size, Padding As Integer, Enabled As Boolean, Background As Color)
        If Image Is Nothing OrElse Image.Width <= 0 OrElse Image.Height <= 0 Then Return
        Dim W As Integer = SurfaceSize.Width - Padding * 2
        Dim H As Integer = SurfaceSize.Height - Padding * 2
        If W <= 0 OrElse H <= 0 Then Return
        Dim Scale As Double = Math.Min(W / CDbl(Image.Width), H / CDbl(Image.Height))
        W = Math.Max(1, CInt(Math.Round(Image.Width * Scale)))
        H = Math.Max(1, CInt(Math.Round(Image.Height * Scale)))
        DrawImage(G, Image, New Rectangle((SurfaceSize.Width - W) \ 2, (SurfaceSize.Height - H) \ 2, W, H), Enabled, Background)
    End Sub

    Private Sub PaintExtension()
        If _Updating OrElse _Painting OrElse Not HasExtension() Then Return
        Dim Info As New NativeWindowInfo With {.Size = CUInt(Marshal.SizeOf(GetType(NativeWindowInfo)))}
        If Not GetWindowInfo(_Owner.Handle, Info) Then Return
        Dim X As Integer = Info.ClientBounds.Left - Info.WindowBounds.Left + _Button.Left
        Dim Y As Integer = Info.ClientBounds.Top - Info.WindowBounds.Top - _TopExtension
        If Y < 1 OrElse _Button.Width <= 0 Then Return
        Dim DC As IntPtr = GetWindowDC(_Owner.Handle)
        If DC = IntPtr.Zero Then Return
        _Painting = True
        Try
            Using G As Graphics = Graphics.FromHdc(DC)
                G.SetClip(New Rectangle(X, Y, _Button.Width, _TopExtension))
                RenderSurface(G, New Rectangle(X, Y, _Button.Width, _Button.Height + _TopExtension))
            End Using
        Finally
            ReleaseDC(_Owner.Handle, DC)
            _Painting = False
        End Try
    End Sub

    Private Sub RefreshButton()
        If _Disposed OrElse _Button.IsDisposed Then Return
        _Button.Invalidate()
        PaintExtension()
    End Sub

    Private Sub ButtonInvalidated(Sender As Object, E As InvalidateEventArgs)
        PaintExtension()
    End Sub
    Private Sub ButtonEnter(Sender As Object, E As EventArgs)
        _Hot = True
        RefreshButton()
    End Sub
    Private Sub ButtonLeave(Sender As Object, E As EventArgs)
        _Hot = IsOver(True)
        _Pressed = False
        RefreshButton()
    End Sub
    Private Sub ButtonDown(Sender As Object, E As MouseEventArgs)
        If E.Button = MouseButtons.Left Then
            _Pressed = True
            RefreshButton()
        End If
    End Sub
    Private Sub ButtonUp(Sender As Object, E As MouseEventArgs)
        _Pressed = False
        RefreshButton()
    End Sub
    Private Sub ButtonCaptureChanged(Sender As Object, E As EventArgs)
        If Not _Button.Capture Then
            _Pressed = False
            RefreshButton()
        End If
    End Sub
    Private Sub OwnerDisposed(Sender As Object, E As EventArgs)
        Dispose()
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        If _Disposed Then Return
        _Disposed = True
        RemoveHandler _Button.Paint, AddressOf PaintChild
        RemoveHandler _Button.Invalidated, AddressOf ButtonInvalidated
        RemoveHandler _Button.MouseEnter, AddressOf ButtonEnter
        RemoveHandler _Button.MouseLeave, AddressOf ButtonLeave
        RemoveHandler _Button.MouseDown, AddressOf ButtonDown
        RemoveHandler _Button.MouseUp, AddressOf ButtonUp
        RemoveHandler _Button.MouseCaptureChanged, AddressOf ButtonCaptureChanged
        RemoveHandler _Owner.Disposed, AddressOf OwnerDisposed
    End Sub

    <StructLayout(LayoutKind.Sequential)>
    Private Structure NativeRect
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure
    <StructLayout(LayoutKind.Sequential)>
    Private Structure NativeWindowInfo
        Public Size As UInteger
        Public WindowBounds As NativeRect
        Public ClientBounds As NativeRect
        Public Style As UInteger
        Public ExStyle As UInteger
        Public WindowStatus As UInteger
        Public BorderWidth As UInteger
        Public BorderHeight As UInteger
        Public WindowType As UShort
        Public CreatorVersion As UShort
    End Structure
    <DllImport("user32.dll")>
    Private Shared Function GetWindowInfo(WindowHandle As IntPtr, ByRef Info As NativeWindowInfo) As Boolean
    End Function
    <DllImport("user32.dll")>
    Private Shared Function GetWindowDC(WindowHandle As IntPtr) As IntPtr
    End Function
    <DllImport("user32.dll")>
    Private Shared Function ReleaseDC(WindowHandle As IntPtr, DC As IntPtr) As Integer
    End Function
    <DllImport("user32.dll")>
    Private Shared Function RedrawWindow(WindowHandle As IntPtr, UpdateRect As IntPtr, UpdateRegion As IntPtr, Flags As UInteger) As Boolean
    End Function
End Class
