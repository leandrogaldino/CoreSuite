Imports System.ComponentModel
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

''' <summary>
''' Provides native text editing with placeholder support and improved text spacing.
''' </summary>
Public Class CTextBox
    Inherits TextBox

    Private Const EM_SETCUEBANNER As Integer = &H1501
    Private Const WS_BORDER As Integer = &H800000
    Private Const WS_EX_CLIENTEDGE As Integer = &H200
    Private Const WM_SETFOCUS As Integer = &H7
    Private Const WM_KILLFOCUS As Integer = &H8
    Private Const WM_NCACTIVATE As Integer = &H86
    Private Const EM_GETRECT As Integer = &HB2
    Private Const EM_SETRECT As Integer = &HB3
    Private Const EM_SETRECTNP As Integer = &HB4
    Private Const WM_NCCALCSIZE As Integer = &H83
    Private Const WM_NCHITTEST As Integer = &H84
    Private Const HTCLIENT As Integer = 1
    Private Const SWP_FRAMECHANGED As UInteger = &H20UI
    Private Const SWP_NOCHANGE As UInteger = &H17UI ' NOSIZE, NOMOVE, NOZORDER, NOACTIVATE
    Private Const EM_SETMARGINS As Integer = &HD3
    Private Const EC_BOTHMARGINS As Integer = 3
    Private Const WM_NCPAINT As Integer = &H85
    Private Const WM_PAINT As Integer = &HF
    Private Const RDW_FRAME As UInteger = &H400UI
    Private Const RDW_INVALIDATE As UInteger = &H1UI

    Private _PlaceholderText As String = String.Empty
    Private _HorizontalPadding As Integer = 5
    Private _VerticalPadding As Integer
    Private _MultilinePaddingApplied As Boolean
    Private _ReservedLeft As Integer
    Private _ReservedRight As Integer
    Private _RequestedTopInset As Integer
    Private _AppliedTopInset As Integer
    Private _UpdatingTextLayout As Boolean
    Private _BorderColor As Color = Color.Empty
    Private _FocusedBorderColor As Color = SystemColors.Highlight

    ' A native WS_BORDER is drawn by EDIT itself, sometimes outside WM_NCPAINT.
    ' For the single-line FixedSingle case, reserve its space ourselves so there
    ' is exactly one owner of the outline. Keep the managed BorderStyle intact.
    Private ReadOnly Property OwnsSingleLineFrame As Boolean
        Get
            Return Not Multiline AndAlso BorderStyle = BorderStyle.FixedSingle
        End Get
    End Property

    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim Params As CreateParams = MyBase.CreateParams
            If OwnsSingleLineFrame Then
                Params.Style = Params.Style And Not WS_BORDER
                Params.ExStyle = Params.ExStyle And Not WS_EX_CLIENTEDGE
            End If
            Return Params
        End Get
    End Property

    ''' <summary>Gets or sets the native cue banner displayed when the control is empty and unfocused.</summary>
    <DefaultValue(""), Localizable(True), Category("Appearance"), Description("Text shown when the control is empty and does not have focus.")>
    Public Overridable Property PlaceholderText As String
        Get
            Return _PlaceholderText
        End Get
        Set(Value As String)
            _PlaceholderText = If(Value, String.Empty)
            UpdatePlaceholder()
        End Set
    End Property

    ''' <summary>Gets or sets the text inset on each side, in logical pixels at 96 DPI.</summary>
    <DefaultValue(5), Category("Layout"), Description("Space between the text and each side of the edit area, in logical pixels (96 DPI).")>
    Public Property HorizontalPadding As Integer
        Get
            Return _HorizontalPadding
        End Get
        Set(Value As Integer)
            If Value < 0 OrElse Value > 1024 Then Throw New ArgumentOutOfRangeException(NameOf(Value))
            If _HorizontalPadding = Value Then Return
            _HorizontalPadding = Value
            ApplyTextMargins()
        End Set
    End Property

    ''' <summary>Gets or sets additional top and bottom text padding in multiline mode, in logical pixels at 96 DPI.</summary>
    <DefaultValue(0), Category("Layout"), Description("Additional space above and below the text in multiline mode, in logical pixels (96 DPI).")>
    Public Property VerticalPadding As Integer
        Get
            Return _VerticalPadding
        End Get
        Set(Value As Integer)
            If Value < 0 OrElse Value > 1024 Then Throw New ArgumentOutOfRangeException(NameOf(Value))
            If _VerticalPadding = Value Then Return
            _VerticalPadding = Value
            If Multiline Then UpdateTextLayout()
        End Set
    End Property

    ''' <summary>Gets or sets the normal outer border color. Empty uses the default border color.</summary>
    <DefaultValue(GetType(Color), ""), Category("Appearance"), Description("Normal border color (full outline for FixedSingle, bottom edge for Fixed3D). Empty uses the system border color.")>
    Public Property BorderColor As Color
        Get
            Return _BorderColor
        End Get
        Set(Value As Color)
            If _BorderColor = Value Then Return
            _BorderColor = Value
            InvalidateBorder()
        End Set
    End Property

    ''' <summary>Gets or sets the outer border color when the text box has keyboard focus.</summary>
    <DefaultValue(GetType(Color), "Highlight"), Category("Appearance"), Description("Focused border color (full outline for FixedSingle, bottom edge for Fixed3D). Empty uses the normal border color.")>
    Public Property FocusedBorderColor As Color
        Get
            Return _FocusedBorderColor
        End Get
        Set(Value As Color)
            If _FocusedBorderColor = Value Then Return
            _FocusedBorderColor = Value
            InvalidateBorder()
        End Set
    End Property

    ''' <summary>Reserves additional physical pixels for embedded controls, without replacing the text padding.</summary>
    Protected Sub SetTextMarginReservation(Left As Integer, Right As Integer)
        _ReservedLeft = Math.Max(0, Left)
        _ReservedRight = Math.Max(0, Right)
        ApplyTextMargins()
    End Sub

    Private Sub ApplyTextMargins()
        If Not IsHandleCreated Then Return
        Dim Inset As Integer = CInt(Math.Round(_HorizontalPadding * DeviceDpi / 96.0R, MidpointRounding.AwayFromZero))
        ' 0xFFFF is reserved by EM_SETMARGINS for EC_USEFONTINFO.
        Dim Left As Integer = CInt(Math.Min(65534L, CLng(Inset) + _ReservedLeft))
        Dim Right As Integer = CInt(Math.Min(65534L, CLng(Inset) + _ReservedRight))
        Dim Packed As Integer = Left Or (Right << 16)
        SendNativeMessage(Handle, EM_SETMARGINS, New IntPtr(EC_BOTHMARGINS), New IntPtr(Packed))
        Invalidate()
    End Sub

    Protected Overrides Sub OnHandleCreated(E As EventArgs)
        MyBase.OnHandleCreated(E)
        UpdatePlaceholder()
        If OwnsSingleLineFrame Then RefreshFrame()
        UpdateTextLayout()
    End Sub

    Protected Overrides Sub OnHandleDestroyed(E As EventArgs)
        _RequestedTopInset = 0
        _AppliedTopInset = 0
        _MultilinePaddingApplied = False
        MyBase.OnHandleDestroyed(E)
    End Sub

    Protected Overrides Sub OnFontChanged(E As EventArgs)
        MyBase.OnFontChanged(E)
        UpdateTextLayout()
    End Sub

    Protected Overrides Sub OnSizeChanged(E As EventArgs)
        MyBase.OnSizeChanged(E)
        UpdateTextLayout()
    End Sub

    Protected Overrides Sub OnBorderStyleChanged(E As EventArgs)
        MyBase.OnBorderStyleChanged(E)
        UpdateTextLayout()
    End Sub

    Protected Overrides Sub OnMultilineChanged(E As EventArgs)
        MyBase.OnMultilineChanged(E)
        UpdatePlaceholder()
        UpdateTextLayout()
    End Sub

    Protected Overrides Sub OnDpiChangedAfterParent(E As EventArgs)
        MyBase.OnDpiChangedAfterParent(E)
        UpdateTextLayout()
    End Sub

    Protected Overrides Sub OnGotFocus(E As EventArgs)
        MyBase.OnGotFocus(E)
        InvalidateBorder()
    End Sub

    Protected Overrides Sub OnLostFocus(E As EventArgs)
        MyBase.OnLostFocus(E)
        InvalidateBorder()
    End Sub

    Protected Overrides Sub OnEnabledChanged(E As EventArgs)
        MyBase.OnEnabledChanged(E)
        InvalidateBorder()
    End Sub

    Protected Overrides Sub OnReadOnlyChanged(E As EventArgs)
        MyBase.OnReadOnlyChanged(E)
        InvalidateBorder()
    End Sub

    Protected Overrides Sub OnBackColorChanged(E As EventArgs)
        MyBase.OnBackColorChanged(E)
        InvalidateBorder()
    End Sub

    Protected Overrides Sub OnSystemColorsChanged(E As EventArgs)
        MyBase.OnSystemColorsChanged(E)
        InvalidateBorder()
    End Sub

    Private Sub UpdatePlaceholder()
        If Not IsHandleCreated OrElse Multiline Then Return
        SendCueBanner(Handle, EM_SETCUEBANNER, IntPtr.Zero, If(_PlaceholderText, String.Empty))
    End Sub

    ' Measure the native edit line, then reserve only the unused space above it.
    ' The managed size is unchanged. Fixed3D keeps its original native geometry.
    Private Sub UpdateTextLayout()
        If Not IsHandleCreated OrElse IsDisposed OrElse Disposing OrElse _UpdatingTextLayout Then Return
        _UpdatingTextLayout = True
        Try
            ' Start with native geometry, not a previously reduced client rectangle.
            If _RequestedTopInset <> 0 OrElse _AppliedTopInset <> 0 Then
                _RequestedTopInset = 0
                If Not RefreshFrame() Then Return
            End If
            ApplyTextMargins()
            ' Do not carve the native themed 3D frame. Its rendering depends on
            ' the original client geometry; preserving that frame takes priority.
            If Multiline Then
                ApplyMultilineVerticalPadding()
                Return
            End If
            If BorderStyle = BorderStyle.Fixed3D Then Return

            Dim NativeInfo As New NativeWindowInfo With {.Size = CUInt(Marshal.SizeOf(GetType(NativeWindowInfo)))}
            If Not GetWindowInfo(Handle, NativeInfo) Then Return
            Dim NativeHeight As Integer = NativeInfo.ClientBounds.Bottom - NativeInfo.ClientBounds.Top
            Dim TextRect As NativeRect
            SendRectMessage(Handle, EM_GETRECT, IntPtr.Zero, TextRect)
            If TextRect.Top < 0 OrElse TextRect.Bottom <= TextRect.Top Then Return

            ' EM_GETRECT may describe a larger formatting area than a single line.
            ' Use the native line rectangle when it agrees with the font metrics.
            Dim LineHeight As Integer = FontHeight
            Dim RectHeight As Integer = TextRect.Bottom - TextRect.Top
            If Math.Abs(RectHeight - LineHeight) <= 2 Then LineHeight = RectHeight
            If LineHeight <= 0 OrElse TextRect.Top + LineHeight > NativeHeight Then Return

            Dim FreeSpace As Integer = NativeHeight - LineHeight
            Dim DesiredTop As Integer = (FreeSpace + 1) \ 2
            Dim InitialError As Integer = Math.Abs(DesiredTop - TextRect.Top)
            Dim Offset As Integer = DesiredTop - TextRect.Top
            If Offset <= 0 Then Return

            Dim BestInset As Integer = 0
            Dim BestError As Integer = InitialError
            ' Some EDIT implementations change their own text inset when the client
            ' height changes. Check the actual result; never retain a worse layout.
            Dim MaxInset As Integer = Math.Max(0, FreeSpace - TextRect.Top)
            Dim Candidates As Integer() = {Offset, Offset + 1, Offset - 1, Offset + 2, Offset - 2, MaxInset}
            For Each Candidate As Integer In Candidates
                If Candidate <= 0 OrElse Candidate > MaxInset Then Continue For
                _RequestedTopInset = Candidate
                If Not RefreshFrame() Then Exit For

                Dim ActualInfo As New NativeWindowInfo With {.Size = CUInt(Marshal.SizeOf(GetType(NativeWindowInfo)))}
                If Not GetWindowInfo(Handle, ActualInfo) Then Exit For
                Dim ActualRect As NativeRect
                SendRectMessage(Handle, EM_GETRECT, IntPtr.Zero, ActualRect)
                Dim ActualHeight As Integer = ActualInfo.ClientBounds.Bottom - ActualInfo.ClientBounds.Top
                Dim ActualTop As Integer = ActualInfo.ClientBounds.Top - NativeInfo.ClientBounds.Top + ActualRect.Top
                If ActualRect.Top >= 0 AndAlso ActualRect.Bottom > ActualRect.Top AndAlso
                   ActualRect.Top + LineHeight <= ActualHeight AndAlso
                   ActualInfo.WindowBounds.Bottom - ActualInfo.WindowBounds.Top = NativeInfo.WindowBounds.Bottom - NativeInfo.WindowBounds.Top AndAlso
                   ActualInfo.WindowBounds.Right - ActualInfo.WindowBounds.Left = NativeInfo.WindowBounds.Right - NativeInfo.WindowBounds.Left Then
                    Dim AlignmentError As Integer = Math.Abs(DesiredTop - ActualTop)
                    If AlignmentError < BestError Then
                        BestError = AlignmentError
                        BestInset = Candidate
                    End If
                    If BestError = 0 Then Exit For
                End If
            Next

            If _RequestedTopInset <> BestInset Then
                _RequestedTopInset = BestInset
                RefreshFrame()
            End If
            ApplyTextMargins()
        Finally
            _UpdatingTextLayout = False
            InvalidateBorder()
        End Try
    End Sub

    ' Multiline EDIT supports a native formatting rectangle. Do not change the
    ' window/client geometry or the non-client painting used by single-line fields.
    Private Sub ApplyMultilineVerticalPadding()
        If Not Multiline OrElse Not IsHandleCreated Then Return
        If _VerticalPadding = 0 AndAlso Not _MultilinePaddingApplied Then Return

        Dim Info As New NativeWindowInfo With {.Size = CUInt(Marshal.SizeOf(GetType(NativeWindowInfo)))}
        If Not GetWindowInfo(Handle, Info) Then Return
        Dim ClientWidth As Integer = Info.ClientBounds.Right - Info.ClientBounds.Left
        Dim ClientHeight As Integer = Info.ClientBounds.Bottom - Info.ClientBounds.Top
        If ClientWidth <= 0 OrElse ClientHeight <= 0 Then Return

        ' Start from the full client area each time, not from a previously padded
        ' EM_GETRECT result. EDIT applies its own border and horizontal margins.
        ' Pass an explicit rectangle, since a NULL reset is not always supported.
        Dim Bounds As New NativeRect With {.Right = ClientWidth, .Bottom = ClientHeight}
        If _VerticalPadding = 0 Then
            SendRectMessage(Handle, EM_SETRECT, IntPtr.Zero, Bounds)
            _MultilinePaddingApplied = False
            Return
        End If

        ' Measure the unpadded native formatting area without an intermediate
        ' repaint. Keep at least one line available when the field is small.
        SendRectMessage(Handle, EM_SETRECTNP, IntPtr.Zero, Bounds)
        Dim NativeBounds As NativeRect
        SendRectMessage(Handle, EM_GETRECT, IntPtr.Zero, NativeBounds)
        Dim NativeHeight As Integer = Math.Max(0, NativeBounds.Bottom - NativeBounds.Top)
        Dim Requested As Integer = CInt(Math.Round(_VerticalPadding * DeviceDpi / 96.0R, MidpointRounding.AwayFromZero))
        Dim Inset As Integer = Math.Min(Requested, Math.Max(0, (Math.Min(ClientHeight, NativeHeight) - FontHeight) \ 2))
        Bounds.Top = Inset
        Bounds.Bottom = ClientHeight - Inset
        SendRectMessage(Handle, EM_SETRECT, IntPtr.Zero, Bounds)
        _MultilinePaddingApplied = Inset > 0
    End Sub

    Private Function RefreshFrame() As Boolean
        If Not IsHandleCreated Then Return False
        Return SetWindowPos(Handle, IntPtr.Zero, 0, 0, 0, 0, SWP_NOCHANGE Or SWP_FRAMECHANGED)
    End Function

    Protected Overrides Sub WndProc(ByRef M As Message)
        ' A borderless EDIT must not draw another outline at the shifted origin.
        ' Consume non-client paint only when we own the complete single-line frame.
        If M.Msg = WM_NCPAINT AndAlso OwnsSingleLineFrame AndAlso IsHandleCreated Then
            If PaintBorder() Then
                M.Result = IntPtr.Zero
                Return
            End If
        End If

        MyBase.WndProc(M)
        Select Case M.Msg
            Case WM_NCCALCSIZE
                _AppliedTopInset = 0
                If M.LParam = IntPtr.Zero OrElse Multiline OrElse BorderStyle = BorderStyle.Fixed3D Then Return
                If Not OwnsSingleLineFrame AndAlso _RequestedTopInset <= 0 Then Return
                Dim Bounds As NativeRect = Marshal.PtrToStructure(Of NativeRect)(M.LParam)

                If OwnsSingleLineFrame Then
                    ' Reserve the same single-pixel outline on each side; EDIT
                    ' itself is borderless, but the public BorderStyle is unchanged.
                    Dim HorizontalBorder As Integer = Math.Min(1, Math.Max(0, (Bounds.Right - Bounds.Left - 1) \ 2))
                    Dim VerticalBorder As Integer = Math.Min(1, Math.Max(0, (Bounds.Bottom - Bounds.Top - 1) \ 2))
                    Bounds.Left += HorizontalBorder
                    Bounds.Right -= HorizontalBorder
                    Bounds.Top += VerticalBorder
                    Bounds.Bottom -= VerticalBorder
                End If

                _AppliedTopInset = Math.Min(_RequestedTopInset, Math.Max(0, Bounds.Bottom - Bounds.Top - 1))
                Bounds.Top += _AppliedTopInset
                Marshal.StructureToPtr(Bounds, M.LParam, False)
                M.Result = IntPtr.Zero

            Case WM_NCPAINT, WM_PAINT, WM_SETFOCUS, WM_KILLFOCUS, WM_NCACTIVATE
                PaintBorder()

            Case WM_NCHITTEST
                If _AppliedTopInset <= 0 OrElse Not IsHandleCreated Then Return
                Dim Info As New NativeWindowInfo With {.Size = CUInt(Marshal.SizeOf(GetType(NativeWindowInfo)))}
                If Not GetWindowInfo(Handle, Info) Then Return
                Dim Packed As Long = M.LParam.ToInt64()
                Dim X As Integer = CInt(Packed And &HFFFFL)
                Dim Y As Integer = CInt((Packed >> 16) And &HFFFFL)
                If X > Short.MaxValue Then X -= 65536
                If Y > Short.MaxValue Then Y -= 65536
                If X >= Info.ClientBounds.Left AndAlso X < Info.ClientBounds.Right AndAlso
                   Y >= Info.ClientBounds.Top - _AppliedTopInset AndAlso Y < Info.ClientBounds.Top Then
                    M.Result = New IntPtr(HTCLIENT)
                End If
        End Select
    End Sub

    Private Sub InvalidateBorder()
        If Not IsHandleCreated OrElse IsDisposed OrElse Disposing Then Return
        RedrawWindow(Handle, IntPtr.Zero, IntPtr.Zero, RDW_INVALIDATE Or RDW_FRAME)
    End Sub

    ' Paint only explicit non-client rectangles in window-relative coordinates.
    ' Do not paint into the edit surface or change its native text rendering.
    Private Function PaintBorder() As Boolean
        If Not IsHandleCreated OrElse IsDisposed OrElse Disposing Then Return False
        Dim CurrentColor As Color = Color.Empty
        If Enabled AndAlso BorderStyle <> BorderStyle.None Then
            CurrentColor = _BorderColor
            If Focused AndAlso Not _FocusedBorderColor.IsEmpty Then CurrentColor = _FocusedBorderColor
        End If
        If OwnsSingleLineFrame AndAlso CurrentColor.IsEmpty Then
            CurrentColor = If(Enabled, SystemColors.WindowFrame, SystemColors.ControlDark)
        End If
        If _AppliedTopInset = 0 AndAlso CurrentColor.IsEmpty Then Return False

        Dim WindowHandle As IntPtr = Handle
        Dim Info As New NativeWindowInfo With {.Size = CUInt(Marshal.SizeOf(GetType(NativeWindowInfo)))}
        If Not GetWindowInfo(WindowHandle, Info) Then Return False
        Dim WindowWidth As Integer = Info.WindowBounds.Right - Info.WindowBounds.Left
        Dim WindowHeight As Integer = Info.WindowBounds.Bottom - Info.WindowBounds.Top
        If WindowWidth < 2 OrElse WindowHeight < 2 Then Return False

        Dim ClientLeft As Integer = Math.Max(0, Info.ClientBounds.Left - Info.WindowBounds.Left)
        Dim ClientTop As Integer = Math.Max(0, Info.ClientBounds.Top - Info.WindowBounds.Top)
        Dim ClientRight As Integer = Math.Min(WindowWidth, Info.ClientBounds.Right - Info.WindowBounds.Left)
        Dim ClientBottom As Integer = Math.Min(WindowHeight, Info.ClientBounds.Bottom - Info.WindowBounds.Top)
        Dim NativeTop As Integer = Math.Max(0, ClientTop - _AppliedTopInset)
        Dim DC As IntPtr = GetWindowDC(WindowHandle)
        If DC = IntPtr.Zero Then Return False
        Try
            ' The padding strip is entirely outside the native edit area. It is
            ' not a second border and uses the same background as the text box.
            If _AppliedTopInset > 0 Then
                Dim FillColor As Color = BackColor
                If Not Enabled AndAlso BackColor = SystemColors.Window Then FillColor = SystemColors.Control
                FillNativeRectangle(DC, FillColor, ClientLeft, NativeTop, ClientRight, ClientTop)
            End If
            If CurrentColor.IsEmpty OrElse BorderStyle = BorderStyle.None Then Return True

            Select Case BorderStyle
                Case BorderStyle.FixedSingle
                    FillNativeRectangle(DC, CurrentColor, 0, 0, WindowWidth, 1)
                    FillNativeRectangle(DC, CurrentColor, 0, WindowHeight - 1, WindowWidth, WindowHeight)
                    FillNativeRectangle(DC, CurrentColor, 0, 1, 1, WindowHeight - 1)
                    FillNativeRectangle(DC, CurrentColor, WindowWidth - 1, 1, WindowWidth, WindowHeight - 1)

                Case BorderStyle.Fixed3D
                    ' Recolor the native bottom band across the full window
                    ' width, including both corners. Preserve the other edges.
                    Dim BottomThickness As Integer = Math.Min(NativeTop, WindowHeight - ClientBottom)
                    If BottomThickness > 0 AndAlso BottomThickness < WindowHeight Then
                        FillNativeRectangle(DC, CurrentColor, 0, WindowHeight - BottomThickness, WindowWidth, WindowHeight)
                    End If
            End Select
            Return True
        Finally
            ReleaseDC(WindowHandle, DC)
        End Try
    End Function

    Private Shared Sub FillNativeRectangle(DC As IntPtr, Color As Color, Left As Integer, Top As Integer, Right As Integer, Bottom As Integer)
        If Right <= Left OrElse Bottom <= Top Then Return
        Dim Bounds As New NativeRect With {.Left = Left, .Top = Top, .Right = Right, .Bottom = Bottom}
        Dim Brush As IntPtr = CreateSolidBrush(CUInt(ColorTranslator.ToWin32(Color)))
        If Brush = IntPtr.Zero Then Return
        Try
            FillRect(DC, Bounds, Brush)
        Finally
            DeleteObject(Brush)
        End Try
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
        Public ExtendedStyle As UInteger
        Public WindowStatus As UInteger
        Public WindowBorderWidth As UInteger
        Public WindowBorderHeight As UInteger
        Public WindowType As UShort
        Public CreatorVersion As UShort
    End Structure

    <DllImport("user32.dll", EntryPoint:="SendMessageW", ExactSpelling:=True, CharSet:=CharSet.Unicode)>
    Private Shared Function SendCueBanner(HWnd As IntPtr, Msg As Integer, WParam As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> LParam As String) As IntPtr
    End Function

    <DllImport("user32.dll", EntryPoint:="SendMessageW", ExactSpelling:=True)>
    Private Shared Function SendNativeMessage(HWnd As IntPtr, Msg As Integer, WParam As IntPtr, LParam As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", EntryPoint:="SendMessageW", ExactSpelling:=True)>
    Private Shared Function SendRectMessage(HWnd As IntPtr, Msg As Integer, WParam As IntPtr, ByRef LParam As NativeRect) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetWindowPos(HWnd As IntPtr, HWndInsertAfter As IntPtr, X As Integer, Y As Integer, CX As Integer, CY As Integer, Flags As UInteger) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function RedrawWindow(HWnd As IntPtr, UpdateRect As IntPtr, UpdateRegion As IntPtr, Flags As UInteger) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetWindowInfo(HWnd As IntPtr, ByRef Info As NativeWindowInfo) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetWindowDC(HWnd As IntPtr) As IntPtr
    End Function

    <DllImport("gdi32.dll", ExactSpelling:=True)>
    Private Shared Function CreateSolidBrush(ColorRef As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll", ExactSpelling:=True)>
    Private Shared Function FillRect(DC As IntPtr, ByRef Bounds As NativeRect, Brush As IntPtr) As Integer
    End Function

    <DllImport("gdi32.dll", ExactSpelling:=True)>
    Private Shared Function DeleteObject(Handle As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function ReleaseDC(HWnd As IntPtr, DC As IntPtr) As Integer
    End Function
End Class
