Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
''' <summary>Draws placeholder text for an empty, unfocused TextBox.</summary>
Public NotInheritable Class TextBoxPlaceholder
    Inherits NativeWindow
    Private ReadOnly _Owner As TextBox
    ''' <summary>Gets or sets the placeholder text.</summary>
    Public Property Text As String = String.Empty
    ''' <summary>Attaches placeholder rendering to the specified text box.</summary>
    Public Sub New(Owner As TextBox)
        If Owner Is Nothing Then Throw New ArgumentNullException(NameOf(Owner))
        _Owner = Owner
        AddHandler Owner.HandleCreated, AddressOf Attach
        AddHandler Owner.HandleDestroyed, AddressOf Detach
        AddHandler Owner.GotFocus, AddressOf Refresh
        AddHandler Owner.LostFocus, AddressOf Refresh
        AddHandler Owner.TextChanged, AddressOf Refresh
        If Owner.IsHandleCreated Then AssignHandle(Owner.Handle)
    End Sub
    Private Sub Attach(Sender As Object, E As EventArgs)
        AssignHandle(_Owner.Handle)
    End Sub
    Private Sub Detach(Sender As Object, E As EventArgs)
        ReleaseHandle()
    End Sub
    Private Sub Refresh(Sender As Object, E As EventArgs)
        _Owner.Invalidate()
    End Sub
    Protected Overrides Sub WndProc(ByRef M As Message)
        Dim Paint = M.Msg = &HF
        Dim Print = M.Msg = &H317 OrElse M.Msg = &H318
        Dim Visible = Not String.IsNullOrEmpty(Text) AndAlso Not _Owner.Focused AndAlso _Owner.TextLength = 0
        If Paint AndAlso Visible Then _Owner.Invalidate()
        MyBase.WndProc(M)
        If Not Visible OrElse Not (Paint OrElse Print) Then Return
        If Print AndAlso M.WParam <> IntPtr.Zero Then
            Using Canvas = Graphics.FromHdc(M.WParam)
                Draw(Canvas)
            End Using
        ElseIf Paint Then
            Using Canvas = Graphics.FromHwnd(_Owner.Handle)
                Draw(Canvas)
            End Using
        End If
    End Sub
    Private Sub Draw(Canvas As Graphics)
        Dim Bounds = _Owner.ClientRectangle
        Dim Flags = TextFormatFlags.NoPadding Or TextFormatFlags.Top Or TextFormatFlags.EndEllipsis
        Dim Alignment = _Owner.TextAlign
        If _Owner.RightToLeft = RightToLeft.Yes Then
            Flags = Flags Or TextFormatFlags.RightToLeft
            If Alignment = HorizontalAlignment.Left Then
                Alignment = HorizontalAlignment.Right
            ElseIf Alignment = HorizontalAlignment.Right Then
                Alignment = HorizontalAlignment.Left
            End If
        End If
        Select Case Alignment
            Case HorizontalAlignment.Center
                Flags = Flags Or TextFormatFlags.HorizontalCenter
            Case HorizontalAlignment.Right
                Flags = Flags Or TextFormatFlags.Right
            Case Else
                Flags = Flags Or TextFormatFlags.Left
        End Select
        Bounds.Offset(If(_Owner.TextAlign = HorizontalAlignment.Left, 1, 0), 1)
        TextRenderer.DrawText(Canvas, Text, _Owner.Font, Bounds, SystemColors.GrayText, Flags)
    End Sub
End Class
