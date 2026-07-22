Imports System.Runtime.InteropServices

Friend Class ColorUIWnd
    Inherits NativeWindow

    <StructLayout(LayoutKind.Sequential)>
    Private Structure WINDOWPOS
        Public hwnd As IntPtr
        Public hwndInsertAfter As IntPtr
        Public x As Integer
        Public y As Integer
        Public cx As Integer
        Public cy As Integer
        Public flags As Integer
    End Structure

    Private Const WM_WINDOWPOSCHANGING As Integer = &H46
    Private Const SWP_NOSIZE As Integer = &H1
    Public PreventSizing As Boolean

    Protected Overrides Sub WndProc(ByRef m As Message)
        If PreventSizing AndAlso m.Msg = WM_WINDOWPOSCHANGING Then
            Dim wp As WINDOWPOS = CType(Marshal.PtrToStructure(m.LParam, GetType(WINDOWPOS)), WINDOWPOS)
            wp.flags = wp.flags Or SWP_NOSIZE
            Marshal.StructureToPtr(wp, m.LParam, False)
        End If

        MyBase.WndProc(m)
    End Sub
End Class