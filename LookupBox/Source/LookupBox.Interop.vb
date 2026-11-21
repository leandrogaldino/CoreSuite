Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Friend Module LookupBoxInterop
    Private Const EM_SETMARGINS As Integer = &HD3
    Private Const EM_SETCUEBANNER As Integer = &H1501
    Private Const EC_LEFTMARGIN As Integer = &H1
    Private Const EC_RIGHTMARGIN As Integer = &H2
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Function SendMessage(WindowHandle As IntPtr, Message As Integer, Parameter As IntPtr, Data As IntPtr) As IntPtr
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Unicode, EntryPoint:="SendMessageW")>
    Private Function SendMessageCueBanner(WindowHandle As IntPtr, Message As Integer, Parameter As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> Text As String) As IntPtr
    End Function
    Friend Sub SetCueBanner(Control As TextBox, Text As String)
        If Control Is Nothing OrElse Not Control.IsHandleCreated Then Return
        SendMessageCueBanner(Control.Handle, EM_SETCUEBANNER, IntPtr.Zero, If(Text, String.Empty))
    End Sub
    Friend Sub SetTextMargins(Control As TextBox, LeftMargin As Integer, RightMargin As Integer)
        If Control Is Nothing OrElse Not Control.IsHandleCreated Then Return
        Dim SafeLeftMargin As Integer = Math.Max(0, Math.Min(UShort.MaxValue, LeftMargin))
        Dim SafeRightMargin As Integer = Math.Max(0, Math.Min(UShort.MaxValue, RightMargin))
        Dim MarginData As Integer = SafeLeftMargin Or (SafeRightMargin << 16)
        SendMessage(Control.Handle, EM_SETMARGINS, New IntPtr(EC_LEFTMARGIN Or EC_RIGHTMARGIN), New IntPtr(MarginData))
    End Sub
End Module
