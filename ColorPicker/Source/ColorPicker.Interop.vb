Imports System.Runtime.InteropServices
''' <summary>
''' Provides the native Win32 interop methods required by the <see cref="ColorPicker"/> control.
''' </summary>
Partial Public Class ColorPicker
    ''' <summary>
    ''' Sends the specified message to the given window.
    ''' </summary>
    ''' <param name="hWnd">A handle to the destination window.</param>
    ''' <param name="Msg">The message to send.</param>
    ''' <param name="wParam">Additional message-specific information.</param>
    ''' <param name="lParam">Additional message-specific information.</param>
    ''' <returns>
    ''' The result of the message processing, which depends on the message sent.
    ''' </returns>
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(hWnd As IntPtr, Msg As UInteger, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function
    ''' <summary>
    ''' Determines whether a key is currently pressed or was pressed since the previous call.
    ''' </summary>
    ''' <param name="vKey">The virtual key to query.</param>
    ''' <returns>
    ''' A 16-bit value indicating the key state. The high-order bit is set if the key is currently pressed.
    ''' </returns>
    <DllImport("user32.dll")>
    Private Shared Function GetAsyncKeyState(vKey As Keys) As Short
    End Function
End Class
