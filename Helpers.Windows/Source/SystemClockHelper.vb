Imports System.Runtime.InteropServices

''' <summary>Provides the Windows monotonic system clock.</summary>
Public NotInheritable Class SystemClockHelper
    Private Sub New()
    End Sub
    <DllImport("kernel32.dll", EntryPoint:="GetTickCount64")>
    Private Shared Function GetTickCount64Native() As ULong
    End Function

    ''' <summary>Gets the 64-bit Windows system uptime in milliseconds.</summary>
    Public Shared ReadOnly Property TickCount64 As Long
        Get
            Return CLng(GetTickCount64Native())
        End Get
    End Property

End Class
