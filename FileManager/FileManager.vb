Imports System.IO

Partial Public NotInheritable Class FileManager

    Private Sub New()
    End Sub

#Region "Helpers"

    Private Shared Function CalcPercent(handled As Long, total As Long) As Integer
        If total <= 0 Then Return 0
        Return CInt((handled * 100L) \ total)
    End Function

    Private Shared Function GetDirectorySize(directory As DirectoryInfo) As Long
        Dim size As Long = 0
        For Each file In directory.GetFiles("*.*", SearchOption.AllDirectories)
            size += file.Length
        Next
        Return size
    End Function

#End Region

End Class