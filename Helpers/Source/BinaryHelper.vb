

''' <summary>Converts between binary data and hexadecimal text.</summary>
Public NotInheritable Class BinaryHelper
    Private Sub New()
    End Sub
    ''' <summary>Encodes bytes as uppercase hexadecimal text.</summary>
    Public Shared Function ToHexString(Bytes As Byte()) As String
        If Bytes Is Nothing Then Throw New ArgumentNullException(NameOf(Bytes))
        Return BitConverter.ToString(Bytes).Replace("-", "")
    End Function

    ''' <summary>Decodes hexadecimal text, rejecting invalid characters and lengths.</summary>
    Public Shared Function FromHexString(Text As String) As Byte()
        If Text Is Nothing Then Throw New ArgumentNullException(NameOf(Text))
        If Text.Length Mod 2 <> 0 Then Throw New FormatException("Invalid hexadecimal length.")
        Dim Result(Text.Length \ 2 - 1) As Byte
        For I = 0 To Result.Length - 1
            Dim High = HexDigit(Text(I * 2))
            Dim Low = HexDigit(Text(I * 2 + 1))
            If High < 0 OrElse Low < 0 Then Throw New FormatException("Invalid hexadecimal character.")
            Result(I) = CByte(High * 16 + Low)
        Next
        Return Result
    End Function

    Private Shared Function HexDigit(Value As Char) As Integer
        If Value >= "0"c AndAlso Value <= "9"c Then Return AscW(Value) - AscW("0"c)
        If Value >= "a"c AndAlso Value <= "f"c Then Return AscW(Value) - AscW("a"c) + 10
        If Value >= "A"c AndAlso Value <= "F"c Then Return AscW(Value) - AscW("A"c) + 10
        Return -1
    End Function

End Class
