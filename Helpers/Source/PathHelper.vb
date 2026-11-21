Imports System.IO

''' <summary>Provides filesystem path operations.</summary>
Public NotInheritable Class PathHelper
    Private Sub New()
    End Sub
    ''' <summary>Removes one trailing separator without removing a filesystem root.</summary>
    Public Shared Function TrimEndingDirectorySeparator(Value As String) As String
        If Value Is Nothing Then Throw New ArgumentNullException(NameOf(Value))
        If Value.Length > Path.GetPathRoot(Value).Length AndAlso (Value.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) OrElse Value.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal)) Then Return Value.Substring(0, Value.Length - 1)
        Return Value
    End Function

    ''' <summary>Returns the relative path between two filesystem locations.</summary>
    Public Shared Function GetRelativePath(RelativeTo As String, Target As String) As String
        If String.IsNullOrEmpty(RelativeTo) Then Throw New ArgumentException("Path cannot be empty.", NameOf(RelativeTo))
        If String.IsNullOrEmpty(Target) Then Throw New ArgumentException("Path cannot be empty.", NameOf(Target))
        Dim Root = Path.GetFullPath(RelativeTo)
        Dim FullTarget = Path.GetFullPath(Target)
        Dim Comparison = If(Environment.OSVersion.Platform = PlatformID.Win32NT, StringComparison.OrdinalIgnoreCase, StringComparison.Ordinal)
        If Not String.Equals(Path.GetPathRoot(Root), Path.GetPathRoot(FullTarget), Comparison) Then Return FullTarget
        Dim Separators = New Char() {Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar}
        Dim RootParts = Root.TrimEnd(Separators).Split(Separators, StringSplitOptions.RemoveEmptyEntries)
        Dim TargetParts = FullTarget.TrimEnd(Separators).Split(Separators, StringSplitOptions.RemoveEmptyEntries)
        Dim Common As Integer = 0
        While Common < Math.Min(RootParts.Length, TargetParts.Length) AndAlso String.Equals(RootParts(Common), TargetParts(Common), Comparison)
            Common += 1
        End While
        Dim Parts As New Collections.Generic.List(Of String)
        For I = Common To RootParts.Length - 1
            Parts.Add("..")
        Next
        For I = Common To TargetParts.Length - 1
            Parts.Add(TargetParts(I))
        Next
        If Parts.Count = 0 Then Return "."
        Dim Result = String.Join(Path.DirectorySeparatorChar.ToString(), Parts)
        If FullTarget.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) OrElse FullTarget.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal) Then Result &= Path.DirectorySeparatorChar
        Return Result
    End Function

End Class
