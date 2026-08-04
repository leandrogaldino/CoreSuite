Imports System.Text.RegularExpressions
Friend Module MySqlSql
    Private ReadOnly ParameterNamePattern As New Regex("^@?[A-Za-z_][A-Za-z0-9_]*$", RegexOptions.Compiled Or RegexOptions.CultureInvariant)
    Private ReadOnly RoutineNamePattern As New Regex("^[A-Za-z0-9_$]+(?:\.[A-Za-z0-9_$]+)*$", RegexOptions.Compiled Or RegexOptions.CultureInvariant)
    Private ReadOnly SqlTokenPattern As New Regex("^[A-Za-z0-9_]+$", RegexOptions.Compiled Or RegexOptions.CultureInvariant)
    Friend Function RequireValue(Value As String, ParameterName As String) As String
        If String.IsNullOrWhiteSpace(Value) Then Throw New ArgumentException("The value cannot be null, empty, or whitespace.", ParameterName)
        Return Value.Trim()
    End Function
    Friend Function QuoteIdentifier(Identifier As String, ParameterName As String, Optional AllowWildcard As Boolean = False) As String
        Dim NormalizedIdentifier As String = RequireValue(Identifier, ParameterName)
        Dim Parts As String() = NormalizedIdentifier.Split("."c)
        Dim QuotedParts As New List(Of String)(Parts.Length)
        For Index As Integer = 0 To Parts.Length - 1
            Dim Part As String = Parts(Index).Trim()
            If Part.Length = 0 Then Throw New ArgumentException("The identifier contains an empty part.", ParameterName)
            If AllowWildcard AndAlso Part = "*" AndAlso Index = Parts.Length - 1 Then
                QuotedParts.Add("*")
            Else
                Dim EscapedPart As String = Part.Replace("`", "``", StringComparison.Ordinal)
                QuotedParts.Add("`" & EscapedPart & "`")
            End If
        Next Index
        Return String.Join(".", QuotedParts)
    End Function
    Friend Function QuoteSingleIdentifier(Identifier As String, ParameterName As String) As String
        Dim NormalizedIdentifier As String = RequireValue(Identifier, ParameterName)
        Dim WscapedIdentifier As String = NormalizedIdentifier.Replace("`", "``", StringComparison.Ordinal)
        Return "`" & WscapedIdentifier & "`"
    End Function
    Friend Function ValidateRoutineName(RoutineName As String, ParameterName As String) As String
        Dim NormalizedName As String = RequireValue(RoutineName, ParameterName)
        If Not RoutineNamePattern.IsMatch(NormalizedName) Then Throw New ArgumentException("The routine name must contain only letters, digits, underscores, dollar signs, and optional qualifier separators.", ParameterName)
        Return NormalizedName
    End Function
    Friend Function NormalizeParameterName(ParameterName As String) As String
        Dim NormalizedName As String = RequireValue(ParameterName, NameOf(ParameterName))
        If Not ParameterNamePattern.IsMatch(NormalizedName) Then Throw New ArgumentException("Parameter names must begin with a letter or underscore and contain only letters, digits, and underscores.", NameOf(ParameterName))
        If NormalizedName(0) <> "@"c Then NormalizedName = "@" & NormalizedName
        Return NormalizedName
    End Function
    Friend Function ValidateSqlToken(Value As String, ParameterName As String) As String
        Dim NormalizedValue As String = RequireValue(Value, ParameterName)
        If Not SqlTokenPattern.IsMatch(NormalizedValue) Then Throw New ArgumentException("The value contains characters that are not valid for this SQL token.", ParameterName)
        Return NormalizedValue
    End Function
    Friend Function RequirePositive(Value As Integer, ParameterName As String) As Integer
        If Value <= 0 Then Throw New ArgumentOutOfRangeException(ParameterName, Value, "The value must be greater than zero.")
        Return Value
    End Function
    Friend Function ClampPercentage(CurrentValue As Long, TotalValue As Long) As Integer
        If TotalValue <= 0 Then Return 0
        Dim Percentage As Double = CurrentValue / CDbl(TotalValue) * 100.0R
        Return Math.Clamp(CInt(Math.Truncate(Percentage)), 0, 100)
    End Function
End Module
