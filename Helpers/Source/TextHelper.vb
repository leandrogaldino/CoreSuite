''' <summary>
''' Provides helper methods for text generation, string manipulation,
''' and simple text parsing operations.
''' </summary>
Public Class TextHelper

    ''' <summary>
    ''' Generates a random string composed of multiple sets of characters.
    ''' </summary>
    ''' <param name="SetCount">
    ''' The number of character sets to generate.
    ''' </param>
    ''' <param name="CharsPerSet">
    ''' The number of characters in each set.
    ''' </param>
    ''' <param name="SetSeparator">
    ''' The string used to separate each generated set.
    ''' </param>
    ''' <param name="PossibleChars">
    ''' An optional list defining the allowed character filters.
    ''' If not provided, the default character set will be used.
    ''' </param>
    ''' <returns>
    ''' A randomly generated string composed of the specified character sets.
    ''' </returns>
    Public Shared Function GetRandomString(SetCount As Integer, CharsPerSet As Integer, SetSeparator As String, Optional PossibleChars As List(Of CharFilter) = Nothing) As String
        Dim CharSet As String = GetCharSet(PossibleChars)
        Dim Random As New Random()
        Dim Sets As String() = Enumerable.Range(0, SetCount).Select(Function(i) New String(Enumerable.Repeat(CharSet, CharsPerSet).Select(Function(s) s(Random.Next(s.Length))).ToArray())).ToArray()
        Dim RandomString As String = String.Join(SetSeparator, Sets)
        Return RandomString
    End Function

    ''' <summary>
    ''' Generates a random file name using the current date/time,
    ''' a GUID, and a system-generated random component.
    ''' </summary>
    ''' <param name="Extension">
    ''' An optional file extension to append to the generated file name.
    ''' </param>
    ''' <returns>
    ''' A randomly generated file name string.
    ''' </returns>
    Public Shared Function GetRandomFileName(Optional Extension As String = Nothing) As String
        Dim Filename As New Text.StringBuilder
        Filename.Append(Now.ToString("ddMMyyyyHHmmss"))
        Filename.Append(Guid.NewGuid.ToString("N").ToUpper)
        Filename.Append(IO.Path.GetRandomFileName().Replace(".", Nothing).ToUpper)
        If Not String.IsNullOrEmpty(Extension) Or String.IsNullOrWhiteSpace(Extension) Then
            Filename.Append(Extension)
        End If
        Return Filename.ToString
    End Function

    ''' <summary>
    ''' Builds a character set string based on the specified
    ''' character filter options.
    ''' </summary>
    ''' <param name="PossibleChars">
    ''' A list of character filters that define which characters
    ''' should be included in the resulting character set.
    ''' </param>
    ''' <returns>
    ''' A string containing all characters allowed by the specified filters.
    ''' </returns>
    Private Shared Function GetCharSet(PossibleChars As List(Of CharFilter)) As String
        Dim DistinctChars = PossibleChars.Distinct().ToList()
        Dim CharSet As New Text.StringBuilder
        For Each CharType As CharFilter In DistinctChars
            Select Case CharType
                Case CharFilter.Alphanumeric
                    CharSet.Append("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz")
                Case CharFilter.Numeric
                    CharSet.Append("0123456789")
                Case CharFilter.UppercaseAlphabetic
                    CharSet.Append("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
                Case CharFilter.LowercaseAlphabetic
                    CharSet.Append("abcdefghijklmnopqrstuvwxyz")
                Case CharFilter.SpecialCharacters
                    CharSet.Append("!#$%&'()*+,-./:;<=>?@[]^_`{|}~")
                Case CharFilter.Hexadecimal
                    CharSet.Append("0123456789ABCDEF")
            End Select
        Next
        Return CharSet.ToString()
    End Function

    ''' <summary>
    ''' Extracts the value of a specified key from a JSON string.
    ''' </summary>
    ''' <param name="Json">
    ''' The JSON string to be parsed.
    ''' </param>
    ''' <param name="Key">
    ''' The key whose associated value will be extracted.
    ''' </param>
    ''' <returns>
    ''' The value associated with the specified key,
    ''' or <c>Nothing</c> if the key is not found or parsing fails.
    ''' </returns>
    Public Shared Function ExtractJsonValue(Json As String, Key As String) As String
        Try
            Dim SearchKey = $"""{Key}"":"
            Dim StartIndex = Json.IndexOf(SearchKey)
            If StartIndex = -1 Then Return Nothing
            StartIndex += SearchKey.Length
            StartIndex = Json.IndexOf("""", StartIndex) + 1
            Dim EndIndex = Json.IndexOf("""", StartIndex)
            Return Json.Substring(StartIndex, EndIndex - StartIndex)
        Catch
            Return Nothing
        End Try
    End Function

End Class

''' <summary>
''' Defines character filter options used to compose
''' character sets for random string generation.
''' </summary>
Public Enum CharFilter

    ''' <summary>
    ''' Alphanumeric characters (A–Z, a–z, 0–9).
    ''' </summary>
    Alphanumeric

    ''' <summary>
    ''' Numeric characters (0–9).
    ''' </summary>
    Numeric

    ''' <summary>
    ''' Uppercase alphabetic characters (A–Z).
    ''' </summary>
    UppercaseAlphabetic

    ''' <summary>
    ''' Lowercase alphabetic characters (a–z).
    ''' </summary>
    LowercaseAlphabetic

    ''' <summary>
    ''' Special characters (e.g. !, @, #, $, %, etc.).
    ''' </summary>
    SpecialCharacters

    ''' <summary>
    ''' Hexadecimal characters (0–9, A–F).
    ''' </summary>
    Hexadecimal

End Enum