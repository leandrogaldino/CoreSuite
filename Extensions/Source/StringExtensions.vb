Imports System.Runtime.CompilerServices
Imports System.Text

Namespace Extensions
    Public Module StringExtensions

        ''' <summary>
        ''' Reverses the order of characters in a string.
        ''' </summary>
        ''' <param name="Text">The string to be reversed.</param>
        ''' <returns>
        ''' A new string with characters in reverse order.
        ''' Returns an empty string if the input is null or empty.
        ''' </returns>
        ''' <remarks>
        ''' This method does not alter the content within words; it only reverses the character order.
        ''' Example: "Hello World" becomes "dlroW olleH".
        ''' Useful for simple text manipulations.
        ''' </remarks>
        <Extension>
        Public Function ReverseText(Text As String) As String
            If String.IsNullOrEmpty(Text) Then Return String.Empty
            Return New String(Text.Reverse().ToArray())
        End Function

        ''' <summary>
        ''' Reverses the order of words in a string.
        ''' </summary>
        ''' <param name="Text">The string containing the words to be reversed.</param>
        ''' <returns>
        ''' A new string with the word order reversed.
        ''' Returns an empty string if the input is null or empty.
        ''' </returns>
        ''' <remarks>
        ''' The method uses spaces as delimiters to identify words.
        ''' Words are reversed while keeping characters within each word intact.
        ''' Example: "Hello World" becomes "World Hello".
        ''' Useful for text manipulation and formatting.
        ''' </remarks>
        <Extension>
        Public Function ReverseWords(Text As String) As String
            If String.IsNullOrEmpty(Text) Then Return String.Empty
            Return String.Join(" ", Text.Split(" "c).Reverse())
        End Function

        ''' <summary>
        ''' Counts the number of words in a string.
        ''' </summary>
        ''' <param name="Text">The string to analyze.</param>
        ''' <returns>
        ''' An integer representing the number of words in the provided string.
        ''' Returns 0 if the input is null or empty.
        ''' </returns>
        ''' <remarks>
        ''' The method considers words as text fragments separated by spaces, tabs,
        ''' or line breaks (LF and CR). Useful for text analysis and user input processing.
        ''' </remarks>
        <Extension>
        Public Function CountWords(Text As String) As Integer
            If String.IsNullOrEmpty(Text) Then Return 0
            Return Text.Split(New Char() {" "c, ControlChars.Tab, ControlChars.Lf, ControlChars.Cr}, StringSplitOptions.RemoveEmptyEntries).Length
        End Function

        ''' <summary>
        ''' Removes all accents and diacritics from a string, returning an unaccented version.
        ''' </summary>
        ''' <param name="Text">The string to process.</param>
        ''' <returns>
        ''' A new string without accents or diacritics.
        ''' Returns an empty string if the input is null or empty.
        ''' </returns>
        ''' <remarks>
        ''' This method uses normalization to decompose accented characters and then
        ''' removes the diacritics. Useful when text comparison or storage needs to
        ''' be accent-insensitive (e.g., database searches or name standardization).
        ''' Note that non-accented special characters, like cedilla (ç), are preserved.
        ''' </remarks>
        <Extension()>
        Public Function ToUnaccented(Text As String) As String
            If String.IsNullOrEmpty(Text) Then Return String.Empty
            Dim SbReturn As New StringBuilder()
            Dim ArrayText = Text.Normalize(NormalizationForm.FormD).ToCharArray()
            For Each letter As Char In ArrayText
                If Globalization.CharUnicodeInfo.GetUnicodeCategory(letter) <> Globalization.UnicodeCategory.NonSpacingMark Then
                    SbReturn.Append(letter)
                End If
            Next
            Return SbReturn.ToString()
        End Function

        ''' <summary>
        ''' Returns a new string with the first letter of each word capitalized
        ''' and the remaining letters in lowercase.
        ''' </summary>
        ''' <param name="Text">The string to convert.</param>
        ''' <returns>
        ''' A new string where the first letter of each word is uppercase and the rest lowercase.
        ''' Returns <c>Nothing</c> if the provided text is null or empty.
        ''' </returns>
        ''' <remarks>
        ''' This function uses the system's current culture for conversion, so behavior may vary
        ''' depending on the runtime culture settings. Useful for formatting titles or names.
        ''' Note that words in uppercase followed by numbers or special characters may not
        ''' format correctly due to the rules of the current culture.
        ''' </remarks>
        <Extension()>
        Public Function ToTitle(Text As String) As String
            Return If(Not String.IsNullOrEmpty(Text),
              Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(Text.ToLower()),
              Nothing)
        End Function

        ''' <summary>
        ''' Converts a string to CamelCase format.
        ''' Words are separated by spaces, underscores, or hyphens, with the first word lowercase
        ''' and subsequent words capitalized.
        ''' </summary>
        ''' <param name="Text">The input string to convert to CamelCase.</param>
        ''' <returns>A string converted to CamelCase format.</returns>
        <Extension>
        Public Function ToCamel(Text As String) As String
            Dim words = Text.Split(New Char() {" "c, "_"c, "-"c}, StringSplitOptions.RemoveEmptyEntries)
            Return String.Join("", words.Select(Function(w, i) If(i = 0, w.ToLower(), Char.ToUpper(w(0)) & w.Substring(1).ToLower())))
        End Function

        ''' <summary>
        ''' Replaces the first occurrence of a specific substring in a string
        ''' with another provided substring.
        ''' </summary>
        ''' <param name="Text">The original string where the search will be performed.</param>
        ''' <param name="OldValue">The substring to find in the string for replacement.</param>
        ''' <param name="NewValue">The substring that will replace the first occurrence of <paramref name="OldValue"/>.</param>
        ''' <returns>
        ''' Returns a new string where only the first occurrence of <paramref name="OldValue"/>
        ''' was replaced by <paramref name="NewValue"/>. If the substring is not found,
        ''' returns the original string unchanged.
        ''' </returns>
        ''' <remarks>
        ''' This function is useful when you want to replace only the first occurrence
        ''' of a substring instead of all occurrences as the standard <see cref="String.Replace"/> method does.
        ''' </remarks>
        <Extension()>
        Public Function ReplaceFirst(Text As String, OldValue As String, NewValue As String) As String
            Dim pos As Integer = Text.IndexOf(OldValue)
            If pos < 0 Then
                Return Text
            End If
            Return Text.Substring(0, pos) & NewValue & Text.Substring(pos + OldValue.Length)
        End Function

        ''' <summary>
        ''' Removes extra spaces from a string, leaving only a single space between words.
        ''' </summary>
        ''' <param name="Text">The string to process.</param>
        ''' <returns>A string with extra spaces removed.</returns>
        <Extension()>
        Public Function RemoveExtraSpaces(Text As String) As String
            Return String.Join(" ", Text.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries))
        End Function

    End Module
End Namespace