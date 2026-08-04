Imports System.Text.RegularExpressions

''' <summary>
''' Provides utility methods for formatting and validating
''' common Brazilian data formats, including ZIP codes (CEP),
''' phone numbers, CPF, CNPJ, and email addresses.
''' </summary>
Public Class BrazilianFormatHelper

    ''' <summary>
    ''' If the input string is a valid Brazilian ZIP code (CEP), returns the value formatted with a mask.
    ''' </summary>
    ''' <param name="ZipCode">The ZIP code value.</param>
    ''' <returns>The formatted ZIP code.</returns>
    Public Shared Function GetFormatedZipCode(ZipCode As String) As String
        Dim CleanZipCode As String = ZipCode.Replace(".", Nothing).Replace("-", Nothing).Replace(" ", Nothing)
        If IsValidZipCode(CleanZipCode) Then
            Return String.Format("{0}.{1}-{2}", Mid(CleanZipCode, 1, 2), Mid(CleanZipCode, 3, 3), Mid(CleanZipCode, 6, 3))
        Else
            Return ZipCode
        End If
    End Function

    ''' <summary>
    ''' Determines whether the provided string is a valid Brazilian ZIP code (CEP).
    ''' </summary>
    ''' <param name="ZipCode">The ZIP code value.</param>
    ''' <returns>
    ''' <c>True</c> if the ZIP code is valid; otherwise, <c>False</c>.
    ''' </returns>
    Public Shared Function IsValidZipCode(ZipCode As String) As Boolean
        Dim CleanZipCode As String = ZipCode.Replace(".", Nothing).Replace("-", Nothing).Replace(" ", Nothing)
        Return Regex.IsMatch(CleanZipCode, "^[0-9]{8}$")
    End Function

    ''' <summary>
    ''' If the input string is a valid phone number, returns the value formatted with the appropriate mask.
    ''' </summary>
    ''' <param name="PhoneNumber">The phone number value.</param>
    ''' <returns>The formatted phone number.</returns>
    Public Shared Function GetFormatedPhoneNumber(PhoneNumber As String) As String
        Dim Phone As String = PhoneNumber.Replace("(", Nothing).Replace(")", Nothing).Replace("-", Nothing).Replace(" ", Nothing)
        Dim SpecialPhones As New List(Of String) From {
                                                            "0300", "0500", "0800", "0900"
                                                      }
        Dim FixedPhones As New List(Of String) From {
                                                        "11", "12", "13", "14", "15", "16", "17", "18", "19",
                                                        "21", "22", "24", "27", "28", "31", "32", "33", "34",
                                                        "35", "37", "38", "41", "42", "43", "44", "45", "46",
                                                        "47", "48", "49", "51", "53", "54", "55", "61", "62",
                                                        "63", "64", "65", "66", "67", "68", "69", "71", "73",
                                                        "74", "75", "77", "79", "81", "82", "83", "84", "85",
                                                        "86", "87", "88", "89", "91", "92", "93", "94", "95",
                                                         "96", "97", "98", "99"
                                                    }
        Dim CellPhones As New List(Of String) From {
                                                        "119", "129", "139", "149", "159", "169", "179", "189", "199",
                                                        "219", "229", "249", "279", "289", "319", "329", "339", "349",
                                                        "359", "379", "389", "419", "429", "439", "449", "459", "469",
                                                        "479", "489", "499", "519", "539", "549", "559", "619", "629",
                                                        "639", "649", "659", "669", "679", "689", "699", "719", "739",
                                                        "749", "759", "779", "799", "819", "829", "839", "849", "859",
                                                        "869", "879", "889", "899", "919", "929", "939", "949", "959",
                                                         "969", "979", "989", "999"
                                                    }
        If GetWhichPhoneFormat(Phone) = PhoneFormat.FixedPhone Then
            Phone = String.Format("({0}) {1}-{2}",
                                      Mid(Phone, 1, 2),
                                      Mid(Phone, 3, 4),
                                      Mid(Phone, 7, 4)
                                  )
        ElseIf GetWhichPhoneFormat(Phone) = PhoneFormat.SpecialPhone Then
            Phone = String.Format("{0}-{1}-{2}",
                          Mid(Phone, 1, 4),
                          Mid(Phone, 5, 3),
                          Mid(Phone, 8, 4)
                      )
        ElseIf GetWhichPhoneFormat(Phone) = PhoneFormat.CellPhone Then
            Phone = String.Format("({0}) {1} {2}-{3}",
                          Mid(Phone, 1, 2),
                          Mid(Phone, 3, 1),
                          Mid(Phone, 4, 4),
                          Mid(Phone, 8, 4)
                      )
        End If
        Return Phone
    End Function

    ''' <summary>
    ''' Determines the phone number type based on its format.
    ''' </summary>
    ''' <param name="PhoneNumber">The phone number value.</param>
    ''' <returns>The detected <see cref="PhoneFormat"/>.</returns>
    Public Shared Function GetWhichPhoneFormat(PhoneNumber As String) As PhoneFormat
        Dim Phone As String = PhoneNumber.Replace("(", Nothing).Replace(")", Nothing).Replace("-", Nothing).Replace(" ", Nothing)
        Dim SpecialPhones As New List(Of String) From {
                                                    "0300", "0500", "0800", "0900"
                                              }
        Dim FixedPhones As New List(Of String) From {
                                                        "11", "12", "13", "14", "15", "16", "17", "18", "19",
                                                        "21", "22", "24", "27", "28", "31", "32", "33", "34",
                                                        "35", "37", "38", "41", "42", "43", "44", "45", "46",
                                                        "47", "48", "49", "51", "53", "54", "55", "61", "62",
                                                        "63", "64", "65", "66", "67", "68", "69", "71", "73",
                                                        "74", "75", "77", "79", "81", "82", "83", "84", "85",
                                                        "86", "87", "88", "89", "91", "92", "93", "94", "95",
                                                         "96", "97", "98", "99"
                                                    }
        Dim CellPhones As New List(Of String) From {
                                                        "119", "129", "139", "149", "159", "169", "179", "189", "199",
                                                        "219", "229", "249", "279", "289", "319", "329", "339", "349",
                                                        "359", "379", "389", "419", "429", "439", "449", "459", "469",
                                                        "479", "489", "499", "519", "539", "549", "559", "619", "629",
                                                        "639", "649", "659", "669", "679", "689", "699", "719", "739",
                                                        "749", "759", "779", "799", "819", "829", "839", "849", "859",
                                                        "869", "879", "889", "899", "919", "929", "939", "949", "959",
                                                         "969", "979", "989", "999"}
        If Phone.Length = 10 And FixedPhones.Contains(Strings.Left(Phone, 2)) Then
            Return PhoneFormat.FixedPhone
        Else
            If Phone.Length = 11 And SpecialPhones.Contains(Strings.Left(Phone, 4)) Then
                Return PhoneFormat.SpecialPhone
            ElseIf Phone.Length = 11 And CellPhones.Contains(Strings.Left(Phone, 3)) Then
                Return PhoneFormat.CellPhone
            End If
        End If
        Return PhoneFormat.InvalidPhone
    End Function

    ''' <summary>
    ''' If the input string is a valid CPF or CNPJ, returns the value formatted with a mask.
    ''' </summary>
    ''' <param name="Document">The document value.</param>
    ''' <returns>The formatted document.</returns>
    Public Shared Function GetFormatedDocument(Document As String) As String
        Dim CleanDocument As String = Document.Replace(".", "").Replace("-", "").Replace("/", "")
        If IsNumeric(CleanDocument) Then
            If CleanDocument.Length = 14 Then
                CleanDocument = CleanDocument.Substring(0, 2) & "." & CleanDocument.Substring(2, 3) & "." & CleanDocument.Substring(5, 3) & "/" & CleanDocument.Substring(8, 4) & "-" & CleanDocument.Substring(12, 2)
            ElseIf CleanDocument.Length = 11 Then
                CleanDocument = CleanDocument.Substring(0, 3) & "." & CleanDocument.Substring(3, 3) & "." & CleanDocument.Substring(6, 3) & "-" & CleanDocument.Substring(9, 2)
            End If
            Return CleanDocument
        End If
        Return Document
    End Function

    ''' <summary>
    ''' Determines whether the provided string is a valid email address.
    ''' </summary>
    ''' <param name="Email">The email address.</param>
    ''' <returns>
    ''' <c>True</c> if the email address is valid; otherwise, <c>False</c>.
    ''' </returns>
    Public Shared Function IsValidEmail(Email As String) As Boolean
        Dim ValidEmailPattern As String = "^(?!\.)(""([^""\r\\]|\\[""\r\\])*""|" & "([-a-z0-9!#$%&'*+/=?^_{|}~]|(?<!\.)\.)*)(?<!\.)" & "@[a-z0-9][\w\.-]*[a-z0-9]\.[a-z][a-z\.]*[a-z]$"
        Return New Regex(ValidEmailPattern, RegexOptions.IgnoreCase).IsMatch(Email)
    End Function

    ''' <summary>
    ''' Determines whether the provided string is a valid Brazilian legal entity document (CNPJ).
    ''' </summary>
    ''' <param name="Document">The document value.</param>
    ''' <returns>
    ''' <c>True</c> if the document is valid; otherwise, <c>False</c>.
    ''' </returns>
    Public Shared Function IsValidLegalEntityDocument(Document As String) As Boolean
        Dim DadosArray() As String
        Dim Number(13) As Integer
        Dim Sum As Integer
        Dim Result1 As Integer
        Dim Result2 As Integer
        Dim FormatedDocument As String = GetFormatedDocument(Document)
        If IsNumeric(FormatedDocument.Replace(".", "").Replace("-", "").Replace("/", "")) Then
            If FormatedDocument.Length = 18 Then
                DadosArray = {"111.111.111-11", "222.222.222-22", "333.333.333-33", "444.444.444-44",
                                      "555.555.555-55", "666.666.666-66", "777.777.777-77", "888.888.888-88", "999.999.999-99"}
                FormatedDocument = Trim(FormatedDocument)
                For i = 0 To DadosArray.Length - 1
                    If FormatedDocument.Length <> 18 Or DadosArray(i).Equals(FormatedDocument) Then
                        Return False
                    End If
                Next
                FormatedDocument = FormatedDocument.Substring(0, 2) + FormatedDocument.Substring(3, 3) + FormatedDocument.Substring(7, 3) + FormatedDocument.Substring(11, 4) + FormatedDocument.Substring(16)
                If FormatedDocument = "00000000000000" Then Return False
                For i = 0 To Number.Length - 1
                    Number(i) = CInt(FormatedDocument.Substring(i, 1))
                Next
                Sum = Number(0) * 5 + Number(1) * 4 + Number(2) * 3 + Number(3) * 2 + Number(4) * 9 + Number(5) * 8 + Number(6) * 7 +
                       Number(7) * 6 + Number(8) * 5 + Number(9) * 4 + Number(10) * 3 + Number(11) * 2
                Sum -= (11 * (Int(Sum / 11)))
                If Sum = 0 Or Sum = 1 Then
                    Result1 = 0
                Else
                    Result1 = 11 - Sum
                End If
                If Result1 = Number(12) Then
                    Sum = Number(0) * 6 + Number(1) * 5 + Number(2) * 4 + Number(3) * 3 + Number(4) * 2 + Number(5) * 9 + Number(6) * 8 +
                                 Number(7) * 7 + Number(8) * 6 + Number(9) * 5 + Number(10) * 4 + Number(11) * 3 + Number(12) * 2
                    Sum -= (11 * (Int(Sum / 11)))
                    If Sum = 0 Or Sum = 1 Then
                        Result2 = 0
                    Else
                        Result2 = 11 - Sum
                    End If
                    If Result2 = Number(13) Then
                        Return True
                    Else
                        Return False
                    End If
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Determines whether the provided string is a valid Brazilian natural person document (CPF).
    ''' </summary>
    ''' <param name="Document">The document value.</param>
    ''' <returns>
    ''' <c>True</c> if the document is valid; otherwise, <c>False</c>.
    ''' </returns>
    Public Shared Function IsValidNaturalEntityDocument(Document As String) As Boolean
        Dim DadosArray() As String
        Dim N1, N2 As Integer
        Dim Number(13) As Integer
        Dim FormatedDocument As String = GetFormatedDocument(Document)
        If FormatedDocument.Length = 14 Then
            DadosArray = {"111.111.111-11", "222.222.222-22", "333.333.333-33", "444.444.444-44",
                                  "555.555.555-55", "666.666.666-66", "777.777.777-77", "888.888.888-88", "999.999.999-99"}
            FormatedDocument = Trim(FormatedDocument)
            For i = 0 To DadosArray.Length - 1
                If FormatedDocument.Length <> 14 Or DadosArray(i).Equals(FormatedDocument) Then
                    Return False
                End If
            Next
            FormatedDocument = FormatedDocument.Replace(".", Nothing).Replace("-", Nothing).Replace("/", Nothing)
            If FormatedDocument = "00000000000" Then Return False
            For x = 0 To 1
                N1 = 0
                For i = 0 To 8 + x
                    N1 += Val(FormatedDocument.Substring(i, 1)) * (10 + x - i)
                Next
                N2 = 11 - (N1 - (Int(N1 / 11) * 11))
                If N2 = 10 Or N2 = 11 Then N2 = 0
                If N2 <> Val(FormatedDocument.Substring(9 + x, 1)) Then
                    Return False
                End If
            Next
            Return True
        Else
            Return False
        End If
    End Function

End Class

''' <summary>
''' Defines the supported phone number formats.
''' </summary>
Public Enum PhoneFormat

    ''' <summary>
    ''' Fixed (landline) phone number.
    ''' </summary>
    FixedPhone

    ''' <summary>
    ''' Mobile (cellular) phone number.
    ''' </summary>
    CellPhone

    ''' <summary>
    ''' Special service phone number (e.g. 0800, 0300).
    ''' </summary>
    SpecialPhone

    ''' <summary>
    ''' Invalid or unsupported phone number format.
    ''' </summary>
    InvalidPhone

End Enum