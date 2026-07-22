Imports System.ComponentModel
Imports System.Resources

''' <summary>
''' Busca o nome da categoria diretamente do arquivo de recursos (.resx) do projeto.
''' </summary>
Public Class LocalizedCategoryAttribute
    Inherits CategoryAttribute

    Public Sub New(resourceKey As String)
        MyBase.New(resourceKey)
    End Sub

    Protected Overrides Function GetLocalizedString(value As String) As String
        If String.IsNullOrEmpty(value) Then Return value

        Try
            ' SUBSTITUA "Mensagens" PELO NOME DO SEU ARQUIVO DE RECURSOS (.resx)
            ' Exemplo: Se seu recurso chama "Resources.resx", use GetType(Resources)
            Dim rm As New ResourceManager(GetType(Globalization))

            Dim localizedValue As String = rm.GetString(value)

            ' Se encontrou no .resx, retorna a tradução; senão, mantém o texto original
            If Not String.IsNullOrEmpty(localizedValue) Then
                Return localizedValue
            End If
        Catch
            ' Fallback em caso de erro na leitura do arquivo de recursos
        End Try

        Return MyBase.GetLocalizedString(value)
    End Function
End Class