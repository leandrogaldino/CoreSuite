Imports System.CodeDom.Compiler
Imports System.Reflection
Imports System.IO
Public Class CodeCompiler
    Implements IDisposable
    Private _CodeCompiler As CodeDomProvider
    Private _CompilerAssemblies As New CompilerParameters
    Private _CompileResults As CompilerResults
    Private _StrCode As String
    Private _Code As String
    Private _File As FileInfo
    Private ReadOnly _Path As String
    Private _CompilerHelper As CompilerHelper
    Private _DisposedValue As Boolean

    ''' <summary>
    ''' Linguagens suportadas pelo compilador.
    ''' </summary>
    Public Enum [Languages]
        CPlusPlus
        CSharp
        JavaScript
        VisualBasic
    End Enum


    ''' <summary>
    ''' Compila um código em tempo de execução.
    ''' </summary>
    ''' <param name="Code">O código a ser compilado.</param>
    ''' <param name="CompilerHelper">O objeto que fornece todas as informações necessárias para compilar o código.</param>
    Public Sub New(ByVal Code As String, ByVal CompilerHelper As CompilerHelper)
        _CompilerHelper = CompilerHelper
        _Code = Code
        _File = Nothing
        _Path = Nothing
    End Sub

    ''' <summary>
    ''' Compila um código em tempo de execução.
    ''' </summary>
    ''' <param name="File">O arquivo que contém o código a ser compilado.</param>
    ''' <param name="CompilerHelper">O objeto que fornece todas as informações necessárias para compilar o código.</param>
    Public Sub New(ByVal File As FileInfo, ByVal CompilerHelper As CompilerHelper)
        _CompilerHelper = CompilerHelper
        _File = File
        _Code = Nothing
        _Path = Nothing
    End Sub

    ''' <summary>
    ''' Compila o código.
    ''' </summary>
    ''' <returns></returns>
    Public Function Compile() As Object


        'Instanciando o compilador em na linguagem informada no CompilarHelper.
        _CodeCompiler = CodeDomProvider.CreateProvider(GetLanguageString(_CompilerHelper.Language))

        ''Passando as referencias informadas no CompilerHelper para serem utilizadas pelo compilados. Todas as referências utilizadas no código devem estar contidas aqui
        ''Para referenciar a propria o proprio aplicativo, utilizar: Assembly GetEntryAssembly() Location
        ''_CompilerAssemblies = New CompilerParameters
        _CompilerAssemblies = New CompilerParameters
        For Each Reference In _CompilerHelper.ReferencedAssemblies
            _CompilerAssemblies.ReferencedAssemblies.Add(Reference)
        Next Reference

        ''_CompilerAssemblies = _CompilerHelper.ReferencedAssemblies.Distinct

        ''Especifica se o código será compilado na memória.
        _CompilerAssemblies.GenerateInMemory = _CompilerHelper.GenerateInMemory

        ''Obtem o código a ser compilado.
        If _Code IsNot Nothing Then
            _StrCode = _Code
        ElseIf _Path IsNot Nothing Then
            _StrCode = IO.File.ReadAllText(_Path)
        Else
            _StrCode = _File.OpenText.ReadToEnd()
        End If

        'Instancia o objeto que receberá o retorno do código e parâmetros passados para o compilador. Esse objeto será utilizado para verificar possíveis erros no código.
        _CompileResults = _CodeCompiler.CompileAssemblyFromSource(_CompilerAssemblies, _StrCode)

        'Retorna um erro caso alguma linha do código tenha erro.
        If _CompileResults.Errors.HasErrors Then
            Throw New Exception("O código possui um erro na linha " & _CompileResults.Errors(0).Line.ToString & ", " & _CompileResults.Errors(0).ErrorText)
        End If

        'Instancia o objeto que recebera o Assembly referente ao código compilado.
        Dim objAssembly As System.Reflection.Assembly = _CompileResults.CompiledAssembly

        'Cria um objeto que recebe a instancia da classe referenciada no código.
        Dim objTheClass As Object = objAssembly.CreateInstance(_CompilerHelper.ClassName)

        'Retorna um erro caso a classe especificada no CompilerHelper não exista no código.
        If objTheClass Is Nothing Then
            Throw New Exception(String.Format("A classe {0} não foi encontrada no código.{1}{2},{3}", _CompilerHelper.ClassName, vbCrLf, _CompileResults.Errors(0).Line.ToString, _CompileResults.Errors(0).ErrorText))
        End If

        'Variavel que irá retornar o objeto do código a ser compilado (se houver um)
        Dim Result As Object
        'Executa o código, e atribui o retorno dele (se houver) a variavel Result.
        Try
            Result = objTheClass.GetType.InvokeMember(_CompilerHelper.MethodName, BindingFlags.InvokeMethod, Nothing, objTheClass, _CompilerHelper.Parameters)
        Catch ex As Exception
            If ex.InnerException IsNot Nothing Then
                'Se algum método chamado dentro do método principal do relatório resultar em um erro, essa exceção será gerada.
                Throw New Exception("O código retornou um erro: " & ex.InnerException.Message)
            Else
                'Se o método informado no _CompilerHelper não existir, essa exceção será gerada.
                Throw New Exception(String.Format("O método {0}.{1} não foi encontrado no código", _CompilerHelper.ClassName, _CompilerHelper.MethodName))
            End If
        End Try

        'Retorna o objeto resultante do código (se houver).
        Return Result
    End Function

    'Retorna a linguagem na propriedade Language do CompilerHelper.
    Private Function GetLanguageString(ByVal Language As Languages) As String
        Select Case Language
            Case Is = Languages.CPlusPlus
                Return "CPP"
            Case Is = Languages.CSharp
                Return "CSharp"
            Case Is = Languages.JavaScript
                Return "JavaScript"
            Case Else
                Return "VisualBasic"
        End Select
    End Function

    'Classe para passar as informações para o compilador.
    Public Class CompilerHelper
        Public Property Language As Languages
        Public Property ReferencedAssemblies As String()
        Public Property Parameters As Object()
        Public Property GenerateInMemory As Boolean
        Public Property ClassName As String
        Public Property MethodName As String
    End Class

    'Implementação do Dispose.
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not _DisposedValue Then
            If disposing Then
                _CompilerAssemblies = Nothing
                _StrCode = Nothing
                _CompileResults = Nothing
                _Code = Nothing
                _File = Nothing
                _CompilerHelper = Nothing
                _CodeCompiler.Dispose()
            End If
        End If
        _DisposedValue = True
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
End Class
