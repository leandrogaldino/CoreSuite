Imports System.IO
Imports System.Reflection
Imports System.Runtime.Loader
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CSharp
Imports Microsoft.CodeAnalysis.VisualBasic
''' <summary>
''' Compiles and executes Visual Basic or C# source code at runtime using Roslyn.
''' </summary>
Public NotInheritable Class CodeCompiler
    Private ReadOnly _ReferencePaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _LoadPaths As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    ''' <summary>
    ''' Initializes a new instance of the <see cref="CodeCompiler"/> class.
    ''' </summary>
    ''' <remarks>
    ''' Platform assemblies and assemblies already loaded by the current application are automatically added as compilation references.
    ''' </remarks>
    Public Sub New()
        AddPlatformReferences()
        AddLoadedAssemblyReferences()
    End Sub
    ''' <summary>
    ''' Adds an assembly file as a compilation reference.
    ''' </summary>
    ''' <param name="AssemblyPath">The full or relative path of the assembly file.</param>
    ''' <exception cref="ArgumentException">The assembly path is empty.</exception>
    ''' <exception cref="FileNotFoundException">The assembly file does not exist.</exception>
    Public Sub AddReference(AssemblyPath As String)
        If String.IsNullOrWhiteSpace(AssemblyPath) Then Throw New ArgumentException("The assembly path cannot be empty.", NameOf(AssemblyPath))
        AddReferencePath(AssemblyPath, True)
    End Sub
    ''' <summary>
    ''' Adds a loaded assembly as a compilation reference.
    ''' </summary>
    ''' <param name="ReferenceAssembly">The assembly to reference.</param>
    ''' <exception cref="ArgumentNullException">The specified assembly is <see langword="Nothing"/>.</exception>
    ''' <exception cref="ArgumentException">The assembly does not have a physical file location.</exception>
    Public Sub AddReference(ReferenceAssembly As Assembly)
        ArgumentNullException.ThrowIfNull(ReferenceAssembly)
        If ReferenceAssembly.IsDynamic OrElse String.IsNullOrWhiteSpace(ReferenceAssembly.Location) Then Throw New ArgumentException("The assembly must have a valid physical location.", NameOf(ReferenceAssembly))
        Dim ResolveAtRuntime As Boolean = AssemblyLoadContext.GetLoadContext(ReferenceAssembly) IsNot AssemblyLoadContext.Default
        AddReferencePath(ReferenceAssembly.Location, ResolveAtRuntime)
    End Sub
    ''' <summary>
    ''' Adds the assembly that defines a type as a compilation reference.
    ''' </summary>
    ''' <param name="ReferenceType">A type defined by the assembly to reference.</param>
    ''' <exception cref="ArgumentNullException">The specified type is <see langword="Nothing"/>.</exception>
    Public Sub AddReference(ReferenceType As System.Type)
        ArgumentNullException.ThrowIfNull(ReferenceType)
        AddReference(ReferenceType.Assembly)
    End Sub
    ''' <summary>
    ''' Compiles source code and returns a loaded assembly that can execute its public methods.
    ''' </summary>
    ''' <param name="SourceCode">The Visual Basic or C# source code to compile.</param>
    ''' <param name="Language">The programming language used by the source code.</param>
    ''' <param name="AssemblyName">The optional name assigned to the generated assembly.</param>
    ''' <returns>A disposable object that represents the compiled assembly.</returns>
    ''' <exception cref="ArgumentException">The source code is empty.</exception>
    ''' <exception cref="ArgumentOutOfRangeException">The specified language is invalid.</exception>
    ''' <exception cref="CodeCompilationException">The source code contains compilation errors.</exception>
    Public Function Compile(SourceCode As String, Language As CodeLanguage, Optional AssemblyName As String = Nothing) As CompiledCode
        If String.IsNullOrWhiteSpace(SourceCode) Then Throw New ArgumentException("The source code cannot be empty.", NameOf(SourceCode))
        If Not [Enum].IsDefined(GetType(CodeLanguage), Language) Then Throw New ArgumentOutOfRangeException(NameOf(Language))
        If String.IsNullOrWhiteSpace(AssemblyName) Then AssemblyName = "DynamicCode_" & Guid.NewGuid().ToString("N")
        AddLoadedAssemblyReferences()
        Dim References As IEnumerable(Of MetadataReference) = _ReferencePaths.Select(Function(ReferencePath) MetadataReference.CreateFromFile(ReferencePath))
        Dim RoslynCompilation As Compilation
        Select Case Language
            Case CodeLanguage.VisualBasic
                Dim ParseOptions As New VisualBasicParseOptions(Microsoft.CodeAnalysis.VisualBasic.LanguageVersion.Latest)
                Dim SourceSyntaxTree As SyntaxTree = VisualBasicSyntaxTree.ParseText(SourceCode, ParseOptions, "DynamicCode.vb")
                Dim CompilationOptions As New VisualBasicCompilationOptions(OutputKind.DynamicallyLinkedLibrary, optimizationLevel:=OptimizationLevel.Release, optionStrict:=Microsoft.CodeAnalysis.VisualBasic.OptionStrict.On, concurrentBuild:=True)
                RoslynCompilation = VisualBasicCompilation.Create(AssemblyName, New SyntaxTree() {SourceSyntaxTree}, References, CompilationOptions)
            Case CodeLanguage.CSharp
                Dim ParseOptions As New CSharpParseOptions(Microsoft.CodeAnalysis.CSharp.LanguageVersion.Latest)
                Dim SourceSyntaxTree As SyntaxTree = CSharpSyntaxTree.ParseText(SourceCode, ParseOptions, "DynamicCode.cs")
                Dim CompilationOptions As New CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary, optimizationLevel:=OptimizationLevel.Release, concurrentBuild:=True)
                RoslynCompilation = CSharpCompilation.Create(AssemblyName, New SyntaxTree() {SourceSyntaxTree}, References, CompilationOptions)
            Case Else
                Throw New ArgumentOutOfRangeException(NameOf(Language))
        End Select
        Using AssemblyStream As New MemoryStream()
            Dim EmitResult = RoslynCompilation.Emit(AssemblyStream)
            If Not EmitResult.Success Then
                Dim Errors As String() = EmitResult.Diagnostics.Where(Function(Item) Item.Severity = DiagnosticSeverity.Error).Select(Function(Item) Item.ToString()).ToArray()
                Throw New CodeCompilationException(Errors)
            End If
            AssemblyStream.Position = 0
            Dim LoadContext As New CompilerAssemblyLoadContext(_LoadPaths)
            Try
                Dim RuntimeAssembly As Assembly = LoadContext.LoadFromStream(AssemblyStream)
                Return New CompiledCode(LoadContext, RuntimeAssembly)
            Catch
                LoadContext.Unload()
                Throw
            End Try
        End Using
    End Function
    ''' <summary>
    ''' Compiles source code, invokes a public method and releases the generated assembly.
    ''' </summary>
    ''' <param name="SourceCode">The Visual Basic or C# source code to compile.</param>
    ''' <param name="Language">The programming language used by the source code.</param>
    ''' <param name="TypeName">The full name of the type that contains the method.</param>
    ''' <param name="MethodName">The name of the method to invoke.</param>
    ''' <param name="Parameters">The arguments passed to the method.</param>
    ''' <returns>The value returned by the invoked method, or <see langword="Nothing"/> when the method does not return a value.</returns>
    Public Function Execute(SourceCode As String, Language As CodeLanguage, TypeName As String, MethodName As String, ParamArray Parameters As Object()) As Object
        Using RuntimeCode As CompiledCode = Compile(SourceCode, Language)
            Return RuntimeCode.Invoke(TypeName, MethodName, Parameters)
        End Using
    End Function
    ''' <summary>
    ''' Compiles a Visual Basic or C# source file.
    ''' </summary>
    ''' <param name="FilePath">The path of the .vb or .cs source file.</param>
    ''' <returns>A disposable object that represents the compiled assembly.</returns>
    Public Function CompileFile(FilePath As String) As CompiledCode
        ValidateSourceFile(FilePath)
        Return CompileSourceFile(FilePath, GetLanguageFromFileExtension(FilePath))
    End Function
    ''' <summary>
    ''' Compiles a source file using the specified programming language.
    ''' </summary>
    ''' <param name="FilePath">The source file path.</param>
    ''' <param name="Language">The programming language used by the source file.</param>
    ''' <returns>A disposable object that represents the compiled assembly.</returns>
    Public Function CompileFile(FilePath As String, Language As CodeLanguage) As CompiledCode
        ValidateSourceFile(FilePath)
        Return CompileSourceFile(FilePath, Language)
    End Function
    ''' <summary>
    ''' Compiles a Visual Basic or C# source file, invokes a public method and releases the generated assembly.
    ''' </summary>
    ''' <param name="FilePath">The path of the .vb or .cs source file.</param>
    ''' <param name="TypeName">The full name of the type that contains the method.</param>
    ''' <param name="MethodName">The name of the method to invoke.</param>
    ''' <param name="Parameters">The arguments passed to the method.</param>
    ''' <returns>The value returned by the invoked method, or <see langword="Nothing"/> when the method does not return a value.</returns>
    Public Function ExecuteFile(FilePath As String, TypeName As String, MethodName As String, ParamArray Parameters As Object()) As Object
        Using RuntimeCode As CompiledCode = CompileFile(FilePath)
            Return RuntimeCode.Invoke(TypeName, MethodName, Parameters)
        End Using
    End Function
    ''' <summary>
    ''' Compiles a source file, invokes a public method and releases the generated assembly.
    ''' </summary>
    ''' <param name="FilePath">The source file path.</param>
    ''' <param name="Language">The programming language used by the source file.</param>
    ''' <param name="TypeName">The full name of the type that contains the method.</param>
    ''' <param name="MethodName">The name of the method to invoke.</param>
    ''' <param name="Parameters">The arguments passed to the method.</param>
    ''' <returns>The value returned by the invoked method, or <see langword="Nothing"/> when the method does not return a value.</returns>
    Public Function ExecuteFile(FilePath As String, Language As CodeLanguage, TypeName As String, MethodName As String, ParamArray Parameters As Object()) As Object
        Using RuntimeCode As CompiledCode = CompileFile(FilePath, Language)
            Return RuntimeCode.Invoke(TypeName, MethodName, Parameters)
        End Using
    End Function
    Private Function CompileSourceFile(FilePath As String, Language As CodeLanguage) As CompiledCode
        Dim AssemblyName As String = Path.GetFileNameWithoutExtension(FilePath) & "_" & Guid.NewGuid().ToString("N")
        Return Compile(File.ReadAllText(FilePath), Language, AssemblyName)
    End Function
    Private Sub AddPlatformReferences()
        Dim TrustedPlatformAssemblies As String = TryCast(AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES"), String)
        If Not String.IsNullOrWhiteSpace(TrustedPlatformAssemblies) Then
            For Each AssemblyPath As String In TrustedPlatformAssemblies.Split(New Char() {Path.PathSeparator}, StringSplitOptions.RemoveEmptyEntries)
                AddReferencePath(AssemblyPath, False)
            Next AssemblyPath
            Return
        End If
        AddReferencePath(GetType(Object).Assembly.Location, False)
        AddReferencePath(GetType(Enumerable).Assembly.Location, False)
        AddReferencePath(GetType(Console).Assembly.Location, False)
    End Sub
    Private Sub AddLoadedAssemblyReferences()
        For Each LoadedAssembly As Assembly In AppDomain.CurrentDomain.GetAssemblies()
            Try
                If LoadedAssembly.IsDynamic OrElse String.IsNullOrWhiteSpace(LoadedAssembly.Location) Then Continue For
                Dim ResolveAtRuntime As Boolean = AssemblyLoadContext.GetLoadContext(LoadedAssembly) IsNot AssemblyLoadContext.Default
                AddReferencePath(LoadedAssembly.Location, ResolveAtRuntime)
            Catch Ex As NotSupportedException
            End Try
        Next LoadedAssembly
    End Sub
    Private Sub AddReferencePath(AssemblyPath As String, ResolveAtRuntime As Boolean)
        Dim FullPath As String = Path.GetFullPath(AssemblyPath)
        If Not File.Exists(FullPath) Then Throw New FileNotFoundException("The referenced assembly was not found.", FullPath)
        _ReferencePaths.Add(FullPath)
        If Not ResolveAtRuntime Then Return
        Dim ReferenceName As AssemblyName = System.Reflection.AssemblyName.GetAssemblyName(FullPath)
        If String.IsNullOrWhiteSpace(ReferenceName.Name) Then Throw New BadImageFormatException("The referenced file does not contain a valid .NET assembly.", FullPath)
        _LoadPaths(ReferenceName.Name) = FullPath
    End Sub
    Private Shared Function GetLanguageFromFileExtension(FilePath As String) As CodeLanguage
        Select Case Path.GetExtension(FilePath).ToLowerInvariant()
            Case ".vb"
                Return CodeLanguage.VisualBasic
            Case ".cs"
                Return CodeLanguage.CSharp
            Case Else
                Throw New NotSupportedException("Only .vb and .cs source files are supported.")
        End Select
    End Function
    Private Shared Sub ValidateSourceFile(FilePath As String)
        If String.IsNullOrWhiteSpace(FilePath) Then Throw New ArgumentException("The source file path cannot be empty.", NameOf(FilePath))
        If Not File.Exists(FilePath) Then Throw New FileNotFoundException("The source file was not found.", FilePath)
    End Sub
End Class