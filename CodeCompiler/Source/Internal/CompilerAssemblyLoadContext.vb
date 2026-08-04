Imports System.IO
Imports System.Reflection
Imports System.Runtime.Loader

''' <summary>
''' Provides an isolated and collectible load context for dynamically compiled assemblies.
''' </summary>
Friend NotInheritable Class CompilerAssemblyLoadContext
    Inherits AssemblyLoadContext
    Private ReadOnly _LoadPaths As Dictionary(Of String, String)
    Private ReadOnly _SearchDirectories As String()
    ''' <summary>
    ''' Initializes a new collectible assembly load context.
    ''' </summary>
    ''' <param name="LoadPaths">The assembly paths available for dependency resolution.</param>
    Public Sub New(LoadPaths As IReadOnlyDictionary(Of String, String))
        MyBase.New("CodeCompiler_" & Guid.NewGuid().ToString("N"), True)
        _LoadPaths = LoadPaths.ToDictionary(Function(Item) Item.Key, Function(Item) Item.Value, StringComparer.OrdinalIgnoreCase)
        _SearchDirectories = _LoadPaths.Values.Select(Function(Item) Path.GetDirectoryName(Item)).Where(Function(Item) Not String.IsNullOrWhiteSpace(Item)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray()
    End Sub
    ''' <summary>
    ''' Resolves an assembly requested by the dynamically compiled assembly.
    ''' </summary>
    ''' <param name="RequestedAssemblyName">The identity of the requested assembly.</param>
    ''' <returns>The resolved assembly, or <see langword="Nothing"/> when it cannot be resolved.</returns>
    Protected Overrides Function Load(RequestedAssemblyName As AssemblyName) As Assembly
        Dim SharedAssembly As Assembly = AssemblyLoadContext.Default.Assemblies.FirstOrDefault(Function(Item) AssemblyName.ReferenceMatchesDefinition(Item.GetName(), RequestedAssemblyName))
        If SharedAssembly IsNot Nothing Then Return SharedAssembly
        If String.IsNullOrWhiteSpace(RequestedAssemblyName.Name) Then Return Nothing
        Dim AssemblyPath As String = Nothing
        If _LoadPaths.TryGetValue(RequestedAssemblyName.Name, AssemblyPath) Then Return LoadFromAssemblyPath(AssemblyPath)
        For Each DirectoryPath As String In _SearchDirectories
            Dim CandidatePath As String = Path.Combine(DirectoryPath, RequestedAssemblyName.Name & ".dll")
            If File.Exists(CandidatePath) Then Return LoadFromAssemblyPath(CandidatePath)
        Next DirectoryPath
        Return Nothing
    End Function
End Class