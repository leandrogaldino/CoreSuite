Imports System.IO
Imports System.Reflection

' .NET Framework cannot unload one assembly while exposing its live Assembly/types.
' Preserve direct invocation and object identity; dispose only removes this resolver.
Friend NotInheritable Class CompiledAssemblyScope
    Implements IDisposable
    Private ReadOnly _LoadPaths As Dictionary(Of String, String)
    Private ReadOnly _SearchDirectories As String()
    Private ReadOnly _OwnedAssemblies As New HashSet(Of Assembly)
    Public Sub New(LoadPaths As IReadOnlyDictionary(Of String, String))
        _LoadPaths = LoadPaths.ToDictionary(Function(Item) Item.Key, Function(Item) Item.Value, StringComparer.OrdinalIgnoreCase)
        _SearchDirectories = _LoadPaths.Values.Select(Function(Item) Path.GetDirectoryName(Item)).Where(Function(Item) Not String.IsNullOrWhiteSpace(Item)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray()
        AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf Resolve
    End Sub
    Public Function LoadFromStream(Stream As MemoryStream) As Assembly
        Dim Result = Assembly.Load(Stream.ToArray())
        SyncLock _OwnedAssemblies
            _OwnedAssemblies.Add(Result)
        End SyncLock
        Return Result
    End Function
    Private Function Resolve(Sender As Object, Args As ResolveEventArgs) As Assembly
        SyncLock _OwnedAssemblies
            If Args.RequestingAssembly Is Nothing OrElse Not _OwnedAssemblies.Contains(Args.RequestingAssembly) Then Return Nothing
        End SyncLock
        Dim Requested = New AssemblyName(Args.Name)
        Dim SharedAssembly = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(Function(Item) AssemblyName.ReferenceMatchesDefinition(Item.GetName(), Requested))
        If SharedAssembly IsNot Nothing Then Return SharedAssembly
        Dim AssemblyPath As String = Nothing
        If Not _LoadPaths.TryGetValue(Requested.Name, AssemblyPath) Then
            AssemblyPath = _SearchDirectories.Select(Function(Directory) Path.Combine(Directory, Requested.Name & ".dll")).FirstOrDefault(Function(Candidate) File.Exists(Candidate))
        End If
        If AssemblyPath Is Nothing Then Return Nothing
        Dim Result = Assembly.LoadFrom(AssemblyPath)
        SyncLock _OwnedAssemblies
            _OwnedAssemblies.Add(Result)
        End SyncLock
        Return Result
    End Function
    Public Sub Dispose() Implements IDisposable.Dispose
        RemoveHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf Resolve
        SyncLock _OwnedAssemblies
            _OwnedAssemblies.Clear()
        End SyncLock
    End Sub
End Class
