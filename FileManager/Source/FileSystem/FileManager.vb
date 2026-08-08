Imports System.IO
Imports System.Threading
''' <summary>
''' Provides asynchronous services for copying and deleting files and directories with progress reporting and cancellation support.
''' </summary>
''' <remarks>
''' <para>
''' The service performs potentially expensive directory enumeration and deletion operations on background threads and uses asynchronous file streams when copying file contents.
''' </para>
''' <para>
''' Progress can be reported through optional <see cref="IProgress(Of Integer)"/> parameters. A <see cref="Progress(Of Integer)"/> instance captures the caller's synchronization context when it is created, allowing UI callers to receive progress updates on their original context.
''' </para>
''' <para>
''' Directory enumeration skips reparse points to avoid following symbolic links, junctions, and other redirected directory structures.
''' </para>
''' </remarks>
Public NotInheritable Class FileManager
    ''' <summary>
    ''' Prevents creation of <see cref="FileManager"/> instances because the service exposes only shared operations.
    ''' </summary>
    Private Sub New()
    End Sub
    ''' <summary>
    ''' Defines the buffer size, in bytes, used when asynchronously reading and writing file streams.
    ''' </summary>
    ''' <remarks>
    ''' The configured value corresponds to 128 KB.
    ''' </remarks>
    Private Const StreamBufferSize As Integer = 131072
    ''' <summary>
    ''' Provides the path comparer appropriate for the current operating system.
    ''' </summary>
    ''' <remarks>
    ''' Path comparisons are case-insensitive on Windows and case-sensitive on other supported operating systems.
    ''' </remarks>
    Private Shared ReadOnly FileSystemPathComparer As StringComparer = If(OperatingSystem.IsWindows(), StringComparer.OrdinalIgnoreCase, StringComparer.Ordinal)
    ''' <summary>
    ''' Defines the string comparison mode used when determining path ancestry and path equality.
    ''' </summary>
    ''' <remarks>
    ''' Path comparisons are case-insensitive on Windows and case-sensitive on other supported operating systems.
    ''' </remarks>
    Private Shared ReadOnly FileSystemPathComparison As StringComparison = If(OperatingSystem.IsWindows(), StringComparison.OrdinalIgnoreCase, StringComparison.Ordinal)
    ''' <summary>
    ''' Asynchronously deletes the specified files and reports the accumulated deletion progress.
    ''' </summary>
    ''' <param name="Files">
    ''' The collection of files to delete. Duplicate paths are processed only once.
    ''' </param>
    ''' <param name="Progress">
    ''' An optional progress reporter that receives completion percentages from 0 through 100.
    ''' </param>
    ''' <param name="CancellationToken">
    ''' A token that may be used to cancel the operation.
    ''' </param>
    ''' <returns>
    ''' A task that represents the asynchronous deletion operation.
    ''' </returns>
    ''' <exception cref="ArgumentNullException">
    ''' <paramref name="Files"/> is <see langword="Nothing"/>.
    ''' </exception>
    ''' <exception cref="ArgumentException">
    ''' The collection contains a <see langword="Nothing"/> item or an invalid path.
    ''' </exception>
    ''' <exception cref="FileNotFoundException">
    ''' One of the specified files does not exist when the deletion plan is created.
    ''' </exception>
    ''' <exception cref="OperationCanceledException">
    ''' The operation is canceled through <paramref name="CancellationToken"/>.
    ''' </exception>
    ''' <exception cref="UnauthorizedAccessException">
    ''' The caller does not have permission to access or delete one of the files.
    ''' </exception>
    ''' <exception cref="IOException">
    ''' An I/O error occurs while accessing or deleting a file.
    ''' </exception>
    Public Shared Async Function DeleteFilesAsync(Files As IEnumerable(Of FileInfo), Optional Progress As IProgress(Of Integer) = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task
        ArgumentNullException.ThrowIfNull(Files)
        Dim FileList As List(Of FileInfo) = GetDistinctFiles(Files)
        Dim Entries As List(Of FileDeleteEntry) = Await Task.Run(Function()
                                                                     Dim Results As New List(Of FileDeleteEntry)
                                                                     For Each CurrentFile As FileInfo In FileList
                                                                         CancellationToken.ThrowIfCancellationRequested()
                                                                         CurrentFile.Refresh()
                                                                         If Not CurrentFile.Exists Then Throw New FileNotFoundException($"File '{CurrentFile.FullName}' was not found.", CurrentFile.FullName)
                                                                         Results.Add(New FileDeleteEntry(NormalizePath(CurrentFile.FullName), CurrentFile.Length))
                                                                     Next
                                                                     Return Results
                                                                 End Function, CancellationToken)
        Await Task.Run(Sub()
                           Dim TotalSize As Long = Entries.Sum(Function(CurrentFile) CurrentFile.Length)
                           Dim HandledSize As Long = 0
                           Dim ProcessedItems As Long = 0
                           Dim Reporter As New ProgressThrottle()
                           For Each CurrentFile As FileDeleteEntry In Entries
                               CancellationToken.ThrowIfCancellationRequested()
                               File.Delete(CurrentFile.Path)
                               HandledSize += CurrentFile.Length
                               ProcessedItems += 1
                               If Reporter.ShouldReport(ProcessedItems = Entries.Count) Then ReportProgress(Progress, TotalSize, HandledSize, ProcessedItems, Entries.Count)
                           Next
                           If Entries.Count = 0 Then ReportProgress(Progress, 0, 0, 0, 0)
                       End Sub, CancellationToken)
    End Function
    ''' <summary>
    ''' Asynchronously copies a file to the specified destination and reports copy progress.
    ''' </summary>
    ''' <param name="Source">
    ''' Information describing the source file to copy.
    ''' </param>
    ''' <param name="Destination">
    ''' Information describing the destination file to create or overwrite.
    ''' </param>
    ''' <param name="Progress">
    ''' An optional progress reporter that receives completion percentages from 0 through 100.
    ''' </param>
    ''' <param name="CancellationToken">
    ''' A token that may be used to cancel the operation.
    ''' </param>
    ''' <returns>
    ''' A task that represents the asynchronous file copy operation.
    ''' </returns>
    ''' <remarks>
    ''' The destination parent directory is created automatically when it does not already exist. An existing destination file is overwritten.
    ''' </remarks>
    ''' <exception cref="ArgumentNullException">
    ''' <paramref name="Source"/> or <paramref name="Destination"/> is <see langword="Nothing"/>.
    ''' </exception>
    ''' <exception cref="FileNotFoundException">
    ''' The source file does not exist.
    ''' </exception>
    ''' <exception cref="IOException">
    ''' The source and destination identify the same file, or another I/O error occurs during the copy.
    ''' </exception>
    ''' <exception cref="UnauthorizedAccessException">
    ''' The caller does not have permission to read the source or write the destination.
    ''' </exception>
    ''' <exception cref="OperationCanceledException">
    ''' The operation is canceled through <paramref name="CancellationToken"/>.
    ''' </exception>
    Public Shared Async Function CopyFileAsync(Source As FileInfo, Destination As FileInfo, Optional Progress As IProgress(Of Integer) = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task
        ArgumentNullException.ThrowIfNull(Source)
        ArgumentNullException.ThrowIfNull(Destination)
        Source.Refresh()
        If Not Source.Exists Then Throw New FileNotFoundException($"Source file '{Source.FullName}' was not found.", Source.FullName)
        Dim SourcePath As String = NormalizePath(Source.FullName)
        Dim DestinationPath As String = NormalizePath(Destination.FullName)
        If AreSamePath(SourcePath, DestinationPath) Then Throw New IOException("The source and destination files cannot be the same.")
        Dim TotalSize As Long = Source.Length
        Dim Reporter As New ProgressThrottle()
        Dim CopiedSize As Long = Await CopyFileCoreAsync(SourcePath, DestinationPath, Sub(CurrentCopiedSize As Long)
                                                                                          If Reporter.ShouldReport() Then ReportProgress(Progress, TotalSize, CurrentCopiedSize, 0, 1)
                                                                                      End Sub, CancellationToken)
        ReportProgress(Progress, TotalSize, CopiedSize, 1, 1)
    End Function
    ''' <summary>
    ''' Asynchronously copies multiple directory trees and reports their combined progress.
    ''' </summary>
    ''' <param name="Directories">
    ''' The collection of source and destination directory mappings to process.
    ''' </param>
    ''' <param name="Progress">
    ''' An optional progress reporter that receives completion percentages from 0 through 100.
    ''' </param>
    ''' <param name="CancellationToken">
    ''' A token that may be used to cancel the operation.
    ''' </param>
    ''' <returns>
    ''' A task that represents the asynchronous directory copy operation.
    ''' </returns>
    ''' <remarks>
    ''' <para>
    ''' A complete copy plan is created before file contents are copied so that the total byte count can be calculated.
    ''' </para>
    ''' <para>
    ''' Empty directories are created at the destination. Existing destination files are overwritten.
    ''' </para>
    ''' </remarks>
    ''' <exception cref="ArgumentNullException">
    ''' <paramref name="Directories"/> is <see langword="Nothing"/>, or a request has no source or destination.
    ''' </exception>
    ''' <exception cref="ArgumentException">
    ''' The collection contains a <see langword="Nothing"/> request or an invalid path.
    ''' </exception>
    ''' <exception cref="DirectoryNotFoundException">
    ''' A source directory does not exist.
    ''' </exception>
    ''' <exception cref="IOException">
    ''' A destination is equal to or located inside its corresponding source directory, or another I/O error occurs.
    ''' </exception>
    ''' <exception cref="UnauthorizedAccessException">
    ''' The caller does not have permission to enumerate, read, create, or write one of the directories or files.
    ''' </exception>
    ''' <exception cref="OperationCanceledException">
    ''' The operation is canceled through <paramref name="CancellationToken"/>.
    ''' </exception>
    Public Shared Async Function CopyDirectoriesAsync(Directories As IEnumerable(Of CopyDirectoryInfo), Optional Progress As IProgress(Of Integer) = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task
        ArgumentNullException.ThrowIfNull(Directories)
        Dim Requests As List(Of CopyDirectoryInfo) = Directories.ToList()
        If Requests.Any(Function(CurrentRequest) CurrentRequest Is Nothing) Then Throw New ArgumentException("The directory collection cannot contain null items.", NameOf(Directories))
        Dim Plans As List(Of DirectoryCopyPlan) = Await Task.Run(Function()
                                                                     Dim Results As New List(Of DirectoryCopyPlan)
                                                                     For Each CurrentRequest As CopyDirectoryInfo In Requests
                                                                         CancellationToken.ThrowIfCancellationRequested()
                                                                         Results.Add(BuildDirectoryCopyPlan(CurrentRequest, CancellationToken))
                                                                     Next
                                                                     Return Results
                                                                 End Function, CancellationToken)
        Dim TotalSize As Long = Plans.Sum(Function(CurrentPlan) CurrentPlan.TotalSize)
        Await ExecuteDirectoryCopyPlansAsync(Plans, TotalSize, 0, Progress, CancellationToken)
    End Function
    ''' <summary>
    ''' Asynchronously copies a single directory tree and returns the number of bytes copied.
    ''' </summary>
    ''' <param name="CopyInfo">
    ''' The source and destination directory mapping to process.
    ''' </param>
    ''' <param name="TotalSize">
    ''' The optional total size, in bytes, of a larger combined operation. When zero, the method calculates the effective total from the current copy plan.
    ''' </param>
    ''' <param name="HandledSize">
    ''' The optional number of bytes already handled by a larger combined operation.
    ''' </param>
    ''' <param name="Progress">
    ''' An optional progress reporter that receives completion percentages from 0 through 100.
    ''' </param>
    ''' <param name="CancellationToken">
    ''' A token that may be used to cancel the operation.
    ''' </param>
    ''' <returns>
    ''' A task whose result contains the total number of bytes copied for the specified directory.
    ''' </returns>
    ''' <remarks>
    ''' The <paramref name="TotalSize"/> and <paramref name="HandledSize"/> parameters allow the method to participate in a larger progress calculation.
    ''' </remarks>
    ''' <exception cref="ArgumentNullException">
    ''' <paramref name="CopyInfo"/> is <see langword="Nothing"/>, or its source or destination is <see langword="Nothing"/>.
    ''' </exception>
    ''' <exception cref="ArgumentOutOfRangeException">
    ''' <paramref name="TotalSize"/> or <paramref name="HandledSize"/> is negative.
    ''' </exception>
    ''' <exception cref="DirectoryNotFoundException">
    ''' The source directory does not exist.
    ''' </exception>
    ''' <exception cref="IOException">
    ''' The destination is equal to or located inside the source directory, or another I/O error occurs.
    ''' </exception>
    ''' <exception cref="UnauthorizedAccessException">
    ''' The caller does not have permission to enumerate, read, create, or write one of the directories or files.
    ''' </exception>
    ''' <exception cref="OperationCanceledException">
    ''' The operation is canceled through <paramref name="CancellationToken"/>.
    ''' </exception>
    Public Shared Async Function CopyDirectoryAsync(CopyInfo As CopyDirectoryInfo, Optional TotalSize As Long = 0, Optional HandledSize As Long = 0, Optional Progress As IProgress(Of Integer) = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task(Of Long)
        ArgumentNullException.ThrowIfNull(CopyInfo)
        If TotalSize < 0 Then Throw New ArgumentOutOfRangeException(NameOf(TotalSize), "The total size cannot be negative.")
        If HandledSize < 0 Then Throw New ArgumentOutOfRangeException(NameOf(HandledSize), "The handled size cannot be negative.")
        Dim Plan As DirectoryCopyPlan = Await Task.Run(Function() BuildDirectoryCopyPlan(CopyInfo, CancellationToken), CancellationToken)
        Dim EffectiveTotalSize As Long = If(TotalSize > 0, TotalSize, HandledSize + Plan.TotalSize)
        Return Await ExecuteDirectoryCopyPlansAsync(New List(Of DirectoryCopyPlan) From {Plan}, EffectiveTotalSize, HandledSize, Progress, CancellationToken)
    End Function
    ''' <summary>
    ''' Asynchronously deletes the contents of multiple directories and optionally deletes their root directories.
    ''' </summary>
    ''' <param name="Directories">
    ''' The collection of directory deletion requests to process.
    ''' </param>
    ''' <param name="Progress">
    ''' An optional progress reporter that receives completion percentages from 0 through 100.
    ''' </param>
    ''' <param name="CancellationToken">
    ''' A token that may be used to cancel the operation.
    ''' </param>
    ''' <returns>
    ''' A task that represents the asynchronous directory deletion operation.
    ''' </returns>
    ''' <remarks>
    ''' <para>
    ''' Files are deleted before directories. Child directories are then deleted from deepest to shallowest so that each directory is empty before removal.
    ''' </para>
    ''' <para>
    ''' Reparse points are not recursively enumerated.
    ''' </para>
    ''' </remarks>
    ''' <exception cref="ArgumentNullException">
    ''' <paramref name="Directories"/> is <see langword="Nothing"/>, or a request does not specify a directory.
    ''' </exception>
    ''' <exception cref="ArgumentException">
    ''' The collection contains a <see langword="Nothing"/> request, an invalid path, or overlapping root directories.
    ''' </exception>
    ''' <exception cref="UnauthorizedAccessException">
    ''' The caller does not have permission to enumerate or delete one of the files or directories.
    ''' </exception>
    ''' <exception cref="IOException">
    ''' An I/O error occurs while enumerating or deleting a file or directory.
    ''' </exception>
    ''' <exception cref="OperationCanceledException">
    ''' The operation is canceled through <paramref name="CancellationToken"/>.
    ''' </exception>
    Public Shared Async Function DeleteDirectoriesAsync(Directories As IEnumerable(Of DeleteDirectoryInfo), Optional Progress As IProgress(Of Integer) = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task
        ArgumentNullException.ThrowIfNull(Directories)
        Dim Requests As List(Of DeleteDirectoryInfo) = Directories.ToList()
        If Requests.Any(Function(CurrentRequest) CurrentRequest Is Nothing) Then Throw New ArgumentException("The directory collection cannot contain null items.", NameOf(Directories))
        Dim Plan As DirectoryDeletePlan = Await Task.Run(Function() BuildDirectoryDeletePlan(Requests, CancellationToken), CancellationToken)
        Await Task.Run(Sub()
                           Dim HandledSize As Long = 0
                           Dim ProcessedItems As Long = 0
                           Dim Reporter As New ProgressThrottle()
                           For Each CurrentFile As FileDeleteEntry In Plan.Files
                               CancellationToken.ThrowIfCancellationRequested()
                               File.Delete(CurrentFile.Path)
                               HandledSize += CurrentFile.Length
                               ProcessedItems += 1
                               If Reporter.ShouldReport(ProcessedItems = Plan.Files.Count) Then ReportProgress(Progress, Plan.TotalSize, HandledSize, ProcessedItems, Plan.Files.Count)
                           Next
                           For Each CurrentDirectory As String In Plan.Directories.OrderByDescending(Function(DirectoryPath) DirectoryPath.Length)
                               CancellationToken.ThrowIfCancellationRequested()
                               If System.IO.Directory.Exists(CurrentDirectory) Then System.IO.Directory.Delete(CurrentDirectory, False)
                           Next
                           For Each CurrentRoot As DirectoryDeleteRoot In Plan.Roots
                               CancellationToken.ThrowIfCancellationRequested()
                               If CurrentRoot.DeleteRoot AndAlso System.IO.Directory.Exists(CurrentRoot.Path) Then System.IO.Directory.Delete(CurrentRoot.Path, False)
                           Next
                           ReportProgress(Progress, Plan.TotalSize, HandledSize, Plan.Files.Count, Plan.Files.Count)
                       End Sub, CancellationToken)
    End Function
    ''' <summary>
    ''' Asynchronously deletes the contents of a directory while preserving explicitly excluded files and directories.
    ''' </summary>
    ''' <param name="Directory">
    ''' The root directory whose contents will be deleted.
    ''' </param>
    ''' <param name="ExceptDirectories">
    ''' An optional collection of directories to preserve. The required ancestor directories are also preserved.
    ''' </param>
    ''' <param name="ExceptFiles">
    ''' An optional collection of files to preserve. The required parent directories are also preserved.
    ''' </param>
    ''' <param name="Progress">
    ''' An optional progress reporter that receives completion percentages from 0 through 100.
    ''' </param>
    ''' <param name="CancellationToken">
    ''' A token that may be used to cancel the operation.
    ''' </param>
    ''' <returns>
    ''' A task that represents the asynchronous content deletion operation.
    ''' </returns>
    ''' <remarks>
    ''' The root directory itself is never deleted. If the root directory is included in <paramref name="ExceptDirectories"/>, the method returns without deleting any content.
    ''' </remarks>
    ''' <exception cref="ArgumentNullException">
    ''' <paramref name="Directory"/> is <see langword="Nothing"/>.
    ''' </exception>
    ''' <exception cref="DirectoryNotFoundException">
    ''' The root directory does not exist.
    ''' </exception>
    ''' <exception cref="ArgumentException">
    ''' An exclusion collection contains a <see langword="Nothing"/> item, an invalid path, or an item located outside the root directory.
    ''' </exception>
    ''' <exception cref="UnauthorizedAccessException">
    ''' The caller does not have permission to enumerate or delete one of the files or directories.
    ''' </exception>
    ''' <exception cref="IOException">
    ''' An I/O error occurs while enumerating or deleting a file or directory.
    ''' </exception>
    ''' <exception cref="OperationCanceledException">
    ''' The operation is canceled through <paramref name="CancellationToken"/>.
    ''' </exception>
    Public Shared Async Function DeleteDirectoryContentAsync(Directory As DirectoryInfo, Optional ExceptDirectories As IEnumerable(Of DirectoryInfo) = Nothing, Optional ExceptFiles As IEnumerable(Of FileInfo) = Nothing, Optional Progress As IProgress(Of Integer) = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task
        ArgumentNullException.ThrowIfNull(Directory)
        Directory.Refresh()
        If Not Directory.Exists Then Throw New DirectoryNotFoundException($"Directory '{Directory.FullName}' was not found.")
        Dim RootPath As String = NormalizePath(Directory.FullName)
        Dim ExcludedDirectoryPaths As HashSet(Of String) = GetExcludedDirectoryPaths(RootPath, ExceptDirectories)
        Dim ExcludedFilePaths As HashSet(Of String) = GetExcludedFilePaths(RootPath, ExceptFiles)
        If ExcludedDirectoryPaths.Contains(RootPath) Then
            ReportProgress(Progress, 0, 0, 0, 0)
            Return
        End If
        Await Task.Run(Sub()
                           Dim Options As EnumerationOptions = CreateEnumerationOptions()
                           Dim FilesToDelete As List(Of String) = System.IO.Directory.EnumerateFiles(RootPath, "*", Options).Select(AddressOf NormalizePath).Where(Function(FilePath) Not ShouldPreserveFile(FilePath, ExcludedDirectoryPaths, ExcludedFilePaths)).ToList()
                           Dim DirectoriesToDelete As List(Of String) = System.IO.Directory.EnumerateDirectories(RootPath, "*", Options).Select(AddressOf NormalizePath).Where(Function(DirectoryPath) Not ShouldPreserveDirectory(DirectoryPath, ExcludedDirectoryPaths, ExcludedFilePaths)).OrderByDescending(Function(DirectoryPath) DirectoryPath.Length).ToList()
                           Dim TotalItems As Long = FilesToDelete.Count + DirectoriesToDelete.Count
                           Dim ProcessedItems As Long = 0
                           Dim Reporter As New ProgressThrottle()
                           For Each CurrentFilePath As String In FilesToDelete
                               CancellationToken.ThrowIfCancellationRequested()
                               File.Delete(CurrentFilePath)
                               ProcessedItems += 1
                               If Reporter.ShouldReport(ProcessedItems = TotalItems) Then ReportProgress(Progress, 0, 0, ProcessedItems, TotalItems)
                           Next
                           For Each CurrentDirectoryPath As String In DirectoriesToDelete
                               CancellationToken.ThrowIfCancellationRequested()
                               If System.IO.Directory.Exists(CurrentDirectoryPath) Then System.IO.Directory.Delete(CurrentDirectoryPath, False)
                               ProcessedItems += 1
                               If Reporter.ShouldReport(ProcessedItems = TotalItems) Then ReportProgress(Progress, 0, 0, ProcessedItems, TotalItems)
                           Next
                           If TotalItems = 0 Then ReportProgress(Progress, 0, 0, 0, 0)
                       End Sub, CancellationToken)
    End Function
    ''' <summary>
    ''' Creates destination directories and executes one or more previously generated directory copy plans.
    ''' </summary>
    ''' <param name="Plans">
    ''' The directory copy plans to execute.
    ''' </param>
    ''' <param name="TotalSize">
    ''' The total number of bytes represented by the complete operation.
    ''' </param>
    ''' <param name="InitialHandledSize">
    ''' The number of bytes considered handled before the supplied plans begin.
    ''' </param>
    ''' <param name="Progress">
    ''' The optional progress reporter that receives completion percentages from 0 through 100.
    ''' </param>
    ''' <param name="CancellationToken">
    ''' A token that may be used to cancel the operation.
    ''' </param>
    ''' <returns>
    ''' A task whose result contains the total number of bytes copied by the supplied plans.
    ''' </returns>
    ''' <exception cref="OperationCanceledException">
    ''' The operation is canceled through <paramref name="CancellationToken"/>.
    ''' </exception>
    ''' <exception cref="UnauthorizedAccessException">
    ''' The caller does not have permission to create or write a destination directory or file.
    ''' </exception>
    ''' <exception cref="IOException">
    ''' An I/O error occurs while creating a directory or copying a file.
    ''' </exception>
    Private Shared Async Function ExecuteDirectoryCopyPlansAsync(Plans As IReadOnlyList(Of DirectoryCopyPlan), TotalSize As Long, InitialHandledSize As Long, Progress As IProgress(Of Integer), CancellationToken As CancellationToken) As Task(Of Long)
        Await Task.Run(Sub()
                           For Each CurrentPlan As DirectoryCopyPlan In Plans
                               CancellationToken.ThrowIfCancellationRequested()
                               System.IO.Directory.CreateDirectory(CurrentPlan.DestinationRoot)
                               For Each CurrentDirectory As String In CurrentPlan.Directories
                                   CancellationToken.ThrowIfCancellationRequested()
                                   System.IO.Directory.CreateDirectory(CurrentDirectory)
                               Next
                           Next
                       End Sub, CancellationToken)
        Dim TotalItems As Long = Plans.Sum(Function(CurrentPlan) CLng(CurrentPlan.Files.Count))
        Dim ProcessedItems As Long = 0
        Dim CopiedSize As Long = 0
        Dim Reporter As New ProgressThrottle()
        For Each CurrentPlan As DirectoryCopyPlan In Plans
            For Each CurrentFile As FileCopyEntry In CurrentPlan.Files
                CancellationToken.ThrowIfCancellationRequested()
                Dim BaseHandledSize As Long = InitialHandledSize + CopiedSize
                Dim CurrentFileCopiedSize As Long = Await CopyFileCoreAsync(CurrentFile.SourcePath, CurrentFile.DestinationPath, Sub(CurrentCopiedSize As Long)
                                                                                                                                     If Reporter.ShouldReport() Then ReportProgress(Progress, TotalSize, BaseHandledSize + CurrentCopiedSize, ProcessedItems, TotalItems)
                                                                                                                                 End Sub, CancellationToken)
                CopiedSize += CurrentFileCopiedSize
                ProcessedItems += 1
                ReportProgress(Progress, TotalSize, InitialHandledSize + CopiedSize, ProcessedItems, TotalItems)
            Next
        Next
        If TotalItems = 0 Then ReportProgress(Progress, TotalSize, InitialHandledSize + CopiedSize, 0, 0)
        Return CopiedSize
    End Function
    ''' <summary>
    ''' Asynchronously copies the contents of one file to another file.
    ''' </summary>
    ''' <param name="SourcePath">
    ''' The normalized absolute path of the source file.
    ''' </param>
    ''' <param name="DestinationPath">
    ''' The normalized absolute path of the destination file.
    ''' </param>
    ''' <param name="ProgressCallback">
    ''' A callback that receives the accumulated number of bytes copied from the current file.
    ''' </param>
    ''' <param name="CancellationToken">
    ''' A token that may be used to cancel the operation.
    ''' </param>
    ''' <returns>
    ''' A task whose result contains the total number of bytes copied.
    ''' </returns>
    ''' <remarks>
    ''' The destination parent directory is created automatically and an existing destination file is overwritten.
    ''' </remarks>
    ''' <exception cref="DirectoryNotFoundException">
    ''' The destination parent directory cannot be determined.
    ''' </exception>
    ''' <exception cref="FileNotFoundException">
    ''' The source file does not exist.
    ''' </exception>
    ''' <exception cref="UnauthorizedAccessException">
    ''' The caller does not have permission to read the source or write the destination.
    ''' </exception>
    ''' <exception cref="IOException">
    ''' An I/O error occurs while opening, reading, writing, or flushing a stream.
    ''' </exception>
    ''' <exception cref="OperationCanceledException">
    ''' The operation is canceled through <paramref name="CancellationToken"/>.
    ''' </exception>
    Private Shared Async Function CopyFileCoreAsync(SourcePath As String, DestinationPath As String, ProgressCallback As Action(Of Long), CancellationToken As CancellationToken) As Task(Of Long)
        Dim DestinationDirectoryPath As String = Path.GetDirectoryName(DestinationPath)
        If String.IsNullOrWhiteSpace(DestinationDirectoryPath) Then Throw New DirectoryNotFoundException($"The destination directory for '{DestinationPath}' could not be determined.")
        System.IO.Directory.CreateDirectory(DestinationDirectoryPath)
        Dim SourceOptions As New FileStreamOptions With {.Mode = FileMode.Open, .Access = FileAccess.Read, .Share = FileShare.Read, .BufferSize = StreamBufferSize, .Options = FileOptions.Asynchronous Or FileOptions.SequentialScan}
        Dim DestinationOptions As New FileStreamOptions With {.Mode = FileMode.Create, .Access = FileAccess.Write, .Share = FileShare.None, .BufferSize = StreamBufferSize, .Options = FileOptions.Asynchronous Or FileOptions.SequentialScan}
        Dim Buffer(StreamBufferSize - 1) As Byte
        Dim CopiedSize As Long = 0
        Using SourceStream As New FileStream(SourcePath, SourceOptions)
            Using DestinationStream As New FileStream(DestinationPath, DestinationOptions)
                Do
                    CancellationToken.ThrowIfCancellationRequested()
                    Dim BytesRead As Integer = Await SourceStream.ReadAsync(Buffer.AsMemory(0, Buffer.Length), CancellationToken).ConfigureAwait(False)
                    If BytesRead = 0 Then Exit Do
                    Await DestinationStream.WriteAsync(Buffer.AsMemory(0, BytesRead), CancellationToken).ConfigureAwait(False)
                    CopiedSize += BytesRead
                    If ProgressCallback IsNot Nothing Then ProgressCallback(CopiedSize)
                Loop
                Await DestinationStream.FlushAsync(CancellationToken).ConfigureAwait(False)
            End Using
        End Using
        Return CopiedSize
    End Function
    ''' <summary>
    ''' Builds a complete copy plan for a source directory and its destination.
    ''' </summary>
    ''' <param name="CopyInfo">
    ''' The source and destination directory mapping used to build the plan.
    ''' </param>
    ''' <param name="CancellationToken">
    ''' A token that may be used to cancel directory enumeration.
    ''' </param>
    ''' <returns>
    ''' A plan containing the destination directories, files, paths, and file sizes required by the copy operation.
    ''' </returns>
    ''' <remarks>
    ''' Reparse points are skipped. The method also prevents the destination from being equal to or located inside the source directory.
    ''' </remarks>
    ''' <exception cref="ArgumentNullException">
    ''' The source or destination in <paramref name="CopyInfo"/> is <see langword="Nothing"/>.
    ''' </exception>
    ''' <exception cref="DirectoryNotFoundException">
    ''' The source directory does not exist.
    ''' </exception>
    ''' <exception cref="IOException">
    ''' The source and destination are equal or the destination is inside the source directory.
    ''' </exception>
    ''' <exception cref="OperationCanceledException">
    ''' The operation is canceled through <paramref name="CancellationToken"/>.
    ''' </exception>
    Private Shared Function BuildDirectoryCopyPlan(CopyInfo As CopyDirectoryInfo, CancellationToken As CancellationToken) As DirectoryCopyPlan
        ArgumentNullException.ThrowIfNull(CopyInfo.Source)
        ArgumentNullException.ThrowIfNull(CopyInfo.Destination)
        CopyInfo.Source.Refresh()
        If Not CopyInfo.Source.Exists Then Throw New DirectoryNotFoundException($"Source directory '{CopyInfo.Source.FullName}' was not found.")
        Dim SourceRoot As String = NormalizePath(CopyInfo.Source.FullName)
        Dim DestinationRoot As String = NormalizePath(CopyInfo.Destination.FullName)
        If AreSamePath(SourceRoot, DestinationRoot) Then Throw New IOException("The source and destination directories cannot be the same.")
        If IsPathInside(DestinationRoot, SourceRoot) Then Throw New IOException("The destination directory cannot be located inside the source directory.")
        Dim Options As EnumerationOptions = CreateEnumerationOptions()
        Dim DestinationDirectories As New List(Of String)
        Dim Files As New List(Of FileCopyEntry)
        For Each SourceDirectoryPath As String In System.IO.Directory.EnumerateDirectories(SourceRoot, "*", Options)
            CancellationToken.ThrowIfCancellationRequested()
            Dim RelativePath As String = Path.GetRelativePath(SourceRoot, SourceDirectoryPath)
            DestinationDirectories.Add(NormalizePath(Path.Combine(DestinationRoot, RelativePath)))
        Next
        For Each SourceFilePath As String In System.IO.Directory.EnumerateFiles(SourceRoot, "*", Options)
            CancellationToken.ThrowIfCancellationRequested()
            Dim SourceFile As New FileInfo(SourceFilePath)
            Dim RelativePath As String = Path.GetRelativePath(SourceRoot, SourceFilePath)
            Dim DestinationFilePath As String = NormalizePath(Path.Combine(DestinationRoot, RelativePath))
            Files.Add(New FileCopyEntry(NormalizePath(SourceFilePath), DestinationFilePath, SourceFile.Length))
        Next
        Return New DirectoryCopyPlan(DestinationRoot, DestinationDirectories, Files)
    End Function
    ''' <summary>
    ''' Builds a deletion plan for the supplied directory deletion requests.
    ''' </summary>
    ''' <param name="Requests">
    ''' The validated directory deletion requests to include in the plan.
    ''' </param>
    ''' <param name="CancellationToken">
    ''' A token that may be used to cancel directory enumeration.
    ''' </param>
    ''' <returns>
    ''' A plan containing the roots, files, directories, and total file size represented by the operation.
    ''' </returns>
    ''' <remarks>
    ''' Missing root directories are ignored. Existing root directories may not overlap.
    ''' </remarks>
    ''' <exception cref="ArgumentNullException">
    ''' A request does not specify a directory.
    ''' </exception>
    ''' <exception cref="ArgumentException">
    ''' Two or more existing root directories overlap.
    ''' </exception>
    ''' <exception cref="OperationCanceledException">
    ''' The operation is canceled through <paramref name="CancellationToken"/>.
    ''' </exception>
    Private Shared Function BuildDirectoryDeletePlan(Requests As IReadOnlyList(Of DeleteDirectoryInfo), CancellationToken As CancellationToken) As DirectoryDeletePlan
        Dim Roots As New List(Of DirectoryDeleteRoot)
        For Each CurrentRequest As DeleteDirectoryInfo In Requests
            CancellationToken.ThrowIfCancellationRequested()
            ArgumentNullException.ThrowIfNull(CurrentRequest.Directory)
            CurrentRequest.Directory.Refresh()
            If CurrentRequest.Directory.Exists Then Roots.Add(New DirectoryDeleteRoot(NormalizePath(CurrentRequest.Directory.FullName), CurrentRequest.DeleteRoot))
        Next
        ValidateNonOverlappingRoots(Roots)
        Dim Files As New Dictionary(Of String, FileDeleteEntry)(FileSystemPathComparer)
        Dim Directories As New HashSet(Of String)(FileSystemPathComparer)
        Dim Options As EnumerationOptions = CreateEnumerationOptions()
        For Each CurrentRoot As DirectoryDeleteRoot In Roots
            For Each CurrentFilePath As String In System.IO.Directory.EnumerateFiles(CurrentRoot.Path, "*", Options)
                CancellationToken.ThrowIfCancellationRequested()
                Dim NormalizedFilePath As String = NormalizePath(CurrentFilePath)
                If Not Files.ContainsKey(NormalizedFilePath) Then
                    Dim CurrentFile As New FileInfo(NormalizedFilePath)
                    Files.Add(NormalizedFilePath, New FileDeleteEntry(NormalizedFilePath, CurrentFile.Length))
                End If
            Next
            For Each CurrentDirectoryPath As String In System.IO.Directory.EnumerateDirectories(CurrentRoot.Path, "*", Options)
                CancellationToken.ThrowIfCancellationRequested()
                Directories.Add(NormalizePath(CurrentDirectoryPath))
            Next
        Next
        Return New DirectoryDeletePlan(Roots, Files.Values.ToList(), Directories.ToList())
    End Function
    ''' <summary>
    ''' Creates a collection of files with duplicate paths removed.
    ''' </summary>
    ''' <param name="Files">
    ''' The files to normalize and filter.
    ''' </param>
    ''' <returns>
    ''' A list containing one <see cref="FileInfo"/> instance for each distinct normalized path.
    ''' </returns>
    ''' <exception cref="ArgumentException">
    ''' The collection contains a <see langword="Nothing"/> item or an invalid path.
    ''' </exception>
    Private Shared Function GetDistinctFiles(Files As IEnumerable(Of FileInfo)) As List(Of FileInfo)
        Dim Results As New Dictionary(Of String, FileInfo)(FileSystemPathComparer)
        For Each CurrentFile As FileInfo In Files
            If CurrentFile Is Nothing Then Throw New ArgumentException("The file collection cannot contain null items.", NameOf(Files))
            Dim FilePath As String = NormalizePath(CurrentFile.FullName)
            If Not Results.ContainsKey(FilePath) Then Results.Add(FilePath, New FileInfo(FilePath))
        Next
        Return Results.Values.ToList()
    End Function
    ''' <summary>
    ''' Creates a normalized set of directory paths that must be excluded from a content deletion operation.
    ''' </summary>
    ''' <param name="RootPath">
    ''' The normalized absolute path of the root directory.
    ''' </param>
    ''' <param name="Directories">
    ''' The optional collection of directories to preserve.
    ''' </param>
    ''' <returns>
    ''' A path set using the comparer appropriate for the current operating system.
    ''' </returns>
    ''' <exception cref="ArgumentException">
    ''' The collection contains a <see langword="Nothing"/> item, an invalid path, or a directory outside <paramref name="RootPath"/>.
    ''' </exception>
    Private Shared Function GetExcludedDirectoryPaths(RootPath As String, Directories As IEnumerable(Of DirectoryInfo)) As HashSet(Of String)
        Dim Results As New HashSet(Of String)(FileSystemPathComparer)
        If Directories Is Nothing Then Return Results
        For Each CurrentDirectory As DirectoryInfo In Directories
            If CurrentDirectory Is Nothing Then Throw New ArgumentException("The excluded directory collection cannot contain null items.", NameOf(Directories))
            Dim DirectoryPath As String = NormalizePath(CurrentDirectory.FullName)
            If Not IsSameOrChildPath(DirectoryPath, RootPath) Then Throw New ArgumentException($"Excluded directory '{DirectoryPath}' is not located inside '{RootPath}'.", NameOf(Directories))
            Results.Add(DirectoryPath)
        Next
        Return Results
    End Function
    ''' <summary>
    ''' Creates a normalized set of file paths that must be excluded from a content deletion operation.
    ''' </summary>
    ''' <param name="RootPath">
    ''' The normalized absolute path of the root directory.
    ''' </param>
    ''' <param name="Files">
    ''' The optional collection of files to preserve.
    ''' </param>
    ''' <returns>
    ''' A path set using the comparer appropriate for the current operating system.
    ''' </returns>
    ''' <exception cref="ArgumentException">
    ''' The collection contains a <see langword="Nothing"/> item, an invalid path, or a file outside <paramref name="RootPath"/>.
    ''' </exception>
    Private Shared Function GetExcludedFilePaths(RootPath As String, Files As IEnumerable(Of FileInfo)) As HashSet(Of String)
        Dim Results As New HashSet(Of String)(FileSystemPathComparer)
        If Files Is Nothing Then Return Results
        For Each CurrentFile As FileInfo In Files
            If CurrentFile Is Nothing Then Throw New ArgumentException("The excluded file collection cannot contain null items.", NameOf(Files))
            Dim FilePath As String = NormalizePath(CurrentFile.FullName)
            If Not IsPathInside(FilePath, RootPath) Then Throw New ArgumentException($"Excluded file '{FilePath}' is not located inside '{RootPath}'.", NameOf(Files))
            Results.Add(FilePath)
        Next
        Return Results
    End Function
    ''' <summary>
    ''' Determines whether a file must be preserved during a directory content deletion operation.
    ''' </summary>
    ''' <param name="FilePath">
    ''' The normalized absolute path of the file being evaluated.
    ''' </param>
    ''' <param name="ExcludedDirectories">
    ''' The normalized directory paths that must be preserved.
    ''' </param>
    ''' <param name="ExcludedFiles">
    ''' The normalized file paths that must be preserved.
    ''' </param>
    ''' <returns>
    ''' <see langword="True"/> when the file is explicitly excluded or is located inside an excluded directory; otherwise, <see langword="False"/>.
    ''' </returns>
    Private Shared Function ShouldPreserveFile(FilePath As String, ExcludedDirectories As HashSet(Of String), ExcludedFiles As HashSet(Of String)) As Boolean
        If ExcludedFiles.Contains(FilePath) Then Return True
        For Each ExcludedDirectory As String In ExcludedDirectories
            If IsSameOrChildPath(FilePath, ExcludedDirectory) Then Return True
        Next
        Return False
    End Function
    ''' <summary>
    ''' Determines whether a directory must be preserved during a directory content deletion operation.
    ''' </summary>
    ''' <param name="DirectoryPath">
    ''' The normalized absolute path of the directory being evaluated.
    ''' </param>
    ''' <param name="ExcludedDirectories">
    ''' The normalized directory paths that must be preserved.
    ''' </param>
    ''' <param name="ExcludedFiles">
    ''' The normalized file paths that must be preserved.
    ''' </param>
    ''' <returns>
    ''' <see langword="True"/> when the directory is excluded, contains an excluded item, or is located inside an excluded directory; otherwise, <see langword="False"/>.
    ''' </returns>
    Private Shared Function ShouldPreserveDirectory(DirectoryPath As String, ExcludedDirectories As HashSet(Of String), ExcludedFiles As HashSet(Of String)) As Boolean
        For Each ExcludedDirectory As String In ExcludedDirectories
            If IsSameOrChildPath(DirectoryPath, ExcludedDirectory) OrElse IsSameOrChildPath(ExcludedDirectory, DirectoryPath) Then Return True
        Next
        For Each ExcludedFile As String In ExcludedFiles
            If IsSameOrChildPath(ExcludedFile, DirectoryPath) Then Return True
        Next
        Return False
    End Function
    ''' <summary>
    ''' Validates that no deletion root is equal to, contains, or is contained by another deletion root.
    ''' </summary>
    ''' <param name="Roots">
    ''' The normalized deletion roots to validate.
    ''' </param>
    ''' <exception cref="ArgumentException">
    ''' Two root paths overlap.
    ''' </exception>
    Private Shared Sub ValidateNonOverlappingRoots(Roots As List(Of DirectoryDeleteRoot))
        For FirstIndex As Integer = 0 To Roots.Count - 2
            For SecondIndex As Integer = FirstIndex + 1 To Roots.Count - 1
                Dim FirstPath As String = Roots(FirstIndex).Path
                Dim SecondPath As String = Roots(SecondIndex).Path
                If IsSameOrChildPath(FirstPath, SecondPath) OrElse IsSameOrChildPath(SecondPath, FirstPath) Then Throw New ArgumentException($"Directory deletion requests cannot overlap: '{FirstPath}' and '{SecondPath}'.")
            Next
        Next
    End Sub
    ''' <summary>
    ''' Reports the current completion percentage when a progress reporter was supplied.
    ''' </summary>
    ''' <param name="Progress">
    ''' The optional progress reporter that receives values from 0 through 100.
    ''' </param>
    ''' <param name="TotalSize">
    ''' The total number of bytes represented by the operation.
    ''' </param>
    ''' <param name="HandledSize">
    ''' The number of bytes already processed.
    ''' </param>
    ''' <param name="ProcessedItems">
    ''' The number of completely processed files or items.
    ''' </param>
    ''' <param name="TotalItems">
    ''' The total number of files or items represented by the operation.
    ''' </param>
    Private Shared Sub ReportProgress(Progress As IProgress(Of Integer), TotalSize As Long, HandledSize As Long, ProcessedItems As Long, TotalItems As Long)
        If Progress Is Nothing Then Return
        Dim PercentCompleted As Integer
        If TotalItems > 0 AndAlso ProcessedItems >= TotalItems Then
            PercentCompleted = 100
        ElseIf TotalSize > 0 Then
            PercentCompleted = CInt(Math.Clamp(Math.Floor(Math.Max(0, HandledSize) / CDbl(TotalSize) * 100.0), 0.0, 100.0))
        ElseIf TotalItems > 0 Then
            PercentCompleted = CInt(Math.Clamp(Math.Floor(Math.Max(0, ProcessedItems) / CDbl(TotalItems) * 100.0), 0.0, 100.0))
        Else
            PercentCompleted = 100
        End If
        Progress.Report(PercentCompleted)
    End Sub
    ''' <summary>
    ''' Creates the standard options used when recursively enumerating files and directories.
    ''' </summary>
    ''' <returns>
    ''' Enumeration options configured for recursive traversal, strict access-error handling, and reparse-point exclusion.
    ''' </returns>
    ''' <remarks>
    ''' Inaccessible entries cause an exception because <see cref="EnumerationOptions.IgnoreInaccessible"/> is disabled.
    ''' </remarks>
    Private Shared Function CreateEnumerationOptions() As EnumerationOptions
        Return New EnumerationOptions With {.RecurseSubdirectories = True, .IgnoreInaccessible = False, .ReturnSpecialDirectories = False, .AttributesToSkip = FileAttributes.ReparsePoint}
    End Function
    ''' <summary>
    ''' Converts a path to a normalized absolute path and removes its trailing directory separator.
    ''' </summary>
    ''' <param name="PathValue">
    ''' The path to normalize.
    ''' </param>
    ''' <returns>
    ''' The normalized absolute path.
    ''' </returns>
    ''' <exception cref="ArgumentException">
    ''' <paramref name="PathValue"/> is empty, contains only white-space characters, or has an invalid format.
    ''' </exception>
    ''' <exception cref="NotSupportedException">
    ''' The path format is not supported.
    ''' </exception>
    ''' <exception cref="PathTooLongException">
    ''' The path exceeds a limit supported by the current platform or file system.
    ''' </exception>
    Private Shared Function NormalizePath(PathValue As String) As String
        If String.IsNullOrWhiteSpace(PathValue) Then Throw New ArgumentException("The path cannot be empty.", NameOf(PathValue))
        Return Path.TrimEndingDirectorySeparator(Path.GetFullPath(PathValue))
    End Function
    ''' <summary>
    ''' Determines whether two normalized file-system paths represent the same location.
    ''' </summary>
    ''' <param name="FirstPath">
    ''' The first normalized path to compare.
    ''' </param>
    ''' <param name="SecondPath">
    ''' The second normalized path to compare.
    ''' </param>
    ''' <returns>
    ''' <see langword="True"/> when the paths are equal according to the current operating system path rules; otherwise, <see langword="False"/>.
    ''' </returns>
    Private Shared Function AreSamePath(FirstPath As String, SecondPath As String) As Boolean
        Return FileSystemPathComparer.Equals(FirstPath, SecondPath)
    End Function
    ''' <summary>
    ''' Determines whether one normalized path is located strictly inside another normalized path.
    ''' </summary>
    ''' <param name="CandidatePath">
    ''' The path that may be a descendant.
    ''' </param>
    ''' <param name="ParentPath">
    ''' The path that may be the parent.
    ''' </param>
    ''' <returns>
    ''' <see langword="True"/> when <paramref name="CandidatePath"/> is located below <paramref name="ParentPath"/>; otherwise, <see langword="False"/>.
    ''' </returns>
    ''' <remarks>
    ''' Equal paths are not considered a parent-child relationship.
    ''' </remarks>
    Private Shared Function IsPathInside(CandidatePath As String, ParentPath As String) As Boolean
        If AreSamePath(CandidatePath, ParentPath) Then Return False
        Dim ParentPathWithSeparator As String = ParentPath
        If Not ParentPathWithSeparator.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) AndAlso Not ParentPathWithSeparator.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal) Then ParentPathWithSeparator &= Path.DirectorySeparatorChar
        Return CandidatePath.StartsWith(ParentPathWithSeparator, FileSystemPathComparison)
    End Function
    ''' <summary>
    ''' Determines whether one normalized path is equal to or located inside another normalized path.
    ''' </summary>
    ''' <param name="CandidatePath">
    ''' The path that may be equal to or a descendant of the parent.
    ''' </param>
    ''' <param name="ParentPath">
    ''' The path that may be equal to or an ancestor of the candidate.
    ''' </param>
    ''' <returns>
    ''' <see langword="True"/> when the paths are equal or the candidate is located below the parent; otherwise, <see langword="False"/>.
    ''' </returns>
    Private Shared Function IsSameOrChildPath(CandidatePath As String, ParentPath As String) As Boolean
        Return AreSamePath(CandidatePath, ParentPath) OrElse IsPathInside(CandidatePath, ParentPath)
    End Function
End Class