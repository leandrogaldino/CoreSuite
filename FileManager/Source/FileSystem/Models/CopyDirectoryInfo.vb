Imports System.IO
''' <summary>
''' Describes a source directory and its corresponding copy destination.
''' </summary>
Public NotInheritable Class CopyDirectoryInfo
    ''' <summary>
    ''' Initializes a new empty instance of the <see cref="CopyDirectoryInfo"/> class.
    ''' </summary>
    ''' <remarks>
    ''' The <see cref="Source"/> and <see cref="Destination"/> properties must be assigned before the instance is used in a copy operation.
    ''' </remarks>
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="CopyDirectoryInfo"/> class with the specified source and destination.
    ''' </summary>
    ''' <param name="Source">
    ''' The source directory to copy.
    ''' </param>
    ''' <param name="Destination">
    ''' The destination directory that will receive the copied content.
    ''' </param>
    Public Sub New(Source As DirectoryInfo, Destination As DirectoryInfo)
        Me.Source = Source
        Me.Destination = Destination
    End Sub
    ''' <summary>
    ''' Gets or sets the source directory to copy.
    ''' </summary>
    ''' <value>
    ''' A <see cref="DirectoryInfo"/> representing the source directory.
    ''' </value>
    Public Property Source As DirectoryInfo
    ''' <summary>
    ''' Gets or sets the destination directory that will receive the copied content.
    ''' </summary>
    ''' <value>
    ''' A <see cref="DirectoryInfo"/> representing the destination directory.
    ''' </value>
    Public Property Destination As DirectoryInfo
End Class