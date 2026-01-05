Imports System.Drawing

''' <summary>
''' Provides helper methods for common image operations such as cloning,
''' recoloring, color analysis, and resizing.
''' </summary>
Public Class ImageHelper

    ''' <summary>
    ''' Creates and returns a copy of an image loaded from the specified file path.
    ''' </summary>
    ''' <param name="ImagePath">
    ''' The file system path of the image to be loaded.
    ''' </param>
    ''' <returns>
    ''' A copy of the loaded <see cref="Image"/>;
    ''' or a fallback error image if the operation fails.
    ''' </returns>
    Public Shared Function GetCopyImage(ImagePath As String) As Image
        Dim Bmp As Bitmap
        Try
            Using Img As Image = Image.FromFile(ImagePath)
                Bmp = New Bitmap(Img)
                Return Bmp
            End Using
        Catch ex As Exception
            Return My.Resources.ImageError
        End Try
    End Function

    ''' <summary>
    ''' Returns a recolored copy of an image by replacing a specific color value
    ''' with another.
    ''' </summary>
    ''' <param name="Image">
    ''' The source image to be recolored.
    ''' </param>
    ''' <param name="FromColor">
    ''' The ARGB color value to be replaced.
    ''' </param>
    ''' <param name="ToColor">
    ''' The ARGB color value that will replace the original color.
    ''' </param>
    ''' <returns>
    ''' A new <see cref="Bitmap"/> instance with the specified color replaced.
    ''' </returns>
    Public Shared Function GetRecoloredImage(Image As Image, FromColor As Integer, ToColor As Integer) As Bitmap
        Dim Bmp As Bitmap = Image.Clone
        For x As Integer = 0 To Bmp.Width - 1
            For y As Integer = 0 To Bmp.Height - 1
                If Bmp.GetPixel(x, y) = Color.FromArgb(FromColor) Then
                    Bmp.SetPixel(x, y, Color.FromArgb(ToColor))
                End If
            Next
        Next
        Return Bmp
    End Function

    ''' <summary>
    ''' Returns a recolored copy of an image using a base color while preserving
    ''' the original alpha channel values.
    ''' </summary>
    ''' <param name="Image">
    ''' The source image to be recolored.
    ''' </param>
    ''' <param name="BaseColor">
    ''' The base color applied to all pixels while keeping their original transparency.
    ''' </param>
    ''' <returns>
    ''' A new <see cref="Image"/> instance recolored using the specified base color.
    ''' </returns>
    Public Shared Function GetRecoloredImage(Image As Image, BaseColor As Color) As Image
        Dim Img As Image = Image.Clone
        Dim Colors As List(Of Color) = GetImageColors(Img)
        For i = 0 To Colors.Count - 1
            Img = GetRecoloredImage(Img, Colors(i).ToArgb, Color.FromArgb(Colors(i).A, BaseColor).ToArgb)
        Next
        Return Img
    End Function

    ''' <summary>
    ''' Retrieves all distinct colors used in an image.
    ''' </summary>
    ''' <param name="Image">
    ''' The image to be analyzed.
    ''' </param>
    ''' <returns>
    ''' A list of distinct <see cref="Color"/> values found in the image.
    ''' </returns>
    Public Shared Function GetImageColors(Image As Image) As List(Of Color)
        Dim Lst As New List(Of Color)
        Dim Bmp As Bitmap = Image
        For x As Integer = 0 To Bmp.Width - 1
            For y As Integer = 0 To Bmp.Height - 1
                Lst.Add(Bmp.GetPixel(x, y))
            Next
        Next
        Return Lst.Distinct().ToList
    End Function

    ''' <summary>
    ''' Resizes an image while maintaining its aspect ratio, based on the specified
    ''' maximum width and/or height constraints.
    ''' </summary>
    ''' <param name="Image">
    ''' The original image to be resized.
    ''' </param>
    ''' <param name="MaxWidth">
    ''' The maximum allowed width. If zero or less, width is not constrained.
    ''' </param>
    ''' <param name="MaxHeight">
    ''' The maximum allowed height. If zero or less, height is not constrained.
    ''' </param>
    ''' <returns>
    ''' A resized <see cref="Image"/> instance, or the original image if resizing
    ''' is not required.
    ''' </returns>
    Public Shared Function GetResizedImage(Image As Image, Optional MaxWidth As Integer = 0, Optional MaxHeight As Integer = 0) As Image
        If MaxWidth <= 0 AndAlso MaxHeight <= 0 Then
            Return Image
        End If
        Dim NeedResize As Boolean = (MaxWidth > 0 AndAlso Image.Width > MaxWidth) OrElse (MaxHeight > 0 AndAlso Image.Height > MaxHeight)
        If NeedResize Then
            Dim WidthScale As Single = If(MaxWidth > 0, MaxWidth / Image.Width, Single.MaxValue)
            Dim HeightScale As Single = If(MaxHeight > 0, MaxHeight / Image.Height, Single.MaxValue)
            Dim Scale As Single = Math.Min(WidthScale, HeightScale)
            Dim NewWidth As Integer = CInt(Image.Width * Scale)
            Dim NewHeight As Integer = CInt(Image.Height * Scale)
            Dim NewImage As New Bitmap(Image, NewWidth, NewHeight)
            Image.Dispose()
            Return NewImage
        Else
            Return Image
        End If
    End Function

End Class