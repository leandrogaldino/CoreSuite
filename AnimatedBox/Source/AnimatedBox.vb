Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.Windows.Forms

''' <summary>
''' Defines how animation frames are scaled inside the <see cref="AnimatedBox"/> control.
''' </summary>
Public Enum AnimationScaleMode

    ''' <summary>
    ''' Draws the image using its original size, centered within the control.
    ''' </summary>
    Normal

    ''' <summary>
    ''' Stretches the image to completely fill the control,
    ''' ignoring the original aspect ratio.
    ''' </summary>
    Fill

    ''' <summary>
    ''' Scales the image proportionally to fit inside the control
    ''' while keeping its aspect ratio and centering it.
    ''' </summary>
    Centrer

End Enum

''' <summary>
''' A WinForms control that renders frame-based animations using images or GIF files.
''' </summary>
''' <remarks>
''' This control supports animated image sequences and GIF extraction,
''' offering high-quality rendering and customizable scaling behavior.
''' <para>
''' The animation timing is driven by a <see cref="Timer"/> combined with
''' a <see cref="Stopwatch"/> for accurate frame delay handling.
''' </para>
''' </remarks>
Public Class AnimatedBox
    Inherits Panel

    ''' <summary>
    ''' Internal representation of a single animation frame.
    ''' </summary>
    Private Class AnimationFrameInfo

        ''' <summary>
        ''' Gets or sets the image associated with this frame.
        ''' </summary>
        Public Property Image As Image

        ''' <summary>
        ''' Gets or sets the delay (in seconds) before advancing to the next frame.
        ''' </summary>
        Public Property Delay As Double

        ''' <summary>
        ''' Initializes a new instance of the <see cref="AnimationFrameInfo"/> class.
        ''' </summary>
        ''' <param name="Image">The image to be displayed.</param>
        ''' <param name="Delay">
        ''' The frame delay in seconds. Default value is 0.03 (≈33 FPS).
        ''' </param>
        Public Sub New(Image As Image, Optional Delay As Double = 0.03)
            Me.Image = Image
            Me.Delay = Delay
        End Sub

    End Class

    Private _FrameIndex As Integer
    Private _LastFrameTime As Double
    Private ReadOnly _Frames As List(Of AnimationFrameInfo)
    Private ReadOnly _Watch As Stopwatch
    Private ReadOnly _frameTimer As Timer

    ''' <summary>
    ''' Gets or sets the scaling behavior used when rendering animation frames.
    ''' </summary>
    Public Property ScaleMode As AnimationScaleMode

    ''' <summary>
    ''' Loads a sequence of images to be used as animation frames.
    ''' </summary>
    ''' <param name="Images">
    ''' A list of <see cref="Image"/> objects representing the animation frames.
    ''' </param>
    Public Sub LoadImages(Images As List(Of Image))
        _Frames.Clear()
        Images.ForEach(Sub(x)
                           _Frames.Add(New AnimationFrameInfo(x))
                       End Sub)
    End Sub

    ''' <summary>
    ''' Loads and extracts frames from a GIF image.
    ''' </summary>
    ''' <param name="Gif">The GIF image to be decoded into frames.</param>
    Public Sub LoadGif(Gif As Image)
        Dim Images As List(Of Image) = ExtractGifFrames(Gif)
        _Frames.Clear()
        Images.ForEach(Sub(x)
                           _Frames.Add(New AnimationFrameInfo(x))
                       End Sub)
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the <see cref="AnimatedBox"/> control.
    ''' </summary>
    Public Sub New()
        _Frames = New List(Of AnimationFrameInfo)()
        _Watch = New Stopwatch()
        ScaleMode = AnimationScaleMode.Centrer
        SetStyle(ControlStyles.UserPaint Or ControlStyles.SupportsTransparentBackColor Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer, True)
        DoubleBuffered = True
        _frameTimer = New Timer()
        AddHandler _frameTimer.Tick, AddressOf FrameTimer_Tick
    End Sub

    ''' <summary>
    ''' Handles frame advancement based on elapsed time and frame delay.
    ''' </summary>
    Private Sub FrameTimer_Tick(sender As Object, e As EventArgs)
        If _Frames.Count = 0 Then Return
        Dim TotalSeconds = _Watch.Elapsed.TotalSeconds
        Dim Frame = _Frames(_FrameIndex)
        If TotalSeconds - _LastFrameTime >= Frame.Delay Then
            _LastFrameTime = TotalSeconds
            _FrameIndex = (_FrameIndex + 1) Mod _Frames.Count
            Invalidate()
        End If
    End Sub

    ''' <summary>
    ''' Starts the animation playback.
    ''' </summary>
    Public Sub StartAnimation()
        _FrameIndex = 0
        _LastFrameTime = 0
        _Watch.Restart()
        _frameTimer.Interval = 15
        _frameTimer.Start()
    End Sub

    ''' <summary>
    ''' Stops the animation playback.
    ''' </summary>
    Public Sub StopAnimation()
        _frameTimer.Stop()
        _Watch.Stop()
    End Sub

    ''' <summary>
    ''' Renders the current animation frame.
    ''' </summary>
    ''' <param name="e">Paint event data.</param>
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)

        If _Frames.Count = 0 Then Return

        Dim Frame = _Frames(_FrameIndex)

        e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic
        e.Graphics.SmoothingMode = SmoothingMode.HighQuality
        e.Graphics.PixelOffsetMode = PixelOffsetMode.Half

        Select Case ScaleMode
            Case AnimationScaleMode.Normal
                Dim x = (Width - Frame.Image.Width) \ 2
                Dim y = (Height - Frame.Image.Height) \ 2
                e.Graphics.DrawImage(Frame.Image, x, y, Frame.Image.Width, Frame.Image.Height)

            Case AnimationScaleMode.Fill
                e.Graphics.DrawImage(Frame.Image, 0, 0, Width, Height)

            Case AnimationScaleMode.Centrer
                Dim scaleX = Width / Frame.Image.Width
                Dim scaleY = Height / Frame.Image.Height
                Dim scale = Math.Min(scaleX, scaleY)

                Dim w = Frame.Image.Width * scale
                Dim h = Frame.Image.Height * scale
                Dim x = (Width - w) / 2
                Dim y = (Height - h) / 2

                e.Graphics.DrawImage(Frame.Image, CSng(x), CSng(y), CSng(w), CSng(h))
        End Select
    End Sub

    ''' <summary>
    ''' Extracts individual frames from a GIF image.
    ''' </summary>
    ''' <param name="Gif">The source GIF image.</param>
    ''' <returns>A list containing all extracted frames.</returns>
    Private Function ExtractGifFrames(Gif As Image) As List(Of Image)
        Dim List As New List(Of Image)()
        Dim Dimension = New FrameDimension(Gif.FrameDimensionsList(0))
        Dim FrameCount = Gif.GetFrameCount(Dimension)
        For i As Integer = 0 To FrameCount - 1
            Gif.SelectActiveFrame(Dimension, i)
            Dim Bmp = New Bitmap(Gif.Width, Gif.Height)
            Using g = Graphics.FromImage(Bmp)
                g.DrawImage(Gif, Point.Empty)
            End Using
            List.Add(Bmp)
        Next

        Return List
    End Function

    ''' <summary>
    ''' Releases all resources used by the control.
    ''' </summary>
    ''' <param name="disposing">
    ''' True to release managed resources; otherwise, false.
    ''' </param>
    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            _frameTimer.Stop()
            _Watch.Stop()

            If _Frames IsNot Nothing Then
                For Each frame In _Frames
                    frame.Image?.Dispose()
                Next
                _Frames.Clear()
            End If
        End If

        MyBase.Dispose(disposing)
    End Sub

End Class