# CoreSuite.AnimatedBox

**A lightweight Windows Forms control for displaying frame-based animations from image sequences or GIF files.**

> \[!NOTE]
> `CoreSuite.AnimatedBox` is one of the controls included in the \*\*CoreSuite\*\* solution. It targets .NET 8 for Windows and is implemented entirely with standard Windows Forms and `System.Drawing` APIs.

## Overview

`AnimatedBox` is a reusable Windows Forms control that renders a sequence of images as an animation. Frames can be loaded from an existing `List(Of Image)` or extracted from a GIF image.

The control uses a Windows Forms `Timer` to update the interface and a `Stopwatch` to measure elapsed time more accurately. Rendering is double-buffered and uses high-quality interpolation, smoothing, and pixel positioning settings.

Because `AnimatedBox` inherits from `Panel`, it also supports familiar layout features such as anchoring, docking, borders, background colors, and transparent backgrounds.

## Key features

* Display animations created from individual `Image` objects.
* Extract and display frames from GIF images.
* Start and stop playback programmatically.
* Restart playback from the first frame.
* Render images at their original size.
* Stretch images to fill the entire control.
* Scale images proportionally while preserving their aspect ratio.
* Center rendered frames automatically.
* Use high-quality bicubic interpolation.
* Reduce flickering through optimized double buffering.
* Support transparent background colors.
* Release the currently loaded frame images when the control is disposed.
* Use only built-in .NET and Windows Forms APIs.

## Requirements

* Windows
* Windows Forms
* .NET 8 for Windows (`net8.0-windows`)
* A reference to the `CoreSuite.AnimatedBox` package, project, or assembly

The project currently has no external NuGet dependencies.

## Installation

### NuGet Package Manager

```powershell
Install-Package CoreSuite.AnimatedBox
```

### .NET CLI

```bash
dotnet add package CoreSuite.AnimatedBox
```

### Project reference

When using the CoreSuite source code directly, add a reference to the project:

```xml
<ItemGroup>
  <ProjectReference Include="..\\AnimatedBox\\AnimatedBox.vbproj" />
</ItemGroup>
```

## Namespace

All public types are available through the following namespace:

```vb
Imports CoreSuite.Controls
```

## Adding the control to a form

After installing or referencing the project, add `AnimatedBox` to a Windows Forms form through the Visual Studio Toolbox or create it programmatically.

```vb
Imports CoreSuite.Controls
Public Class MainForm
    Private ReadOnly LoadingAnimation As New AnimatedBox()
    Public Sub New()
        InitializeComponent()
        LoadingAnimation.Dock = DockStyle.Fill
        LoadingAnimation.ScaleMode = AnimationScaleMode.Centrer
        Controls.Add(LoadingAnimation)
    End Sub
End Class
```

The control does not load or start an animation automatically. Frames must be supplied at runtime through `LoadImages` or `LoadGif`, followed by a call to `StartAnimation`.

## Quick start with a GIF

The following example loads a GIF when the form opens and starts playback:

```vb
Imports CoreSuite.Controls
Public Class MainForm
    Private Sub MainForm\_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Using Gif As Image = Image.FromFile("Assets\\loading.gif")
            AnimatedBox1.LoadGif(Gif)
        End Using
        AnimatedBox1.ScaleMode = AnimationScaleMode.Centrer
        AnimatedBox1.StartAnimation()
    End Sub
    Private Sub MainForm\_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        AnimatedBox1.StopAnimation()
    End Sub
End Class
```

`LoadGif` extracts independent bitmap frames from the supplied GIF. The source GIF can therefore be disposed after `LoadGif` finishes.

## Quick start with individual images

Use `LoadImages` when animation frames are stored as separate files or created dynamically:

```vb
Imports CoreSuite.Controls
Public Class MainForm
    Private Sub MainForm\_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Frames As New List(Of Image) From {Image.FromFile("Assets\\frame-01.png"), Image.FromFile("Assets\\frame-02.png"), Image.FromFile("Assets\\frame-03.png")}
        AnimatedBox1.LoadImages(Frames)
        AnimatedBox1.ScaleMode = AnimationScaleMode.Centrer
        AnimatedBox1.StartAnimation()
    End Sub
End Class
```

> \[!IMPORTANT]
> `LoadImages` stores the supplied `Image` instances directly. The currently loaded images are disposed when the `AnimatedBox` itself is disposed. Do not reuse those same image instances elsewhere after assigning them to the control. Pass cloned images when the original instances must remain owned by another component.

Example using clones:

```vb
Dim Frames As New List(Of Image) From {DirectCast(SourceImage1.Clone(), Image), DirectCast(SourceImage2.Clone(), Image)}
AnimatedBox1.LoadImages(Frames)
AnimatedBox1.StartAnimation()
```

## Starting and stopping playback

### Start an animation

```vb
AnimatedBox1.StartAnimation()
```

`StartAnimation` performs the following actions:

1. Resets the current frame to index `0`.
2. Resets the accumulated frame timing.
3. Restarts the internal `Stopwatch`.
4. Configures the Windows Forms timer with a 15-millisecond polling interval.
5. Starts frame advancement.

Calling `StartAnimation` after `StopAnimation` restarts playback from the first frame. It does not resume from the previously displayed frame.

### Stop an animation

```vb
AnimatedBox1.StopAnimation()
```

`StopAnimation` stops the timer and the stopwatch. The last rendered frame remains visible until the animation is restarted, new frames are loaded, or the control is repainted without frames.

### Replace an animation safely

Stop playback before replacing the current frame collection, then start it again:

```vb
AnimatedBox1.StopAnimation()
Using Gif As Image = Image.FromFile("Assets\\success.gif")
    AnimatedBox1.LoadGif(Gif)
End Using
AnimatedBox1.StartAnimation()
```

This sequence ensures that playback restarts from the first frame of the newly loaded animation.

## Scale modes

The `ScaleMode` property determines how each frame is drawn inside the control.

```vb
AnimatedBox1.ScaleMode = AnimationScaleMode.Centrer
```

|Mode|Behavior|
|-|-|
|`Normal`|Draws the image at its original size and centers it inside the control. Images larger than the control can be clipped.|
|`Fill`|Stretches the image to occupy the complete control area. The original aspect ratio is not preserved.|
|`Centrer`|Scales the image proportionally until it fits inside the control, preserves its aspect ratio, and centers it.|

The default value is:

```vb
AnimationScaleMode.Centrer
```

> \[!NOTE]
> The public enum member is named `Centrer` in the current API. The README uses that exact spelling so examples can be copied directly into an application.

### Original image size

```vb
AnimatedBox1.ScaleMode = AnimationScaleMode.Normal
```

Use `Normal` for pixel-perfect rendering when frame dimensions already match the desired display size.

### Fill the complete control

```vb
AnimatedBox1.ScaleMode = AnimationScaleMode.Fill
```

Use `Fill` when occupying all available space is more important than preserving the source aspect ratio.

### Preserve the aspect ratio

```vb
AnimatedBox1.ScaleMode = AnimationScaleMode.Centrer
```

Use `Centrer` for most loading animations, illustrations, status indicators, and other content that should not be distorted.

## Docking and layout

Because the control inherits from `Panel`, it can be combined with standard Windows Forms layout properties.

### Fill a container

```vb
AnimatedBox1.Dock = DockStyle.Fill
AnimatedBox1.ScaleMode = AnimationScaleMode.Centrer
```

### Keep a fixed size and center the frames

```vb
AnimatedBox1.Size = New Size(128, 128)
AnimatedBox1.Anchor = AnchorStyles.None
AnimatedBox1.ScaleMode = AnimationScaleMode.Normal
```

### Transparent background

```vb
AnimatedBox1.BackColor = Color.Transparent
```

Transparent backgrounds are supported through the control style configuration. The visual result still depends on the parent control and normal Windows Forms transparency behavior.

## Typical loading overlay

`AnimatedBox` can be used inside an overlay panel while a task is running:

```vb
Private Async Sub LoadButton\_Click(sender As Object, e As EventArgs) Handles LoadButton.Click
    LoadingPanel.Visible = True
    AnimatedBox1.StartAnimation()
    Try
        Await LoadDataAsync()
    Finally
        AnimatedBox1.StopAnimation()
        LoadingPanel.Visible = False
    End Try
End Sub
```

The animation timer runs on the Windows Forms UI thread. Long-running synchronous work on that same thread prevents the animation from repainting. Use asynchronous operations or move CPU-intensive work away from the UI thread.

## GIF behavior

`LoadGif` extracts every frame from the first frame dimension reported by the GIF and converts each one into a separate `Bitmap`.

```vb
Using Gif As Image = Image.FromFile("Assets\\processing.gif")
    AnimatedBox1.LoadGif(Gif)
End Using
```

The current implementation does not read the frame-delay metadata stored in the GIF. Every extracted frame receives the default `AnimationFrame` delay of `0.03` seconds, which is approximately 33 frames per second.

This means GIFs with variable frame timing or a different intended playback speed are displayed using a uniform delay.

## Frame timing

Each loaded image is internally represented by an `AnimationFrame`.

```vb
Dim Frame As New AnimationFrame(MyImage, 0.05)
```

The `Delay` value is expressed in seconds:

|Delay|Approximate rate|
|-:|-:|
|`0.016`|60 FPS|
|`0.03`|33 FPS|
|`0.04`|25 FPS|
|`0.05`|20 FPS|
|`0.1`|10 FPS|

|
The default delay is:

```vb
0.03
```

> \[!IMPORTANT]
> Although `AnimationFrame` is public, the current `AnimatedBox` loading methods accept only `List(Of Image)` or a GIF. They create frames internally with the default delay. Custom `AnimationFrame.Delay` values cannot currently be passed to the control through its public loading API.

## Rendering quality

The control configures the graphics context with the following rendering options:

* `InterpolationMode.HighQualityBicubic`
* `SmoothingMode.HighQuality`
* `PixelOffsetMode.Half`
* `ControlStyles.OptimizedDoubleBuffer`
* `ControlStyles.AllPaintingInWmPaint`
* `ControlStyles.UserPaint`

These settings improve resized-image quality and reduce visible flickering. Bicubic scaling can consume more processing time than lower-quality interpolation, especially for large controls or high-resolution frames.

## API reference

### `AnimatedBox`

Represents a Windows Forms panel that displays frame-based animations.

```vb
Public Class AnimatedBox
    Inherits Panel
```

### Constructor

```vb
Public Sub New()
```

Creates an empty control, enables optimized painting, initializes the animation timer and stopwatch, and sets `ScaleMode` to `AnimationScaleMode.Centrer`.

### `ScaleMode`

```vb
Public Property ScaleMode As AnimationScaleMode
```

Gets or sets the scaling behavior used when rendering animation frames.

|Property detail|Value|
|-|-|
|Category|`AnimatedBox`|
|Default assigned by constructor|`AnimationScaleMode.Centrer`|
|Change notification event|None|

Changing this property affects the next repaint. Call `Invalidate` manually only when immediate repainting is required while playback is stopped.

```vb
AnimatedBox1.ScaleMode = AnimationScaleMode.Fill
AnimatedBox1.Invalidate()
```

### `LoadImages`

```vb
Public Sub LoadImages(Images As List(Of Image))
```

Clears the internal frame list and adds one frame for each supplied image. Every new frame uses the default delay of `0.03` seconds.

```vb
Dim Frames As New List(Of Image) From {Image.FromFile("frame1.png"), Image.FromFile("frame2.png")}
AnimatedBox1.LoadImages(Frames)
```

The method does not automatically start playback and does not reset playback timing. Use the following sequence when replacing a running animation:

```vb
AnimatedBox1.StopAnimation()
AnimatedBox1.LoadImages(Frames)
AnimatedBox1.StartAnimation()
```

`Images` must not be `Nothing`. An empty list is accepted and results in no frame being rendered.

### `LoadGif`

```vb
Public Sub LoadGif(Gif As Image)
```

Extracts all frames from the supplied GIF, clears the internal frame collection, and stores the extracted bitmaps as animation frames.

```vb
Using Gif As Image = Image.FromFile("animation.gif")
    AnimatedBox1.LoadGif(Gif)
End Using
AnimatedBox1.StartAnimation()
```

The supplied `Gif` must not be `Nothing`. The source image is not retained by the control and is not disposed by `LoadGif`.

### `StartAnimation`

```vb
Public Sub StartAnimation()
```

Starts playback from the first loaded frame. The method can be called when no frames are loaded, but the control will remain visually empty until frames are supplied.

### `StopAnimation`

```vb
Public Sub StopAnimation()
```

Stops playback without clearing the current frames.

### `OnPaint`

```vb
Protected Overrides Sub OnPaint(e As PaintEventArgs)
```

Draws the current frame according to `ScaleMode`. Derived controls can override this method to add borders, overlays, labels, or custom visual effects.

```vb
Public Class StatusAnimatedBox
    Inherits AnimatedBox
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        TextRenderer.DrawText(e.Graphics, "Loading...", Font, ClientRectangle, ForeColor, TextFormatFlags.Bottom Or TextFormatFlags.HorizontalCenter)
    End Sub
End Class
```

### `Dispose`

```vb
Protected Overrides Sub Dispose(disposing As Boolean)
```

Stops the internal timer and stopwatch, disposes every image in the currently loaded frame collection, clears the collection, and then invokes the base implementation.

## `AnimationFrame`

Represents one animation frame and its display duration.

```vb
Public Class AnimationFrame
```

### Constructor

```vb
Public Sub New(Image As Image, Optional Delay As Double = 0.03)
```

|Parameter|Description|
|-|-|
|`Image`|Image displayed by the frame.|
|`Delay`|Number of seconds before playback advances to the next frame. The default is `0.03`.|

### `Image`

```vb
Public Property Image As Image
```

Gets or sets the image associated with the frame.

### `Delay`

```vb
Public Property Delay As Double
```

Gets or sets the frame duration in seconds.

The current class does not validate the value. Applications creating `AnimationFrame` objects directly should use a positive finite number.

## `AnimationScaleMode`

Defines how animation frames are rendered inside an `AnimatedBox`.

```vb
Public Enum AnimationScaleMode
    Normal
    Fill
    Centrer
End Enum
```

|Member|Description|
|-|-|
|`Normal`|Keeps the original image dimensions and centers the frame.|
|`Fill`|Stretches the frame to the complete client area without preserving its aspect ratio.|
|`Centrer`|Fits the frame proportionally inside the client area and centers it.|

## Standard panel functionality

Since `AnimatedBox` inherits from `Panel`, it also supports standard members such as:

* `Anchor`
* `AutoScroll`
* `BackColor`
* `BackgroundImage`
* `BorderStyle`
* `Controls`
* `Dock`
* `Enabled`
* `Location`
* `Margin`
* `Padding`
* `Size`
* `Visible`
* `Click`
* `Paint`
* `Resize`
* `VisibleChanged`

No additional setup is required for these members.

## Threading considerations

`AnimatedBox` is a Windows Forms control and follows the normal Windows Forms threading model.

* Create and access the control from the UI thread.
* Call `LoadImages`, `LoadGif`, `StartAnimation`, and `StopAnimation` from the UI thread.
* Do not update the control directly from a worker thread.
* Keep the UI thread responsive so timer ticks and repaint operations can execute.

Use `Invoke` or `BeginInvoke` when animation state must be changed from background work:

```vb
BeginInvoke(Sub() AnimatedBox1.StopAnimation())
```

## Resource management

### Images loaded with `LoadImages`

The image objects are retained directly and disposed when the control is disposed. Treat the control as the owner of the currently loaded images.

### GIF images loaded with `LoadGif`

The source GIF is used only during extraction. The extracted bitmap frames are owned by the control, while the original GIF remains owned by the caller.

### Replacing frames

The current implementation clears the previous internal collection when another animation is loaded. Applications that frequently replace sequences created through `LoadImages` should manage source-image ownership carefully and avoid keeping unnecessary references.

## Current limitations

* GIF frame-delay metadata is not read.
* All frames loaded through `LoadImages` or `LoadGif` use a `0.03`-second delay.
* Custom `AnimationFrame` objects cannot be loaded directly into `AnimatedBox` through the current public API.
* Playback always loops and cannot currently be configured to run only once.
* `StartAnimation` always restarts from the first frame.
* There are no public pause, resume, current-frame, frame-count, playback-speed, or completion members.
* There are no animation lifecycle events.
* Frames are loaded into memory in full rather than streamed on demand.
* The control does not perform argument validation for `Nothing` images or invalid frame delays.

## Recommended usage practices

* Use `Centrer` when images must retain their proportions.
* Keep frame dimensions close to the control size to reduce scaling work.
* Stop playback before loading a different animation.
* Use asynchronous work to keep the UI responsive.
* Clone images before loading them when another component must continue using the originals.
* Dispose the form or control normally so the currently loaded frame images are released.
* Avoid very large GIFs or long sequences when memory use is a concern.

## Package information

|Item|Value|
|-|-|
|Package|`CoreSuite.AnimatedBox`|
|Namespace|`CoreSuite.Controls`|
|Assembly|`CoreSuite.AnimatedBox`|
|Target framework|`net8.0-windows`|
|UI framework|Windows Forms|
|Language|Visual Basic .NET|
|External dependencies|None|

## License

CoreSuite is distributed under the MIT License.

