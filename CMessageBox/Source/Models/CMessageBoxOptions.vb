''' <summary>
''' Represents the configuration options used to customize the appearance and behavior of the CMessageBox.
''' </summary>
''' <remarks>
''' This class allows customization of message box appearance, icons, fonts, colors,
''' exception handling behavior and additional information included in error reports.
''' </remarks>
Public Class CMessageBoxOptions

    ''' <summary>
    ''' Gets or sets whether detailed exception information should be displayed to the user.
    ''' </summary>
    Public Property ShowExceptionDetails As Boolean = False

    ''' <summary>
    ''' Gets or sets the email configuration used to send exception notifications.
    ''' </summary>
    Public Property ExceptionEmail As CMessageBoxExceptionEmail

    ''' <summary>
    ''' Gets or sets the font used for the message box title.
    ''' </summary>
    Public Property TitleFont As Font = New Font("Segoe UI", 11.25, FontStyle.Bold)

    ''' <summary>
    ''' Gets or sets the foreground color used for the message box title.
    ''' </summary>
    Public Property TitleForeColor As Color = SystemColors.ControlText

    ''' <summary>
    ''' Gets or sets the font used for the message body.
    ''' </summary>
    Public Property MessageFont As Font = New Font("Segoe UI", 9.75)

    ''' <summary>
    ''' Gets or sets the foreground color used for the message body.
    ''' </summary>
    Public Property MessageForeColor As Color = SystemColors.ControlText

    ''' <summary>
    ''' Gets or sets additional information to be included in exception reports.
    ''' </summary>
    ''' <remarks>
    ''' This collection can be used to attach custom contextual data to error reports,
    ''' helping with troubleshooting and diagnostics.
    ''' </remarks>
    Public Property AdditionalInformations As Dictionary(Of String, Object)

    ''' <summary>
    ''' Gets or sets the image displayed for error messages.
    ''' </summary>
    ''' <remarks>
    ''' The image must have exactly 64x64 pixels. Assigning an image with a different
    ''' size will throw an <see cref="ArgumentException"/>.
    ''' </remarks>
    ''' <exception cref="ArgumentException">
    ''' Thrown when the assigned image dimensions are different from 64x64 pixels.
    ''' </exception>
    Public Property ErrorImage As Image

    ''' <summary>
    ''' Gets or sets the image displayed for success messages.
    ''' </summary>
    ''' <remarks>
    ''' The image must have exactly 64x64 pixels. Assigning an image with a different
    ''' size will throw an <see cref="ArgumentException"/>.
    ''' </remarks>
    ''' <exception cref="ArgumentException">
    ''' Thrown when the assigned image dimensions are different from 64x64 pixels.
    ''' </exception>
    Public Property SuccessImage As Image

    ''' <summary>
    ''' Gets or sets the image displayed for informational messages.
    ''' </summary>
    ''' <remarks>
    ''' The image must have exactly 64x64 pixels. Assigning an image with a different
    ''' size will throw an <see cref="ArgumentException"/>.
    ''' </remarks>
    ''' <exception cref="ArgumentException">
    ''' Thrown when the assigned image dimensions are different from 64x64 pixels.
    ''' </exception>
    Public Property InformationImage As Image

    ''' <summary>
    ''' Gets or sets the image displayed for warning messages.
    ''' </summary>
    ''' <remarks>
    ''' The image must have exactly 64x64 pixels. Assigning an image with a different
    ''' size will throw an <see cref="ArgumentException"/>.
    ''' </remarks>
    ''' <exception cref="ArgumentException">
    ''' Thrown when the assigned image dimensions are different from 64x64 pixels.
    ''' </exception>
    Public Property WarningImage As Image

    ''' <summary>
    ''' Gets or sets the image displayed for question messages.
    ''' </summary>
    ''' <remarks>
    ''' The image must have exactly 64x64 pixels. Assigning an image with a different
    ''' size will throw an <see cref="ArgumentException"/>.
    ''' </remarks>
    ''' <exception cref="ArgumentException">
    ''' Thrown when the assigned image dimensions are different from 64x64 pixels.
    ''' </exception>
    Public Property QuestionImage As Image

End Class