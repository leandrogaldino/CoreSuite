''' <summary>
''' Provides helper methods for converting <see cref="Date"/> values
''' to and from Unix epoch time expressed in milliseconds.
''' </summary>
Public Class DateTimeHelper

    ''' <summary>
    ''' Converts the number of milliseconds since the Unix Epoch
    ''' (1970-01-01 00:00:00 UTC) to a date and time in the Brasília time zone.
    ''' </summary>
    ''' <param name="Milliseconds">
    ''' The number of milliseconds since the Unix Epoch (UTC).
    ''' </param>
    ''' <returns>
    ''' A <see cref="Date"/> representing the corresponding date and time
    ''' in the Brasília time zone (UTC-3).
    ''' </returns>
    Public Shared Function DateFromMilliseconds(Milliseconds As Long) As Date
        Dim Epoch As New Date(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)
        Dim UtcDate = Epoch.AddMilliseconds(Milliseconds)
        Dim Tz As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time")
        Return TimeZoneInfo.ConvertTimeFromUtc(UtcDate, Tz)
    End Function


    ''' <summary>
    ''' Converts a date and time in the Brasília time zone to the number of
    ''' milliseconds since the Unix Epoch (1970-01-01 00:00:00 UTC).
    ''' </summary>
    ''' <param name="Date">
    ''' The date and time assumed to be in the Brasília time zone (UTC-3).
    ''' </param>
    ''' <returns>
    ''' The number of milliseconds since the Unix Epoch that represents
    ''' the equivalent instant in UTC.
    ''' </returns>
    Public Shared Function MillisecondsFromDate(ByVal [Date] As Date) As Long
        Dim Tz As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time")
        Dim BsbDate = DateTime.SpecifyKind([Date], DateTimeKind.Unspecified)
        Dim UtcDate = TimeZoneInfo.ConvertTimeToUtc(BsbDate, Tz)
        Dim Epoch As New Date(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)
        Return CLng((UtcDate - Epoch).TotalMilliseconds)
    End Function

    ''' <summary>
    ''' Gets the current date and time in the Brasília time zone,
    ''' regardless of the system's local time zone configuration.
    ''' </summary>
    ''' <returns>
    ''' The current date and time in Brasília (UTC-3).
    ''' </returns>
    Public Shared Function Now() As Date
        Dim Tz = TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time")
        Dim BsbNow = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, Tz)
        Return BsbNow
    End Function
End Class
