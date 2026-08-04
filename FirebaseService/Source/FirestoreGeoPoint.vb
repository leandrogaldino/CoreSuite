''' <summary>
''' Represents a geographic coordinate stored in a Cloud Firestore <c>geoPointValue</c>.
''' </summary>
Public NotInheritable Class FirestoreGeoPoint
    ''' <summary>
    ''' Gets the latitude in degrees, from -90 to 90.
    ''' </summary>
    Public ReadOnly Property Latitude As Double
    ''' <summary>
    ''' Gets the longitude in degrees, from -180 to 180.
    ''' </summary>
    Public ReadOnly Property Longitude As Double
    ''' <summary>
    ''' Initializes a geographic coordinate.
    ''' </summary>
    ''' <param name="Latitude">The latitude in degrees.</param>
    ''' <param name="Longitude">The longitude in degrees.</param>
    Public Sub New(Latitude As Double, Longitude As Double)
        If Double.IsNaN(Latitude) OrElse Double.IsInfinity(Latitude) OrElse Latitude < -90 OrElse Latitude > 90 Then Throw New ArgumentOutOfRangeException(NameOf(Latitude))
        If Double.IsNaN(Longitude) OrElse Double.IsInfinity(Longitude) OrElse Longitude < -180 OrElse Longitude > 180 Then Throw New ArgumentOutOfRangeException(NameOf(Longitude))
        Me.Latitude = Latitude
        Me.Longitude = Longitude
    End Sub
End Class
