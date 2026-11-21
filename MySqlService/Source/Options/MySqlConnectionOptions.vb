Imports MySql.Data.MySqlClient
''' <summary>Connection settings owned by CoreSuite. A service snapshots these settings at construction.</summary>
Public NotInheritable Class MySqlConnectionOptions
    Private ReadOnly Builder As New MySqlConnectionStringBuilder()
    Friend Function BuildConnectionString() As String
        Return Builder.ConnectionString
    End Function
    Public Property [PipeName] As String
        Get
            Return Builder.PipeName
        End Get
        Set(value As String)
            Builder.PipeName = value
        End Set
    End Property
    Public Property [UseCompression] As Boolean
        Get
            Return Builder.UseCompression
        End Get
        Set(value As Boolean)
            Builder.UseCompression = value
        End Set
    End Property
    Public Property [AllowBatch] As Boolean
        Get
            Return Builder.AllowBatch
        End Get
        Set(value As Boolean)
            Builder.AllowBatch = value
        End Set
    End Property
    Public Property [Logging] As Boolean
        Get
            Return Builder.Logging
        End Get
        Set(value As Boolean)
            Builder.Logging = value
        End Set
    End Property
    Public Property [SharedMemoryName] As String
        Get
            Return Builder.SharedMemoryName
        End Get
        Set(value As String)
            Builder.SharedMemoryName = value
        End Set
    End Property
    Public Property [DefaultCommandTimeout] As UInteger
        Get
            Return Builder.DefaultCommandTimeout
        End Get
        Set(value As UInteger)
            Builder.DefaultCommandTimeout = value
        End Set
    End Property
    Public Property [ConnectionTimeout] As UInteger
        Get
            Return Builder.ConnectionTimeout
        End Get
        Set(value As UInteger)
            Builder.ConnectionTimeout = value
        End Set
    End Property
    Public Property [AllowLoadLocalInfile] As Boolean
        Get
            Return Builder.AllowLoadLocalInfile
        End Get
        Set(value As Boolean)
            Builder.AllowLoadLocalInfile = value
        End Set
    End Property
    Public Property [AllowLoadLocalInfileInPath] As String
        Get
            Return Builder.AllowLoadLocalInfileInPath
        End Get
        Set(value As String)
            Builder.AllowLoadLocalInfileInPath = value
        End Set
    End Property
    Public Property [PersistSecurityInfo] As Boolean
        Get
            Return Builder.PersistSecurityInfo
        End Get
        Set(value As Boolean)
            Builder.PersistSecurityInfo = value
        End Set
    End Property
    Public Property [IntegratedSecurity] As Boolean
        Get
            Return Builder.IntegratedSecurity
        End Get
        Set(value As Boolean)
            Builder.IntegratedSecurity = value
        End Set
    End Property
    Public Property [AllowPublicKeyRetrieval] As Boolean
        Get
            Return Builder.AllowPublicKeyRetrieval
        End Get
        Set(value As Boolean)
            Builder.AllowPublicKeyRetrieval = value
        End Set
    End Property
    Public Property [DefaultAuthenticationPlugin] As String
        Get
            Return Builder.DefaultAuthenticationPlugin
        End Get
        Set(value As String)
            Builder.DefaultAuthenticationPlugin = value
        End Set
    End Property
    Public Property [OciConfigFile] As String
        Get
            Return Builder.OciConfigFile
        End Get
        Set(value As String)
            Builder.OciConfigFile = value
        End Set
    End Property
    Public Property [OciConfigProfile] As String
        Get
            Return Builder.OciConfigProfile
        End Get
        Set(value As String)
            Builder.OciConfigProfile = value
        End Set
    End Property
    Public Property [KerberosAuthMode] As MySqlConnectionKerberosAuthMode
        Get
            Return CType(CInt(Builder.KerberosAuthMode), MySqlConnectionKerberosAuthMode)
        End Get
        Set(value As MySqlConnectionKerberosAuthMode)
            If Not [Enum].IsDefined(GetType(MySqlConnectionKerberosAuthMode), value) Then Throw New ArgumentOutOfRangeException(NameOf(value))
            Builder.KerberosAuthMode = CType(CInt(value), Global.MySql.Data.MySqlClient.KerberosAuthMode)
        End Set
    End Property
    Public Property [OpenIdIdentityToken] As String
        Get
            Return Builder.OpenIdIdentityToken
        End Get
        Set(value As String)
            Builder.OpenIdIdentityToken = value
        End Set
    End Property
    Public Property [AllowZeroDateTime] As Boolean
        Get
            Return Builder.AllowZeroDateTime
        End Get
        Set(value As Boolean)
            Builder.AllowZeroDateTime = value
        End Set
    End Property
    Public Property [ConvertZeroDateTime] As Boolean
        Get
            Return Builder.ConvertZeroDateTime
        End Get
        Set(value As Boolean)
            Builder.ConvertZeroDateTime = value
        End Set
    End Property
    Public Property [UseUsageAdvisor] As Boolean
        Get
            Return Builder.UseUsageAdvisor
        End Get
        Set(value As Boolean)
            Builder.UseUsageAdvisor = value
        End Set
    End Property
    Public Property [ProcedureCacheSize] As UInteger
        Get
            Return Builder.ProcedureCacheSize
        End Get
        Set(value As UInteger)
            Builder.ProcedureCacheSize = value
        End Set
    End Property
    Public Property [UsePerformanceMonitor] As Boolean
        Get
            Return Builder.UsePerformanceMonitor
        End Get
        Set(value As Boolean)
            Builder.UsePerformanceMonitor = value
        End Set
    End Property
    Public Property [AutoEnlist] As Boolean
        Get
            Return Builder.AutoEnlist
        End Get
        Set(value As Boolean)
            Builder.AutoEnlist = value
        End Set
    End Property
    Public Property [IncludeSecurityAsserts] As Boolean
        Get
            Return Builder.IncludeSecurityAsserts
        End Get
        Set(value As Boolean)
            Builder.IncludeSecurityAsserts = value
        End Set
    End Property
    Public Property [RespectBinaryFlags] As Boolean
        Get
            Return Builder.RespectBinaryFlags
        End Get
        Set(value As Boolean)
            Builder.RespectBinaryFlags = value
        End Set
    End Property
    Public Property [TreatTinyAsBoolean] As Boolean
        Get
            Return Builder.TreatTinyAsBoolean
        End Get
        Set(value As Boolean)
            Builder.TreatTinyAsBoolean = value
        End Set
    End Property
    Public Property [AllowUserVariables] As Boolean
        Get
            Return Builder.AllowUserVariables
        End Get
        Set(value As Boolean)
            Builder.AllowUserVariables = value
        End Set
    End Property
    Public Property [InteractiveSession] As Boolean
        Get
            Return Builder.InteractiveSession
        End Get
        Set(value As Boolean)
            Builder.InteractiveSession = value
        End Set
    End Property
    Public Property [FunctionsReturnString] As Boolean
        Get
            Return Builder.FunctionsReturnString
        End Get
        Set(value As Boolean)
            Builder.FunctionsReturnString = value
        End Set
    End Property
    Public Property [UseAffectedRows] As Boolean
        Get
            Return Builder.UseAffectedRows
        End Get
        Set(value As Boolean)
            Builder.UseAffectedRows = value
        End Set
    End Property
    Public Property [OldGuids] As Boolean
        Get
            Return Builder.OldGuids
        End Get
        Set(value As Boolean)
            Builder.OldGuids = value
        End Set
    End Property
    Public Property [SqlServerMode] As Boolean
        Get
            Return Builder.SqlServerMode
        End Get
        Set(value As Boolean)
            Builder.SqlServerMode = value
        End Set
    End Property
    Public Property [TableCaching] As Boolean
        Get
            Return Builder.TableCaching
        End Get
        Set(value As Boolean)
            Builder.TableCaching = value
        End Set
    End Property
    Public Property [DefaultTableCacheAge] As Integer
        Get
            Return Builder.DefaultTableCacheAge
        End Get
        Set(value As Integer)
            Builder.DefaultTableCacheAge = value
        End Set
    End Property
    Public Property [CheckParameters] As Boolean
        Get
            Return Builder.CheckParameters
        End Get
        Set(value As Boolean)
            Builder.CheckParameters = value
        End Set
    End Property
    Public Property [Replication] As Boolean
        Get
            Return Builder.Replication
        End Get
        Set(value As Boolean)
            Builder.Replication = value
        End Set
    End Property
    Public Property [ExceptionInterceptors] As String
        Get
            Return Builder.ExceptionInterceptors
        End Get
        Set(value As String)
            Builder.ExceptionInterceptors = value
        End Set
    End Property
    Public Property [CommandInterceptors] As String
        Get
            Return Builder.CommandInterceptors
        End Get
        Set(value As String)
            Builder.CommandInterceptors = value
        End Set
    End Property
    Public Property [RewriteBatchedStatements] As Boolean
        Get
            Return Builder.RewriteBatchedStatements
        End Get
        Set(value As Boolean)
            Builder.RewriteBatchedStatements = value
        End Set
    End Property
    Public Property [ConnectionLifeTime] As UInteger
        Get
            Return Builder.ConnectionLifeTime
        End Get
        Set(value As UInteger)
            Builder.ConnectionLifeTime = value
        End Set
    End Property
    Public Property [Pooling] As Boolean
        Get
            Return Builder.Pooling
        End Get
        Set(value As Boolean)
            Builder.Pooling = value
        End Set
    End Property
    Public Property [MinimumPoolSize] As UInteger
        Get
            Return Builder.MinimumPoolSize
        End Get
        Set(value As UInteger)
            Builder.MinimumPoolSize = value
        End Set
    End Property
    Public Property [MaximumPoolSize] As UInteger
        Get
            Return Builder.MaximumPoolSize
        End Get
        Set(value As UInteger)
            Builder.MaximumPoolSize = value
        End Set
    End Property
    Public Property [ConnectionReset] As Boolean
        Get
            Return Builder.ConnectionReset
        End Get
        Set(value As Boolean)
            Builder.ConnectionReset = value
        End Set
    End Property
    Public Property [CacheServerProperties] As Boolean
        Get
            Return Builder.CacheServerProperties
        End Get
        Set(value As Boolean)
            Builder.CacheServerProperties = value
        End Set
    End Property
    Public Property [TreatBlobsAsUTF8] As Boolean
        Get
            Return Builder.TreatBlobsAsUTF8
        End Get
        Set(value As Boolean)
            Builder.TreatBlobsAsUTF8 = value
        End Set
    End Property
    Public Property [BlobAsUTF8IncludePattern] As String
        Get
            Return Builder.BlobAsUTF8IncludePattern
        End Get
        Set(value As String)
            Builder.BlobAsUTF8IncludePattern = value
        End Set
    End Property
    Public Property [BlobAsUTF8ExcludePattern] As String
        Get
            Return Builder.BlobAsUTF8ExcludePattern
        End Get
        Set(value As String)
            Builder.BlobAsUTF8ExcludePattern = value
        End Set
    End Property
    Public Property [UseDefaultCommandTimeoutForEF] As Boolean
        Get
            Return Builder.UseDefaultCommandTimeoutForEF
        End Get
        Set(value As Boolean)
            Builder.UseDefaultCommandTimeoutForEF = value
        End Set
    End Property
    Public Property [Server] As String
        Get
            Return Builder.Server
        End Get
        Set(value As String)
            Builder.Server = value
        End Set
    End Property
    Public Property [Database] As String
        Get
            Return Builder.Database
        End Get
        Set(value As String)
            Builder.Database = value
        End Set
    End Property
    Public Property [ConnectionProtocol] As MySqlConnectionConnectionProtocol
        Get
            Return CType(CInt(Builder.ConnectionProtocol), MySqlConnectionConnectionProtocol)
        End Get
        Set(value As MySqlConnectionConnectionProtocol)
            If Not [Enum].IsDefined(GetType(MySqlConnectionConnectionProtocol), value) Then Throw New ArgumentOutOfRangeException(NameOf(value))
            Builder.ConnectionProtocol = CType(CInt(value), Global.MySql.Data.MySqlClient.MySqlConnectionProtocol)
        End Set
    End Property
    Public Property [Port] As UInteger
        Get
            Return Builder.Port
        End Get
        Set(value As UInteger)
            Builder.Port = value
        End Set
    End Property
    Public Property [DnsSrv] As Boolean
        Get
            Return Builder.DnsSrv
        End Get
        Set(value As Boolean)
            Builder.DnsSrv = value
        End Set
    End Property
    Public Property [UserID] As String
        Get
            Return Builder.UserID
        End Get
        Set(value As String)
            Builder.UserID = value
        End Set
    End Property
    Public Property [Password] As String
        Get
            Return Builder.Password
        End Get
        Set(value As String)
            Builder.Password = value
        End Set
    End Property
    Public Property [Password2] As String
        Get
            Return Builder.Password2
        End Get
        Set(value As String)
            Builder.Password2 = value
        End Set
    End Property
    Public Property [Password3] As String
        Get
            Return Builder.Password3
        End Get
        Set(value As String)
            Builder.Password3 = value
        End Set
    End Property
    Public Property [CertificateFile] As String
        Get
            Return Builder.CertificateFile
        End Get
        Set(value As String)
            Builder.CertificateFile = value
        End Set
    End Property
    Public Property [CertificatePassword] As String
        Get
            Return Builder.CertificatePassword
        End Get
        Set(value As String)
            Builder.CertificatePassword = value
        End Set
    End Property
    Public Property [CertificateStoreLocation] As MySqlConnectionCertificateStoreLocation
        Get
            Return CType(CInt(Builder.CertificateStoreLocation), MySqlConnectionCertificateStoreLocation)
        End Get
        Set(value As MySqlConnectionCertificateStoreLocation)
            If Not [Enum].IsDefined(GetType(MySqlConnectionCertificateStoreLocation), value) Then Throw New ArgumentOutOfRangeException(NameOf(value))
            Builder.CertificateStoreLocation = CType(CInt(value), Global.MySql.Data.MySqlClient.MySqlCertificateStoreLocation)
        End Set
    End Property
    Public Property [CertificateThumbprint] As String
        Get
            Return Builder.CertificateThumbprint
        End Get
        Set(value As String)
            Builder.CertificateThumbprint = value
        End Set
    End Property
    Public Property [SslMode] As MySqlConnectionSslMode
        Get
            Return CType(CInt(Builder.SslMode), MySqlConnectionSslMode)
        End Get
        Set(value As MySqlConnectionSslMode)
            If Not [Enum].IsDefined(GetType(MySqlConnectionSslMode), value) Then Throw New ArgumentOutOfRangeException(NameOf(value))
            Builder.SslMode = CType(CInt(value), Global.MySql.Data.MySqlClient.MySqlSslMode)
        End Set
    End Property
    Public Property [SslCa] As String
        Get
            Return Builder.SslCa
        End Get
        Set(value As String)
            Builder.SslCa = value
        End Set
    End Property
    Public Property [TlsVersion] As String
        Get
            Return Builder.TlsVersion
        End Get
        Set(value As String)
            Builder.TlsVersion = value
        End Set
    End Property
    Public Property [SslKey] As String
        Get
            Return Builder.SslKey
        End Get
        Set(value As String)
            Builder.SslKey = value
        End Set
    End Property
    Public Property [SslCert] As String
        Get
            Return Builder.SslCert
        End Get
        Set(value As String)
            Builder.SslCert = value
        End Set
    End Property
    Public Property [Keepalive] As UInteger
        Get
            Return Builder.Keepalive
        End Get
        Set(value As UInteger)
            Builder.Keepalive = value
        End Set
    End Property
    Public Property [CharacterSet] As String
        Get
            Return Builder.CharacterSet
        End Get
        Set(value As String)
            Builder.CharacterSet = value
        End Set
    End Property
    Public Property [BrowsableConnectionString] As Boolean
        Get
            Return Builder.BrowsableConnectionString
        End Get
        Set(value As Boolean)
            Builder.BrowsableConnectionString = value
        End Set
    End Property
End Class
