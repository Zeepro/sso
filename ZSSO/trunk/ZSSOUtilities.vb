Imports System.Net
Imports System.Web.Http
Imports System.IO
Imports System.Runtime.Caching
Imports System.Web.Script.Serialization
Imports System.Security.Cryptography
Imports System.Data.SqlClient
Imports System.Globalization
Imports BCrypt.Net.BCrypt
Imports System.Net.Cache

Public Class ZSSOUtilities
    Public Shared emailExpression As New Regex("^[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,4})$")
    Public Shared oSerializer As New JavaScriptSerializer

    Public Shared Function Login(oConnexion As SqlConnection, Email As String, Password As String) As Boolean
        Dim QueryString = "SELECT TOP 1 * " & _
            "FROM Account " & _
            "WHERE Email=@email"

        Using oSqlCmdSelect As New SqlCommand(QueryString, oConnexion)
            oSqlCmdSelect.Parameters.AddWithValue("@email", Email)

            Try
                Using QueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()
                    Dim AccountHash = ""

                    If QueryResult.Read() Then
                        AccountHash = QueryResult(QueryResult.GetOrdinal("Password"))
                    End If

                    If Not BCrypt.Net.BCrypt.Verify(Password, AccountHash) Then
                        Return False
                    End If
                End Using
            Catch ex As Exception
                Return False
            End Try
        End Using
        Return True
    End Function

    Public Shared Function getConnectionString() As String
        Dim rootWebConfig As System.Configuration.Configuration
        rootWebConfig = System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration("/Web")
        Dim connString As System.Configuration.ConnectionStringSettings = Nothing
        If (rootWebConfig.ConnectionStrings.ConnectionStrings.Count > 0) Then
            connString = rootWebConfig.ConnectionStrings.ConnectionStrings("ZSSODb")
        End If
        If IsNothing(connString) Then
            Return ""
        End If
        Return connString.ToString()
    End Function

    Public Shared Function WriteLog(Text As String)
        Using oConnexion As New SqlConnection(getConnectionString())
            oConnexion.Open()
            Dim QueryString As String = "INSERT INTO Logs VALUES (DEFAULT, @text)"

            Using oSqlCmdInsert As New SqlCommand(QueryString, oConnexion)
                Try
                    oSqlCmdInsert.Parameters.AddWithValue("@text", Text)
                    oSqlCmdInsert.ExecuteNonQuery()
                Catch ex As Exception

                End Try
            End Using
        End Using
        Return Nothing
    End Function

    Public Shared Function CheckRequests(Ip As String, Type As String)
        Dim HttpCache As Caching.Cache = HttpRuntime.Cache
        Dim cachedCounterByIp As Int32

        Try
            cachedCounterByIp = CInt(HttpCache("request_" + Type + "_" + Ip))
        Catch
            cachedCounterByIp = 0
        End Try

        cachedCounterByIp = cachedCounterByIp + 1
        HttpCache.Insert("request_" + Type + "_" + Ip, cachedCounterByIp, Nothing, DateTime.Now.AddSeconds(5.0), TimeSpan.Zero)
        Return cachedCounterByIp
    End Function

    Public Shared Function GetLocation(Ip As String) As Dictionary(Of String, String)
        Dim LocationData As Dictionary(Of String, String)
        Try
            Dim rssReq As WebRequest = WebRequest.Create("https://freegeoip.net/json/" + Ip)
            'TODO: Using rssReq AS New...
            Using respStream As Stream = rssReq.GetResponse().GetResponseStream()
                Dim response As String = New StreamReader(respStream).ReadToEnd()

                LocationData = ZSSOUtilities.oSerializer.Deserialize(Of Dictionary(Of String, String))(response)
            End Using
            Return LocationData
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Shared Function CalculateDistanceBetweenCoordinates(ServerLatitudeString As String, ServerLongitudeString As String, PrinterLatitudeString As String, PrinterLongitudeString As String)
        Dim ciClone As CultureInfo = CType(CultureInfo.InvariantCulture.Clone(), CultureInfo)
        ciClone.NumberFormat.NumberDecimalSeparator = "."

        Dim ServerLatitude As Double = CDbl(ServerLatitudeString)
        Dim ServerLongitude As Double = CDbl(ServerLongitudeString)
        Dim PrinterLatitude As Double = CDbl(PrinterLatitudeString)
        Dim PrinterLongitude As Double = CDbl(PrinterLongitudeString)

        Dim theta As Double = ServerLongitude - PrinterLongitude
        Dim dist As Double = Math.Sin(deg2rad(ServerLatitude)) * Math.Sin(deg2rad(PrinterLatitude)) + Math.Cos(deg2rad(ServerLatitude)) * Math.Cos(deg2rad(PrinterLatitude)) * Math.Cos(deg2rad(theta))
        dist = Math.Acos(dist)
        dist = rad2deg(dist)
        dist = dist * 60 * 1.1515 'miles
        ' dist = dist * 1.609344 'kilometers
        Return dist
    End Function

    Private Shared Function deg2rad(ByVal deg As Double) As Double
        Return (deg * Math.PI / 180.0)
    End Function

    Private Shared Function rad2deg(ByVal rad As Double) As Double
        Return rad / Math.PI * 180.0
    End Function
End Class
