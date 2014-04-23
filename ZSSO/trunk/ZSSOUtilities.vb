Imports System.Net
Imports System.Web.Http
Imports System.IO
Imports System.Runtime.Caching
Imports System.Web.Script.Serialization
Imports System.Security.Cryptography
Imports System.Data.SqlClient
Imports System.Globalization

Public Class ZSSOUtilities
    Public Shared emailExpression As New Regex("^[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,4})$")
    Public Shared oSerializer As New JavaScriptSerializer

    Public Shared Function Login(oConnexion As SqlConnection, Email As String, Password As String) As Boolean
        Dim QueryString = "SELECT * " & _
            "FROM Account " & _
            "WHERE Email=@email"

        Using oSqlCmdSelect As New SqlCommand(QueryString, oConnexion)
            oSqlCmdSelect.Parameters.AddWithValue("@email", Email)

            Try
                Using md5Hash As MD5 = MD5.Create()

                    Dim QueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()
                    Dim AccountSalt = ""
                    Dim AccountHash = ""

                    While QueryResult.Read()
                        AccountSalt = QueryResult(QueryResult.GetOrdinal("Salt"))
                        AccountHash = QueryResult(QueryResult.GetOrdinal("Password"))
                    End While

                    Dim HashToCheck As String = ZSSOUtilities.GetMd5Hash(md5Hash, Password + AccountSalt)

                    If (AccountSalt.Length > 0 And AccountHash.Length > 0 And (String.Compare(AccountHash, HashToCheck))) Or AccountSalt.Length = 0 Or AccountHash.Length = 0 Then
                        Return False
                    End If
                End Using

            Catch ex As Exception
                Return False
            End Try
        End Using
        Return True
    End Function

    Public Shared Function WriteLog(Text As String)
        Using oConnexion As New SqlConnection("Data Source=(LocalDB)\v11.0;AttachDbFilename=C:\Users\ZPFr1\Desktop\zsso\ZSSO\trunk\App_Data\Database1.mdf;Integrated Security=True;MultipleActiveResultSets=True")
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

    Public Shared Function CheckRequests(Ip As String)
        Dim cacheMemory As ObjectCache = MemoryCache.Default
        Dim cachedCounterByIp As Int32

        Dim cachedRequests = TryCast(cacheMemory("requests"), Dictionary(Of String, Integer))
        If IsNothing(cachedRequests) Then
            cachedRequests = New Dictionary(Of String, Integer)
        End If
        If cachedRequests.ContainsKey(Ip) Then
            cachedCounterByIp = CInt(cachedRequests(Ip))
            If IsNothing(cachedCounterByIp) Then
                cachedCounterByIp = 0
            End If
        End If
        cachedCounterByIp = cachedCounterByIp + 1
        cachedRequests(Ip) = cachedCounterByIp
        cacheMemory.Set("requests", cachedRequests, DateTime.Now.AddSeconds(5.0), Nothing)
        Return cachedCounterByIp
    End Function

    Public Shared Function GetLocation(Ip As String) As Dictionary(Of String, String)
        Try
            Dim rssReq As WebRequest = WebRequest.Create("https://freegeoip.net/json/" + Ip)
            Dim respStream As Stream = rssReq.GetResponse().GetResponseStream()
            Dim response As String = New StreamReader(respStream).ReadToEnd()
            Dim LocationData As Dictionary(Of String, String)

            LocationData = ZSSOUtilities.oSerializer.Deserialize(Of Dictionary(Of String, String))(response)

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
        '        dist = dist * 1.609344 'kilometers
        Return dist
    End Function

    Private Shared Function deg2rad(ByVal deg As Double) As Double
        Return (deg * Math.PI / 180.0)
    End Function

    Private Shared Function rad2deg(ByVal rad As Double) As Double
        Return rad / Math.PI * 180.0
    End Function

    Public Shared Function GetMd5Hash(ByVal md5Hash As MD5, ByVal input As String) As String

        ' Convert the input string to a byte array and compute the hash.
        Dim data As Byte() = md5Hash.ComputeHash(Encoding.UTF8.GetBytes(input))

        ' Create a new Stringbuilder to collect the bytes
        ' and create a string.
        Dim sBuilder As New StringBuilder()

        ' Loop through each byte of the hashed data 
        ' and format each one as a hexadecimal string.
        Dim i As Integer
        For i = 0 To data.Length - 1
            sBuilder.Append(data(i).ToString("x2"))
        Next i

        ' Return the hexadecimal string.
        Return sBuilder.ToString()

    End Function 'GetMd5Hash
End Class
