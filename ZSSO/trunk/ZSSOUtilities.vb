﻿Imports System.Net
Imports System.Web.Http
Imports System.IO
Imports System.Runtime.Caching
Imports System.Web.Script.Serialization
Imports System.Security.Cryptography
Imports System.Data.SqlClient
Imports System.Globalization
Imports BCrypt.Net.BCrypt
Imports System.Net.Cache
Imports System.Net.Mail

Public Class ZSSOUtilities
    Public Shared oEmailRegex As New Regex("^[-0-9a-zA-Z.+_]+@[-0-9a-zA-Z.+_]+\.[a-zA-Z]{2,4}$")
    Public Shared oSerializer As New JavaScriptSerializer
    Public Shared oAdminEmail As String = "zeepromac@zeepro.com"
    Public Shared oSerialEmail As String = "zeeproserial@zeepro.com"
    Public Shared oZoombaiAdminEmail As String = "changepassword@zeepro.com"
    Private Shared oSha1 As New SHA1CryptoServiceProvider
    Private Shared sKey As String = "zeepro"

    Public Shared Function Login(oConnection As SqlConnection, sEmail As String, sPassword As String) As Boolean
        Dim sQuery = "SELECT TOP 1 * " & _
            "FROM Account " & _
            "WHERE Email=@email AND Deleted IS NULL"

        Using oSqlCmdSelect As New SqlCommand(sQuery, oConnection)
            oSqlCmdSelect.Parameters.AddWithValue("@email", sEmail.ToLower)
            Try
                Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()
                    If oQueryResult.Read() Then
                        Dim bConfirmed As Boolean = oQueryResult(oQueryResult.GetOrdinal("Confirmed"))
                        If BCrypt.Net.BCrypt.Verify(sPassword, oQueryResult(oQueryResult.GetOrdinal("Password"))) And bConfirmed Then
                            Return True
                        End If
                    End If
                End Using
            Catch
            End Try
        End Using

        Return False
    End Function

    Public Shared Function WriteLog(sText As String)
        Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
            oConnection.Open()
            Dim sQuery As String = "INSERT INTO Logs (LogText) VALUES (@text)"

            Using oSqlCmdInsert As New SqlCommand(sQuery, oConnection)
                Try
                    oSqlCmdInsert.Parameters.AddWithValue("@text", sText)
                    oSqlCmdInsert.ExecuteNonQuery()
                Catch ex As Exception
                End Try
            End Using

            sQuery = "DELETE FROM Logs WHERE LogDate <= @threemonth"

            Using oSqlCmdDelete As New SqlCommand(sQuery, oConnection)
                Try
                    oSqlCmdDelete.Parameters.AddWithValue("@threemonth", Date.Today.AddMonths(-3))
                    oSqlCmdDelete.ExecuteNonQuery()
                Catch ex As Exception
                End Try
            End Using

        End Using
        Return Nothing
    End Function

    Public Shared Function CheckRequests(sIp As String, sType As String)
        Dim oHttpCache As Caching.Cache = HttpRuntime.Cache
        Dim iCachedCounterByIp As Int32

        Try
            iCachedCounterByIp = CInt(oHttpCache("request_" & sType & "_" & sIp))
        Catch
            iCachedCounterByIp = 0
        End Try

        iCachedCounterByIp = iCachedCounterByIp + 1
        oHttpCache.Insert("request_" & sType & "_" & sIp, iCachedCounterByIp, Nothing, DateTime.Now.AddSeconds(5.0), TimeSpan.Zero)
        Return iCachedCounterByIp
    End Function

    Public Shared Function GetLocation(sIp As String) As Dictionary(Of String, String)
        Dim arLocationData As Dictionary(Of String, String)
        Dim oHttpCache As Caching.Cache = HttpRuntime.Cache

        Try
            Dim oRequest As WebRequest = WebRequest.Create("http://ip-api.com/json/" & sIp)
            Using oResponseStream As Stream = oRequest.GetResponse().GetResponseStream()
                Dim sResponse As String = New StreamReader(oResponseStream).ReadToEnd()
                Dim arData As Dictionary(Of String, String) = ZSSOUtilities.oSerializer.Deserialize(Of Dictionary(Of String, String))(sResponse)
                arLocationData = New Dictionary(Of String, String)
                arLocationData("latitude") = arData("lat")
                arLocationData("longitude") = arData("lon")
            End Using
            Return arLocationData
        Catch ex As Exception
            If IsNothing(oHttpCache.Get("getlocation_service1")) Then
                Dim oHtmlEmail As New Mail
                oHtmlEmail.sReceiver = "iterr@zeepro.fr"
                oHtmlEmail.sSubject = "[SSO] Location Service 1 is down"
                oHtmlEmail.sBody = "The Location service1 (http://ip-api.com/json/" & sIp & ") is down. Please check the logs."
                oHtmlEmail.Send()
                oHttpCache.Insert("getlocation_service1", "Service1 is down", Nothing, DateTime.Now.AddMinutes(60.0), TimeSpan.Zero)
            End If
            ZSSOUtilities.WriteLog("GetLocation : NOK (service1) : " & ex.Message)
            Try
                Dim oRequest As WebRequest = WebRequest.Create("https://freegeoip.net/json/" & sIp)
                Using oResponseStream As Stream = oRequest.GetResponse().GetResponseStream()
                    Dim sResponse As String = New StreamReader(oResponseStream).ReadToEnd()

                    arLocationData = ZSSOUtilities.oSerializer.Deserialize(Of Dictionary(Of String, String))(sResponse)
                End Using
                Return arLocationData
            Catch ex2 As Exception
                If IsNothing(oHttpCache.Get("getlocation_service2")) Then
                    Dim oHtmlEmail As New Mail
                    oHtmlEmail.sReceiver = "iterr@zeepro.fr"
                    oHtmlEmail.sSubject = "[SSO] Location Service 2 is down"
                    oHtmlEmail.sBody = "The Location service2 (https://freegeoip.net/json/" & sIp & ") is down. Please check the logs."
                    oHtmlEmail.Send()
                    oHttpCache.Insert("getlocation_service2", "Service2 is down", Nothing, DateTime.Now.AddMinutes(60.0), TimeSpan.Zero)
                End If
                ZSSOUtilities.WriteLog("GetLocation : NOK (service2) : " & ex2.Message)
                Try
                    Dim oRequest As WebRequest = WebRequest.Create("http://www.telize.com/geoip/" & sIp)
                    Using oResponseStream As Stream = oRequest.GetResponse().GetResponseStream()
                        Dim sResponse As String = New StreamReader(oResponseStream).ReadToEnd()
                        arLocationData = ZSSOUtilities.oSerializer.Deserialize(Of Dictionary(Of String, String))(sResponse)
                    End Using
                    Return arLocationData
                Catch ex3 As Exception
                    If IsNothing(oHttpCache.Get("getlocation_service3")) Then
                        Dim oHtmlEmail As New Mail
                        oHtmlEmail.sReceiver = "iterr@zeepro.fr"
                        oHtmlEmail.sSubject = "[SSO] Location Service 3 is down"
                        oHtmlEmail.sBody = "The Location service3 (http://www.telize.com/geoip/" & sIp & ") is down. Please check the logs."
                        oHtmlEmail.Send()
                        oHttpCache.Insert("getlocation_service3", "Service3 is down", Nothing, DateTime.Now.AddMinutes(60.0), TimeSpan.Zero)
                    End If
                    ZSSOUtilities.WriteLog("GetLocation : NOK (service3) : " & ex3.Message)
                End Try
            End Try
        End Try
        Return Nothing
    End Function

    Public Shared Function CalculateDistanceBetweenCoordinates(sServerLatitude As String, sServerLongitude As String, sPrinterLatitude As String, sPrinterLongitude As String)
        Dim ciClone As CultureInfo = CType(CultureInfo.InvariantCulture.Clone(), CultureInfo)
        ciClone.NumberFormat.NumberDecimalSeparator = "."

        Dim dServerLatitude As Double = CDbl(sServerLatitude)
        Dim dServerLongitude As Double = CDbl(sServerLongitude)
        Dim dPrinterLatitude As Double = CDbl(sPrinterLatitude)
        Dim dPrinterLongitude As Double = CDbl(sPrinterLongitude)

        Dim dDistance As Double = Math.Sin(deg2rad(dServerLatitude)) * Math.Sin(deg2rad(dPrinterLatitude)) + Math.Cos(deg2rad(dServerLatitude)) * Math.Cos(deg2rad(dPrinterLatitude)) * Math.Cos(deg2rad(dServerLongitude - dPrinterLongitude))
        dDistance = Math.Acos(dDistance)
        dDistance = rad2deg(dDistance)
        dDistance = dDistance * 60 * 1.1515 'miles
        ' dist = dist * 1.609344 'kilometers
        Return dDistance
    End Function

    Private Shared Function deg2rad(ByVal dDegree As Double) As Double
        Return (dDegree * Math.PI / 180.0)
    End Function

    Private Shared Function rad2deg(ByVal dRadian As Double) As Double
        Return dRadian / Math.PI * 180.0
    End Function

    Public Shared Function GenerateRangeCode(sRangeStart As String, sRangeEnd As String)
        Dim bytesToHash() As Byte = System.Text.Encoding.ASCII.GetBytes(sRangeStart & sRangeEnd & sKey)
        Dim sHash As String = BitConverter.ToString(oSha1.ComputeHash(bytesToHash)).Replace("-", String.Empty)
        Return sRangeStart & sRangeEnd & sHash.Substring(0, 6)
    End Function

    Public Shared Function CheckRangeCode(sRangeCode As String) As Boolean
        If sRangeCode.Length <> 30 Then
            Return False
        End If

        Dim sRangeStart As String = sRangeCode.Substring(0, 12)
        Dim sRangeEnd As String = sRangeCode.Substring(12, 12)
        Dim sHashedCheck As String = GenerateRangeCode(sRangeStart, sRangeEnd)

        If String.Compare(sHashedCheck, sRangeCode) <> 0 Then
            Return False
        End If
        Return True
    End Function

    Public Shared Function GetRangeStart(oConnection As SqlConnection) As String
        Dim sQuery As String = "SELECT TOP 1 Value FROM LastAddress"
        Dim iRangeStart As Long = 0

        Using oSqlCmdSelect As New SqlCommand(sQuery, oConnection)

            Try
                Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()
                    If oQueryResult.Read() Then
                        iRangeStart = oQueryResult(oQueryResult.GetOrdinal("Value"))
                    End If
                End Using
            Catch ex As Exception
                Throw ex
            End Try
        End Using
        Return Hex(iRangeStart)
    End Function

    Public Shared Function CheckRangeStart(oConnection As SqlConnection, sRangeStart As String) As Boolean
        Dim sDbRangeStart As String

        sDbRangeStart = GetRangeStart(oConnection)

        If String.Compare(sDbRangeStart, sRangeStart) <> 0 Then
            Return False
        End If
        Return True
    End Function

    Public Shared Function SearchSerial(sSerial As String)
        Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
            oConnection.Open()

            Dim sQuery = "SELECT TOP 1 Serial " & _
                "FROM Printer " & _
                "WHERE Serial=@serial"

            Using oSqlCmdSelect As New SqlCommand(sQuery, oConnection)

                oSqlCmdSelect.Parameters.AddWithValue("@serial", sSerial.ToUpper)

                Try
                    Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()

                        If oQueryResult.HasRows Then
                            Return True
                        End If
                    End Using
                Catch ex As Exception
                End Try
            End Using
        End Using
        Return False
    End Function
End Class

Public NotInheritable Class Mail
    Public Property sSubject As String
    Public Property sBody As String
    Public Property sReceiver As String

    Public Sub Send()
        Dim oSmtpServer As New SmtpClient()
        Dim oMail As New MailMessage()
        oSmtpServer.UseDefaultCredentials = False
        oSmtpServer.Credentials = New Net.NetworkCredential("service-informatique@zee3dcompany.com", "uBXf9JhuFAg7FfeJVAVvkA")
        oSmtpServer.Port = 587
        oSmtpServer.EnableSsl = True
        oSmtpServer.Host = "smtp.mandrillapp.com"

        oMail = New MailMessage()
        oMail.From = New MailAddress("zim@zeepro.com")
        oMail.To.Add(sReceiver)
        oMail.Subject = sSubject
        oMail.IsBodyHtml = True
        oMail.Body = sBody
        oSmtpServer.Send(oMail)

    End Sub
End Class