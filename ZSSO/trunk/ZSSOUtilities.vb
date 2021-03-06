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
Imports Microsoft.WindowsAzure.Storage
Imports Microsoft.WindowsAzure.Storage.Auth
Imports Microsoft.WindowsAzure.Storage.Blob


Public Class ZSSOUtilities
    Public Shared oEmailRegex As New Regex("^[-0-9a-zA-Z.+_]+@[-0-9a-zA-Z.+_]+\.[a-zA-Z]{2,4}$")
    Public Shared rPartArchive As New Regex("(.*)\.zip\.(\d{1,2}|100)\.(\d{1,2}|100)$", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public Shared oSerializer As New JavaScriptSerializer

    Public Shared oAdminEmail As String = "zeepromac@zeepro.com"
    Public Shared oSerialEmail As String = "zeeproserial@zeepro.com"
    Public Shared oZoombaiAdminEmail As String = "changepassword@zeepro.com"
    Private Shared oSha1 As New SHA1CryptoServiceProvider
    Private Shared sKey As String = "zeepro"

    Public Shared Function Login(oConnection As SqlConnection, sEmail As String, sPassword As String, Optional lMustBeConfirmed As Boolean = True) As Boolean
        Dim sQuery = "SELECT TOP 1 * " & _
            "FROM Account " & _
            "WHERE Email=@email AND Deleted IS NULL"

        Using oSqlCmdSelect As New SqlCommand(sQuery, oConnection)
            oSqlCmdSelect.Parameters.AddWithValue("@email", sEmail.ToLower)
            Try
                Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteReader()
                    If oQueryResult.Read() Then
                        Dim bConfirmed As Boolean = oQueryResult(oQueryResult.GetOrdinal("Confirmed"))
                        If BCrypt.Net.BCrypt.Verify(sPassword, oQueryResult(oQueryResult.GetOrdinal("Password"))) AndAlso (bConfirmed OrElse Not lMustBeConfirmed) Then
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
        Static dGeo As New Dictionary(Of String, Dictionary(Of String, String))

        arLocationData = HttpRuntime.Cache(sIp)

        If Not arLocationData Is Nothing Then
            Return arLocationData
        Else
            Try
                Dim oRequest As WebRequest = WebRequest.Create("http://ip-api.com/json/" & sIp)
                Using oResponseStream As Stream = oRequest.GetResponse().GetResponseStream()
                    Dim sResponse As String = New StreamReader(oResponseStream).ReadToEnd()
                    Dim arData As Dictionary(Of String, String) = ZSSOUtilities.oSerializer.Deserialize(Of Dictionary(Of String, String))(sResponse)
                    arLocationData = New Dictionary(Of String, String)
                    arLocationData("latitude") = arData("lat")
                    arLocationData("longitude") = arData("lon")
                End Using
                HttpRuntime.Cache.Insert("geo_" & sIp, arLocationData, Nothing, DateTime.Now.AddDays(1), TimeSpan.Zero)
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
                    HttpRuntime.Cache.Insert("geo_" & sIp, arLocationData, Nothing, DateTime.Now.AddDays(1), TimeSpan.Zero)
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
                        HttpRuntime.Cache.Insert("geo_" & sIp, arLocationData, Nothing, DateTime.Now.AddDays(1), TimeSpan.Zero)
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
        End If

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

    Public Shared Function SearchAccountEmail(sToken As String, Optional sSerial As String = Nothing) As String

        Dim sAccountEmail As String
        Dim sQuery As String

        Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
            oConnection.Open()


            If sSerial Is Nothing Then
                sQuery = "DELETE TokenId WHERE date < DATEADD(day, -1, GETDATE());" & _
                    "SELECT TOP 1 TokenId.email " & _
                    "FROM TokenId " & _
                    "WHERE TokenId.token = @token"

                Using oSqlCmdSelect As New SqlCommand(sQuery, oConnection)
                    oSqlCmdSelect.Parameters.AddWithValue("@token", sToken)
                    sAccountEmail = oSqlCmdSelect.ExecuteScalar()
                End Using
            Else
                sQuery = "DELETE TokenId WHERE date < DATEADD(day, -1, GETDATE());" & _
                    "SELECT TOP 1 TokenId.email " & _
                    "FROM TokenId " & _
                    "INNER JOIN AccountPrinterAssociation " & _
                    "ON TokenId.email = AccountPrinterAssociation.email " & _
                    "WHERE TokenId.token = @token AND AccountPrinterAssociation.serial = @serial AND AccountPrinterAssociation.deleted IS NULL AND (AccountPrinterAssociation.accountrestriction IS NULL OR AccountPrinterAssociation.accountrestriction = 0)"

                Using oSqlCmdSelect As New SqlCommand(sQuery, oConnection)
                    oSqlCmdSelect.Parameters.AddWithValue("@token", sToken)
                    oSqlCmdSelect.Parameters.AddWithValue("@serial", sSerial)
                    sAccountEmail = oSqlCmdSelect.ExecuteScalar()
                End Using
            End If

        End Using

        Return sAccountEmail
    End Function

    Public Shared Function SearchService(sService As String)
        Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
            oConnection.Open()

            Using oSqlCmdSelect As New SqlCommand("SELECT TOP 1 service FROM Service WHERE service = @service", oConnection)

                oSqlCmdSelect.Parameters.AddWithValue("@service", sService)

                Try
                    Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteScalar.ExecuteReader()
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

    Public Shared Function SearchTicket(sTicket As String)
        Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
            oConnection.Open()

            Using oSqlCmdSelect As New SqlCommand("SELECT TOP 1 ticket FROM Ticket WHERE ticket = @ticket", oConnection)

                oSqlCmdSelect.Parameters.AddWithValue("@ticket", sTicket)

                Try
                    Using oQueryResult As SqlDataReader = oSqlCmdSelect.ExecuteScalar.ExecuteReader()
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

    Shared Function sha1(ByVal strToHash As String) As String
        Dim sha1Obj As New System.Security.Cryptography.SHA1CryptoServiceProvider
        Dim bytesToHash() As Byte = System.Text.Encoding.ASCII.GetBytes(strToHash)

        bytesToHash = sha1Obj.ComputeHash(bytesToHash)

        Dim strResult As String = ""

        For Each b As Byte In bytesToHash
            strResult += b.ToString("x2")
        Next

        Return strResult
    End Function

    Public Shared Sub SendStat(ByVal oStateInfo As Object)

        Using oWebclient As New WebClient
            oWebclient.UploadValues(System.Configuration.ConfigurationManager.AppSettings("statURL"), _
                                    DirectCast(oStateInfo, NameValueCollection))
        End Using
    End Sub

    Public Shared Sub Store3DFile(ByVal oStateInfo As Object)
        Dim sId As String = CStr(DirectCast(oStateInfo, NameValueCollection)("id"))
        Dim sPath As String = DirectCast(oStateInfo, NameValueCollection)("path")
        Dim oStorageAccount As CloudStorageAccount = CloudStorageAccount.Parse(System.Configuration.ConfigurationManager.AppSettings("StorageConnectionString"))
        Dim oBlobClient As CloudBlobClient
        Dim oContainer As CloudBlobContainer
        Dim oBlob As CloudBlockBlob
        Dim oEtag As Dictionary(Of String, String)

        Try
            oBlobClient = oStorageAccount.CreateCloudBlobClient()
            oContainer = oBlobClient.GetContainerReference("3dmodel")

            Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                oConnection.Open()

                Dim oSqlCmd As New SqlCommand("UPDATE Model3Dfile SET url1 = @url1, " & _
                                              "description = @description, " & _
                                              "etag1 = @etag1, " & _
                                              "url2 = @url2, " & _
                                              "etag2 = @etag2, " & _
                                              "img = @img " & _
                                              "WHERE model = @model", _
                        oConnection)

                oSqlCmd.Parameters.AddWithValue("@description", System.DBNull.Value)
                oSqlCmd.Parameters.AddWithValue("@url1", System.DBNull.Value)
                oSqlCmd.Parameters.AddWithValue("@etag1", System.DBNull.Value)
                oSqlCmd.Parameters.AddWithValue("@url2", System.DBNull.Value)
                oSqlCmd.Parameters.AddWithValue("@etag2", System.DBNull.Value)

                Using oRead As New StreamReader(sPath & "\metadata.json")
                    oEtag = ZSSOUtilities.oSerializer.Deserialize(Of Dictionary(Of String, String))(oRead.ReadToEnd())
                End Using

                If oEtag.ContainsKey("description") Then
                    oSqlCmd.Parameters("@description").Value = oEtag("description")
                End If

                For Each sFile As String In {"model.amf", "model.stl", "model1.stl"}
                    oBlob = oContainer.GetBlockBlobReference(sId & "." & sha1("zeepro" & sId & "." & sFile & ".zip") & ".zip")
                    oBlob.DeleteIfExists(DeleteSnapshotsOption.IncludeSnapshots)
                    If File.Exists(sPath & "\" & sFile) Then
                        Using oZip As New Ionic.Zip.ZipFile
                            oZip.AddFile(sPath & "\" & sFile, "")
                            oZip.Save(sPath & "\" & sFile & ".zip")
                        End Using
                        oBlob.UploadFromFile(sPath & "\" & sFile & ".zip", FileMode.Open)
                        oSqlCmd.Parameters("@url1").Value = oBlob.Uri.AbsoluteUri
                        If oEtag.ContainsKey("file1etag") Then
                            oSqlCmd.Parameters("@etag1").Value = oEtag("file1etag")
                        End If
                    End If
                Next

                oBlob = oContainer.GetBlockBlobReference(sId & "." & sha1("zeepro" & sId & ".model2.stl.zip") & ".zip")
                oBlob.DeleteIfExists(DeleteSnapshotsOption.IncludeSnapshots)
                If File.Exists(sPath & "\model2.stl") Then
                    Using oZip As New Ionic.Zip.ZipFile
                        oZip.AddFile(sPath & "\model2.stl", "")
                        oZip.Save(sPath & "\model2.stl.zip")
                    End Using
                    oBlob.UploadFromFile(sPath & "\model2.stl.zip", FileMode.Open)
                    oSqlCmd.Parameters("@url2").Value = oBlob.Uri.AbsoluteUri
                    If oEtag.ContainsKey("file2etag") Then
                        oSqlCmd.Parameters("@etag2").Value = oEtag("file2etag")
                    End If
                End If

                oBlob = oContainer.GetBlockBlobReference(sId & "." & sha1("zeepro" & sId & ".image.png") & ".png")
                oBlob.DeleteIfExists(DeleteSnapshotsOption.IncludeSnapshots)
                oBlob.UploadFromFile(sPath & "/image.png", FileMode.Open)
                oSqlCmd.Parameters.AddWithValue("@img", oBlob.Uri.AbsoluteUri)

                oSqlCmd.Parameters.AddWithValue("@model", DirectCast(oStateInfo, NameValueCollection)("id"))

                oSqlCmd.ExecuteNonQuery()
            End Using

            Directory.Delete(sPath, True)
        Catch ex As Exception
            ZSSOUtilities.WriteLog("Store3DFile: " & ex.Message)
            Return
        End Try
    End Sub

    Public Shared Sub StorePrint(ByVal oStateInfo As Object)
        Dim sId As String = CStr(DirectCast(oStateInfo, NameValueCollection)("id"))
        Dim sPath As String = DirectCast(oStateInfo, NameValueCollection)("path")
        Dim sDate As String = DirectCast(oStateInfo, NameValueCollection)("date")
        Dim oStorageAccount As CloudStorageAccount = CloudStorageAccount.Parse(System.Configuration.ConfigurationManager.AppSettings("StorageConnectionString"))
        Dim oBlobClient As CloudBlobClient
        Dim oContainer As CloudBlobContainer
        Dim oBlob As CloudBlockBlob
        Dim oEtag As Dictionary(Of String, String)

        Try
            oBlobClient = oStorageAccount.CreateCloudBlobClient()
            oContainer = oBlobClient.GetContainerReference("3dprint")

            Using oConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ZSSODb").ConnectionString)
                oConnection.Open()

                Dim oSqlCmd As New SqlCommand("UPDATE ModelPrint SET description = @description, " & _
                                              "gcode = @gcode, " & _
                                              "gcodeetag = @gcodeetag, " & _
                                              "[time-lapse] = @timelapse, " & _
                                              "[time-lapseetag] = @timelapseetag, " & _
                                              "img = @img " & _
                                              "WHERE model = @model AND date = @date", _
                        oConnection)

                Using oRead As New StreamReader(sPath & "\metadata.json")
                    oEtag = ZSSOUtilities.oSerializer.Deserialize(Of Dictionary(Of String, String))(oRead.ReadToEnd())
                End Using

                If oEtag.ContainsKey("description") Then
                    oSqlCmd.Parameters.AddWithValue("@description", oEtag("description"))
                Else
                    oSqlCmd.Parameters.AddWithValue("@description", System.DBNull.Value)
                End If

                oBlob = oContainer.GetBlockBlobReference(sId & "." & sDate & "." & sha1("zeepro" & sId & "." & sDate & ".print.gcode.zip") & ".zip")
                oBlob.DeleteIfExists(DeleteSnapshotsOption.IncludeSnapshots)
                Using oZip As New Ionic.Zip.ZipFile
                    oZip.AddFile(sPath & "\print.gcode", "")
                    oZip.Save(sPath & "\print.gcode.zip")
                End Using
                oBlob.UploadFromFile(sPath & "\print.gcode.zip", FileMode.Open)
                oSqlCmd.Parameters.AddWithValue("@gcode", oBlob.Uri.AbsoluteUri)

                If oEtag.ContainsKey("gcodeetag") Then
                    oSqlCmd.Parameters.AddWithValue("@gcodeetag", oEtag("gcodeetag"))
                Else
                    oSqlCmd.Parameters.AddWithValue("@gcodeetag", System.DBNull.Value)
                End If

                If File.Exists(sPath & "/time-lapse.mp4") Then
                    oBlob = oContainer.GetBlockBlobReference(sId & "." & sDate & "." & sha1("zeepro" & sId & "." & sDate & ".time-lapse.mp4") & ".mp4")
                    oBlob.DeleteIfExists(DeleteSnapshotsOption.IncludeSnapshots)
                    oBlob.UploadFromFile(sPath & "/time-lapse.mp4", FileMode.Open)
                    oSqlCmd.Parameters.AddWithValue("@timelapse", oBlob.Uri.AbsoluteUri)
                Else
                    oSqlCmd.Parameters.AddWithValue("@timelapse", System.DBNull.Value)
                End If

                If oEtag.ContainsKey("time-lapseetag") Then
                    oSqlCmd.Parameters.AddWithValue("@timelapseetag", oEtag("time-lapseetag"))
                Else
                    oSqlCmd.Parameters.AddWithValue("@timelapseetag", System.DBNull.Value)
                End If

                oBlob = oContainer.GetBlockBlobReference(sId & "." & sDate & "." & sha1("zeepro" & sId & "." & sDate & ".image.jpg") & ".jpg")
                oBlob.DeleteIfExists(DeleteSnapshotsOption.IncludeSnapshots)
                oBlob.UploadFromFile(sPath & "/image.jpg", FileMode.Open)
                oSqlCmd.Parameters.AddWithValue("@img", oBlob.Uri.AbsoluteUri)

                oSqlCmd.Parameters.AddWithValue("@model", DirectCast(oStateInfo, NameValueCollection)("id"))
                oSqlCmd.Parameters.AddWithValue("@date", sDate)

                oSqlCmd.ExecuteNonQuery()
            End Using

            Directory.Delete(sPath, True)
        Catch ex As Exception
            ZSSOUtilities.WriteLog("Store3DFile: " & ex.Message)
            Return
        End Try
    End Sub
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